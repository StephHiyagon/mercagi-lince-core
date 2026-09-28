(function (global) {
  "use strict";

  const MAX_PRODUCTS = 100;
  const encoder = new TextEncoder();

  function parseMercagiMessage(input) {
    const text = String(input || "").replace(/\r\n?/g, "\n");
    const lines = text.split("\n");
    const rows = [];
    const errors = [];

    lines.forEach((rawLine, index) => {
      const line = rawLine.trim();
      if (!line) return;

      const codeMatch = line.toUpperCase().match(/\bACOM-[A-Z0-9]+\b/);
      const quantityMatch = line.match(/(?:^|[\s*_])x\s*(\d{1,4})(?=\s|\*|\||$)/i);
      const looksLikeProduct = Boolean(codeMatch || quantityMatch || /^✅/.test(line));

      if (!looksLikeProduct) return;
      if (!codeMatch && !quantityMatch && /^✅/.test(line)) return;

      if (!codeMatch || !quantityMatch) {
        const missing = [!quantityMatch ? "cantidad" : null, !codeMatch ? "código" : null].filter(Boolean).join(" y ");
        errors.push({ line: index + 1, message: `No se pudo identificar ${missing}.` });
        return;
      }

      const quantity = Number(quantityMatch[1]);
      if (!Number.isInteger(quantity) || quantity < 1) {
        errors.push({ line: index + 1, message: "La cantidad debe ser un número entero mayor que cero." });
        return;
      }

      rows.push({ code: codeMatch[0], quantity });
    });

    if (rows.length > MAX_PRODUCTS) {
      errors.push({ line: null, message: `El pedido tiene ${rows.length} productos. El máximo permitido es ${MAX_PRODUCTS}.` });
    }

    return {
      rows: rows.slice(0, MAX_PRODUCTS),
      errors,
      totalDetected: rows.length,
      isValid: rows.length > 0 && rows.length <= MAX_PRODUCTS && errors.length === 0,
    };
  }

  function xmlEscape(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function u16(value) {
    const bytes = new Uint8Array(2);
    new DataView(bytes.buffer).setUint16(0, value, true);
    return bytes;
  }

  function u32(value) {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    return bytes;
  }

  function concatBytes(parts) {
    const length = parts.reduce((sum, part) => sum + part.length, 0);
    const output = new Uint8Array(length);
    let offset = 0;
    parts.forEach((part) => {
      output.set(part, offset);
      offset += part.length;
    });
    return output;
  }

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) crc = crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function dosDateTime(date) {
    const year = Math.max(1980, date.getFullYear());
    return {
      time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
      date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    };
  }

  function makeZip(entries) {
    const localParts = [];
    const centralParts = [];
    const now = dosDateTime(new Date());
    let offset = 0;

    entries.forEach((entry) => {
      const name = encoder.encode(entry.name);
      const data = typeof entry.data === "string" ? encoder.encode(entry.data) : entry.data;
      const crc = crc32(data);
      const flags = 0x0800;

      const local = concatBytes([
        u32(0x04034b50), u16(20), u16(flags), u16(0), u16(now.time), u16(now.date),
        u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data,
      ]);
      localParts.push(local);

      const central = concatBytes([
        u32(0x02014b50), u16(20), u16(20), u16(flags), u16(0), u16(now.time), u16(now.date),
        u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0),
        u16(0), u32(0), u32(offset), name,
      ]);
      centralParts.push(central);
      offset += local.length;
    });

    const localData = concatBytes(localParts);
    const centralData = concatBytes(centralParts);
    const end = concatBytes([
      u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length),
      u32(centralData.length), u32(localData.length), u16(0),
    ]);
    return concatBytes([localData, centralData, end]);
  }

  function inlineCell(ref, value, style) {
    const styleAttr = style ? ` s="${style}"` : "";
    return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
  }

  function createLinceWorkbook(rows) {
    if (!Array.isArray(rows) || rows.length < 1 || rows.length > MAX_PRODUCTS) {
      throw new Error(`El archivo requiere entre 1 y ${MAX_PRODUCTS} productos.`);
    }

    const header = `<row r="1">${inlineCell("A1", "CODIGO")}${inlineCell("B1", "CANTIDAD")}${inlineCell("C1", "TIPO")}${inlineCell("D1", "PRECIO")}</row>`;
    const body = rows.map((row, index) => {
      const r = index + 2;
      return `<row r="${r}">${inlineCell(`A${r}`, row.code, 1)}<c r="B${r}" s="1"><v>${Number(row.quantity)}</v></c>${inlineCell(`C${r}`, "")}${inlineCell(`D${r}`, "")}</row>`;
    }).join("");
    const lastRow = rows.length + 1;

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D${lastRow}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="4" width="13" customWidth="1"/></cols>
  <sheetData>${header}${body}</sheetData>
</worksheet>`;

    const entries = [
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>` },
      { name: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Mercagi a LINCE CORE</Application></Properties>` },
      { name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Pedido LINCE CORE</dc:title><dc:creator>Invetsa Bolivia</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created></cp:coreProperties>` },
      { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Hoja1" sheetId="1" r:id="rId1"/></sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: "xl/styles.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="2"><border/><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>` },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml },
    ];

    return makeZip(entries);
  }

  function safeFileStamp(date) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
  }

  global.MercagiCore = { MAX_PRODUCTS, parseMercagiMessage, createLinceWorkbook, safeFileStamp };
})(globalThis);
