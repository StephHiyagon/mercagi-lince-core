(function () {
  "use strict";

  const { parseMercagiMessage, createLinceWorkbook, safeFileStamp } = globalThis.MercagiCore;
  const elements = {
    input: document.getElementById("messageInput"),
    characterCount: document.getElementById("characterCount"),
    statusChip: document.getElementById("statusChip"),
    statusText: document.getElementById("statusText"),
    alertBox: document.getElementById("alertBox"),
    emptyState: document.getElementById("emptyState"),
    tableWrap: document.getElementById("tableWrap"),
    previewBody: document.getElementById("previewBody"),
    clearButton: document.getElementById("clearButton"),
    downloadButton: document.getElementById("downloadButton"),
    exampleButton: document.getElementById("exampleButton"),
    toast: document.getElementById("toast"),
  };

  const exampleMessage = `▶️  *_Invetsa-Bolivia_*  ◀️

*------------------------------*
📌    *_Detalle del pedido_*
*------------------------------*

✅ *x2* | Alimento Húmedo para Perros/Gatos a/d , Hill's Prescription Diet | ACOM-00537 | *65 Bs*
✅ *x2* | Urinary Care c/d Perros Bolsa de 3.9 kg, Hill's Prescription Diet | ACOM-00511 | *620 Bs*
✅ *x2* | Kidney Care k/d Para Perros Bolsa de 8 kg, Hill's Prescription Diet | ACOM-00503 | *1230 Bs*
✅ *x2* | Digestive Care i/d Gatos - Lata de 156gr, Hill's Prescription Diet | ACOM-00105 | *65 Bs*

*------------------------------*
*_✅ Total a Pagar: 3960 Bs_*`;

  let currentResult = parseMercagiMessage("");
  let toastTimer;

  function escapeHtml(value) {
    return String(value).replace(/[&<>\"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  }

  function showToast(message) {
    clearTimeout(toastTimer);
    elements.toast.textContent = message;
    elements.toast.hidden = false;
    toastTimer = setTimeout(() => { elements.toast.hidden = true; }, 3600);
  }

  function render() {
    const text = elements.input.value;
    currentResult = parseMercagiMessage(text);
    elements.characterCount.textContent = `${text.length.toLocaleString("es-BO")} caracteres`;
    elements.clearButton.disabled = text.length === 0;

    const hasRows = currentResult.rows.length > 0;
    elements.emptyState.hidden = hasRows;
    elements.tableWrap.hidden = !hasRows;
    elements.downloadButton.disabled = !currentResult.isValid;

    elements.previewBody.innerHTML = currentResult.rows.map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.code)}</td>
        <td>${row.quantity}</td>
        <td class="muted-cell">—</td>
        <td class="muted-cell">—</td>
      </tr>`).join("");

    if (!text.trim()) {
      elements.statusChip.dataset.state = "empty";
      elements.statusText.textContent = "Esperando mensaje";
      elements.alertBox.hidden = true;
      return;
    }

    if (currentResult.errors.length) {
      elements.statusChip.dataset.state = "error";
      elements.statusText.textContent = "Revisar mensaje";
      const items = currentResult.errors.slice(0, 6).map((error) => `<li>${error.line ? `Línea ${error.line}: ` : ""}${escapeHtml(error.message)}</li>`).join("");
      const remaining = currentResult.errors.length - 6;
      elements.alertBox.innerHTML = `<strong>No se puede generar el Excel todavía.</strong><ul>${items}${remaining > 0 ? `<li>Hay ${remaining} observaciones adicionales.</li>` : ""}</ul>`;
      elements.alertBox.hidden = false;
      return;
    }

    if (hasRows) {
      elements.statusChip.dataset.state = "valid";
      elements.statusText.textContent = `${currentResult.rows.length} ${currentResult.rows.length === 1 ? "producto" : "productos"}`;
      elements.alertBox.hidden = true;
    } else {
      elements.statusChip.dataset.state = "error";
      elements.statusText.textContent = "Sin productos";
      elements.alertBox.innerHTML = "No se detectaron líneas con el formato <strong>xCantidad</strong> y <strong>ACOM-código</strong>.";
      elements.alertBox.hidden = false;
    }
  }

  function downloadWorkbook() {
    if (!currentResult.isValid) return;
    try {
      const bytes = createLinceWorkbook(currentResult.rows);
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `Pedido_Mercagi_LINCE_CORE_${safeFileStamp(new Date())}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      showToast(`Excel generado con ${currentResult.rows.length} productos.`);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "No se pudo generar el Excel.");
    }
  }

  elements.input.addEventListener("input", render);
  elements.downloadButton.addEventListener("click", downloadWorkbook);
  elements.clearButton.addEventListener("click", () => {
    elements.input.value = "";
    render();
    elements.input.focus();
  });
  elements.exampleButton.addEventListener("click", () => {
    elements.input.value = exampleMessage;
    render();
    elements.input.focus();
  });

  function registerWebMcpTools() {
    const context = document.modelContext;
    if (!context?.registerTool) return;

    const reportError = () => {};
    try {
      void Promise.resolve(context.registerTool({
        name: "parse_mercagi_order",
        title: "Procesar pedido de Mercagi",
        description: "Pega y procesa un mensaje de pedido de Mercagi para mostrar los productos que pasarán al Excel de LINCE CORE.",
        inputSchema: {
          type: "object",
          properties: { message: { type: "string", description: "Mensaje completo recibido por WhatsApp." } },
          required: ["message"],
          additionalProperties: false,
        },
        annotations: { readOnlyHint: false, untrustedContentHint: true },
        execute(input) {
          if (!input || typeof input.message !== "string") throw new Error("El mensaje es obligatorio.");
          elements.input.value = input.message;
          render();
          return { products: currentResult.rows.length, errors: currentResult.errors, readyToDownload: currentResult.isValid };
        },
      })).catch(reportError);

      void Promise.resolve(context.registerTool({
        name: "download_lince_excel",
        title: "Descargar Excel de LINCE CORE",
        description: "Descarga el Excel del pedido de Mercagi que está validado y visible en la aplicación.",
        inputSchema: { type: "object", properties: {}, additionalProperties: false },
        annotations: { readOnlyHint: false, untrustedContentHint: false },
        execute() {
          if (!currentResult.isValid) throw new Error("Primero procesa un mensaje válido.");
          downloadWorkbook();
          return { downloaded: true, products: currentResult.rows.length };
        },
      })).catch(reportError);
    } catch (error) {
      reportError(error);
    }
  }

  render();
  registerWebMcpTools();
})();
