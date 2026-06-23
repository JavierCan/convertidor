(() => {
  "use strict";

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));
  const MAX_PREVIEW_XML = 30;
  const MAX_PREVIEW_ROWS = 20;
  const MAX_PREVIEW_COLUMNS = 25;
  const XLSX_SHEET_ROW_LIMIT = 250000;
  const STORAGE_KEY = "cfdi_excel_layouts_v2";
  const DEFAULT_LAYOUT_KEY = "cfdi_excel_default_layout_v2";

  const state = {
    uploadedFiles: [],
    xmlEntries: [],
    analysis: { total: 0, valid: 0, concepts: 0, errors: [], types: {} },
    representation: "concept",
    outputMode: "consolidated",
    excludeTechnical: true,
    convertNumbers: true,
    includeNewColumns: true,
    previewRows: [],
    availableColumns: [],
    columns: [],
    layouts: loadLayouts(),
    activeLayoutId: localStorage.getItem(DEFAULT_LAYOUT_KEY) || "",
    generatedBlob: null,
    generatedName: "",
    dragIndex: null
  };

  const els = {
    fileInput: $("#fileInput"),
    dropzone: $("#dropzone"),
    fileSummary: $("#fileSummary"),
    metricXml: $("#metricXml"),
    metricValid: $("#metricValid"),
    metricConcepts: $("#metricConcepts"),
    metricErrors: $("#metricErrors"),
    errorNotice: $("#errorNotice"),
    viewErrors: $("#viewErrors"),
    errorList: $("#errorList"),
    errorsDialog: $("#errorsDialog"),
    helpDialog: $("#helpDialog"),
    openHelp: $("#openHelp"),
    outputModeSection: $("#outputModeSection"),
    outputModeOptions: $("#outputModeOptions"),
    singleFileNotice: $("#singleFileNotice"),
    fileCountTag: $("#fileCountTag"),
    excludeTechnical: $("#excludeTechnical"),
    convertNumbers: $("#convertNumbers"),
    includeNewColumns: $("#includeNewColumns"),
    selectedCount: $("#selectedCount"),
    layoutSelect: $("#layoutSelect"),
    applyLayout: $("#applyLayout"),
    importLayout: $("#importLayout"),
    exportLayout: $("#exportLayout"),
    layoutFileInput: $("#layoutFileInput"),
    layoutName: $("#layoutName"),
    setAsDefault: $("#setAsDefault"),
    saveLayout: $("#saveLayout"),
    deleteLayout: $("#deleteLayout"),
    columnSearch: $("#columnSearch"),
    selectRecommended: $("#selectRecommended"),
    selectAll: $("#selectAll"),
    clearAll: $("#clearAll"),
    columnList: $("#columnList"),
    sampleTag: $("#sampleTag"),
    refreshPreview: $("#refreshPreview"),
    previewTable: $("#previewTable"),
    previewEmpty: $("#previewEmpty"),
    previewDescription: $("#previewDescription"),
    finalXmlCount: $("#finalXmlCount"),
    finalRepresentation: $("#finalRepresentation"),
    finalColumnCount: $("#finalColumnCount"),
    finalOutput: $("#finalOutput"),
    processButton: $("#processButton"),
    processButtonTitle: $("#processButtonTitle"),
    processButtonSubtitle: $("#processButtonSubtitle"),
    progressCard: $("#progressCard"),
    progressTitle: $("#progressTitle"),
    progressPercent: $("#progressPercent"),
    progressBar: $("#progressBar"),
    progressDetail: $("#progressDetail"),
    resultCard: $("#resultCard"),
    resultText: $("#resultText"),
    downloadButton: $("#downloadButton"),
    sideFiles: $("#sideFiles"),
    sideXml: $("#sideXml"),
    sideLayout: $("#sideLayout"),
    toastContainer: $("#toastContainer")
  };

  function showToast(message, type = "") {
    const toast = document.createElement("div");
    toast.className = `toast ${type ? `toast--${type}` : ""}`;
    toast.textContent = message;
    els.toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 3600);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("es-MX").format(value || 0);
  }

  function localName(nodeOrName) {
    const name = typeof nodeOrName === "string"
      ? nodeOrName
      : (nodeOrName.localName || nodeOrName.nodeName || "");
    return name.includes(":") ? name.split(":").pop() : name;
  }

  function normalizeColumn(value) {
    return String(value ?? "")
      .trim()
      .replace(/\s+/g, "_")
      .replace(/[^A-Za-z0-9_áéíóúÁÉÍÓÚñÑ.-]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function addUnique(target, key, value) {
    let candidate = key;
    let index = 2;
    while (Object.prototype.hasOwnProperty.call(target, candidate) && target[candidate] !== value) {
      candidate = `${key}_${index++}`;
    }
    target[candidate] = value;
  }

  function parseXmlDocument(xmlText) {
    const xml = new DOMParser().parseFromString(xmlText, "application/xml");
    const parserError = xml.querySelector("parsererror");
    if (parserError) {
      throw new Error(parserError.textContent.trim().replace(/\s+/g, " ").slice(0, 260));
    }
    return xml;
  }

  function identifyXmlType(root) {
    const rootName = localName(root);
    const version = root.getAttribute("Version") || root.getAttribute("version") || "";
    const names = new Set(Array.from(root.getElementsByTagName("*")).map(localName));
    if (rootName === "Comprobante") {
      if (names.has("Pagos") || names.has("Pago") || names.has("DoctoRelacionado")) {
        return `CFDI ${version || "sin versión"} - Complemento de pago`;
      }
      return `CFDI ${version || "sin versión"}`;
    }
    return `XML genérico - ${rootName}`;
  }

  function flattenElement(node, path, output, options = {}) {
    const name = localName(node);
    if (options.skipNames?.has(name)) return;

    const currentPath = path || name;
    Array.from(node.attributes || []).forEach((attribute) => {
      const attributeName = localName(attribute.name);
      if (options.excludeTechnical && ["Sello", "Certificado", "SelloCFD", "SelloSAT"].includes(attributeName)) return;
      addUnique(output, normalizeColumn(`${currentPath}_${attributeName}`), attribute.value);
    });

    const elementChildren = Array.from(node.children || []);
    const text = (node.textContent || "").trim();
    if (!elementChildren.length && text) {
      addUnique(output, normalizeColumn(`${currentPath}_Texto`), text);
    }

    const counts = new Map();
    elementChildren.forEach((child) => {
      const childName = localName(child);
      counts.set(childName, (counts.get(childName) || 0) + 1);
    });

    const occurrences = new Map();
    elementChildren.forEach((child) => {
      const childName = localName(child);
      const occurrence = (occurrences.get(childName) || 0) + 1;
      occurrences.set(childName, occurrence);
      const suffix = counts.get(childName) > 1 ? `_${occurrence}` : "";
      flattenElement(
        child,
        normalizeColumn(`${currentPath}_${childName}${suffix}`),
        output,
        options
      );
    });
  }

  function findElements(root, name) {
    return Array.from(root.getElementsByTagName("*")).filter((node) => localName(node) === name);
  }

  function rowsFromXml(entry, representation = state.representation) {
    const xml = parseXmlDocument(entry.text);
    const root = xml.documentElement;
    const concepts = findElements(root, "Concepto");
    const base = {
      Archivo_XML: entry.name,
      Tipo_XML: identifyXmlType(root)
    };

    flattenElement(root, localName(root), base, {
      skipNames: new Set(["Conceptos"]),
      excludeTechnical: state.excludeTechnical
    });

    if (representation === "document") {
      const row = { ...base, Total_Conceptos: concepts.length };
      concepts.forEach((concept, index) => {
        flattenElement(
          concept,
          `Concepto_${index + 1}`,
          row,
          { excludeTechnical: state.excludeTechnical }
        );
      });
      return [row];
    }

    if (!concepts.length) return [base];

    return concepts.map((concept, index) => {
      const row = { ...base, Concepto_Indice: index + 1 };
      flattenElement(
        concept,
        "Concepto",
        row,
        { excludeTechnical: state.excludeTechnical }
      );
      return row;
    });
  }

  async function collectXmlEntries(files) {
    const entries = [];
    const errors = [];

    for (let fileIndex = 0; fileIndex < files.length; fileIndex += 1) {
      const file = files[fileIndex];
      updateProgress(
        ((fileIndex + 0.15) / Math.max(files.length, 1)) * 100,
        "Leyendo archivos",
        file.name
      );

      const lowerName = file.name.toLowerCase();
      if (lowerName.endsWith(".xml")) {
        try {
          entries.push({ name: file.name, text: await file.text() });
        } catch (error) {
          errors.push({ file: file.name, error: error.message });
        }
      } else if (lowerName.endsWith(".zip")) {
        try {
          const zip = await JSZip.loadAsync(file);
          const zipEntries = Object.values(zip.files)
            .filter((item) => !item.dir && item.name.toLowerCase().endsWith(".xml"));

          for (let index = 0; index < zipEntries.length; index += 1) {
            const item = zipEntries[index];
            try {
              entries.push({
                name: `${file.name.replace(/\.zip$/i, "")}/${item.name}`,
                text: await item.async("string")
              });
            } catch (error) {
              errors.push({ file: `${file.name}/${item.name}`, error: error.message });
            }
          }
        } catch (error) {
          errors.push({ file: file.name, error: `ZIP inválido: ${error.message}` });
        }
      }
    }

    return { entries, errors };
  }

  function updateProgress(percent, title, detail) {
    const cleanPercent = Math.max(0, Math.min(100, Math.round(percent)));
    els.progressCard.classList.remove("is-hidden");
    els.progressTitle.textContent = title;
    els.progressPercent.textContent = `${cleanPercent}%`;
    els.progressBar.style.width = `${cleanPercent}%`;
    els.progressDetail.textContent = detail || "";
  }

  function hideProgress() {
    els.progressCard.classList.add("is-hidden");
    els.progressBar.style.width = "0%";
  }

  function setSectionEnabled(stepNumber, enabled) {
    const section = $(`[data-step="${stepNumber}"]`);
    section?.classList.toggle("is-disabled", !enabled);
  }

  function updateStepper() {
    const hasFiles = state.analysis.valid > 0;
    const hasColumns = state.columns.some((column) => column.include);

    $$(".step").forEach((step) => {
      const number = Number(step.dataset.stepTarget);
      const complete =
        (number === 1 && hasFiles) ||
        (number === 2 && hasFiles) ||
        (number === 3 && hasFiles && hasColumns);
      step.classList.toggle("is-complete", complete);
      step.classList.toggle(
        "is-active",
        (number === 1 && !hasFiles) ||
        (number === 2 && hasFiles && !state.availableColumns.length) ||
        (number === 3 && state.availableColumns.length && !state.generatedBlob) ||
        (number === 4 && !!state.generatedBlob)
      );
    });
  }

  function updateSummaryUI() {
    const { total, valid, concepts, errors } = state.analysis;
    els.metricXml.textContent = formatNumber(total);
    els.metricValid.textContent = formatNumber(valid);
    els.metricConcepts.textContent = formatNumber(concepts);
    els.metricErrors.textContent = formatNumber(errors.length);
    els.sideFiles.textContent = formatNumber(state.uploadedFiles.length);
    els.sideXml.textContent = formatNumber(total);
    els.fileCountTag.textContent = `${formatNumber(total)} ${total === 1 ? "XML" : "XML"}`;
    els.fileSummary.classList.toggle("is-hidden", total === 0);
    els.errorNotice.classList.toggle("is-hidden", errors.length === 0);

    const single = total === 1;
    els.outputModeOptions.classList.toggle("is-hidden", single);
    els.singleFileNotice.classList.toggle("is-hidden", !single);
    if (single) state.outputMode = "individual";

    els.finalXmlCount.textContent = formatNumber(total);
    els.finalRepresentation.textContent = state.representation === "concept" ? "Por concepto" : "Por comprobante";
    els.finalColumnCount.textContent = formatNumber(getSelectedColumns().length);
    els.finalOutput.textContent =
      total === 1 ? "Excel individual" :
      state.outputMode === "consolidated" ? "Excel consolidado" : "ZIP de Excel";

    const enabled = valid > 0 && getSelectedColumns().length > 0;
    els.processButton.disabled = !enabled;
    els.processButtonTitle.textContent =
      total === 1 ? "Generar Excel" :
      state.outputMode === "consolidated" ? "Generar Excel consolidado" : "Generar ZIP de Excel";
    els.processButtonSubtitle.textContent =
      enabled ? `${formatNumber(valid)} comprobantes válidos · ${formatNumber(getSelectedColumns().length)} columnas` :
      "Carga CFDI y selecciona al menos una columna";

    setSectionEnabled(2, valid > 0);
    setSectionEnabled(3, valid > 0);
    setSectionEnabled(4, valid > 0);
    updateStepper();
  }

  async function handleFiles(files) {
    state.uploadedFiles = Array.from(files || []);
    state.generatedBlob = null;
    state.generatedName = "";
    els.resultCard.classList.add("is-hidden");

    if (!state.uploadedFiles.length) {
      resetState();
      return;
    }

    updateProgress(4, "Preparando carga", "Leyendo archivos seleccionados...");

    const collected = await collectXmlEntries(state.uploadedFiles);
    state.xmlEntries = collected.entries;
    state.analysis = {
      total: state.xmlEntries.length,
      valid: 0,
      concepts: 0,
      errors: [...collected.errors],
      types: {}
    };

    for (let index = 0; index < state.xmlEntries.length; index += 1) {
      const entry = state.xmlEntries[index];
      updateProgress(
        18 + ((index + 1) / Math.max(state.xmlEntries.length, 1)) * 58,
        "Validando CFDI",
        `${index + 1} de ${state.xmlEntries.length}: ${entry.name}`
      );

      try {
        const xml = parseXmlDocument(entry.text);
        const root = xml.documentElement;
        const type = identifyXmlType(root);
        state.analysis.valid += 1;
        state.analysis.concepts += findElements(root, "Concepto").length;
        state.analysis.types[type] = (state.analysis.types[type] || 0) + 1;
      } catch (error) {
        state.analysis.errors.push({ file: entry.name, error: error.message });
      }
    }

    updateProgress(82, "Preparando vista", "Detectando campos contables...");
    updateSummaryUI();
    await buildPreviewAndColumns(true);
    updateProgress(100, "Carga terminada", "Tus CFDI están listos para configurar.");
    setTimeout(hideProgress, 650);

    document.querySelector("#step2").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function resetState() {
    state.uploadedFiles = [];
    state.xmlEntries = [];
    state.analysis = { total: 0, valid: 0, concepts: 0, errors: [], types: {} };
    state.previewRows = [];
    state.availableColumns = [];
    state.columns = [];
    renderColumnList();
    renderPreview();
    updateSummaryUI();
  }

  async function buildPreviewAndColumns(applySavedLayout = false) {
    const validRows = [];
    const columnOrder = [];
    const columnSet = new Set();
    const sampleEntries = state.xmlEntries.slice(0, MAX_PREVIEW_XML);
    let accepted = 0;

    for (const entry of sampleEntries) {
      try {
        const rows = rowsFromXml(entry);
        rows.forEach((row) => {
          Object.keys(row).forEach((column) => {
            if (!columnSet.has(column)) {
              columnSet.add(column);
              columnOrder.push(column);
            }
          });
          if (validRows.length < MAX_PREVIEW_ROWS) validRows.push(row);
        });
        accepted += 1;
      } catch (_) {
        // Ya se reporta durante el análisis.
      }
    }

    state.previewRows = validRows;
    state.availableColumns = columnOrder;

    const previousByName = new Map(state.columns.map((column) => [column.name, column]));
    state.columns = columnOrder.map((name, index) => ({
      name,
      include: previousByName.has(name) ? previousByName.get(name).include : true,
      order: index + 1
    }));

    if (applySavedLayout) {
      const defaultId = localStorage.getItem(DEFAULT_LAYOUT_KEY);
      if (defaultId && state.layouts[defaultId]) {
        applyLayoutObject(state.layouts[defaultId], false);
      }
    }

    els.sampleTag.textContent = `Muestra: ${accepted} XML`;
    els.previewDescription.textContent =
      `Vista basada en ${accepted} de ${state.analysis.valid} comprobantes válidos.`;

    renderColumnList();
    renderPreview();
    updateSummaryUI();
  }

  function getSelectedColumns() {
    return state.columns
      .filter((column) => column.include)
      .sort((a, b) => a.order - b.order)
      .map((column) => column.name);
  }

  function renderColumnList() {
    const query = els.columnSearch.value.trim().toLowerCase();
    const sorted = [...state.columns].sort((a, b) => a.order - b.order);
    const filtered = sorted.filter((column) => column.name.toLowerCase().includes(query));

    els.columnList.innerHTML = "";
    filtered.forEach((column) => {
      const actualIndex = state.columns.findIndex((item) => item.name === column.name);
      const item = document.createElement("div");
      item.className = "column-item";
      item.draggable = true;
      item.dataset.index = String(actualIndex);
      item.innerHTML = `
        <div class="column-item__handle" title="Arrastrar">⋮⋮</div>
        <input type="checkbox" ${column.include ? "checked" : ""} aria-label="Incluir ${escapeHtml(column.name)}">
        <div class="column-item__name" title="${escapeHtml(column.name)}">${escapeHtml(column.name)}</div>
        <div class="column-item__actions">
          <button class="icon-mini" type="button" data-action="up" title="Mover arriba">↑</button>
          <button class="icon-mini" type="button" data-action="down" title="Mover abajo">↓</button>
        </div>
      `;

      const checkbox = $("input[type=checkbox]", item);
      checkbox.addEventListener("change", () => {
        column.include = checkbox.checked;
        renderPreview();
        updateSummaryUI();
      });

      $("[data-action=up]", item).addEventListener("click", () => moveColumn(column.name, -1));
      $("[data-action=down]", item).addEventListener("click", () => moveColumn(column.name, 1));

      item.addEventListener("dragstart", () => {
        state.dragIndex = state.columns.findIndex((entry) => entry.name === column.name);
        item.classList.add("is-dragging");
      });
      item.addEventListener("dragend", () => {
        item.classList.remove("is-dragging");
        state.dragIndex = null;
      });
      item.addEventListener("dragover", (event) => event.preventDefault());
      item.addEventListener("drop", (event) => {
        event.preventDefault();
        const targetIndex = state.columns.findIndex((entry) => entry.name === column.name);
        if (state.dragIndex === null || targetIndex < 0 || state.dragIndex === targetIndex) return;
        const [moved] = state.columns.splice(state.dragIndex, 1);
        state.columns.splice(targetIndex, 0, moved);
        normalizeColumnOrder();
        renderColumnList();
        renderPreview();
      });

      els.columnList.appendChild(item);
    });

    els.selectedCount.textContent = formatNumber(getSelectedColumns().length);
    els.finalColumnCount.textContent = formatNumber(getSelectedColumns().length);
  }

  function normalizeColumnOrder() {
    state.columns.forEach((column, index) => {
      column.order = index + 1;
    });
  }

  function moveColumn(name, direction) {
    const sorted = [...state.columns].sort((a, b) => a.order - b.order);
    const index = sorted.findIndex((column) => column.name === name);
    const targetIndex = index + direction;
    if (index < 0 || targetIndex < 0 || targetIndex >= sorted.length) return;
    [sorted[index], sorted[targetIndex]] = [sorted[targetIndex], sorted[index]];
    state.columns = sorted;
    normalizeColumnOrder();
    renderColumnList();
    renderPreview();
  }

  function recommendedColumn(columnName) {
    return [
      /^Archivo_XML$/i,
      /UUID/i,
      /Serie/i,
      /Folio/i,
      /Fecha(?!.*Timbrado)/i,
      /Emisor.*Rfc/i,
      /Emisor.*Nombre/i,
      /Receptor.*Rfc/i,
      /Receptor.*Nombre/i,
      /SubTotal/i,
      /Total(?!.*Conceptos)/i,
      /Moneda/i,
      /TipoDeComprobante/i,
      /MetodoPago/i,
      /FormaPago/i,
      /Concepto.*ClaveProdServ/i,
      /Concepto.*Descripcion/i,
      /Concepto.*Cantidad/i,
      /Concepto.*ValorUnitario/i,
      /Concepto.*Importe/i,
      /Impuesto/i,
      /TasaOCuota/i
    ].some((pattern) => pattern.test(columnName));
  }

  function renderPreview() {
    const columns = getSelectedColumns().slice(0, MAX_PREVIEW_COLUMNS);
    const rows = state.previewRows.slice(0, MAX_PREVIEW_ROWS);
    const thead = $("thead", els.previewTable);
    const tbody = $("tbody", els.previewTable);

    if (!columns.length) {
      els.previewTable.classList.add("is-hidden");
      els.previewEmpty.classList.remove("is-hidden");
      thead.innerHTML = "";
      tbody.innerHTML = "";
      updateSummaryUI();
      return;
    }

    els.previewTable.classList.remove("is-hidden");
    els.previewEmpty.classList.add("is-hidden");
    thead.innerHTML = `<tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;
    tbody.innerHTML = rows.map((row) => (
      `<tr>${columns.map((column) => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>`
    )).join("");

    if (getSelectedColumns().length > MAX_PREVIEW_COLUMNS) {
      els.previewDescription.textContent =
        `Mostrando las primeras ${MAX_PREVIEW_COLUMNS} de ${getSelectedColumns().length} columnas seleccionadas.`;
    }

    updateSummaryUI();
  }

  function loadLayouts() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    } catch (_) {
      return {};
    }
  }

  function persistLayouts() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.layouts));
    renderLayoutSelect();
  }

  function renderLayoutSelect() {
    const current = els.layoutSelect.value || state.activeLayoutId;
    const options = Object.entries(state.layouts)
      .sort(([, a], [, b]) => String(a.name).localeCompare(String(b.name), "es"))
      .map(([id, layout]) => `<option value="${escapeHtml(id)}">${escapeHtml(layout.name)}</option>`)
      .join("");

    els.layoutSelect.innerHTML = `<option value="">Sin layout guardado</option>${options}`;
    if (state.layouts[current]) els.layoutSelect.value = current;
    els.sideLayout.textContent = state.layouts[state.activeLayoutId]?.name || "Ninguno";
  }

  function currentLayoutPayload() {
    return {
      version: 2,
      name: els.layoutName.value.trim() || "Mi layout CFDI",
      representation: state.representation,
      outputMode: state.outputMode,
      excludeTechnical: state.excludeTechnical,
      convertNumbers: state.convertNumbers,
      includeNewColumns: state.includeNewColumns,
      columns: state.columns
        .slice()
        .sort((a, b) => a.order - b.order)
        .map((column) => ({
          name: column.name,
          include: column.include,
          order: column.order
        })),
      savedAt: new Date().toISOString()
    };
  }

  function applyLayoutObject(layout, notify = true) {
    if (!layout || !Array.isArray(layout.columns)) {
      showToast("El layout no tiene un formato válido.", "error");
      return;
    }

    state.representation = layout.representation === "document" ? "document" : "concept";
    state.outputMode = layout.outputMode === "individual" ? "individual" : "consolidated";
    state.excludeTechnical = layout.excludeTechnical !== false;
    state.convertNumbers = layout.convertNumbers !== false;
    state.includeNewColumns = layout.includeNewColumns !== false;

    $(`input[name=representation][value=${state.representation}]`).checked = true;
    if ($(`input[name=outputMode][value=${state.outputMode}]`)) {
      $(`input[name=outputMode][value=${state.outputMode}]`).checked = true;
    }
    els.excludeTechnical.checked = state.excludeTechnical;
    els.convertNumbers.checked = state.convertNumbers;
    els.includeNewColumns.checked = state.includeNewColumns;
    updateChoiceStyles();

    const layoutMap = new Map(layout.columns.map((column) => [column.name, column]));
    const orderedNames = layout.columns
      .slice()
      .sort((a, b) => Number(a.order) - Number(b.order))
      .map((column) => column.name);

    const allNames = [...orderedNames, ...state.availableColumns.filter((name) => !orderedNames.includes(name))];
    state.columns = allNames.map((name, index) => {
      const saved = layoutMap.get(name);
      return {
        name,
        include: saved ? saved.include !== false : state.includeNewColumns,
        order: index + 1
      };
    });

    state.activeLayoutId = Object.keys(state.layouts).find((id) => state.layouts[id] === layout) || state.activeLayoutId;
    els.layoutName.value = layout.name || "";
    renderColumnList();
    renderPreview();
    updateSummaryUI();
    if (notify) showToast(`Layout “${layout.name || "sin nombre"}” aplicado.`, "success");
  }

  function saveCurrentLayout() {
    if (!state.columns.length) {
      showToast("Primero carga CFDI y configura las columnas.", "error");
      return;
    }

    const payload = currentLayoutPayload();
    const existingId = els.layoutSelect.value;
    const id = existingId || `layout_${Date.now()}`;
    state.layouts[id] = payload;
    state.activeLayoutId = id;

    if (els.setAsDefault.checked) {
      localStorage.setItem(DEFAULT_LAYOUT_KEY, id);
    }

    persistLayouts();
    els.layoutSelect.value = id;
    els.sideLayout.textContent = payload.name;
    showToast("Layout guardado. Se reutilizará en futuras sesiones.", "success");
  }

  function deleteSelectedLayout() {
    const id = els.layoutSelect.value;
    if (!id || !state.layouts[id]) {
      showToast("Selecciona un layout para eliminarlo.", "error");
      return;
    }

    const name = state.layouts[id].name;
    delete state.layouts[id];
    if (localStorage.getItem(DEFAULT_LAYOUT_KEY) === id) {
      localStorage.removeItem(DEFAULT_LAYOUT_KEY);
    }
    if (state.activeLayoutId === id) state.activeLayoutId = "";
    persistLayouts();
    showToast(`Layout “${name}” eliminado.`);
  }

  function exportLayout() {
    const id = els.layoutSelect.value;
    const layout = id && state.layouts[id] ? state.layouts[id] : currentLayoutPayload();
    const blob = new Blob([JSON.stringify(layout, null, 2)], { type: "application/json" });
    downloadBlob(blob, `${normalizeColumn(layout.name || "layout_cfdi")}.json`);
  }

  async function importLayoutFile(file) {
    try {
      const payload = JSON.parse(await file.text());
      if (!Array.isArray(payload.columns)) throw new Error("Falta la lista de columnas.");
      const id = `layout_${Date.now()}`;
      state.layouts[id] = payload;
      persistLayouts();
      els.layoutSelect.value = id;
      state.activeLayoutId = id;
      applyLayoutObject(payload);
      showToast("Layout importado correctamente.", "success");
    } catch (error) {
      showToast(`No se pudo importar el layout: ${error.message}`, "error");
    }
  }

  function updateChoiceStyles() {
    $$(".choice-card").forEach((card) => {
      const input = $("input", card);
      card.classList.toggle("is-selected", input.checked);
    });
    $$(".segmented__option").forEach((option) => {
      const input = $("input", option);
      option.classList.toggle("is-selected", input.checked);
    });
  }

  function shouldConvertToNumber(columnName, value) {
    if (!state.convertNumbers || typeof value !== "string") return false;
    const numericColumn = /(Total|SubTotal|Importe|Cantidad|ValorUnitario|Descuento|Base|TasaOCuota|TipoCambio)$/i.test(columnName);
    return numericColumn && /^-?\d+(?:\.\d+)?$/.test(value.trim());
  }

  function normalizeExcelValue(columnName, value) {
    return shouldConvertToNumber(columnName, value) ? Number(value) : (value ?? "");
  }

  function finalColumnsForRows(documentRows) {
    const selected = getSelectedColumns();
    if (!state.includeNewColumns) return selected;

    const seen = new Set(selected);
    const finalColumns = [...selected];

    documentRows.forEach(({ rows }) => {
      rows.forEach((row) => {
        Object.keys(row).forEach((column) => {
          if (!seen.has(column)) {
            seen.add(column);
            finalColumns.push(column);
          }
        });
      });
    });

    return finalColumns;
  }

  function createWorkbookFromRows(rows, columns, baseSheetName = "CFDI") {
    const workbook = XLSX.utils.book_new();
    if (!rows.length) {
      const worksheet = XLSX.utils.aoa_to_sheet([columns]);
      XLSX.utils.book_append_sheet(workbook, worksheet, baseSheetName.slice(0, 31));
      return workbook;
    }

    for (let start = 0, sheetIndex = 1; start < rows.length; start += XLSX_SHEET_ROW_LIMIT, sheetIndex += 1) {
      const chunk = rows.slice(start, start + XLSX_SHEET_ROW_LIMIT);
      const aoa = [
        columns,
        ...chunk.map((row) => columns.map((column) => normalizeExcelValue(column, row[column])))
      ];
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      worksheet["!autofilter"] = { ref: worksheet["!ref"] };
      worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };
      worksheet["!cols"] = columns.map((column) => ({ wch: Math.min(Math.max(column.length + 2, 12), 38) }));
      const sheetName = (sheetIndex === 1 ? baseSheetName : `${baseSheetName}_${sheetIndex}`).slice(0, 31);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    return workbook;
  }

  function workbookToBlob(workbook) {
    const data = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
    return new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
  }

  async function processForExport() {
    if (!state.analysis.valid || !getSelectedColumns().length) return;

    state.generatedBlob = null;
    els.resultCard.classList.add("is-hidden");
    updateProgress(2, "Preparando exportación", "Revisando la configuración...");

    const documentRows = [];
    const errors = [];
    const totalEntries = state.xmlEntries.length;

    for (let index = 0; index < totalEntries; index += 1) {
      const entry = state.xmlEntries[index];
      updateProgress(
        5 + ((index + 1) / Math.max(totalEntries, 1)) * 58,
        "Procesando CFDI",
        `${index + 1} de ${totalEntries}: ${entry.name}`
      );

      try {
        documentRows.push({ name: entry.name, rows: rowsFromXml(entry) });
      } catch (error) {
        errors.push({ file: entry.name, error: error.message });
      }

      if (index % 20 === 0) await new Promise((resolve) => setTimeout(resolve, 0));
    }

    const columns = finalColumnsForRows(documentRows);
    if (!columns.length) throw new Error("No hay columnas seleccionadas.");

    if (state.analysis.total === 1 || state.outputMode === "consolidated") {
      updateProgress(72, "Creando Excel", "Organizando la tabla consolidada...");
      const allRows = documentRows.flatMap((document) => document.rows);
      const workbook = createWorkbookFromRows(
        allRows,
        columns,
        state.analysis.total === 1 ? "CFDI" : "CFDI_Consolidado"
      );

      if (errors.length) {
        const errorSheet = XLSX.utils.json_to_sheet(errors.map((error) => ({
          Archivo: error.file,
          Error: error.error
        })));
        XLSX.utils.book_append_sheet(workbook, errorSheet, "Errores");
      }

      state.generatedBlob = workbookToBlob(workbook);
      state.generatedName = state.analysis.total === 1
        ? `${normalizeColumn(state.xmlEntries[0].name.replace(/\.xml$/i, "").split("/").pop())}_convertido.xlsx`
        : "CFDI_Consolidado.xlsx";
    } else {
      updateProgress(70, "Creando archivos individuales", "Preparando el paquete ZIP...");
      const outputZip = new JSZip();
      const names = new Map();

      for (let index = 0; index < documentRows.length; index += 1) {
        const document = documentRows[index];
        updateProgress(
          70 + ((index + 1) / Math.max(documentRows.length, 1)) * 24,
          "Creando archivos individuales",
          `${index + 1} de ${documentRows.length}: ${document.name}`
        );

        const workbook = createWorkbookFromRows(document.rows, columns, "CFDI");
        const workbookData = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
        const base = `${normalizeColumn(document.name.replace(/\.xml$/i, "").split("/").pop())}_convertido.xlsx`;
        const count = (names.get(base) || 0) + 1;
        names.set(base, count);
        const finalName = count === 1 ? base : base.replace(/\.xlsx$/i, `_${count}.xlsx`);
        outputZip.file(finalName, workbookData);
      }

      if (errors.length) {
        const csv = [
          "Archivo,Error",
          ...errors.map((error) => `"${String(error.file).replaceAll('"', '""')}","${String(error.error).replaceAll('"', '""')}"`)
        ].join("\n");
        outputZip.file("reporte_errores.csv", "\ufeff" + csv);
      }

      state.generatedBlob = await outputZip.generateAsync(
        { type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } },
        (metadata) => updateProgress(94 + metadata.percent * 0.05, "Comprimiendo ZIP", `${Math.round(metadata.percent)}%`)
      );
      state.generatedName = "CFDI_Individuales.zip";
    }

    updateProgress(100, "Proceso terminado", "El archivo está listo para descargar.");
    els.resultText.textContent =
      `${formatNumber(documentRows.length)} comprobantes procesados · ${formatNumber(columns.length)} columnas · ${formatNumber(errors.length)} observaciones.`;
    els.resultCard.classList.remove("is-hidden");
    updateStepper();
    setTimeout(hideProgress, 650);
    els.resultCard.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function renderErrors() {
    els.errorList.innerHTML = state.analysis.errors.length
      ? state.analysis.errors.map((error) => `
          <div class="error-row">
            <strong>${escapeHtml(error.file)}</strong>
            <small>${escapeHtml(error.error)}</small>
          </div>
        `).join("")
      : "<p>No hay observaciones.</p>";
  }

  // Upload.
  els.fileInput.addEventListener("change", (event) => handleFiles(event.target.files));
  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.add("is-dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.remove("is-dragover");
    });
  });
  els.dropzone.addEventListener("drop", (event) => handleFiles(event.dataTransfer.files));

  // Representation and output.
  $$("input[name=representation]").forEach((input) => {
    input.addEventListener("change", async () => {
      state.representation = input.value;
      updateChoiceStyles();
      await buildPreviewAndColumns(false);
    });
  });
  $$("input[name=outputMode]").forEach((input) => {
    input.addEventListener("change", () => {
      state.outputMode = input.value;
      updateChoiceStyles();
      updateSummaryUI();
    });
  });

  els.excludeTechnical.addEventListener("change", async () => {
    state.excludeTechnical = els.excludeTechnical.checked;
    await buildPreviewAndColumns(false);
  });
  els.convertNumbers.addEventListener("change", () => {
    state.convertNumbers = els.convertNumbers.checked;
  });
  els.includeNewColumns.addEventListener("change", () => {
    state.includeNewColumns = els.includeNewColumns.checked;
  });

  // Columns.
  els.columnSearch.addEventListener("input", renderColumnList);
  els.selectAll.addEventListener("click", () => {
    state.columns.forEach((column) => { column.include = true; });
    renderColumnList();
    renderPreview();
  });
  els.clearAll.addEventListener("click", () => {
    state.columns.forEach((column) => { column.include = false; });
    renderColumnList();
    renderPreview();
  });
  els.selectRecommended.addEventListener("click", () => {
    state.columns.forEach((column) => { column.include = recommendedColumn(column.name); });
    renderColumnList();
    renderPreview();
    showToast("Se seleccionaron los campos contables principales.", "success");
  });
  els.refreshPreview.addEventListener("click", renderPreview);

  // Layouts.
  els.saveLayout.addEventListener("click", saveCurrentLayout);
  els.applyLayout.addEventListener("click", () => {
    const layout = state.layouts[els.layoutSelect.value];
    if (!layout) {
      showToast("Selecciona un layout guardado.", "error");
      return;
    }
    state.activeLayoutId = els.layoutSelect.value;
    applyLayoutObject(layout);
  });
  els.deleteLayout.addEventListener("click", deleteSelectedLayout);
  els.exportLayout.addEventListener("click", exportLayout);
  els.importLayout.addEventListener("click", () => els.layoutFileInput.click());
  els.layoutFileInput.addEventListener("change", (event) => {
    if (event.target.files?.[0]) importLayoutFile(event.target.files[0]);
    event.target.value = "";
  });
  els.layoutSelect.addEventListener("change", () => {
    const layout = state.layouts[els.layoutSelect.value];
    els.layoutName.value = layout?.name || "";
  });

  // Dialogs and navigation.
  els.openHelp.addEventListener("click", () => els.helpDialog.showModal());
  els.viewErrors.addEventListener("click", () => {
    renderErrors();
    els.errorsDialog.showModal();
  });
  $$("[data-close-dialog]").forEach((button) => {
    button.addEventListener("click", () => {
      document.getElementById(button.dataset.closeDialog)?.close();
    });
  });
  $$(".step").forEach((step) => {
    step.addEventListener("click", () => {
      document.querySelector(`[data-step="${step.dataset.stepTarget}"]`)
        ?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });

  // Export.
  els.processButton.addEventListener("click", async () => {
    try {
      await processForExport();
    } catch (error) {
      hideProgress();
      showToast(`No se pudo generar el archivo: ${error.message}`, "error");
    }
  });
  els.downloadButton.addEventListener("click", () => {
    if (state.generatedBlob) {
      downloadBlob(state.generatedBlob, state.generatedName);
    }
  });

  // Initialization.
  renderLayoutSelect();
  updateChoiceStyles();
  updateSummaryUI();
})();
