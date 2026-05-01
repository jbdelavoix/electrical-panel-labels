const SUPPORTED_LANGUAGES = ["fr", "en", "de", "es"];
let I18N = Object.fromEntries(SUPPORTED_LANGUAGES.map((language) => [language, {}]));
let PAPER_FORMATS = [];
let ICON_DEFINITIONS = [];

async function loadTranslations() {
  await Promise.all(SUPPORTED_LANGUAGES.map(async (language) => {
    try {
      const response = await fetch(`i18n/${language}.json`, { cache: "no-cache" });
      if (!response.ok) throw new Error(`Failed to load ${language}`);
      I18N[language] = await response.json();
    } catch {
      I18N[language] = {};
    }
  }));
}

async function loadPaperFormats() {
  try {
    const response = await fetch("config/paper-formats.json", { cache: "no-cache" });
    if (!response.ok) throw new Error("Failed to load paper formats");
    const payload = await response.json();
    PAPER_FORMATS = Array.isArray(payload?.formats) ? payload.formats : [];
  } catch {
    PAPER_FORMATS = [];
  }
}

async function loadIconDefinitions() {
  try {
    const response = await fetch("config/icons.json", { cache: "no-cache" });
    if (!response.ok) throw new Error("Failed to load icons");
    const payload = await response.json();
    ICON_DEFINITIONS = Array.isArray(payload?.icons) ? payload.icons : [];
  } catch {
    ICON_DEFINITIONS = [];
  }
}

const COLOR_OPTIONS = [
  { value: "", labelKey: "colorNone" },
  { value: "blue", labelKey: "colorBlue" },
  { value: "ochre", labelKey: "colorOchre" },
  { value: "purple", labelKey: "colorPurple" },
  { value: "red", labelKey: "colorRed" },
  { value: "green", labelKey: "colorGreen" },
  { value: "orange", labelKey: "colorOrange" },
  { value: "brown", labelKey: "colorBrown" },
  { value: "teal", labelKey: "colorTeal" },
];


const state = {
  language: "fr",
  themeMode: "light",
  projectTitle: "",
  paperFormat: "a4",
  moduleWidthMm: 18,
  labelHeightMm: 30,
  modulesPerRow: 13,
  selectedRowIndex: 0,
  selectedModuleIndex: 0,
  text90Mode: false,
  permalinkUrl: "",
  rowScrollLeftByIndex: {},
  rows: [{ modules: [] }],
};

const PRINT_MARGIN_CM = 2.5;
const PRINT_CHUNK_LABEL_CM = 0.45;
const PRINT_CHUNK_GAP_CM = 0.25;
const DEFAULT_PERMALINK_QUERY = "?json&enc=gz-b64&paper=a4&mode=standard&data=H4sIAAAAAAAAE62QTWrDMBCFrxJmbRf6t9EuuKUEEgikdFNnIdsTWWU8cmUpKQTfpz6HL1bctLYpSdpFhRZi3vdmRm8PpTUvmDoQsOxeDgIoZYkWBMgbCKAoLYjL6wCs2VUgnvdDpTCZJzwUdaFAQJVrpAwCIJOCgDu92bSNRXYaCT4NIK7q4BsnrXKXeEp6x31KUlupMOaVJMMxr6YLGBwledXDS6srrE6QZ3tHuSwSi9VvnY9wCWnOqp58MoRurEttwy2yG4hF1Mnj-4U6oxRhaLtNe_yxbahtrC8degv1usP_L_GN5OGT02gyY3cBp-X7t4P8c_MjU_HMUJIFDn1JvnqMecaZT502PJq_k1WuWYWFTHPNg2cutziZa1Y4hjUPK6za9zTHmHtolNztH4KLOexPzJHxpbcY80PbcNtYSdj_K4DUEAiwmI2iWdfr-gNJ3fioTQMAAA";

function getAvailablePaperFormats() {
  return PAPER_FORMATS;
}

function getPaperFormatValues() {
  return getAvailablePaperFormats().map((format) => format.id);
}

function getPaperFormatById(paperId) {
  return getAvailablePaperFormats().find((format) => format.id === paperId) || getAvailablePaperFormats()[0];
}

function formatPaperOptionLabel(format) {
  if (!format) return "";
  const label = String(format.id || "").toUpperCase();
  return `${label} (${format.standard}, ${format.widthCm} x ${format.heightCm} cm)`;
}

function applyThemePreference() {
  document.documentElement.classList.toggle("forced-dark", state.themeMode === "dark");
  document.documentElement.classList.toggle("forced-light", state.themeMode === "light");
}

function text(key) {
  return I18N[state.language]?.[key] || I18N.en?.[key] || key;
}

function textWithVars(key, vars) {
  let value = text(key);
  for (const [token, tokenValue] of Object.entries(vars)) {
    value = value.replace(`{${token}}`, String(tokenValue));
  }
  return value;
}

function getIconLabel(definition) {
  if (!definition) return "-";
  const iconKey = String(definition.value || "");
  return I18N[state.language]?.icons?.[iconKey]
    || I18N.en?.icons?.[iconKey]
    || definition.value
    || "-";
}

function detectLanguageFromBrowser() {
  const rawLanguage = (navigator.language || "en").toLowerCase();
  const matchedLanguage = SUPPORTED_LANGUAGES.find((language) => rawLanguage.startsWith(language));
  return matchedLanguage || "en";
}

async function copyTextToClipboard(value) {
  const safeValue = String(value || "");
  if (!safeValue) return false;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(safeValue);
      return true;
    }
  } catch {
    // Fallback below.
  }
  try {
    const temp = document.createElement("textarea");
    temp.value = safeValue;
    temp.setAttribute("readonly", "");
    temp.style.position = "fixed";
    temp.style.opacity = "0";
    document.body.appendChild(temp);
    temp.select();
    temp.setSelectionRange(0, safeValue.length);
    const ok = document.execCommand("copy");
    temp.remove();
    return ok;
  } catch {
    return false;
  }
}

function getModuleWidthMm() {
  const parsed = Number(state.moduleWidthMm);
  if (Number.isNaN(parsed)) return 18;
  return Math.max(10, Math.min(30, parsed));
}

function getModuleWidthCm() {
  return getModuleWidthMm() / 10;
}

function applyModuleWidthCss() {
  document.documentElement.style.setProperty("--module-width", `${getModuleWidthMm()}mm`);
}

function getLabelHeightMm() {
  const parsed = Number(state.labelHeightMm);
  if (Number.isNaN(parsed)) return 30;
  return Math.max(20, Math.min(80, parsed));
}

function getLabelHeightCm() {
  return getLabelHeightMm() / 10;
}

function applyLabelHeightCss() {
  const labelHeightMm = getLabelHeightMm();
  const iconHeightMm = Math.max(8, Math.min(24, (0.4 * labelHeightMm) - 2));
  const locationHeightMm = Math.max(8, labelHeightMm - iconHeightMm);
  document.documentElement.style.setProperty("--label-height", `${labelHeightMm}mm`);
  document.documentElement.style.setProperty("--icon-height", `${iconHeightMm.toFixed(2)}mm`);
  document.documentElement.style.setProperty("--location-height", `${locationHeightMm.toFixed(2)}mm`);
}

function trimLocationByModule(locationValue, moduleSize) {
  return String(locationValue || "");
}

function getDisplayLocationHtml(module, moduleSize) {
  return trimLocationByModule(module.loc || "", moduleSize).replaceAll("\n", "<br>");
}

function renderLocationCellHtml(module, moduleSpan) {
  if (state.text90Mode) {
    const locationText = trimLocationByModule(module.loc || "", moduleSpan).replaceAll("\n", "<br>");
    return `<span class="text-rotate-wrap text-rotate-wrap-multiline">${locationText}</span>`;
  }
  return getDisplayLocationHtml(module, moduleSpan);
}

function normalizeModule(module) {
  if (!module.img) module.img = "";
  if (module.mod) module.mod = Number(module.mod);
  if (!module.mod || Number.isNaN(module.mod)) module.mod = 1;
  if (!module.col) module.col = "";
  module.loc = trimLocationByModule(module.loc || "", module.mod);
}

function getOccupiedSlots(modules) {
  let occupied = modules.length;
  for (const mod of modules) occupied += Math.max(1, Number(mod.mod || 1)) - 1;
  return occupied;
}

function createEmptyModule() {
  return { loc: "", mod: 1, img: "", col: "" };
}

function moduleColorClass(module) {
  return module.col || "";
}

function getRowModulesPerRow(row) {
  const localSize = Number(row?.mpr);
  if (!Number.isNaN(localSize) && localSize >= 1 && localSize <= 24) return localSize;
  return state.modulesPerRow;
}

function ensureRowSize(modules, targetModulesPerRow = state.modulesPerRow) {
  let occupied = getOccupiedSlots(modules);
  while (occupied > targetModulesPerRow && modules.length > 1) {
    const removed = modules.pop();
    occupied -= Math.max(1, Number(removed.mod || 1));
  }
  while (occupied < targetModulesPerRow) {
    modules.push(createEmptyModule());
    occupied += 1;
  }
}

function normalizeRows(rows) {
  for (const row of rows) {
    row.modules = row.modules || [];
    row.modules.forEach(normalizeModule);
    ensureRowSize(row.modules, getRowModulesPerRow(row));
  }
}

function getSelectedModule() {
  const row = state.rows[state.selectedRowIndex] || state.rows[0];
  return row.modules[state.selectedModuleIndex] || row.modules[0];
}

function getEffectiveModuleSize(module, rowIndex, moduleIndex) {
  return Math.max(1, Number(module.mod || 1));
}

function buildDisplayModules(row, rowIndex) {
  const modules = row.modules || [];
  return modules.map((module, index) => ({ module, sourceIndex: index }));
}

function getMaxSpanForModule(rowIndex, moduleIndex) {
  const row = state.rows[rowIndex];
  if (!row) return 1;
  const rowModulesPerRow = getRowModulesPerRow(row);
  let occupiedBefore = 0;
  for (let index = 0; index < moduleIndex; index += 1) {
    const module = row.modules[index];
    occupiedBefore += Math.max(1, Number(module?.mod || 1));
  }
  const maxByWidth = rowModulesPerRow - occupiedBefore;
  return Math.max(1, Math.min(6, maxByWidth));
}

function consumeSlotsAfterItems(items, startIndex, slotsToConsume) {
  let remaining = slotsToConsume;
  let cursor = startIndex + 1;
  while (remaining > 0 && cursor < items.length) {
    const currentItem = items[cursor];
    const currentSize = Math.max(1, Number(currentItem.module?.mod || 1));
    if (currentSize <= remaining) {
      items.splice(cursor, 1);
      remaining -= currentSize;
    } else {
      currentItem.module.mod = currentSize - remaining;
      remaining = 0;
    }
  }
}

function applyModuleSizeInItems(items, moduleIndex, targetSize) {
  const selectedItem = items[moduleIndex];
  if (!selectedItem) return;
  const previousSize = Math.max(1, Number(selectedItem.module?.mod || 1));
  const nextSize = Math.max(1, Number(targetSize || 1));
  if (nextSize === previousSize) return;

  if (nextSize > previousSize) {
    consumeSlotsAfterItems(items, moduleIndex, nextSize - previousSize);
  } else {
    const slotsToInsert = previousSize - nextSize;
    const inserted = Array.from({ length: slotsToInsert }, () => ({ module: createEmptyModule(), sourceIndex: null }));
    items.splice(moduleIndex + 1, 0, ...inserted);
  }

  selectedItem.module.mod = nextSize;
}

function applyModuleSizeInRow(row, moduleIndex, targetSize) {
  const items = row.modules.map((module, index) => ({ module, sourceIndex: index }));
  applyModuleSizeInItems(items, moduleIndex, targetSize);
  row.modules = items.map((item) => item.module);
}

function moduleClass(module, rowIndex, moduleIndex) {
  const widthClass = `mod-${getEffectiveModuleSize(module, rowIndex, moduleIndex)}`;
  const selected = rowIndex === state.selectedRowIndex && moduleIndex === state.selectedModuleIndex ? " is-selected" : "";
  return widthClass + selected;
}

function getRenderedItemsForRow(row, rowIndex) {
  const rowModulesPerRow = getRowModulesPerRow(row);
  const displayModules = buildDisplayModules(row, rowIndex);
  const renderedItems = [];
  let occupiedSlots = 0;
  for (let displayIndex = 0; displayIndex < displayModules.length; displayIndex += 1) {
    if (occupiedSlots >= rowModulesPerRow) break;
    const displayItem = displayModules[displayIndex];
    const module = displayItem.module;
    const sourceIndex = displayItem.sourceIndex;
    const requestedSpan = sourceIndex === null
      ? Math.max(1, Number(module.mod || 1))
      : getEffectiveModuleSize(module, rowIndex, sourceIndex);
    const moduleSpan = Math.min(requestedSpan, rowModulesPerRow - occupiedSlots);
    renderedItems.push({ module, sourceIndex, moduleSpan });
    occupiedSlots += moduleSpan;
  }
  return renderedItems;
}

function buildModuleTable(items, rowIndex, options = {}) {
  const slotCount = Math.max(1, Number(options.slotCount || 1));
  const interactive = Boolean(options.interactive);

  const table = document.createElement("table");
  table.className = "module-table";
  table.style.width = `${(getModuleWidthMm() * slotCount).toFixed(2)}mm`;
  table.style.minWidth = `${(getModuleWidthMm() * slotCount).toFixed(2)}mm`;

  const colgroup = document.createElement("colgroup");
  for (let colIndex = 0; colIndex < slotCount; colIndex += 1) {
    const col = document.createElement("col");
    col.style.width = `${getModuleWidthMm()}mm`;
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);

  const trIcon = document.createElement("tr");
  trIcon.className = "cell-icon";
  const trLoc = document.createElement("tr");
  trLoc.className = "cell-location";

  for (const item of items) {
    const module = item.module;
    const sourceIndex = item.sourceIndex;
    const moduleSpan = item.moduleSpan;
    const baseClass = sourceIndex === null ? `mod-${moduleSpan}` : moduleClass(module, rowIndex, sourceIndex);

    const tdIcon = document.createElement("td");
    tdIcon.className = `${baseClass} ${moduleColorClass(module)}`;
    tdIcon.colSpan = moduleSpan;
    tdIcon.innerHTML = module.img ? `<i data-lucide="${module.img}" class="inline-block w-8 h-8"></i>` : "";

    const tdLoc = document.createElement("td");
    tdLoc.className = `${baseClass} ${moduleColorClass(module)}`;
    tdLoc.colSpan = moduleSpan;
    if (interactive) {
      tdLoc.dataset.rowIndex = String(rowIndex);
      tdLoc.dataset.moduleIndex = String(sourceIndex ?? -1);
      tdLoc.dataset.moduleSpan = String(moduleSpan);
    }
    tdLoc.innerHTML = renderLocationCellHtml(module, moduleSpan);
    if (interactive && sourceIndex !== null) {
      tdIcon.onclick = (event) => {
        const previousTop = event.currentTarget instanceof Element
          ? event.currentTarget.getBoundingClientRect().top
          : null;
        selectModule(rowIndex, sourceIndex);
        renderAllPreserveViewport(previousTop);
      };
      tdLoc.onclick = (event) => {
        const previousTop = event.currentTarget instanceof Element
          ? event.currentTarget.getBoundingClientRect().top
          : null;
        selectModule(rowIndex, sourceIndex);
        renderAllPreserveViewport(previousTop);
      };
    }

    if (!state.text90Mode) {
      trIcon.appendChild(tdIcon);
    }
    trLoc.appendChild(tdLoc);

  }

  if (!state.text90Mode) {
    table.append(trIcon);
  }
  table.append(trLoc);
  return table;
}

function splitItemsForPrint(items, maxSlotsPerRow) {
  const totalSlots = items.reduce((sum, item) => sum + item.moduleSpan, 0);

  // If the row needs exactly 2 chunks, prefer the most balanced split
  // (same number of slots on each side) without exceeding max width.
  if (totalSlots > maxSlotsPerRow && totalSlots <= (2 * maxSlotsPerRow)) {
    let bestSplitIndex = -1;
    let bestDelta = Number.POSITIVE_INFINITY;
    let leftSlots = 0;
    for (let splitIndex = 1; splitIndex < items.length; splitIndex += 1) {
      leftSlots += items[splitIndex - 1].moduleSpan;
      const rightSlots = totalSlots - leftSlots;
      if (leftSlots <= maxSlotsPerRow && rightSlots <= maxSlotsPerRow) {
        const delta = Math.abs(leftSlots - rightSlots);
        if (delta < bestDelta) {
          bestDelta = delta;
          bestSplitIndex = splitIndex;
        }
      }
    }
    if (bestSplitIndex > 0) {
      return [items.slice(0, bestSplitIndex), items.slice(bestSplitIndex)];
    }
  }

  const chunks = [];
  let currentChunk = [];
  let currentSlots = 0;
  for (const item of items) {
    if (currentSlots > 0 && currentSlots + item.moduleSpan > maxSlotsPerRow) {
      chunks.push(currentChunk);
      currentChunk = [];
      currentSlots = 0;
    }
    currentChunk.push(item);
    currentSlots += item.moduleSpan;
  }
  if (currentChunk.length > 0) {
    chunks.push(currentChunk);
  }
  return chunks;
}

function getPaperSizeForOrientation(orientation) {
  const paper = getPaperFormatById(state.paperFormat);
  const width = Number(paper?.widthCm || 21);
  const height = Number(paper?.heightCm || 29.7);
  const pageWidth = orientation === "portrait" ? Math.min(width, height) : Math.max(width, height);
  const pageHeight = orientation === "portrait" ? Math.max(width, height) : Math.min(width, height);
  return { pageWidth, pageHeight };
}

function getMaxSlotsForOrientation(orientation) {
  const { pageWidth } = getPaperSizeForOrientation(orientation);
  const usableWidth = Math.max(0, pageWidth - (2 * PRINT_MARGIN_CM));
  return Math.max(1, Math.floor(usableWidth / getModuleWidthCm()));
}

function getPrintedRowHeightCm() {
  return getLabelHeightCm();
}

function buildPrintPlan() {
  const rowItems = state.rows.map((row, rowIndex) => getRenderedItemsForRow(row, rowIndex));
  const maxRowModulesPerRow = state.rows.reduce(
    (maxValue, row) => Math.max(maxValue, getRowModulesPerRow(row)),
    state.modulesPerRow,
  );
  const portraitCapacity = getMaxSlotsForOrientation("portrait");
  const landscapeCapacity = getMaxSlotsForOrientation("landscape");
  const canFitFullRowInPortrait = maxRowModulesPerRow <= portraitCapacity;
  const canFitFullRowInLandscape = maxRowModulesPerRow <= landscapeCapacity;

  // Never pick an orientation that would force unnecessary horizontal splitting
  // if another orientation can keep a full row intact.
  let orientations = ["portrait", "landscape"];
  if (!canFitFullRowInPortrait && canFitFullRowInLandscape) {
    orientations = ["landscape"];
  } else if (!canFitFullRowInLandscape && canFitFullRowInPortrait) {
    orientations = ["portrait"];
  }

  let bestPlan = null;
  for (const orientation of orientations) {
    const maxSlotsPerChunk = getMaxSlotsForOrientation(orientation);
    const { pageHeight } = getPaperSizeForOrientation(orientation);
    const usableHeight = Math.max(0, pageHeight - (2 * PRINT_MARGIN_CM));
    const chunkHeight = getPrintedRowHeightCm() + PRINT_CHUNK_LABEL_CM;
    const chunksPerPage = Math.max(1, Math.floor((usableHeight + PRINT_CHUNK_GAP_CM) / (chunkHeight + PRINT_CHUNK_GAP_CM)));
    const totalChunks = rowItems.reduce((sum, items) => sum + splitItemsForPrint(items, maxSlotsPerChunk).length, 0);
    const estimatedPages = Math.max(1, Math.ceil(totalChunks / chunksPerPage));
    const candidate = { orientation, maxSlotsPerChunk, chunksPerPage, totalChunks, estimatedPages };
    if (
      !bestPlan
      || candidate.estimatedPages < bestPlan.estimatedPages
      || (candidate.estimatedPages === bestPlan.estimatedPages && candidate.totalChunks < bestPlan.totalChunks)
      || (candidate.estimatedPages === bestPlan.estimatedPages
        && candidate.totalChunks === bestPlan.totalChunks
        && candidate.maxSlotsPerChunk > bestPlan.maxSlotsPerChunk)
    ) {
      bestPlan = candidate;
    }
  }
  return bestPlan || { orientation: "landscape", maxSlotsPerChunk: getMaxSlotsForOrientation("landscape"), chunksPerPage: 1, totalChunks: 1, estimatedPages: 1 };
}

function renderLabels() {
  const root = document.getElementById("labels-root");
  const rowsRoot = document.getElementById("rows-root");
  const editorPanel = document.getElementById("editor-panel-container");
  const editorPanelParking = document.getElementById("editor-panel-parking");
  if (!root || !rowsRoot) return;
  if (editorPanel && editorPanelParking) {
    editorPanelParking.appendChild(editorPanel);
  }
  const existingScrollContainers = rowsRoot.querySelectorAll("[data-row-scroll-index]");
  existingScrollContainers.forEach((element) => {
    const rowIndex = Number(element.getAttribute("data-row-scroll-index"));
    if (!Number.isNaN(rowIndex)) {
      state.rowScrollLeftByIndex[rowIndex] = element.scrollLeft;
    }
  });
  root.className = `w-full px-4 pb-4 ${state.text90Mode ? "mode-text-90" : ""}`;
  rowsRoot.innerHTML = "";
  const printProjectTitle = document.createElement("div");
  printProjectTitle.className = "print-only";
  printProjectTitle.style.fontSize = "14pt";
  printProjectTitle.style.fontWeight = "700";
  printProjectTitle.style.marginBottom = "0.25cm";
  printProjectTitle.textContent = String(state.projectTitle || "").trim() || text("project");
  rowsRoot.appendChild(printProjectTitle);
  const printPlan = buildPrintPlan();
  let usedChunksOnCurrentPrintPage = 0;

  state.rows.forEach((row, rowIndex) => {
    const rowModulesPerRow = getRowModulesPerRow(row);
    const wrapper = document.createElement("section");
    wrapper.className = "labels-row border p-3 mt-4";
    const rowHeader = document.createElement("div");
    rowHeader.className = "hide-print flex items-center justify-between mb-2";
    const rowLeft = document.createElement("div");
    rowLeft.className = "flex items-center gap-2";
    const rowTitle = document.createElement("span");
    rowTitle.className = "text-sm font-semibold";
    rowTitle.textContent = textWithVars("rowTitle", { row: rowIndex + 1 });
    const rowModulesLabel = document.createElement("span");
    rowModulesLabel.className = "text-xs text-slate-600";
    rowModulesLabel.textContent = text("rowSize");
    const rowModulesInput = document.createElement("input");
    rowModulesInput.type = "number";
    rowModulesInput.min = "1";
    rowModulesInput.max = "24";
    rowModulesInput.value = String(rowModulesPerRow);
    rowModulesInput.className = "ui-input w-20 px-2 py-1 text-sm";
    rowModulesInput.onchange = (event) => {
      const nextSize = Number(event.target.value);
      if (Number.isNaN(nextSize) || nextSize < 1 || nextSize > 24) {
        event.target.value = String(getRowModulesPerRow(row));
        return;
      }
      row.mpr = nextSize;
      ensureRowSize(row.modules, getRowModulesPerRow(row));
      if (rowIndex === state.selectedRowIndex) {
        state.selectedModuleIndex = Math.max(0, Math.min(state.selectedModuleIndex, row.modules.length - 1));
      }
      renderAll();
    };
    rowLeft.appendChild(rowTitle);
    rowLeft.appendChild(rowModulesLabel);
    rowLeft.appendChild(rowModulesInput);
    const rowDeleteBtn = document.createElement("button");
    rowDeleteBtn.type = "button";
    rowDeleteBtn.className = "ui-btn ui-input w-8 h-8 inline-flex items-center justify-center";
    rowDeleteBtn.setAttribute("title", text("deleteRowShort"));
    rowDeleteBtn.setAttribute("aria-label", text("deleteRowShort"));
    rowDeleteBtn.innerHTML = `<i data-lucide="trash-2" class="w-4 h-4"></i>`;
    rowDeleteBtn.onclick = () => {
      state.rows.splice(rowIndex, 1);
      if (state.rows.length === 0) state.rows.push({ modules: [] });
      normalizeRows(state.rows);
      state.selectedRowIndex = Math.max(0, Math.min(state.selectedRowIndex, state.rows.length - 1));
      const selectedRow = state.rows[state.selectedRowIndex];
      state.selectedModuleIndex = Math.max(0, Math.min(state.selectedModuleIndex, (selectedRow?.modules?.length || 1) - 1));
      renderAll();
    };
    rowHeader.appendChild(rowLeft);
    rowHeader.appendChild(rowDeleteBtn);
    wrapper.appendChild(rowHeader);
    const scrollContainer = document.createElement("div");
    scrollContainer.className = "row-scroll screen-only";
    scrollContainer.setAttribute("data-row-scroll-index", String(rowIndex));
    scrollContainer.addEventListener("scroll", () => {
      state.rowScrollLeftByIndex[rowIndex] = scrollContainer.scrollLeft;
    });
    const renderedItems = getRenderedItemsForRow(row, rowIndex);

    const table = buildModuleTable(renderedItems, rowIndex, {
      slotCount: rowModulesPerRow,
      interactive: true,
    });
    scrollContainer.appendChild(table);
    wrapper.appendChild(scrollContainer);

    const printContainer = document.createElement("div");
    printContainer.className = "print-only print-split";
    const printChunks = splitItemsForPrint(renderedItems, printPlan.maxSlotsPerChunk);
    const rowChunkCount = printChunks.length;
    if (
      rowChunkCount <= printPlan.chunksPerPage
      && usedChunksOnCurrentPrintPage > 0
      && (usedChunksOnCurrentPrintPage + rowChunkCount) > printPlan.chunksPerPage
    ) {
      wrapper.classList.add("sheet-row-break");
      usedChunksOnCurrentPrintPage = 0;
    }
    for (let chunkIndex = 0; chunkIndex < printChunks.length; chunkIndex += 1) {
      const chunk = printChunks[chunkIndex];
      const chunkSlots = chunk.reduce((total, item) => total + item.moduleSpan, 0);
      const chunkWrapper = document.createElement("div");
      chunkWrapper.className = "print-chunk";
      const chunkLabel = document.createElement("div");
      chunkLabel.className = "print-chunk-label";
      chunkLabel.textContent = textWithVars("printLinePart", {
        row: rowIndex + 1,
        part: chunkIndex + 1,
        total: printChunks.length,
      });
      chunkWrapper.appendChild(chunkLabel);
      chunkWrapper.appendChild(buildModuleTable(chunk, rowIndex, {
        slotCount: chunkSlots,
        interactive: false,
      }));
      printContainer.appendChild(chunkWrapper);
    }
    usedChunksOnCurrentPrintPage += rowChunkCount;
    if (usedChunksOnCurrentPrintPage >= printPlan.chunksPerPage) {
      usedChunksOnCurrentPrintPage %= printPlan.chunksPerPage;
    }
    wrapper.appendChild(printContainer);

    rowsRoot.appendChild(wrapper);
    if (rowIndex === state.selectedRowIndex && editorPanel) {
      rowsRoot.appendChild(editorPanel);
    }
    const savedScrollLeft = state.rowScrollLeftByIndex[rowIndex];
    if (typeof savedScrollLeft === "number") {
      scrollContainer.scrollLeft = savedScrollLeft;
    }
  });
}

function selectModule(rowIndex, moduleIndex) {
  state.selectedRowIndex = rowIndex;
  state.selectedModuleIndex = moduleIndex;
}

function renderAllPreserveViewport(previousTop = null) {
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  renderAll();
  if (typeof previousTop === "number") {
    const selectedCell = document.querySelector(".cell-location td.is-selected, .cell-icon td.is-selected");
    if (selectedCell) {
      const nextTop = selectedCell.getBoundingClientRect().top;
      window.scrollBy(0, nextTop - previousTop);
      return;
    }
  }
  window.scrollTo(scrollX, scrollY);
}

function renderEditor() {
  const module = getSelectedModule();
  document.getElementById("editor-title").textContent = `${text("editor")} #${state.selectedRowIndex + 1}.${state.selectedModuleIndex + 1}`;
  document.getElementById("icon-label").textContent = text("icon");
  document.getElementById("location-custom-label").textContent = text("locationCustom");
  document.getElementById("size-label").textContent = text("size");
  document.getElementById("color-label").textContent = text("color");

  setupIconPicker(module);

  const locationInput = document.getElementById("location-input");
  locationInput.value = module.loc || "";
  locationInput.rows = 3;
  locationInput.oninput = (e) => {
    module.loc = trimLocationByModule(e.target.value, module.mod || 1);
    if (module.loc !== e.target.value) {
      locationInput.value = module.loc;
    }
    // Keep focus/caret: update only selected preview cell instead of re-rendering rows/editor.
    const selectedLocationCell = document.querySelector(".cell-location td.is-selected");
    if (selectedLocationCell) {
      const moduleSpan = Math.max(1, Number(selectedLocationCell.dataset.moduleSpan || module.mod || 1));
      selectedLocationCell.innerHTML = renderLocationCellHtml(module, moduleSpan);
    }
    refreshPermalink();
  };

  const sizeSelect = document.getElementById("size-select");
  const maxSize = getMaxSpanForModule(state.selectedRowIndex, state.selectedModuleIndex);
  sizeSelect.innerHTML = Array.from({ length: maxSize }, (_, i) => i + 1).map((size) => `<option value="${size}">${size}</option>`).join("");
  const selectedSize = Math.min(maxSize, Number(module.mod || 1));
  sizeSelect.value = String(selectedSize);
  sizeSelect.onchange = (e) => {
    const row = state.rows[state.selectedRowIndex];
    const currentModule = row.modules[state.selectedModuleIndex];
    const finalSize = Math.min(getMaxSpanForModule(state.selectedRowIndex, state.selectedModuleIndex), Number(e.target.value));
    applyModuleSizeInRow(row, state.selectedModuleIndex, finalSize);
    currentModule.mod = finalSize;
    currentModule.loc = trimLocationByModule(currentModule.loc, currentModule.mod);
    ensureRowSize(row.modules, getRowModulesPerRow(row));
    renderAll();
  };

  const colorSelect = document.getElementById("color-select");
  colorSelect.innerHTML = COLOR_OPTIONS.map((option) => `<option value="${option.value}">${text(option.labelKey)}</option>`).join("");
  colorSelect.value = module.col || "";
  colorSelect.onchange = (e) => {
    module.col = e.target.value || "";
    renderAll();
  };

}

function setupIconPicker(module) {
  const button = document.getElementById("icon-picker-button");
  const panel = document.getElementById("icon-picker-panel");
  const value = document.getElementById("icon-picker-value");
  const list = document.getElementById("icon-picker-list");
  const search = document.getElementById("icon-picker-search");
  if (!button || !panel || !value || !list || !search) return;

  const currentIcon = module.img || "";
  const currentDefinition = ICON_DEFINITIONS.find((definition) => definition.icon === currentIcon);
  const safeLabel = currentDefinition ? getIconLabel(currentDefinition) : "-";
  if (currentIcon) {
    value.innerHTML = `<span class="inline-flex items-center gap-2"><i data-lucide="${currentIcon}" class="w-5 h-5"></i><span class="truncate">${safeLabel}</span></span>`;
  } else {
    value.innerHTML = `<span class="inline-flex items-center gap-2"><span class="truncate">${safeLabel}</span></span>`;
  }

  const sortedDefinitions = [...ICON_DEFINITIONS].sort((a, b) => {
    const aLabel = getIconLabel(a);
    const bLabel = getIconLabel(b);
    return aLabel.localeCompare(bLabel, state.language, { sensitivity: "base" });
  });

  list.innerHTML = sortedDefinitions.map((definition) => {
    const icon = definition.icon || "";
    const label = getIconLabel(definition);
    const iconHtml = icon ? `<i data-lucide="${icon}" class="w-4 h-4"></i>` : "";
    const searchTokens = [
      definition.value || "",
      I18N.fr?.icons?.[definition.value || ""] || "",
      I18N.en?.icons?.[definition.value || ""] || "",
      I18N.de?.icons?.[definition.value || ""] || "",
      I18N.es?.icons?.[definition.value || ""] || "",
      ...(Array.isArray(definition.aliases) ? definition.aliases : []),
    ].join(" ").toLowerCase();
    return `<button type="button" data-icon="${icon}" data-search="${searchTokens}" class="w-full px-2 py-1.5 text-left hover:bg-slate-100 flex items-center gap-2">${iconHtml}<span>${label}</span></button>`;
  }).join("");

  const applySearchFilter = () => {
    const filter = (search.value || "").trim().toLowerCase();
    list.querySelectorAll("button[data-icon]").forEach((item) => {
      const rowText = ((item.getAttribute("data-search") || "") + " " + (item.textContent || "")).toLowerCase();
      item.classList.toggle("hidden", Boolean(filter) && !rowText.includes(filter));
    });
  };

  search.oninput = () => {
    applySearchFilter();
  };
  search.placeholder = text("iconSearchPlaceholder");

  button.onclick = () => {
    panel.classList.toggle("hidden");
    if (!panel.classList.contains("hidden")) {
      search.value = "";
      applySearchFilter();
      search.focus();
    }
    if (window.lucide?.createIcons) window.lucide.createIcons();
  };

  list.querySelectorAll("button[data-icon]").forEach((item) => {
    item.onclick = () => {
      const selectedCell = document.querySelector(".cell-location td.is-selected, .cell-icon td.is-selected");
      const previousTop = selectedCell ? selectedCell.getBoundingClientRect().top : null;
      module.img = item.getAttribute("data-icon") || "";
      panel.classList.add("hidden");
      renderAllPreserveViewport(previousTop);
    };
  });
}

function serializeRowsAsJsonData(rows) {
  const clone = rows.map((row) => ({
    ...(Number(row.mpr) !== state.modulesPerRow ? { mpr: getRowModulesPerRow(row) } : {}),
    modules: (row.modules || []).map((module) => {
      const payload = {};
      const icon = String(module.img || "");
      const location = String(module.loc || "");
      const color = String(module.col || "");
      const moduleSize = Number(module.mod || 1);
      if (icon) payload.img = icon;
      if (location) payload.loc = location;
      if (moduleSize !== 1) payload.mod = moduleSize;
      if (color) payload.col = color;
      return payload;
    }),
  }));
  return JSON.stringify({
    project: state.projectTitle || text("project"),
    paper: state.paperFormat,
    ...(getModuleWidthMm() !== 18 ? { mw: getModuleWidthMm() } : {}),
    ...(getLabelHeightMm() !== 30 ? { lh: getLabelHeightMm() } : {}),
    mpr: state.modulesPerRow,
    rows: clone,
  });
}

function deserializeJsonData(data) {
  const parsed = JSON.parse(data);
  if (typeof parsed.project === "string" && parsed.project.trim()) {
    state.projectTitle = parsed.project.trim();
  }
  if (getPaperFormatValues().includes(parsed.paper)) {
    state.paperFormat = parsed.paper;
  }
  if (parsed.mw !== undefined) {
    const parsedWidth = Number(parsed.mw);
    if (!Number.isNaN(parsedWidth) && parsedWidth >= 10 && parsedWidth <= 30) {
      state.moduleWidthMm = parsedWidth;
    }
  }
  if (parsed.lh !== undefined) {
    const parsedHeight = Number(parsed.lh);
    if (!Number.isNaN(parsedHeight) && parsedHeight >= 20 && parsedHeight <= 80) {
      state.labelHeightMm = parsedHeight;
    }
  }
  const decodedRows = parsed.rows;
  if (!Array.isArray(decodedRows)) {
    throw new Error(text("invalidJsonSchema"));
  }
  if (parsed.mpr !== undefined) {
    const nextMpr = Number(parsed.mpr);
    if (!Number.isNaN(nextMpr) && nextMpr >= 1 && nextMpr <= 24) {
      state.modulesPerRow = nextMpr;
    }
  }
  for (const row of decodedRows) {
    if (row.mpr !== undefined) {
      const rowMpr = Number(row.mpr);
      row.mpr = !Number.isNaN(rowMpr) && rowMpr >= 1 && rowMpr <= 24 ? rowMpr : undefined;
    }
    for (const module of row.modules) {
      module.img = module.img || "";
      module.loc = module.loc || "";
      module.mod = module.mod || 1;
      module.col = module.col || "";
    }
  }
  normalizeRows(decodedRows);
  return decodedRows;
}

function toBase64Url(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function fromBase64Url(value) {
  const normalized = value.replaceAll("-", "+").replaceAll("_", "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function gzipToBase64Url(text) {
  if (typeof CompressionStream === "undefined") return null;
  const input = new TextEncoder().encode(text);
  const stream = new Blob([input]).stream().pipeThrough(new CompressionStream("gzip"));
  const compressedBuffer = await new Response(stream).arrayBuffer();
  return toBase64Url(new Uint8Array(compressedBuffer));
}

async function gunzipFromBase64Url(payload) {
  if (typeof DecompressionStream === "undefined") throw new Error("DecompressionStream unavailable");
  const compressedBytes = fromBase64Url(payload);
  const stream = new Blob([compressedBytes]).stream().pipeThrough(new DecompressionStream("gzip"));
  const decompressedBuffer = await new Response(stream).arrayBuffer();
  return new TextDecoder().decode(decompressedBuffer);
}

async function buildPermalink() {
  const compactJson = serializeRowsAsJsonData(state.rows);
  const compressedPayload = await gzipToBase64Url(compactJson);
  const widthParam = getModuleWidthMm() !== 18 ? `&mw=${getModuleWidthMm()}` : "";
  const heightParam = getLabelHeightMm() !== 30 ? `&lh=${getLabelHeightMm()}` : "";
  if (compressedPayload) {
    return `${window.location.origin}${window.location.pathname}?json&enc=gz-b64&paper=${state.paperFormat}${widthParam}${heightParam}&mode=${state.text90Mode ? "text90" : "standard"}&data=${compressedPayload}`;
  }
  const legacyData = encodeURIComponent(compactJson);
  return `${window.location.origin}${window.location.pathname}?json&paper=${state.paperFormat}${widthParam}${heightParam}&mode=${state.text90Mode ? "text90" : "standard"}&data=${legacyData}`;
}

async function refreshPermalink() {
  state.permalinkUrl = await buildPermalink();
  if (state.permalinkUrl) {
    history.replaceState(null, "", state.permalinkUrl);
  }
}

function renderTopBarText() {
  document.documentElement.lang = state.language;
  document.title = text("pageTitle");
  const seoDescription = text("seoDescription");
  const metaDescription = document.getElementById("meta-description");
  if (metaDescription) metaDescription.setAttribute("content", seoDescription);
  const metaOgTitle = document.getElementById("meta-og-title");
  if (metaOgTitle) metaOgTitle.setAttribute("content", text("pageTitle"));
  const metaOgDescription = document.getElementById("meta-og-description");
  if (metaOgDescription) metaOgDescription.setAttribute("content", seoDescription);
  const metaOgSiteName = document.getElementById("meta-og-site-name");
  if (metaOgSiteName) metaOgSiteName.setAttribute("content", text("appTitle"));
  const metaTwitterTitle = document.getElementById("meta-twitter-title");
  if (metaTwitterTitle) metaTwitterTitle.setAttribute("content", text("pageTitle"));
  const metaTwitterDescription = document.getElementById("meta-twitter-description");
  if (metaTwitterDescription) metaTwitterDescription.setAttribute("content", seoDescription);
  applyThemePreference();
  applyModuleWidthCss();
  applyLabelHeightCss();
  const displayProjectTitle = String(state.projectTitle || "").trim() || text("project");
  document.getElementById("app-title").textContent = displayProjectTitle;
  document.getElementById("app-title").setAttribute("title", text("appTitle"));
  const helpTitle = document.getElementById("help-title");
  if (helpTitle) helpTitle.textContent = text("helpTitle");
  const helpIntro = document.getElementById("help-intro");
  if (helpIntro) helpIntro.textContent = text("helpIntro");
  const projectGroupTitle = document.getElementById("project-group-title");
  if (projectGroupTitle) projectGroupTitle.textContent = text("project");
  const toolbarSummary = document.getElementById("toolbar-summary");
  if (toolbarSummary) toolbarSummary.textContent = text("settingsRibbon");
  const dimensionsGroupTitle = document.getElementById("dimensions-group-title");
  if (dimensionsGroupTitle) dimensionsGroupTitle.textContent = text("labelHeight");
  const printGroupTitle = document.getElementById("print-group-title");
  if (printGroupTitle) printGroupTitle.textContent = text("print");
  document.getElementById("project-label").textContent = text("project");
  const themeButton = document.getElementById("theme-toggle-btn");
  if (themeButton) {
    const isDark = state.themeMode === "dark";
    const nextModeLabel = isDark ? text("themeLight") : text("themeDark");
    themeButton.innerHTML = `<i data-lucide="${isDark ? "sun" : "moon"}" class="w-4 h-4"></i>`;
    themeButton.setAttribute("title", nextModeLabel);
    themeButton.setAttribute("aria-label", nextModeLabel);
  }
  const formatLabel = document.getElementById("format-label");
  if (formatLabel) formatLabel.textContent = text("labelHeight");
  const pageFormatLabel = document.getElementById("page-format-label");
  if (pageFormatLabel) pageFormatLabel.textContent = text("pageFormat");
  const pageFormatSelect = document.getElementById("page-format-select");
  getAvailablePaperFormats().forEach((paperFormat) => {
    const option = pageFormatSelect.querySelector(`option[value="${paperFormat.id}"]`);
    if (option) option.textContent = formatPaperOptionLabel(paperFormat);
  });
  const rowSizeLabel = document.getElementById("row-size-label");
  if (rowSizeLabel) rowSizeLabel.textContent = text("rowSize");
  const moduleWidthLabel = document.getElementById("module-width-label");
  if (moduleWidthLabel) moduleWidthLabel.textContent = text("moduleWidth");
  const text90Label = document.getElementById("text-90-label");
  if (text90Label) text90Label.textContent = text("text90Mode");
  const buttonSpecs = [
    { id: "copy-link-btn", icon: "link-2", label: text("copyLink") },
    { id: "export-json-btn", icon: "save", label: text("save") },
    { id: "import-json-btn", icon: "folder-open", label: text("open") },
    { id: "print-btn", icon: "printer", label: text("print") },
    { id: "add-row-bottom-btn", icon: "plus", label: text("newRowBottom") },
  ];
  buttonSpecs.forEach(({ id, icon, label }) => {
    const button = document.getElementById(id);
    if (!button) return;
    const iconOnly = id === "copy-link-btn" || id === "import-json-btn" || id === "export-json-btn" || id === "print-btn";
    button.innerHTML = iconOnly
      ? `<i data-lucide="${icon}" class="w-4 h-4 shrink-0"></i>`
      : `<i data-lucide="${icon}" class="w-4 h-4 shrink-0"></i><span>${label}</span>`;
    button.setAttribute("title", label);
    button.setAttribute("aria-label", label);
  });
  const projectInput = document.getElementById("project-input");
  if (projectInput) projectInput.value = state.projectTitle || "";
  if (pageFormatSelect) pageFormatSelect.value = state.paperFormat;
  const labelHeightInput = document.getElementById("label-height-input");
  if (labelHeightInput) labelHeightInput.value = String(getLabelHeightMm());
  const rowSizeInput = document.getElementById("row-size-input");
  if (rowSizeInput) rowSizeInput.value = String(state.modulesPerRow);
  const languageSelect = document.getElementById("language-select");
  if (languageSelect) languageSelect.value = state.language;
  const moduleWidthInput = document.getElementById("module-width-input");
  if (moduleWidthInput) moduleWidthInput.value = String(getModuleWidthMm());
  const text90Toggle = document.getElementById("text-90-toggle");
  if (text90Toggle) text90Toggle.checked = state.text90Mode;
}

function renderAll() {
  renderTopBarText();
  renderLabels();
  renderEditor();
  if (window.lucide?.createIcons) window.lucide.createIcons();
  refreshPermalink();
}

function bindEvents() {
  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const panel = document.getElementById("icon-picker-panel");
    if (!panel) return;
    if (target.closest("#icon-picker-button") || target.closest("#icon-picker-panel")) return;
    panel.classList.add("hidden");
  });

  document.getElementById("theme-toggle-btn").onclick = () => {
    state.themeMode = state.themeMode === "dark" ? "light" : "dark";
    localStorage.setItem("panel-theme-mode", state.themeMode);
    applyThemePreference();
    renderTopBarText();
    if (window.lucide?.createIcons) window.lucide.createIcons();
  };

  document.getElementById("label-height-input").onchange = (e) => {
    const nextHeight = Number(e.target.value);
    if (Number.isNaN(nextHeight) || nextHeight < 20 || nextHeight > 80) {
      e.target.value = String(getLabelHeightMm());
      return;
    }
    state.labelHeightMm = nextHeight;
    renderAll();
  };
  document.getElementById("page-format-select").onchange = (e) => {
    const nextFormat = e.target.value;
    if (getPaperFormatValues().includes(nextFormat)) {
      state.paperFormat = nextFormat;
      renderAll();
    }
  };
  document.getElementById("project-input").oninput = (e) => {
    state.projectTitle = String(e.target.value || "").trim();
    renderTopBarText();
    refreshPermalink();
  };
  document.getElementById("row-size-input").onchange = (e) => {
    const nextSize = Number(e.target.value);
    if (Number.isNaN(nextSize) || nextSize < 1 || nextSize > 24) {
      e.target.value = String(state.modulesPerRow);
      return;
    }
    state.modulesPerRow = nextSize;
    normalizeRows(state.rows);
    renderAll();
  };
  const languageSelect = document.getElementById("language-select");
  if (languageSelect) {
    languageSelect.onchange = (event) => {
      const nextLanguage = String(event.target.value || "");
      if (!SUPPORTED_LANGUAGES.includes(nextLanguage)) return;
      state.language = nextLanguage;
      localStorage.setItem("panel-language", state.language);
      renderAll();
    };
  }
  document.getElementById("module-width-input").onchange = (e) => {
    const nextWidth = Number(e.target.value);
    if (Number.isNaN(nextWidth) || nextWidth < 10 || nextWidth > 30) {
      e.target.value = String(getModuleWidthMm());
      return;
    }
    state.moduleWidthMm = nextWidth;
    renderAll();
  };
  document.getElementById("text-90-toggle").onchange = (e) => {
    state.text90Mode = Boolean(e.target.checked);
    renderAll();
  };
  const addRowBottomBtn = document.getElementById("add-row-bottom-btn");
  if (addRowBottomBtn) {
    addRowBottomBtn.onclick = () => {
      const scrollX = window.scrollX;
      const scrollY = window.scrollY;
      state.rows.push({ modules: [] });
      normalizeRows(state.rows);
      state.selectedRowIndex = state.rows.length - 1;
      state.selectedModuleIndex = 0;
      renderAll();
      window.scrollTo(scrollX, scrollY);
    };
  }
  document.getElementById("copy-link-btn").onclick = async () => {
    await refreshPermalink();
    const ok = await copyTextToClipboard(state.permalinkUrl || window.location.href);
    if (!ok) {
      window.alert(text("copyFailed"));
    }
  };
  document.getElementById("export-json-btn").onclick = () => {
    const payload = {
      version: 1,
      project: state.projectTitle || text("project"),
      paper: state.paperFormat,
      modulesPerRow: state.modulesPerRow,
      rows: state.rows.map((row) => ({
        ...(Number(row.mpr) !== state.modulesPerRow ? { modulesPerRow: getRowModulesPerRow(row) } : {}),
        modules: (row.modules || []).map((module) => {
          const out = {};
          const icon = String(module.img || "");
          const location = String(module.loc || "");
          const color = String(module.col || "");
          const moduleSize = Number(module.mod || 1);
          if (icon) out.img = icon;
          if (location) out.loc = location;
          if (moduleSize !== 1) out.mod = moduleSize;
          if (color) out.col = color;
          return out;
        }),
      })),
    };
    if (getModuleWidthMm() !== 18) payload.moduleWidthMm = getModuleWidthMm();
    if (getLabelHeightMm() !== 30) payload.labelHeightMm = getLabelHeightMm();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `panel-labels-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };
  document.getElementById("import-json-btn").onclick = () => document.getElementById("import-file-input").click();
  document.getElementById("import-file-input").onchange = async (event) => {
    try {
      const file = event.target.files?.[0];
      if (!file) return;
      const parsed = JSON.parse(await file.text());
      if (typeof parsed.project === "string" && parsed.project.trim()) {
        state.projectTitle = parsed.project.trim();
      }
      if (getPaperFormatValues().includes(parsed.paper)) {
        state.paperFormat = parsed.paper;
      }
      if (parsed.moduleWidthMm !== undefined) {
        const parsedWidth = Number(parsed.moduleWidthMm);
        if (!Number.isNaN(parsedWidth) && parsedWidth >= 10 && parsedWidth <= 30) {
          state.moduleWidthMm = parsedWidth;
        }
      }
      if (parsed.labelHeightMm !== undefined) {
        const parsedHeight = Number(parsed.labelHeightMm);
        if (!Number.isNaN(parsedHeight) && parsedHeight >= 20 && parsedHeight <= 80) {
          state.labelHeightMm = parsedHeight;
        }
      } else if (parsed.format) {
        // Legacy support
        state.labelHeightMm = parsed.format === "5cm" ? 50 : 30;
      }
      if (parsed.modulesPerRow !== undefined) {
        const nextMpr = Number(parsed.modulesPerRow);
        if (!Number.isNaN(nextMpr) && nextMpr >= 1 && nextMpr <= 24) {
          state.modulesPerRow = nextMpr;
        }
      }
      if (!Array.isArray(parsed.rows)) {
        throw new Error(text("invalidJsonSchema"));
      }
      state.rows = parsed.rows;
      state.rows.forEach((row) => {
        if (row.modulesPerRow !== undefined) {
          const nextRowMpr = Number(row.modulesPerRow);
          row.mpr = !Number.isNaN(nextRowMpr) && nextRowMpr >= 1 && nextRowMpr <= 24 ? nextRowMpr : undefined;
          delete row.modulesPerRow;
        }
      });
      normalizeRows(state.rows);
      renderAll();
    } catch (error) {
      window.alert(`${text("importError")}: ${error.message}`);
    }
  };
  window.addEventListener("afterprint", () => {
    document.body.classList.remove("print-portrait");
    document.body.classList.remove("print-a4", "print-a3", "print-letter", "print-tabloid");
  });

  document.getElementById("print-btn").onclick = () => {
    const printPlan = buildPrintPlan();
    document.body.classList.remove("print-a4", "print-a3", "print-letter", "print-tabloid");
    document.body.classList.add(`print-${state.paperFormat}`);
    document.body.classList.toggle("print-portrait", printPlan.orientation === "portrait");
    window.print();
  };
}

async function bootstrap() {
  const savedLanguage = localStorage.getItem("panel-language");
  state.language = SUPPORTED_LANGUAGES.includes(savedLanguage)
    ? savedLanguage
    : detectLanguageFromBrowser();
  await loadTranslations();
  await loadPaperFormats();
  await loadIconDefinitions();
  if (PAPER_FORMATS.length === 0 || ICON_DEFINITIONS.length === 0) {
    window.alert(text("configLoadError"));
    return;
  }
  const savedThemeMode = localStorage.getItem("panel-theme-mode");
  if (savedThemeMode === "light" || savedThemeMode === "dark") {
    state.themeMode = savedThemeMode;
  } else {
    state.themeMode = window.matchMedia?.("(prefers-color-scheme: dark)")?.matches ? "dark" : "light";
  }
  applyThemePreference();
  const params = new URLSearchParams(window.location.search);
  const defaultParams = new URLSearchParams(DEFAULT_PERMALINK_QUERY);
  const hasCustomData = Boolean(params.get("data"));
  const effectiveParams = hasCustomData ? params : defaultParams;
  const paper = effectiveParams.get("paper");
  const moduleWidthMm = effectiveParams.get("mw");
  const labelHeightMm = effectiveParams.get("lh");
  const mode = effectiveParams.get("mode");
  const data = effectiveParams.get("data");
  const encoding = effectiveParams.get("enc");
  if (getPaperFormatValues().includes(paper)) state.paperFormat = paper;
  if (moduleWidthMm !== null) {
    const parsedWidth = Number(moduleWidthMm);
    if (!Number.isNaN(parsedWidth) && parsedWidth >= 10 && parsedWidth <= 30) {
      state.moduleWidthMm = parsedWidth;
    }
  }
  if (labelHeightMm !== null) {
    const parsedHeight = Number(labelHeightMm);
    if (!Number.isNaN(parsedHeight) && parsedHeight >= 20 && parsedHeight <= 80) {
      state.labelHeightMm = parsedHeight;
    }
  }
  if (mode === "text90") state.text90Mode = true;
  if (data) {
    try {
      if (encoding === "gz-b64") {
        const jsonData = await gunzipFromBase64Url(data);
        state.rows = deserializeJsonData(jsonData);
      } else {
        state.rows = deserializeJsonData(decodeURIComponent(data));
      }
    } catch {
      // Backward compatibility / graceful fallback to legacy format.
      state.rows = deserializeJsonData(decodeURIComponent(data));
    }
  }
  normalizeRows(state.rows);
  const pageFormatSelect = document.getElementById("page-format-select");
  if (pageFormatSelect) pageFormatSelect.value = state.paperFormat;
  bindEvents();
  await refreshPermalink();
  renderAll();
}

document.addEventListener("DOMContentLoaded", () => {
  bootstrap();
});
