// Tableau Photo Gallery (Chunked + Pagination)
// Requirements implemented:
// - Chunk size fixed to 1000
// - Page size fixed to 100
// - Paginator right after message box shows loaded/shown/page
// - Clear Cache does a hard reset + reload so images always come back
// - Viewer is "fit-only" (CSS object-fit: contain). No zoom/pan.

"use strict";

const CHUNK_SIZE = 1000;
const PAGE_SIZE = 100;

if (!window.tableau || !tableau.extensions) {
  const empty = document.getElementById("empty");
  empty.style.display = "grid";
  empty.innerHTML = `
    <div style="max-width:720px">
      <div style="font-weight:800; font-size:14px; margin-bottom:8px;">Tableau Extensions API not loaded</div>
      <div style="font-size:12px; color:#475569;">
        Fix: host <code>tableau.extensions.1.latest.min.js</code> in this same repo and load it via
        <code>&lt;script src="./tableau.extensions.1.latest.min.js"&gt;&lt;/script&gt;</code>.
      </div>
    </div>`;
  throw new Error("Tableau Extensions API not loaded");
}

let dashboardRef = null;
let initPromise = null;

// ---------- DOM ----------
const emptyEl = document.getElementById("empty");
const gridEl = document.getElementById("grid");

const refreshBtn = document.getElementById("refreshBtn");
const clearBtn = document.getElementById("clearBtn");

const loadMoreBtn = document.getElementById("loadMoreBtn");
const prevPageBtn = document.getElementById("prevPageBtn");
const nextPageBtn = document.getElementById("nextPageBtn");

const sheetSelect = document.getElementById("sheetSelect");
const fieldSelect = document.getElementById("fieldSelect");
const searchInput = document.getElementById("searchInput");

const countPill = document.getElementById("countPill");
const pagePill = document.getElementById("pagePill");

// Viewer
const overlayEl = document.getElementById("overlay");
const closeBtn = document.getElementById("closeBtn");
const vPrevBtn = document.getElementById("vPrevBtn");
const vNextBtn = document.getElementById("vNextBtn");
const viewerTitleEl = document.getElementById("viewerTitle");
const viewerCountEl = document.getElementById("viewerCount");
const openTabBtn = document.getElementById("openTabBtn");
const fullImgEl = document.getElementById("fullImg");

// ---------- State ----------
let allUrls = [];
let filteredUrls = [];
let currentPage = 1;
let viewerIndex = 0;
let loadedChunks = 0;
let lastCacheKey = null;
let loading = false;

// ---------- Helpers ----------
function setStatus(msg) {
  console.log("[status]", msg);
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function safeHost(url) {
  try {
    return new URL(url).host;
  } catch {
    return "";
  }
}

function normalizeUrl(raw) {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s) return null;
  const lower = s.toLowerCase();
  if (lower === "null" || lower === "none") return null;
  if (!/^https?:\/\//i.test(s)) return null;
  return s;
}

function showError(html) {
  emptyEl.style.display = "grid";
  emptyEl.innerHTML = html;
}

function hideError() {
  emptyEl.style.display = "none";
  emptyEl.innerHTML = "";
}

function getChosenSheet() {
  return sheetSelect.value || null;
}

function getChosenField() {
  return fieldSelect.value || null;
}

function getTotalPages() {
  return Math.max(1, Math.ceil(filteredUrls.length / PAGE_SIZE));
}

function clampPage(p) {
  return Math.max(1, Math.min(getTotalPages(), p));
}

function getPageSlice() {
  const start = (currentPage - 1) * PAGE_SIZE;
  return filteredUrls.slice(start, start + PAGE_SIZE);
}

// ---------- Tableau init ----------
async function ensureInitialized() {
  if (dashboardRef) return dashboardRef;

  if (!initPromise) {
    initPromise = (async () => {
      await tableau.extensions.initializeAsync();
      dashboardRef = tableau.extensions.dashboardContent.dashboard;
      return dashboardRef;
    })();
  }

  return initPromise;
}

// ---------- UI wiring ----------
function fillSheetSelect() {
  const sheets = dashboardRef?.worksheets || [];
  const current = sheetSelect.value;

  sheetSelect.innerHTML = "";

  sheets.forEach((ws) => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name;
    sheetSelect.appendChild(opt);
  });

  if (current && sheets.some((s) => s.name === current)) sheetSelect.value = current;
  else if (sheets.length) sheetSelect.value = sheets[0].name;
}

function buildFieldOptions(columns) {
  const current = fieldSelect.value;

  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;

  columns.forEach((c) => {
    const label = c.fieldName || c.caption || "";
    if (!label) return;
    const opt = document.createElement("option");
    opt.value = label;
    opt.textContent = label;
    fieldSelect.appendChild(opt);
  });

  // keep selection if still valid
  if (current) fieldSelect.value = current;
}

function findCandidateColumnIndexes(columns, chosenField) {
  if (chosenField) {
    const idx = columns.findIndex(
      (c) => c.fieldName === chosenField || c.caption === chosenField
    );
    return idx >= 0 ? [idx] : [];
  }

  // Priority 1: exact match on "Image URL"
  const exactIdx = columns.findIndex((c) => {
    const name = (c.fieldName || c.caption || "").trim().toLowerCase();
    return name === "image url";
  });
  if (exactIdx >= 0) return [exactIdx];

  // Priority 2: heuristic — column name contains a URL-related keyword
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
    if (hints.some((h) => name.includes(h))) idxs.push(i);
  });

  return idxs.length ? idxs : columns.map((_, i) => i);
}

// Find a single meta column (Location Name, Visit Date, etc.)
function findMetaColumnIndex(columns, exactName, keywords) {
  const exactIdx = columns.findIndex((c) => {
    const name = (c.fieldName || c.caption || "").trim().toLowerCase();
    return name === exactName.toLowerCase();
  });
  if (exactIdx >= 0) return exactIdx;

  for (const kw of keywords) {
    const idx = columns.findIndex((c) => {
      const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
      return name.includes(kw.toLowerCase());
    });
    if (idx >= 0) return idx;
  }
  return -1;
}

function buildLabel(location, date) {
  const loc = String(location || "").trim();
  const dt  = String(date   || "").trim();
  if (loc && dt) return `${loc} (${dt})`;
  return loc || dt || "";
}

function applySearch() {
  const q = (searchInput.value || "").trim().toLowerCase();
  filteredUrls = q
    ? allUrls.filter(({ url, label }) =>
        url.toLowerCase().includes(q) || label.toLowerCase().includes(q))
    : allUrls.slice();
  currentPage = 1;
  renderGrid();
}

function renderGrid() {
  const slice = getPageSlice();
  gridEl.innerHTML = "";

  if (!filteredUrls.length) {
    showError(`
      <div style="max-width:720px">
        <div style="font-weight:800; font-size:14px; margin-bottom:8px;">No images found</div>
        <div style="font-size:12px; color:#475569;">
          Try selecting a different URL field, or click <b>Load more</b>.
        </div>
      </div>
    `);
  } else {
    hideError();
  }

  slice.forEach((item, i) => {
    const { url, label } = item;
    const globalIndex = (currentPage - 1) * PAGE_SIZE + i;

    const card = document.createElement("div");
    card.className = "card";
    card.title = label || url;

    const thumb = document.createElement("div");
    thumb.className = "thumb";

    const img = document.createElement("img");
    img.loading = "lazy";
    img.src = url;

    img.onerror = () => {
      img.remove();
      thumb.style.background = "#1e2229";

      const d = document.createElement("div");
      d.style.padding = "12px";
      d.style.fontSize = "12px";
      d.style.color = "#8891a2";
      d.textContent = "Image failed to load";
      thumb.appendChild(d);
    };

    const badge = document.createElement("div");
    badge.className = "badge";
    badge.textContent = `#${globalIndex + 1}`;

    const caption = document.createElement("div");
    caption.className = "caption";

    const text = document.createElement("div");
    text.className = "text";
    text.textContent = label || safeHost(url);

    const icon = document.createElement("button");
    icon.className = "iconbtn";
    icon.type = "button";
    icon.title = "Open in new tab";
    icon.textContent = "↗";
    icon.addEventListener("click", (e) => {
      e.stopPropagation();
      window.open(url, "_blank", "noopener,noreferrer");
    });

    caption.appendChild(text);
    caption.appendChild(icon);

    thumb.appendChild(img);
    thumb.appendChild(badge);
    thumb.appendChild(caption);

    card.appendChild(thumb);
    card.addEventListener("click", () => openViewer(globalIndex));
    gridEl.appendChild(card);
  });

  // paginator pills
  countPill.textContent = `${allUrls.length.toLocaleString()} images`;
  pagePill.textContent = `Page ${currentPage} / ${getTotalPages()}`;

  prevPageBtn.disabled = currentPage <= 1;
  nextPageBtn.disabled = currentPage >= getTotalPages();
}

// ---------- Viewer (fit-only) ----------
function openViewer(globalIndex) {
  if (!filteredUrls.length) return;

  viewerIndex = Math.max(0, Math.min(filteredUrls.length - 1, globalIndex));
  overlayEl.classList.add("open");
  overlayEl.setAttribute("aria-hidden", "false");
  showViewerIndex(viewerIndex);
}

function closeViewer() {
  overlayEl.classList.remove("open");
  overlayEl.setAttribute("aria-hidden", "true");
  fullImgEl.src = "";
}

function showViewerIndex(idx) {
  if (!filteredUrls.length) return;

  viewerIndex = (idx + filteredUrls.length) % filteredUrls.length;
  const { url, label } = filteredUrls[viewerIndex];

  viewerTitleEl.textContent = label || safeHost(url) || "Image";
  viewerCountEl.textContent = `${viewerIndex + 1} / ${filteredUrls.length}`;

  openTabBtn.onclick = () => window.open(url, "_blank", "noopener,noreferrer");
  fullImgEl.src = url; // fit is handled by CSS object-fit: contain
}

// ---------- Data fetch (row-level preferred over summary) ----------
async function getDataTable(ws, maxRows) {
  try {
    const tables = await ws.getUnderlyingTablesAsync();
    if (tables && tables.length > 0) {
      const tbl = await ws.getUnderlyingTableDataAsync(tables[0].id, {
        maxRows,
        ignoreAliases: false,
        ignoreSelection: true,
        includeAllColumns: true,
      });
      tbl._sourceMode = "row-level";
      return tbl;
    }
  } catch (_) {
    // underlying data not available — fall through
  }
  const tbl = await ws.getSummaryDataAsync({ maxRows });
  tbl._sourceMode = "summary";
  return tbl;
}

// ---------- Data load (chunked) ----------
async function loadChunk({ reset }) {
  if (loading) return;
  loading = true;

  try {
    await ensureInitialized();

    const sheetName = getChosenSheet();
    const chosenField = getChosenField();

    const cacheKey = `${sheetName}::${chosenField || "auto"}`;

    // reset conditions: explicit reset or when user changes sheet/field
    if (reset || cacheKey !== lastCacheKey) {
      lastCacheKey = cacheKey;
      loadedChunks = 0;
      allUrls = [];
      filteredUrls = [];
      currentPage = 1;
      loadMoreBtn.disabled = false;
      loadMoreBtn.textContent = "Load more";
      renderGrid();
    }

    const ws = dashboardRef.worksheets.find((w) => w.name === sheetName);
    if (!ws) {
      showError(`
        <div style="max-width:720px">
          <div style="font-weight:800; font-size:14px; margin-bottom:8px;">Worksheet not found</div>
          <div style="font-size:12px; color:#475569;">${escapeHtml(sheetName || "")}</div>
        </div>
      `);
      setStatus("Worksheet not found");
      return;
    }

    // Each chunk re-requests data up to maxRows (simple + robust approach in Extensions API)
    const maxRows = (loadedChunks + 1) * CHUNK_SIZE;
    setStatus(`Loading up to ${maxRows.toLocaleString()} rows…`);

    const summary = await getDataTable(ws, maxRows);
    const columns = summary.columns || [];
    const rows = summary.data || [];

    buildFieldOptions(columns);

    const idxs = findCandidateColumnIndexes(columns, chosenField);

    // Find location name and visit date columns for the overlay label
    const locationIdx = findMetaColumnIndex(columns, "Location Name", ["location name", "location", "store name", "store"]);
    const dateIdx     = findMetaColumnIndex(columns, "Visit Date",     ["visit date", "date"]);

    const extracted = [];
    for (const row of rows) {
      for (const colIdx of idxs) {
        const cell = row[colIdx];
        const raw  = cell?.value ?? cell?.formattedValue ?? null;
        const url  = normalizeUrl(raw);
        if (!url) continue;

        const location = String(locationIdx >= 0 ? (row[locationIdx]?.formattedValue ?? row[locationIdx]?.value ?? "") : "").trim();
        const date     = String(dateIdx     >= 0 ? (row[dateIdx]?.formattedValue     ?? row[dateIdx]?.value     ?? "") : "").trim();
        extracted.push({ url, label: buildLabel(location, date), location, date });
      }
    }

    // De-dupe by URL string
    const seen = new Set(allUrls.map((item) => item.url));
    for (const item of extracted) {
      if (seen.has(item.url)) continue;
      seen.add(item.url);
      allUrls.push(item);
    }

    // Sort: Visit Date descending, then Location Name ascending
    allUrls.sort((a, b) => {
      const dateCmp = (b.date || "").localeCompare(a.date || "", undefined, { numeric: true, sensitivity: "base" });
      if (dateCmp !== 0) return dateCmp;
      return (a.location || "").localeCompare(b.location || "", undefined, { numeric: true, sensitivity: "base" });
    });

    // Always advance the chunk counter so the next load requests a larger row window
    loadedChunks += 1;

    const candidateNames = idxs.map(i => columns[i]?.fieldName || columns[i]?.caption || `col[${i}]`).join(", ");
    const mode = summary._sourceMode || "unknown";
    setStatus(`[${mode}] Rows: ${rows.length} | Cols: ${columns.length} | Scanning: [${candidateNames || "none"}]\nLoaded ${allUrls.length.toLocaleString()} unique URL(s).`);
    applySearch();

    // Only stop when the API returned fewer rows than requested — data exhausted
    if (rows.length < maxRows) {
      loadMoreBtn.disabled = true;
      loadMoreBtn.textContent = "No more in view";
    }
  } catch (e) {
    console.error(e);
    showError(`
      <div style="max-width:720px">
        <div style="font-weight:800; font-size:14px; margin-bottom:8px;">Error</div>
        <div style="font-size:12px; color:#475569;">${escapeHtml(e?.message || String(e))}</div>
        <div style="font-size:12px; color:#64748b; margin-top:8px;">
          If this mentions permissions, your <code>.trex</code> must include <code>&lt;permission&gt;full data&lt;/permission&gt;</code>.
        </div>
      </div>
    `);
    setStatus("Error");
  } finally {
    loading = false;
  }
}

// ---------- Events ----------
searchInput.addEventListener("input", applySearch);

refreshBtn.addEventListener("click", () => loadChunk({ reset: true }));

clearBtn.addEventListener("click", async () => {
  // Hard reset
  lastCacheKey = null;
  loadedChunks = 0;
  allUrls = [];
  filteredUrls = [];
  currentPage = 1;

  // Reset UI bits
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  setStatus("Cache cleared. Reloading…");
  renderGrid();

  // IMPORTANT: immediately reload so user isn't stuck with an empty state
  await loadChunk({ reset: true });
});

loadMoreBtn.addEventListener("click", () => loadChunk({ reset: false }));

prevPageBtn.addEventListener("click", () => {
  currentPage = clampPage(currentPage - 1);
  renderGrid();
});

nextPageBtn.addEventListener("click", () => {
  currentPage = clampPage(currentPage + 1);
  renderGrid();
});

sheetSelect.addEventListener("change", () => {
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  loadChunk({ reset: true });
});

fieldSelect.addEventListener("change", () => loadChunk({ reset: true }));

// Viewer controls
closeBtn.addEventListener("click", closeViewer);
vPrevBtn.addEventListener("click", () => showViewerIndex(viewerIndex - 1));
vNextBtn.addEventListener("click", () => showViewerIndex(viewerIndex + 1));

window.addEventListener("keydown", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  if (e.key === "Escape") closeViewer();
  if (e.key === "ArrowLeft") showViewerIndex(viewerIndex - 1);
  if (e.key === "ArrowRight") showViewerIndex(viewerIndex + 1);
});

// ---------- Init ----------
async function init() {
  await ensureInitialized();
  fillSheetSelect();
  setStatus("Ready");
  await loadChunk({ reset: true });
}

init().catch((e) => {
  console.error("Init failed:", e);
  showError(`
    <div style="max-width:720px">
      <div style="font-weight:800; font-size:14px; margin-bottom:8px;">Init error</div>
      <div style="font-size:12px; color:#475569;">${escapeHtml(e?.message || String(e))}</div>
    </div>
  `);
  setStatus("Init error");
});