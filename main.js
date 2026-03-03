// Tableau Photo Gallery (Chunked + Pagination) - Viewer fits to widget, no zoom/pan.

if (!window.tableau || !tableau.extensions) {
  const empty = document.getElementById("empty");
  empty.style.display = "grid";
  empty.innerHTML = `
    <div class="errorText">
      Tableau Extensions API not loaded (tableau is undefined).<br/><br/>
      Fix: host <b>tableau.extensions.1.latest.min.js</b> in the same GitHub Pages repo and load it via
      <code>&lt;script src="./tableau.extensions.1.latest.min.js"&gt;&lt;/script&gt;</code>.
    </div>
  `;
  throw new Error("Tableau Extensions API not loaded");
}

let dashboardRef = null;
let initPromise = null;

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

// ---------- DOM ----------
const statusEl = document.getElementById("status");
const emptyEl = document.getElementById("empty");
const gridWrapEl = document.getElementById("gridWrap");
const gridEl = document.getElementById("grid");

const refreshBtn = document.getElementById("refreshBtn");
const clearBtn = document.getElementById("clearBtn");
const loadMoreBtn = document.getElementById("loadMoreBtn");
const prevPageBtn = document.getElementById("prevPageBtn");
const nextPageBtn = document.getElementById("nextPageBtn");

const sheetSelect = document.getElementById("sheetSelect");
const fieldSelect = document.getElementById("fieldSelect");
const chunkSelect = document.getElementById("chunkSelect");
const pageSizeSelect = document.getElementById("pageSizeSelect");
const searchInput = document.getElementById("searchInput");

const countPill = document.getElementById("countPill");
const filteredPill = document.getElementById("filteredPill");
const chunkPill = document.getElementById("chunkPill");
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
function setStatus(msg) { statusEl.value = msg; }
function escapeHtml(str) {
  return String(str).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");
}
function safeHost(url) { try { return new URL(url).host; } catch { return ""; } }
function normalizeUrl(raw) {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s) return null;
  const lower = s.toLowerCase();
  if (lower === "null" || lower === "none") return null;
  if (!/^https?:\/\//i.test(s)) return null;
  return s;
}
function showError(html) { emptyEl.style.display = "grid"; emptyEl.innerHTML = html; }
function hideError() { emptyEl.style.display = "none"; emptyEl.innerHTML = ""; }

function getChunkSize() {
  const n = parseInt(chunkSelect.value, 10);
  return Number.isFinite(n) ? n : 5000;
}
function getPageSize() {
  const n = parseInt(pageSizeSelect.value, 10);
  return Number.isFinite(n) ? n : 200;
}
function getChosenSheet() { return sheetSelect.value || null; }
function getChosenField() { return fieldSelect.value || null; }

function getTotalPages() {
  const ps = getPageSize();
  return Math.max(1, Math.ceil(filteredUrls.length / ps));
}
function clampPage(p) {
  return Math.max(1, Math.min(getTotalPages(), p));
}
function getPageSlice() {
  const ps = getPageSize();
  const start = (currentPage - 1) * ps;
  return filteredUrls.slice(start, start + ps);
}

// ---------- UI ----------
function fillSheetSelect() {
  const sheets = dashboardRef?.worksheets || [];
  const current = sheetSelect.value;

  sheetSelect.innerHTML = "";
  sheets.forEach(ws => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name;
    sheetSelect.appendChild(opt);
  });

  if (current && sheets.some(s => s.name === current)) sheetSelect.value = current;
  else if (sheets.length) sheetSelect.value = sheets[0].name;
}

function buildFieldOptions(columns) {
  const current = fieldSelect.value;
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;

  columns.forEach(c => {
    const label = c.fieldName || c.caption || "";
    if (!label) return;
    const opt = document.createElement("option");
    opt.value = label;
    opt.textContent = label;
    fieldSelect.appendChild(opt);
  });

  if (current) fieldSelect.value = current;
}

function findCandidateColumnIndexes(columns, chosenField) {
  if (chosenField) {
    const idx = columns.findIndex(c => c.fieldName === chosenField || c.caption === chosenField);
    return idx >= 0 ? [idx] : [];
  }
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
    if (hints.some(h => name.includes(h))) idxs.push(i);
  });
  return idxs.length ? idxs : columns.map((_, i) => i);
}

function applySearch() {
  const q = (searchInput.value || "").trim().toLowerCase();
  filteredUrls = q ? allUrls.filter(u => u.toLowerCase().includes(q)) : allUrls.slice();
  currentPage = 1;
  renderGrid();
}

function renderGrid() {
  const slice = getPageSlice();
  gridEl.innerHTML = "";

  if (!filteredUrls.length) {
    showError(`
      <div>
        <div style="font-size:14px; margin-bottom:6px;">No images found.</div>
        <div style="font-size:12px; color:#777;">Try selecting a different URL field or Load more.</div>
      </div>
    `);
  } else {
    hideError();
  }

  slice.forEach((url, i) => {
    const globalIndex = (currentPage - 1) * getPageSize() + i;

    const card = document.createElement("div");
    card.className = "card";
    card.title = url;

    const thumb = document.createElement("div");
    thumb.className = "thumb";

    const img = document.createElement("img");
    img.loading = "lazy";
    img.src = url;
    img.onerror = () => {
      img.remove();
      thumb.style.background = "#dfe6ee";
      const d = document.createElement("div");
      d.style.padding = "12px";
      d.style.fontSize = "12px";
      d.style.color = "#445";
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
    text.textContent = safeHost(url);

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

  countPill.textContent = `${allUrls.length.toLocaleString()} images`;
  filteredPill.textContent = `${filteredUrls.length.toLocaleString()} shown`;
  chunkPill.textContent = `Loaded: ${(loadedChunks * getChunkSize()).toLocaleString()} (chunks: ${loadedChunks})`;
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

  const url = filteredUrls[viewerIndex];
  viewerTitleEl.textContent = safeHost(url);
  viewerCountEl.textContent = `${viewerIndex + 1} / ${filteredUrls.length}`;
  openTabBtn.onclick = () => window.open(url, "_blank", "noopener,noreferrer");

  // Fit is handled by CSS object-fit: contain
  fullImgEl.src = url;
}

// ---------- Data load (chunked) ----------
async function loadChunk({ reset }) {
  if (loading) return;
  loading = true;

  try {
    await ensureInitialized();

    const sheetName = getChosenSheet();
    const chosenField = getChosenField();
    const chunkSize = getChunkSize();

    const cacheKey = `${sheetName}::${chosenField || "auto"}`;
    if (reset || cacheKey !== lastCacheKey) {
      lastCacheKey = cacheKey;
      loadedChunks = 0;
      allUrls = [];
      filteredUrls = [];
      currentPage = 1;
      loadMoreBtn.disabled = false;
      loadMoreBtn.textContent = "Load more";
    }

    const ws = dashboardRef.worksheets.find(w => w.name === sheetName);
    if (!ws) {
      showError(`<div class="errorText">Worksheet not found: ${escapeHtml(sheetName)}</div>`);
      setStatus("Worksheet not found");
      return;
    }

    // Reads more rows each time (safe approach)
    const maxRows = (loadedChunks + 1) * chunkSize;
    setStatus(`Loading up to ${maxRows.toLocaleString()} rows…`);

    const summary = await ws.getSummaryDataAsync({ maxRows });
    const columns = summary.columns || [];
    const rows = summary.data || [];

    buildFieldOptions(columns);
    const idxs = findCandidateColumnIndexes(columns, chosenField);

    const extracted = [];
    for (const row of rows) {
      for (const colIdx of idxs) {
        const cell = row[colIdx];
        const raw = cell?.value ?? cell?.formattedValue ?? null;
        const url = normalizeUrl(raw);
        if (url) extracted.push(url);
      }
    }

    const before = allUrls.length;
    const seen = new Set(allUrls);
    for (const u of extracted) {
      if (seen.has(u)) continue;
      seen.add(u);
      allUrls.push(u);
    }

    if (allUrls.length > before) loadedChunks += 1;

    setStatus(`Loaded ${allUrls.length.toLocaleString()} unique URL(s).`);
    applySearch();

    if (allUrls.length === before) {
      loadMoreBtn.disabled = true;
      loadMoreBtn.textContent = "No more in view";
    }
  } catch (e) {
    console.error(e);
    showError(`
      <div class="errorText">
        Error: ${escapeHtml(e?.message || String(e))}
        <br/><br/>
        If this mentions permissions, your .trex must include
        <code>&lt;permission&gt;full data&lt;/permission&gt;</code>.
      </div>
    `);
    setStatus("Error");
  } finally {
    loading = false;
  }
}

// ---------- Events ----------
searchInput.addEventListener("input", applySearch);

refreshBtn.addEventListener("click", () => loadChunk({ reset:true }));
clearBtn.addEventListener("click", () => {
  lastCacheKey = null;
  loadedChunks = 0;
  allUrls = [];
  filteredUrls = [];
  currentPage = 1;
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  setStatus("Cache cleared");
  renderGrid();
});

loadMoreBtn.addEventListener("click", () => loadChunk({ reset:false }));

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
  loadChunk({ reset:true });
});
fieldSelect.addEventListener("change", () => loadChunk({ reset:true }));
chunkSelect.addEventListener("change", () => loadChunk({ reset:true }));
pageSizeSelect.addEventListener("change", () => { currentPage = 1; renderGrid(); });

// Viewer buttons
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
  await loadChunk({ reset:true });
}

init().catch((e) => {
  console.error("Init failed:", e);
  showError(`<div class="errorText">Init error: ${escapeHtml(e?.message || String(e))}</div>`);
  setStatus("Init error");
});