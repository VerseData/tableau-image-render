// =====================
// Tableau Photo Gallery (Chunked + Pagination + Optional Infinite Scroll)
// - Reads up to CHUNK_SIZE URLs at a time from a worksheet using getSummaryDataAsync(maxRows)
// - Maintains a growing list up to MAX_TOTAL_LOADED (multiple chunks)
// - UI paginates and optionally infinite-scrolls (page-by-page rendering)
// - Search filters the loaded URLs
// - Includes lightbox viewer (zoom/pan + next/prev)
// =====================

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

const sheetSelect = document.getElementById("sheetSelect");
const fieldSelect = document.getElementById("fieldSelect");
const chunkSelect = document.getElementById("chunkSelect");
const pageSizeSelect = document.getElementById("pageSizeSelect");
const searchInput = document.getElementById("searchInput");
const infiniteToggle = document.getElementById("infiniteToggle");

const countPill = document.getElementById("countPill");
const filteredPill = document.getElementById("filteredPill");
const chunkPill = document.getElementById("chunkPill");

const firstBtn = document.getElementById("firstBtn");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");
const lastBtn = document.getElementById("lastBtn");
const pageBox = document.getElementById("pageBox");

// Viewer
const overlayEl = document.getElementById("overlay");
const closeBtn = document.getElementById("closeBtn");
const vPrevBtn = document.getElementById("vPrevBtn");
const vNextBtn = document.getElementById("vNextBtn");
const viewerTitleEl = document.getElementById("viewerTitle");
const viewerCountEl = document.getElementById("viewerCount");
const openTabBtn = document.getElementById("openTabBtn");

const stageEl = document.getElementById("stage");
const viewportEl = document.getElementById("viewport");
const fullImgEl = document.getElementById("fullImg");
const zoomInBtn = document.getElementById("zoomInBtn");
const zoomOutBtn = document.getElementById("zoomOutBtn");
const fitBtn = document.getElementById("fitBtn");
const resetBtn = document.getElementById("resetBtn");

// ---------- State ----------
let allUrls = [];          // full loaded list (deduped)
let filteredUrls = [];     // after search
let currentPage = 1;
let viewerIndex = 0;

// "Chunking" tracking
let loadedChunks = 0;      // how many chunks we have loaded
let lastCacheKey = null;   // used to invalidate when sheet/field changes
let loading = false;

// Viewer pan/zoom
let scale = 1, offsetX = 0, offsetY = 0;
let dragging = false;
let dragStart = { x: 0, y: 0, ox: 0, oy: 0 };

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
function getChosenSheet() { return sheetSelect.value || null; }
function getChosenField() { return fieldSelect.value || null; }
function getChunkSize() {
  const n = parseInt(chunkSelect.value, 10);
  return Number.isFinite(n) ? n : 5000;
}
function getPageSize() {
  const n = parseInt(pageSizeSelect.value, 10);
  return Number.isFinite(n) ? n : 200;
}
function showError(html) {
  emptyEl.style.display = "grid";
  emptyEl.innerHTML = html;
}
function hideError() {
  emptyEl.style.display = "none";
  emptyEl.innerHTML = "";
}

// ---------- UI: Pagination ----------
function getTotalPages() {
  const pageSize = getPageSize();
  return Math.max(1, Math.ceil(filteredUrls.length / pageSize));
}
function clampPage(p) {
  const total = getTotalPages();
  return Math.max(1, Math.min(total, p));
}
function updatePager() {
  const total = getTotalPages();
  currentPage = clampPage(currentPage);

  pageBox.textContent = `${currentPage} / ${total}`;

  firstBtn.disabled = currentPage === 1;
  prevBtn.disabled = currentPage === 1;
  nextBtn.disabled = currentPage === total;
  lastBtn.disabled = currentPage === total;
}
function getPageSlice() {
  const pageSize = getPageSize();
  const start = (currentPage - 1) * pageSize;
  const end = start + pageSize;
  return filteredUrls.slice(start, end);
}

// ---------- UI: Render grid ----------
function renderGrid() {
  const slice = getPageSlice();
  gridEl.innerHTML = "";

  if (!filteredUrls.length) {
    showError(`
      <div>
        <div style="font-size:14px; margin-bottom:6px;">No images found.</div>
        <div style="font-size:12px; color:#777;">
          Try selecting a different URL field, or load more.
        </div>
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
      // fallback
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

    // Open lightbox viewer
    card.addEventListener("click", () => openViewer(globalIndex));

    gridEl.appendChild(card);
  });

  countPill.textContent = `${allUrls.length.toLocaleString()} images`;
  filteredPill.textContent = `${filteredUrls.length.toLocaleString()} shown`;
  chunkPill.textContent = `Loaded: ${(loadedChunks * getChunkSize()).toLocaleString()} (chunks: ${loadedChunks})`;

  updatePager();
}

// ---------- Search ----------
function applySearch() {
  const q = (searchInput.value || "").trim().toLowerCase();
  if (!q) {
    filteredUrls = allUrls.slice();
  } else {
    filteredUrls = allUrls.filter(u => u.toLowerCase().includes(q));
  }
  currentPage = 1;
  renderGrid();
}

// ---------- Worksheet + Field discovery ----------
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

  if (current && sheets.some(s => s.name === current)) {
    sheetSelect.value = current;
  } else if (sheets.length) {
    sheetSelect.value = sheets[0].name;
  }
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

  // Prefer URL-ish columns by name
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
    if (hints.some(h => name.includes(h))) idxs.push(i);
  });

  // If nothing matches, scan all columns
  return idxs.length ? idxs : columns.map((_, i) => i);
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
      // Reset loaded state
      lastCacheKey = cacheKey;
      loadedChunks = 0;
      allUrls = [];
      filteredUrls = [];
      currentPage = 1;
    }

    const ws = dashboardRef.worksheets.find(w => w.name === sheetName);
    if (!ws) {
      showError(`<div class="errorText">Worksheet not found: ${escapeHtml(sheetName)}</div>`);
      setStatus("Worksheet not found");
      return;
    }

    // NOTE: Tableau summary data reads from the current view.
    // We cannot truly "offset" rows in a single view without changing the worksheet (filters) or using
    // underlying data paging APIs (not consistently available).
    //
    // Practical approach:
    // - Read up to (chunksLoaded+1)*chunkSize rows from the sheet view
    // - Extract URLs and dedupe
    // This is safe and works as long as the worksheet view has enough rows and is filtered appropriately.

    const maxRows = (loadedChunks + 1) * chunkSize;
    setStatus(`Loading up to ${maxRows.toLocaleString()} rows…`);

    const summary = await ws.getSummaryDataAsync({ maxRows });
    const columns = summary.columns || [];
    const rows = summary.data || [];

    buildFieldOptions(columns);

    const idxs = findCandidateColumnIndexes(columns, chosenField);

    // Extract URLs
    const extracted = [];
    for (const row of rows) {
      for (const colIdx of idxs) {
        const cell = row[colIdx];
        const raw = cell?.value ?? cell?.formattedValue ?? null;
        const url = normalizeUrl(raw);
        if (url) extracted.push(url);
      }
    }

    // Dedupe and keep order
    const seen = new Set();
    const merged = [];
    // Keep existing first
    for (const u of allUrls) {
      if (seen.has(u)) continue;
      seen.add(u);
      merged.push(u);
    }
    // Add extracted next
    for (const u of extracted) {
      if (seen.has(u)) continue;
      seen.add(u);
      merged.push(u);
    }

    const before = allUrls.length;
    allUrls = merged;

    // Decide if we actually increased enough to count as another chunk.
    // If the sheet doesn't have more rows, this won't grow and "Load more" won't help.
    if (allUrls.length > before) {
      loadedChunks += 1;
    }

    setStatus(`Loaded ${allUrls.length.toLocaleString()} unique URL(s).`);

    applySearch(); // sets filteredUrls + renders page 1

    // Enable/disable load more button heuristically
    // If we didn't add any new URLs this time, we likely hit the end of what's visible in the worksheet.
    if (allUrls.length === before) {
      loadMoreBtn.disabled = true;
      loadMoreBtn.textContent = "No more in view";
    } else {
      loadMoreBtn.disabled = false;
      loadMoreBtn.textContent = "Load more";
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

// ---------- Infinite scroll ----------
function maybeInfiniteAdvance() {
  if (!infiniteToggle.checked) return;

  const total = getTotalPages();
  if (currentPage >= total) return;

  // if near bottom, go next page
  const thresholdPx = 600;
  const remaining = gridWrapEl.scrollHeight - (gridWrapEl.scrollTop + gridWrapEl.clientHeight);
  if (remaining < thresholdPx) {
    currentPage += 1;
    renderGrid();
  }
}

// ---------- Viewer ----------
function applyTransform() {
  fullImgEl.style.transform = `translate(${offsetX}px, ${offsetY}px) scale(${scale})`;
}
function resetView() {
  scale = 1; offsetX = 0; offsetY = 0;
  applyTransform();
}
function fitToView() {
  if (!fullImgEl.naturalWidth || !fullImgEl.naturalHeight) { resetView(); return; }
  const rect = stageEl.getBoundingClientRect();
  const wScale = rect.width / fullImgEl.naturalWidth;
  const hScale = rect.height / fullImgEl.naturalHeight;
  scale = Math.max(0.2, Math.min(8, Math.min(wScale, hScale) * 0.98));
  offsetX = 0; offsetY = 0;
  applyTransform();
}
function zoomAt(x, y, factor) {
  const newScale = Math.max(0.2, Math.min(8, scale * factor));
  factor = newScale / scale;
  offsetX = (offsetX - x) * factor + x;
  offsetY = (offsetY - y) * factor + y;
  scale = newScale;
  applyTransform();
}
function zoomAtCenter(factor) {
  const rect = stageEl.getBoundingClientRect();
  zoomAt(rect.width / 2, rect.height / 2, factor);
}
function openViewer(globalIndex) {
  viewerIndex = Math.max(0, Math.min(filteredUrls.length - 1, globalIndex));
  overlayEl.classList.add("open");
  overlayEl.setAttribute("aria-hidden", "false");
  showViewerIndex(viewerIndex);
}
function closeViewer() {
  overlayEl.classList.remove("open");
  overlayEl.setAttribute("aria-hidden", "true");
}
function showViewerIndex(idx) {
  if (!filteredUrls.length) return;
  viewerIndex = (idx + filteredUrls.length) % filteredUrls.length;

  const url = filteredUrls[viewerIndex];
  viewerTitleEl.textContent = safeHost(url);
  viewerCountEl.textContent = `${viewerIndex + 1} / ${filteredUrls.length}`;

  openTabBtn.onclick = () => window.open(url, "_blank", "noopener,noreferrer");

  fullImgEl.onload = () => fitToView();
  fullImgEl.onerror = () => {
    // skip broken images
    showViewerIndex(viewerIndex + 1);
  };
  fullImgEl.src = url;
  resetView();
}

// Viewer interactions
stageEl.addEventListener("wheel", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  e.preventDefault();
  const rect = stageEl.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  zoomAt(x, y, e.deltaY > 0 ? (1/1.2) : 1.2);
}, { passive:false });

viewportEl.addEventListener("pointerdown", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  dragging = true;
  viewportEl.classList.add("dragging");
  dragStart = { x: e.clientX, y: e.clientY, ox: offsetX, oy: offsetY };
});
window.addEventListener("pointermove", (e) => {
  if (!dragging) return;
  offsetX = dragStart.ox + (e.clientX - dragStart.x);
  offsetY = dragStart.oy + (e.clientY - dragStart.y);
  applyTransform();
});
window.addEventListener("pointerup", () => {
  dragging = false;
  viewportEl.classList.remove("dragging");
});

zoomInBtn.addEventListener("click", () => zoomAtCenter(1.2));
zoomOutBtn.addEventListener("click", () => zoomAtCenter(1/1.2));
fitBtn.addEventListener("click", fitToView);
resetBtn.addEventListener("click", resetView);

closeBtn.addEventListener("click", closeViewer);
vPrevBtn.addEventListener("click", () => showViewerIndex(viewerIndex - 1));
vNextBtn.addEventListener("click", () => showViewerIndex(viewerIndex + 1));

window.addEventListener("keydown", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  if (e.key === "Escape") closeViewer();
  if (e.key === "ArrowLeft") showViewerIndex(viewerIndex - 1);
  if (e.key === "ArrowRight") showViewerIndex(viewerIndex + 1);
  if (e.key === "+" || e.key === "=") zoomAtCenter(1.2);
  if (e.key === "-" || e.key === "_") zoomAtCenter(1/1.2);
  if (e.key.toLowerCase() === "f") fitToView();
  if (e.key === "0") resetView();
});

// ---------- Wire up UI ----------
searchInput.addEventListener("input", () => applySearch());

firstBtn.addEventListener("click", () => { currentPage = 1; renderGrid(); });
prevBtn.addEventListener("click", () => { currentPage = clampPage(currentPage - 1); renderGrid(); });
nextBtn.addEventListener("click", () => { currentPage = clampPage(currentPage + 1); renderGrid(); });
lastBtn.addEventListener("click", () => { currentPage = getTotalPages(); renderGrid(); });

gridWrapEl.addEventListener("scroll", () => maybeInfiniteAdvance());

refreshBtn.addEventListener("click", () => loadChunk({ reset:true }));
clearBtn.addEventListener("click", () => {
  lastCacheKey = null;
  loadedChunks = 0;
  allUrls = [];
  filteredUrls = [];
  currentPage = 1;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  setStatus("Cache cleared");
  renderGrid();
});
loadMoreBtn.addEventListener("click", () => loadChunk({ reset:false }));

sheetSelect.addEventListener("change", () => {
  // new sheet: reset
  lastCacheKey = null;
  loadedChunks = 0;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  loadChunk({ reset:true });
});
fieldSelect.addEventListener("change", () => {
  lastCacheKey = null;
  loadedChunks = 0;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  loadChunk({ reset:true });
});
chunkSelect.addEventListener("change", () => {
  // changing chunk size impacts how many rows we read; reset for simplicity
  lastCacheKey = null;
  loadedChunks = 0;
  loadMoreBtn.disabled = false;
  loadMoreBtn.textContent = "Load more";
  loadChunk({ reset:true });
});
pageSizeSelect.addEventListener("change", () => {
  currentPage = 1;
  renderGrid();
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