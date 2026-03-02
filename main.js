const gridEl = document.getElementById("grid");
const emptyEl = document.getElementById("empty");
const statusEl = document.getElementById("status");
const refreshBtn = document.getElementById("refreshBtn");
const fieldSelect = document.getElementById("fieldSelect");
const limitSelect = document.getElementById("limitSelect");

// Viewer
const overlayEl = document.getElementById("overlay");
const closeBtn = document.getElementById("closeBtn");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");
const viewerTitleEl = document.getElementById("viewerTitle");
const countPillEl = document.getElementById("countPill");
const stageEl = document.getElementById("stage");
const viewportEl = document.getElementById("viewport");
const fullImgEl = document.getElementById("fullImg");
const zoomInBtn = document.getElementById("zoomInBtn");
const zoomOutBtn = document.getElementById("zoomOutBtn");
const fitBtn = document.getElementById("fitBtn");
const resetBtn = document.getElementById("resetBtn");

refreshBtn.addEventListener("click", refreshAll);
fieldSelect.addEventListener("change", refreshAll);
limitSelect.addEventListener("change", refreshAll);

// Viewer events
closeBtn.addEventListener("click", closeViewer);
prevBtn.addEventListener("click", () => showViewerIndex(viewerIndex - 1));
nextBtn.addEventListener("click", () => showViewerIndex(viewerIndex + 1));
zoomInBtn.addEventListener("click", () => zoomAtCenter(1.2));
zoomOutBtn.addEventListener("click", () => zoomAtCenter(1 / 1.2));
fitBtn.addEventListener("click", fitToView);
resetBtn.addEventListener("click", resetView);
window.addEventListener("keydown", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  if (e.key === "Escape") closeViewer();
  if (e.key === "ArrowLeft") showViewerIndex(viewerIndex - 1);
  if (e.key === "ArrowRight") showViewerIndex(viewerIndex + 1);
  if (e.key === "+" || e.key === "=") zoomAtCenter(1.2);
  if (e.key === "-" || e.key === "_") zoomAtCenter(1 / 1.2);
  if (e.key.toLowerCase() === "f") fitToView();
  if (e.key === "0") resetView();
});

let dashboardRef = null;

// Gallery state
let images = []; // [{url, sourceSheet, field}]
let viewerIndex = 0;

// Viewer pan/zoom state
let scale = 1;
let offsetX = 0;
let offsetY = 0;
let dragging = false;
let dragStart = { x: 0, y: 0, ox: 0, oy: 0 };

function setStatus(msg) {
  statusEl.value = msg;
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

function applyTransform() {
  fullImgEl.style.transform = `translate(${offsetX}px, ${offsetY}px) scale(${scale})`;
}

function resetView() {
  scale = 1;
  offsetX = 0;
  offsetY = 0;
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

function fitToView() {
  if (!fullImgEl.naturalWidth || !fullImgEl.naturalHeight) {
    resetView();
    return;
  }
  const rect = stageEl.getBoundingClientRect();
  const wScale = rect.width / fullImgEl.naturalWidth;
  const hScale = rect.height / fullImgEl.naturalHeight;
  scale = Math.max(0.2, Math.min(8, Math.min(wScale, hScale) * 0.98));
  offsetX = 0;
  offsetY = 0;
  applyTransform();
}

// Pan / zoom interactions in viewer
stageEl.addEventListener("wheel", (e) => {
  if (!overlayEl.classList.contains("open")) return;
  e.preventDefault();
  const rect = stageEl.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  zoomAt(x, y, e.deltaY > 0 ? (1 / 1.2) : 1.2);
}, { passive: false });

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

function openViewer(idx) {
  viewerIndex = idx;
  overlayEl.classList.add("open");
  overlayEl.setAttribute("aria-hidden", "false");
  showViewerIndex(viewerIndex);
}

function closeViewer() {
  overlayEl.classList.remove("open");
  overlayEl.setAttribute("aria-hidden", "true");
}

function showViewerIndex(idx) {
  if (!images.length) return;
  viewerIndex = (idx + images.length) % images.length;

  const item = images[viewerIndex];
  viewerTitleEl.textContent = `${item.sourceSheet} • ${item.field}`;
  countPillEl.textContent = `${viewerIndex + 1} / ${images.length}`;

  // Load image
  fullImgEl.onload = () => fitToView();
  fullImgEl.onerror = () => {
    // skip broken image
    // show next automatically
    showViewerIndex(viewerIndex + 1);
  };
  fullImgEl.src = item.url;

  resetView();
}

function renderGrid() {
  gridEl.innerHTML = "";

  if (!images.length) {
    emptyEl.style.display = "grid";
    return;
  }
  emptyEl.style.display = "none";

  images.forEach((item, idx) => {
    const card = document.createElement("div");
    card.className = "card";
    card.title = item.url;

    const thumb = document.createElement("div");
    thumb.className = "thumb";

    const img = document.createElement("img");
    img.loading = "lazy";
    img.referrerPolicy = "no-referrer"; // safe default; remove if it breaks for your host
    img.src = item.url;

    thumb.appendChild(img);

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.innerHTML = `<span class="muted">${escapeHtml(item.sourceSheet)}</span><span>#${idx + 1}</span>`;

    card.appendChild(thumb);
    card.appendChild(meta);

    card.addEventListener("click", () => openViewer(idx));

    gridEl.appendChild(card);
  });
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function getLimit() {
  const n = parseInt(limitSelect.value, 10);
  return Number.isFinite(n) ? n : 100;
}

function getChosenField() {
  const v = fieldSelect.value;
  return v ? v : null; // null = auto-detect
}

function buildFieldOptionsFromMarks(marks) {
  // Build dropdown options based on visible columns in selection
  const set = new Set();
  for (const table of marks.data || []) {
    for (const c of (table.columns || [])) {
      const key = c.caption || c.fieldName;
      if (key) set.add(key);
    }
  }

  // Keep existing selection if possible
  const current = fieldSelect.value;

  // Replace options (leave Auto-detect)
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  [...set].sort().forEach(k => {
    const opt = document.createElement("option");
    opt.value = k;
    opt.textContent = k;
    fieldSelect.appendChild(opt);
  });

  if (current && set.has(current)) fieldSelect.value = current;
}

function findUrlColumnIndexes(columns, chosenField) {
  if (chosenField) {
    const idx = columns.findIndex(c => c.caption === chosenField || c.fieldName === chosenField);
    return idx >= 0 ? [idx] : [];
  }

  // Auto-detect: any column name that hints URL/image
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.caption || ""} ${c.fieldName || ""}`.toLowerCase();
    if (hints.some(h => name.includes(h))) idxs.push(i);
  });

  // If no hints, fall back to scanning all columns
  if (!idxs.length) return columns.map((_, i) => i);
  return idxs;
}

async function collectUrlsFromWorksheet(ws, chosenField, limit, out) {
  try {
    const marks = await ws.getSelectedMarksAsync();
    if (!marks?.data?.length) return;

    // Update field dropdown from first meaningful selection we see
    buildFieldOptionsFromMarks(marks);

    for (const table of marks.data) {
      const columns = table.columns || [];
      const rows = table.data || [];
      if (!columns.length || !rows.length) continue;

      const candidateIdxs = findUrlColumnIndexes(columns, chosenField);

      for (const row of rows) {
        for (const colIdx of candidateIdxs) {
          const cell = row[colIdx];
          const raw = (cell?.value ?? cell?.formattedValue ?? null);
          const url = normalizeUrl(raw);
          if (!url) continue;

          out.push({
            url,
            sourceSheet: ws.name,
            field: columns[colIdx]?.caption || columns[colIdx]?.fieldName || "URL"
          });

          if (out.length >= limit) return;
        }
        if (out.length >= limit) return;
      }
    }
  } catch (e) {
    // Some sheets can throw; ignore and continue
    console.warn(`Skipping sheet "${ws.name}" due to selection read error`, e);
  }
}

async function refreshAll() {
  try {
    setStatus("Collecting images from selections...");
    const limit = getLimit();
    const chosenField = getChosenField();

    const collected = [];

    for (const ws of dashboardRef.worksheets) {
      await collectUrlsFromWorksheet(ws, chosenField, limit, collected);
      if (collected.length >= limit) break;
    }

    // Deduplicate by URL (keeps order)
    const seen = new Set();
    images = collected.filter(item => {
      if (seen.has(item.url)) return false;
      seen.add(item.url);
      return true;
    });

    setStatus(`Found ${images.length} image(s)${images.length >= limit ? " (limit reached)" : ""}`);
    renderGrid();
  } catch (e) {
    console.error("refreshAll error:", e);
    setStatus("Error");
    emptyEl.style.display = "grid";
    emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">Error reading selections: ${escapeHtml(e?.message || String(e))}</div>`;
  }
}

async function init() {
  await tableau.extensions.initializeAsync();
  dashboardRef = tableau.extensions.dashboardContent.dashboard;

  // Refresh when any sheet selection changes
  dashboardRef.worksheets.forEach(ws => {
    ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, refreshAll);
  });

  await refreshAll();
}

init().catch(e => {
  console.error(e);
  setStatus("Init error");
});
