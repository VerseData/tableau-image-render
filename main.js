// ========= CONFIG (EDIT THIS FIELD NAME) =========
const CONFIG = {
  urlFieldCaptionOrFieldName: "Image URL",
  defaultImageUrl: null,

  // Zoom constraints
  minScale: 0.2,
  maxScale: 8.0,
  zoomStep: 1.2
};
// ================================================

const rootEl = document.getElementById("root");
const stageEl = document.getElementById("stage");
const viewportEl = document.getElementById("viewport");
const statusEl = document.getElementById("status");
const imgEl = document.getElementById("img");
const placeholderEl = document.getElementById("placeholder");

const refreshBtn = document.getElementById("refreshBtn");
const zoomInBtn = document.getElementById("zoomInBtn");
const zoomOutBtn = document.getElementById("zoomOutBtn");
const resetBtn = document.getElementById("resetBtn");
const fitBtn = document.getElementById("fitBtn");
const fsBtn = document.getElementById("fsBtn");

refreshBtn.addEventListener("click", refreshAll);
zoomInBtn.addEventListener("click", () => zoomAtCenter(CONFIG.zoomStep));
zoomOutBtn.addEventListener("click", () => zoomAtCenter(1 / CONFIG.zoomStep));
resetBtn.addEventListener("click", resetView);
fitBtn.addEventListener("click", fitToView);
fsBtn.addEventListener("click", toggleFullscreen);

let dashboardRef = null;

// View state
let scale = 1.0;
let offsetX = 0; // px
let offsetY = 0; // px
let isDragging = false;
let dragStart = { x: 0, y: 0, ox: 0, oy: 0 };

function setStatus(msg) {
  statusEl.value = msg;
}

function showPlaceholder(message, isError = false) {
  imgEl.style.display = "none";
  placeholderEl.style.display = "grid";
  placeholderEl.innerHTML = `<div class="${isError ? "err" : ""}">${escapeHtml(message)}</div>`;
}

// function showImage(url) {
//   placeholderEl.style.display = "none";
//   imgEl.style.display = "block";
//   imgEl.src = appendQueryParam(url, "_ts", String(Date.now()));
// }

function showImage(url) {
  placeholderEl.style.display = "none";
  imgEl.style.display = "block";
  imgEl.style.transform = "none";   // disable zoom logic
  imgEl.src = url;                  // no cache busting
}

function appendQueryParam(url, key, value) {
  try {
    const u = new URL(url);
    u.searchParams.set(key, value);
    return u.toString();
  } catch {
    return url;
  }
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function clamp(v, min, max) {
  return Math.min(max, Math.max(min, v));
}

function applyTransform() {
  imgEl.style.transform = `translate(${offsetX}px, ${offsetY}px) scale(${scale})`;
}

function resetView() {
  scale = 1.0;
  offsetX = 0;
  offsetY = 0;
  applyTransform();
}

function zoomAt(pointX, pointY, factor) {
  const newScale = clamp(scale * factor, CONFIG.minScale, CONFIG.maxScale);
  factor = newScale / scale;

  // Convert point to keep anchored during zoom
  // point is in stage coords; our transform is translate + scale around center.
  offsetX = (offsetX - pointX) * factor + pointX;
  offsetY = (offsetY - pointY) * factor + pointY;

  scale = newScale;
  applyTransform();
}

function zoomAtCenter(factor) {
  const rect = stageEl.getBoundingClientRect();
  zoomAt(rect.width / 2, rect.height / 2, factor);
}

function fitToView() {
  // Fit the image within stage by adjusting scale based on natural dimensions.
  // Works best once image is loaded.
  if (!imgEl.naturalWidth || !imgEl.naturalHeight) {
    resetView();
    return;
  }

  const rect = stageEl.getBoundingClientRect();
  const wScale = rect.width / imgEl.naturalWidth;
  const hScale = rect.height / imgEl.naturalHeight;

  // Fit with some padding
  const fitScale = clamp(Math.min(wScale, hScale) * 0.98, CONFIG.minScale, CONFIG.maxScale);

  scale = fitScale;
  offsetX = 0;
  offsetY = 0;
  applyTransform();
}

function toggleFullscreen() {
  // In-iframe fullscreen (always works)
  rootEl.classList.toggle("fullscreen");

  // Try browser fullscreen if allowed by embedding policies
  // If blocked, no problem; in-iframe fullscreen still works.
  try {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen?.().catch(() => {});
    } else {
      document.exitFullscreen?.().catch(() => {});
    }
  } catch {}
}

function onWheel(e) {
  // Zoom on wheel/trackpad with cursor focus
  e.preventDefault();
  const delta = e.deltaY;
  const factor = delta > 0 ? (1 / CONFIG.zoomStep) : CONFIG.zoomStep;

  const rect = stageEl.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;

  zoomAt(x, y, factor);
}

function onPointerDown(e) {
  if (imgEl.style.display === "none") return;
  isDragging = true;
  viewportEl.classList.add("dragging");
  dragStart = { x: e.clientX, y: e.clientY, ox: offsetX, oy: offsetY };
}

function onPointerMove(e) {
  if (!isDragging) return;
  offsetX = dragStart.ox + (e.clientX - dragStart.x);
  offsetY = dragStart.oy + (e.clientY - dragStart.y);
  applyTransform();
}

function onPointerUp() {
  isDragging = false;
  viewportEl.classList.remove("dragging");
}

function onDoubleClick(e) {
  if (imgEl.style.display === "none") return;
  const rect = stageEl.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;

  // Toggle between 1x and 2x around cursor
  const target = scale < 1.8 ? 2.0 : 1.0;
  const factor = target / scale;
  zoomAt(x, y, factor);
}

function onKeyDown(e) {
  if (e.key === "+" || e.key === "=") zoomAtCenter(CONFIG.zoomStep);
  else if (e.key === "-" || e.key === "_") zoomAtCenter(1 / CONFIG.zoomStep);
  else if (e.key === "0") resetView();
  else if (e.key.toLowerCase() === "f") fitToView();
  else if (e.key === "Enter") toggleFullscreen();
}

function findUrlColumnIndex(columns) {
  const target = CONFIG.urlFieldCaptionOrFieldName;
  return columns.findIndex(c => c.caption === target || c.fieldName === target);
}

async function tryGetImageFromWorksheet(worksheet) {
  const marks = await worksheet.getSelectedMarksAsync();
  if (!marks?.data?.length) return null;

  const table = marks.data[0];
  const colIdx = findUrlColumnIndex(table.columns);
  if (colIdx < 0) return null;

  if (!table.data?.length) return null;

  const cell = table.data[0][colIdx];
  const url = cell?.formattedValue || cell?.value;
  if (!url) return null;

  return String(url).trim();
}

async function refreshAll() {
  try {
    setStatus("Checking selections...");

    for (const ws of dashboardRef.worksheets) {
      const url = await tryGetImageFromWorksheet(ws);
      if (url) {
        showImage(url);
        resetView();
        setStatus(`Loaded from "${ws.name}"`);
        return;
      }
    }

    if (CONFIG.defaultImageUrl) {
      showImage(CONFIG.defaultImageUrl);
      resetView();
      setStatus("Showing default image");
    } else {
      showPlaceholder("Select a mark in any worksheet that contains the image URL field.");
      setStatus("Waiting for selection");
    }
  } catch (e) {
    console.error(e);
    showPlaceholder("Error loading image.", true);
    setStatus("Error");
  }
}

async function init() {
  try {
    await tableau.extensions.initializeAsync();
    dashboardRef = tableau.extensions.dashboardContent.dashboard;

    // Listen for selection changes on all sheets
    dashboardRef.worksheets.forEach(ws => {
      ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, refreshAll);
    });

    // Interaction controls
    stageEl.addEventListener("wheel", onWheel, { passive: false });
    viewportEl.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    stageEl.addEventListener("dblclick", onDoubleClick);
    window.addEventListener("keydown", onKeyDown);

    // Fit when image loads (nice default)
    imgEl.addEventListener("load", () => {
      fitToView();
    });
    imgEl.addEventListener("error", () => {
      showPlaceholder("Image failed to load (broken URL, blocked embedding, or auth required).", true);
      setStatus("Image error");
    });

    await refreshAll();
  } catch (e) {
    console.error(e);
    showPlaceholder("Failed to initialize Tableau Extensions API.", true);
    setStatus("Init error");
  }
}

init();
