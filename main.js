let dashboardRef = null;

const statusEl = document.getElementById("status");
const gridEl = document.getElementById("grid");
const emptyEl = document.getElementById("empty");
const refreshBtn = document.getElementById("refreshBtn");
const fieldSelect = document.getElementById("fieldSelect");
const limitSelect = document.getElementById("limitSelect");

// Add a worksheet selector (we'll inject it beside URL field dropdown)
const barEl = document.querySelector(".bar");
const sheetSelect = document.createElement("select");
sheetSelect.id = "sheetSelect";
sheetSelect.style.marginLeft = "6px";
barEl.insertBefore(sheetSelect, fieldSelect);

sheetSelect.addEventListener("change", () => {
  // reset field dropdown when sheet changes
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  refreshAll();
});
refreshBtn.addEventListener("click", refreshAll);
fieldSelect.addEventListener("change", refreshAll);
limitSelect.addEventListener("change", refreshAll);

let images = []; // {url, idx}
let lastCacheKey = null;
let cached = null;

function setStatus(msg) { statusEl.value = msg; }

function escapeHtml(str) {
  return String(str).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");
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

function getLimit() {
  const n = parseInt(limitSelect.value, 10);
  return Number.isFinite(n) ? n : 100;
}

function getChosenField() {
  const v = fieldSelect.value;
  return v ? v : null;
}

function getChosenSheet() {
  const v = sheetSelect.value;
  return v ? v : null;
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
    img.src = item.url;
    thumb.appendChild(img);

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.innerHTML = `<span class="muted">#${idx + 1}</span><span>open</span>`;

    card.appendChild(thumb);
    card.appendChild(meta);

    card.addEventListener("click", () => window.open(item.url, "_blank", "noopener,noreferrer"));
    gridEl.appendChild(card);
  });
}

function fillSheetSelect() {
  const sheets = dashboardRef.worksheets || [];
  const current = sheetSelect.value;
  sheetSelect.innerHTML = "";
  sheets.forEach(ws => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name;
    sheetSelect.appendChild(opt);
  });

  if (current && sheets.some(s => s.name === current)) sheetSelect.value = current;
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
    const idx = columns.findIndex(c => (c.fieldName === chosenField) || (c.caption === chosenField));
    return idx >= 0 ? [idx] : [];
  }
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
    if (hints.some(h => name.includes(h))) idxs.push(i);
  });
  // if no hints, scan all columns
  return idxs.length ? idxs : columns.map((_, i) => i);
}

async function getAllUrlsFromWorksheet(ws, limit, chosenField) {
  // Use Summary Data (what’s in the current view). This respects filters.
  const summary = await ws.getSummaryDataAsync({ maxRows: limit * 5 }); // allow null skipping
  const columns = summary.columns || [];
  const rows = summary.data || [];

  buildFieldOptions(columns);

  const idxs = findCandidateColumnIndexes(columns, chosenField);
  const out = [];

  for (const row of rows) {
    for (const colIdx of idxs) {
      const cell = row[colIdx];
      const raw = (cell?.value ?? cell?.formattedValue ?? null);
      const url = normalizeUrl(raw);
      if (!url) continue;
      out.push(url);
      if (out.length >= limit) return { urls: out, totalRows: rows.length };
    }
  }
  return { urls: out, totalRows: rows.length };
}

async function refreshAll() {
  try {
    setStatus("Loading worksheet data...");
    fillSheetSelect();

    const sheetName = getChosenSheet() || dashboardRef.worksheets?.[0]?.name;
    if (!sheetName) {
      setStatus("No worksheets found");
      return;
    }
    sheetSelect.value = sheetName;

    const ws = dashboardRef.worksheets.find(w => w.name === sheetName);
    if (!ws) {
      setStatus("Worksheet not found");
      return;
    }

    const limit = getLimit();
    const chosenField = getChosenField();

    // Cache key to avoid reloading on minor UI events
    const cacheKey = `${sheetName}::${chosenField || "auto"}::${limit}`;
    if (cacheKey === lastCacheKey && cached) {
      images = cached;
      renderGrid();
      setStatus(`Loaded ${images.length} from cache`);
      return;
    }

    const { urls, totalRows } = await getAllUrlsFromWorksheet(ws, limit, chosenField);

    // Dedup
    const seen = new Set();
    const unique = [];
    for (const u of urls) {
      if (seen.has(u)) continue;
      seen.add(u);
      unique.push({ url: u });
    }

    images = unique;
    cached = images;
    lastCacheKey = cacheKey;

    renderGrid();
    setStatus(`Found ${images.length} image(s) from ${totalRows} row(s) in "${sheetName}"${images.length >= limit ? " (limit reached)" : ""}`);
  } catch (e) {
    console.error(e);
    emptyEl.style.display = "grid";
    emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">Error: ${escapeHtml(e?.message || String(e))}<br/><br/>If this says permissions, you must enable <b>Full Data</b> permission in the .trex.</div>`;
    setStatus("Error");
  }
}

async function init() {
  await tableau.extensions.initializeAsync();
  dashboardRef = tableau.extensions.dashboardContent.dashboard;
  fillSheetSelect();
  await refreshAll();
}

init();
