// =====================
// Tableau Image Gallery (Full Sheet)
// - Reads ALL image URLs from a chosen worksheet via getSummaryDataAsync()
// - Auto-detects URL-like columns OR lets user pick the exact field
// - Skips Null/None/empty values
// - Dedupes URLs
// - Avoids double initializeAsync()
// =====================
if (!window.tableau || !tableau.extensions) {
  document.getElementById("empty").style.display = "grid";
  document.getElementById("empty").innerHTML =
    `<div style="color:#b00020; text-align:center;">
      Tableau Extensions API not loaded.<br/><br/>
      Fix: host <b>tableau.extensions.1.latest.min.js</b> in the same GitHub Pages repo and reference it locally.
    </div>`;
  throw new Error("Tableau Extensions API not loaded (tableau is undefined)");
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
const gridEl = document.getElementById("grid");
const emptyEl = document.getElementById("empty");
const refreshBtn = document.getElementById("refreshBtn");
const fieldSelect = document.getElementById("fieldSelect");
const limitSelect = document.getElementById("limitSelect");

// Inject worksheet selector next to URL field dropdown
const barEl = document.querySelector(".bar");
const sheetLabel = document.createElement("span");
sheetLabel.className = "small";
sheetLabel.textContent = "Sheet:";
sheetLabel.style.marginLeft = "6px";

const sheetSelect = document.createElement("select");
sheetSelect.id = "sheetSelect";
sheetSelect.style.marginLeft = "6px";

barEl.insertBefore(sheetLabel, fieldSelect);
barEl.insertBefore(sheetSelect, fieldSelect);

// ---------- Events ----------
sheetSelect.addEventListener("change", () => {
  fieldSelect.innerHTML = `<option value="">Auto-detect</option>`;
  // invalidate cache when sheet changes
  lastCacheKey = null;
  cached = null;
  refreshAll();
});
refreshBtn.addEventListener("click", () => {
  lastCacheKey = null;
  cached = null;
  refreshAll();
});
fieldSelect.addEventListener("change", () => {
  lastCacheKey = null;
  cached = null;
  refreshAll();
});
limitSelect.addEventListener("change", () => {
  lastCacheKey = null;
  cached = null;
  refreshAll();
});

// ---------- State ----------
let images = []; // [{url}]
let lastCacheKey = null;
let cached = null;

// ---------- Helpers ----------
function setStatus(msg) {
  statusEl.value = msg;
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
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

    // Keep simple for now: open image in new tab.
    // If you want, we can add an in-extension fullscreen viewer with zoom/pan.
    card.addEventListener("click", () =>
      window.open(item.url, "_blank", "noopener,noreferrer")
    );

    gridEl.appendChild(card);
  });
}

function fillSheetSelect() {
  if (!dashboardRef?.worksheets?.length) return;

  const sheets = dashboardRef.worksheets;
  const current = sheetSelect.value;

  sheetSelect.innerHTML = "";
  sheets.forEach((ws) => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name;
    sheetSelect.appendChild(opt);
  });

  if (current && sheets.some((s) => s.name === current)) {
    sheetSelect.value = current;
  }
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

  if (current) fieldSelect.value = current;
}

function findCandidateColumnIndexes(columns, chosenField) {
  if (chosenField) {
    const idx = columns.findIndex(
      (c) => c.fieldName === chosenField || c.caption === chosenField
    );
    return idx >= 0 ? [idx] : [];
  }

  // name-based heuristic first
  const hints = ["url", "image", "img", "photo", "link", "href"];
  const idxs = [];
  columns.forEach((c, i) => {
    const name = `${c.fieldName || ""} ${c.caption || ""}`.toLowerCase();
    if (hints.some((h) => name.includes(h))) idxs.push(i);
  });

  // if nothing matches, scan all columns
  return idxs.length ? idxs : columns.map((_, i) => i);
}

async function getAllUrlsFromWorksheet(ws, limit, chosenField) {
  // Summary data = what's currently in the worksheet view (respects filters)
  // Pull extra rows so Nulls don't cause us to return too few images
  const summary = await ws.getSummaryDataAsync({ maxRows: limit * 10 });

  const columns = summary.columns || [];
  const rows = summary.data || [];

  buildFieldOptions(columns);

  const idxs = findCandidateColumnIndexes(columns, chosenField);
  const out = [];

  for (const row of rows) {
    for (const colIdx of idxs) {
      const cell = row[colIdx];
      // Prefer raw value; formattedValue can be "Null"
      const raw = cell?.value ?? cell?.formattedValue ?? null;
      const url = normalizeUrl(raw);
      if (!url) continue;

      out.push(url);
      if (out.length >= limit) {
        return { urls: out, totalRows: rows.length };
      }
    }
  }

  return { urls: out, totalRows: rows.length };
}

// ---------- Main ----------
async function refreshAll() {
  try {
    await ensureInitialized();

    setStatus("Loading worksheet data...");
    fillSheetSelect();

    const sheetName =
      getChosenSheet() || dashboardRef.worksheets?.[0]?.name || null;

    if (!sheetName) {
      emptyEl.style.display = "grid";
      emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">No worksheets found.</div>`;
      setStatus("No worksheets");
      return;
    }
    sheetSelect.value = sheetName;

    const ws = dashboardRef.worksheets.find((w) => w.name === sheetName);
    if (!ws) {
      emptyEl.style.display = "grid";
      emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">Worksheet not found: ${escapeHtml(
        sheetName
      )}</div>`;
      setStatus("Worksheet not found");
      return;
    }

    const limit = getLimit();
    const chosenField = getChosenField();

    const cacheKey = `${sheetName}::${chosenField || "auto"}::${limit}`;
    if (cacheKey === lastCacheKey && cached) {
      images = cached;
      renderGrid();
      setStatus(`Loaded ${images.length} from cache`);
      return;
    }

    const { urls, totalRows } = await getAllUrlsFromWorksheet(
      ws,
      limit,
      chosenField
    );

    // Deduplicate
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
    setStatus(
      `Found ${images.length} image(s) from ${totalRows} row(s) in "${sheetName}"${
        images.length >= limit ? " (limit reached)" : ""
      }`
    );
  } catch (e) {
    console.error("refreshAll error:", e);
    emptyEl.style.display = "grid";
    emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">
      Error: ${escapeHtml(e?.message || String(e))}
      <br/><br/>
      If this mentions permissions, your .trex must include <b>&lt;permission&gt;full data&lt;/permission&gt;</b>.
    </div>`;
    setStatus("Error");
  }
}

// Kick off once (do NOT call initializeAsync elsewhere)
ensureInitialized()
  .then(() => refreshAll())
  .catch((e) => {
    console.error("Init failed:", e);
    emptyEl.style.display = "grid";
    emptyEl.innerHTML = `<div style="color:#b00020; text-align:center;">Init error: ${escapeHtml(
      e?.message || String(e)
    )}</div>`;
    setStatus("Init error");
  });
