// ========= CONFIG (EDIT THIS FIELD NAME) =========
const CONFIG = {
  // The field (column) that contains the image URL
  urlFieldCaptionOrFieldName: "Image URL",

  // Optional default image
  defaultImageUrl: null
};
// =================================================

const statusEl = document.getElementById("status");
const imgEl = document.getElementById("img");
const placeholderEl = document.getElementById("placeholder");
document.getElementById("refreshBtn").addEventListener("click", refreshAll);

let dashboardRef = null;

function setStatus(msg) {
  statusEl.value = msg;
}

function showPlaceholder(message, isError = false) {
  imgEl.style.display = "none";
  placeholderEl.style.display = "block";
  placeholderEl.innerHTML = `<div style="color:${isError ? "#b00020" : "#555"}">${escapeHtml(message)}</div>`;
}

function showImage(url) {
  placeholderEl.style.display = "none";
  imgEl.style.display = "block";
  imgEl.src = appendQueryParam(url, "_ts", String(Date.now()));
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

function findUrlColumnIndex(columns) {
  const target = CONFIG.urlFieldCaptionOrFieldName;
  return columns.findIndex(c => c.caption === target || c.fieldName === target);
}

async function tryGetImageFromWorksheet(worksheet) {
  const marks = await worksheet.getSelectedMarksAsync();
  if (!marks || !marks.data || marks.data.length === 0) return null;

  const table = marks.data[0];
  const colIdx = findUrlColumnIndex(table.columns);
  if (colIdx < 0) return null;

  if (!table.data || table.data.length === 0) return null;

  const cell = table.data[0][colIdx];
  const url = cell?.formattedValue || cell?.value;
  if (!url) return null;

  return String(url).trim();
}

async function refreshAll() {
  try {
    setStatus("Checking selections...");

    const worksheets = dashboardRef.worksheets;

    for (const ws of worksheets) {
      const url = await tryGetImageFromWorksheet(ws);
      if (url) {
        showImage(url);
        setStatus(`Loaded from "${ws.name}"`);
        return;
      }
    }

    if (CONFIG.defaultImageUrl) {
      showImage(CONFIG.defaultImageUrl);
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

    // Attach selection listeners to ALL worksheets
    dashboardRef.worksheets.forEach(ws => {
      ws.addEventListener(
        tableau.TableauEventType.MarkSelectionChanged,
        refreshAll
      );
    });

    await refreshAll();
  } catch (e) {
    console.error(e);
    showPlaceholder("Failed to initialize Tableau Extensions API.", true);
    setStatus("Init error");
  }
}

init();
