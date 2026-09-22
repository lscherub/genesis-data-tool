const OUTPUT_COLUMNS = [
  "Sku",
  "Product Number",
  "Description",
  "Vendor",
  "Brand",
  "List Cost",
  "Price",
  "Size Desc"
];

const fileInput = document.getElementById("fileInput");
const chooseBtn = document.getElementById("chooseBtn");
const dropZone = document.getElementById("dropZone");
const statusCard = document.getElementById("statusCard");
const fileNameEl = document.getElementById("fileName");
const statusBadge = document.getElementById("statusBadge");
const statusMessage = document.getElementById("statusMessage");
const checks = document.getElementById("checks");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");
const rowCount = document.getElementById("rowCount");
const downloadBtn = document.getElementById("downloadBtn");
const resetBtn = document.getElementById("resetBtn");

let cleanedRows = [];
let originalFileName = "";

chooseBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  fileInput.click();
});

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") fileInput.click();
});

fileInput.addEventListener("change", () => {
  if (fileInput.files?.[0]) processFile(fileInput.files[0]);
});

["dragenter", "dragover"].forEach(type => {
  dropZone.addEventListener(type, e => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach(type => {
  dropZone.addEventListener(type, e => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
  });
});

dropZone.addEventListener("drop", e => {
  const file = e.dataTransfer.files?.[0];
  if (file) processFile(file);
});

downloadBtn.addEventListener("click", downloadExcel);
resetBtn.addEventListener("click", resetApp);

async function processFile(file) {
  resetResultsOnly();
  originalFileName = file.name;
  statusCard.classList.remove("hidden");
  fileNameEl.textContent = file.name;
  setStatus("Reading file…", "Waiting");

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      raw: false
    });

    if (!workbook.SheetNames.length) throw new Error("The workbook does not contain a worksheet.");

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false
    });

    const headerInfo = findHeaderRow(matrix);
    if (!headerInfo) {
      throw new Error("Could not find the POS header row. A row containing “Sku” and “Description” is required.");
    }

    const { index: headerIndex, headers } = headerInfo;
    const normalized = headers.map(normalizeHeader);

    const found = {};
    OUTPUT_COLUMNS.forEach(col => {
      const idx = findColumnIndex(normalized, normalizeHeader(col));
      found[col] = idx;
    });

    const missing = OUTPUT_COLUMNS.filter(col => found[col] === -1);

    if (missing.length) {
      renderChecks(found, missing);
      throw new Error("Required column(s) missing: " + missing.join(", "));
    }

    renderChecks(found, []);
    const dataRows = matrix.slice(headerIndex + 1);

    cleanedRows = dataRows
      .map(row => {
        const output = {};
        OUTPUT_COLUMNS.forEach(col => {
          output[col] = cleanCellValue(row[found[col]]);
        });
        return output;
      })
      .filter(row => Object.values(row).some(v => String(v).trim() !== ""));

    if (!cleanedRows.length) {
      throw new Error("The file contains the correct columns, but no product rows were found.");
    }

    renderPreview();
    setStatus(
      `Successfully found ${cleanedRows.length.toLocaleString()} product${cleanedRows.length === 1 ? "" : "s"} and prepared the cleaned output.`,
      "Ready"
    );
    statusBadge.classList.add("success");
    downloadBtn.disabled = false;
  } catch (error) {
    setStatus(error.message || "Unable to process this file.", "Error");
    statusBadge.classList.add("error");
    downloadBtn.disabled = true;
  }
}

function findHeaderRow(matrix) {
  const maxRows = Math.min(matrix.length, 15);

  for (let r = 0; r < maxRows; r++) {
    const row = matrix[r] || [];
    const normalized = row.map(normalizeHeader);
    const hasSku = normalized.some(v => v === "sku");
    const hasDescription = normalized.some(v => v === "description");

    if (hasSku && hasDescription) {
      return { index: r, headers: row };
    }
  }
  return null;
}

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function findColumnIndex(headers, wanted) {
  return headers.findIndex(h => h === wanted);
}

function cleanCellValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value;
  return String(value).trim();
}

function renderChecks(found, missing) {
  checks.innerHTML = "";
  OUTPUT_COLUMNS.forEach(col => {
    const div = document.createElement("span");
    div.className = "check" + (missing.includes(col) ? " missing" : "");
    div.textContent = missing.includes(col) ? `✕ ${col}` : `✓ ${col}`;
    checks.appendChild(div);
  });
}

function renderPreview() {
  previewSection.classList.remove("hidden");
  rowCount.textContent = `${cleanedRows.length.toLocaleString()} product${cleanedRows.length === 1 ? "" : "s"}`;

  const thead = previewTable.querySelector("thead");
  const tbody = previewTable.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const headerRow = document.createElement("tr");
  OUTPUT_COLUMNS.forEach(col => {
    const th = document.createElement("th");
    th.textContent = col;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  cleanedRows.slice(0, 25).forEach(row => {
    const tr = document.createElement("tr");
    OUTPUT_COLUMNS.forEach(col => {
      const td = document.createElement("td");
      td.textContent = formatPreview(row[col]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function formatPreview(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  return value ?? "";
}

function downloadExcel() {
  if (!cleanedRows.length) return;

  // Build a minimal worksheet: only the 8 required columns and values.
  // We intentionally avoid per-cell styling/formatting because it can
  // substantially increase the size of large XLSX files.
  const matrix = [
    OUTPUT_COLUMNS,
    ...cleanedRows.map(row =>
      OUTPUT_COLUMNS.map(col => {
        const value = row[col];
        if (value instanceof Date && !isNaN(value)) {
          return value.toISOString().slice(0, 10);
        }
        return value ?? "";
      })
    )
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(matrix);

  // Only set practical column widths. No cell-level styles are added.
  worksheet["!cols"] = [
    { wch: 16 },
    { wch: 18 },
    { wch: 42 },
    { wch: 24 },
    { wch: 16 },
    { wch: 13 },
    { wch: 13 },
    { wch: 14 }
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Cleaned POS");

  const date = new Date().toISOString().slice(0, 10);
  const base = originalFileName
    .replace(/\.(xlsx|xls|csv)$/i, "")
    .replace(/[^a-z0-9_-]+/gi, "_")
    .replace(/^_+|_+$/g, "");

  XLSX.writeFile(workbook, `${base || "Genesis_POS"}_Cleaned_${date}.xlsx`, {
    compression: true
  });
}

function setStatus(message, badgeText) {
  statusMessage.textContent = message;
  statusBadge.textContent = badgeText;
  statusBadge.className = "badge";
}

function resetResultsOnly() {
  cleanedRows = [];
  previewSection.classList.add("hidden");
  downloadBtn.disabled = true;
  checks.innerHTML = "";
  statusBadge.className = "badge";
}

function resetApp() {
  fileInput.value = "";
  originalFileName = "";
  cleanedRows = [];
  statusCard.classList.add("hidden");
  previewSection.classList.add("hidden");
  downloadBtn.disabled = true;
  checks.innerHTML = "";
  statusMessage.textContent = "Ready.";
  statusBadge.textContent = "Waiting";
  statusBadge.className = "badge";
}
