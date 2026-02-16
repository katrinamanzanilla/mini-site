const LOCAL_SOURCE_KEY = "projectStatusSheetLink";

const REQUIRED_FILTER_KEYS = {
  system: ["system", "project name"],
  milestone: ["milestone"]
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseGvizResponse(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("Invalid response from Google Sheets.");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function parseGoogleSheetLink(rawValue) {
  const trimmed = String(rawValue ?? "").trim();
  if (!trimmed) {
    throw new Error("Please paste a Google Sheets link.");
  }

  const withProtocol = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  let url;

  try {
    url = new URL(withProtocol);
  } catch {
    throw new Error("Invalid URL. Paste a valid Google Sheets link.");
  }

  const validHost =
    url.hostname === "docs.google.com" ||
    url.hostname.endsWith(".docs.google.com");

  if (!validHost || !url.pathname.includes("/spreadsheets/d/")) {
    throw new Error("Only Google Sheets links are supported.");
  }

  return { url, displayUrl: withProtocol };
}

function getSheetSource(url) {
  const match = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match?.[1]) {
    throw new Error("Unable to read this Google Sheets link.");
  }

  const hashParams = new URLSearchParams(String(url.hash ?? "").replace(/^#/, ""));
  return {
    sheetId: match[1],
    gid: String(url.searchParams.get("gid") || hashParams.get("gid") || "").trim(),
    sheet: String(url.searchParams.get("sheet") || hashParams.get("sheet") || "").trim()
  };
}

function buildSheetApiUrl(source) {
  const params = new URLSearchParams({ tqx: "out:json", headers: "1" });

  if (source.gid) {
    params.set("gid", source.gid);
  } else if (source.sheet) {
    params.set("sheet", source.sheet);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/gviz/tq?${params.toString()}`;
}

function formatCellValue(cell, columnType) {
  if (!cell) return "";

  if (typeof cell.f === "string" && cell.f.trim()) {
    return cell.f;
  }

  if (cell.v == null) {
    return "";
  }

  if (columnType === "date" && typeof cell.v === "string") {
    const parts = cell.v.match(/^Date\((\d+),(\d+),(\d+)/);
    if (parts) {
      const year = Number(parts[1]);
      const month = Number(parts[2]);
      const day = Number(parts[3]);
      const parsed = new Date(year, month, day);
      return parsed.toLocaleDateString();
    }
  }

  return String(cell.v);
}

function mapPayload(payload) {
  const columns = (payload.table?.cols ?? []).map((column, index) => ({
    key: normalizeHeader(column.label) || `column ${index + 1}`,
    label: String(column.label || `Column ${index + 1}`).trim(),
    type: column.type || "string"
  }));

  const rows = (payload.table?.rows ?? []).map((row) => {
    const mapped = {};

    columns.forEach((column, index) => {
      mapped[column.key] = formatCellValue(row.c?.[index], column.type);
    });

    return mapped;
  });

  return { columns, rows };
}

function findColumnKey(columns, aliases) {
  const aliasSet = new Set(aliases.map((alias) => normalizeHeader(alias)));
  const found = columns.find((column) => aliasSet.has(column.key));
  return found?.key || "";
}

function uniqueValues(rows, key) {
  if (!key) return [];

  return [...new Set(rows
    .map((row) => String(row[key] ?? "").trim())
    .filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));
}

function setFeedback(message, isError = false) {
  const feedback = document.getElementById("sheet-source-feedback");
  feedback.textContent = message;
  feedback.style.color = isError ? "#ef4444" : "";
}

const state = {
  columns: [],
  rows: [],
  filteredRows: [],
  filters: {
    system: "",
    milestone: "",
    search: ""
  },
  filterColumns: {
    system: "",
    milestone: "",
    developer: "",
    manager: ""
  }
};

function renderSummary() {
  const totalProjects = document.getElementById("kpi-total-projects");
  const totalMilestones = document.getElementById("kpi-total-milestones");

  const projectValues = uniqueValues(state.filteredRows, state.filterColumns.system);
  totalProjects.textContent = String(projectValues.length);
  totalMilestones.textContent = String(state.filteredRows.length);
}

function renderFilterOptions() {
  const systemSelect = document.getElementById("system-filter");
  const milestoneSelect = document.getElementById("milestone-filter");

  const systemOptions = uniqueValues(state.rows, state.filterColumns.system);
  const milestoneOptions = uniqueValues(state.rows, state.filterColumns.milestone);

  const buildOptions = (values, selected) => {
    const defaultOption = '<option value="">All</option>';
    const options = values.map((value) => {
      const safe = escapeHtml(value);
      const isSelected = value === selected ? " selected" : "";
      return `<option value="${safe}"${isSelected}>${safe}</option>`;
    });
    return [defaultOption, ...options].join("");
  };

  systemSelect.innerHTML = buildOptions(systemOptions, state.filters.system);
  milestoneSelect.innerHTML = buildOptions(milestoneOptions, state.filters.milestone);

  systemSelect.disabled = !systemOptions.length;
  milestoneSelect.disabled = !milestoneOptions.length;
}

function rowMatchesSearch(row, query) {
  if (!query) return true;

  const fields = [
    row[state.filterColumns.system],
    row[state.filterColumns.developer],
    row[state.filterColumns.manager],
    row[state.filterColumns.milestone]
  ];

  return fields
    .map((value) => String(value ?? "").toLowerCase())
    .some((value) => value.includes(query));
}

function applyFilters() {
  const searchQuery = state.filters.search.trim().toLowerCase();

  state.filteredRows = state.rows.filter((row) => {
    if (state.filters.system && row[state.filterColumns.system] !== state.filters.system) {
      return false;
    }

    if (state.filters.milestone && row[state.filterColumns.milestone] !== state.filters.milestone) {
      return false;
    }

    return rowMatchesSearch(row, searchQuery);
  });

  renderTable();
  renderSummary();
}

function renderTable() {
  const head = document.getElementById("status-table-head");
  const body = document.getElementById("status-table-body");

  head.innerHTML = `
    <tr>
      ${state.columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}
    </tr>
  `;

  if (!state.filteredRows.length) {
    body.innerHTML = `<tr><td colspan="${Math.max(state.columns.length, 1)}">No results found</td></tr>`;
    return;
  }

  body.innerHTML = state.filteredRows
    .map((row) => `
      <tr>
        ${state.columns.map((column) => `<td>${escapeHtml(row[column.key])}</td>`).join("")}
      </tr>
    `)
    .join("");
}

function bindFilters() {
  document.getElementById("system-filter").addEventListener("change", (event) => {
    state.filters.system = event.target.value;
    applyFilters();
  });

  document.getElementById("milestone-filter").addEventListener("change", (event) => {
    state.filters.milestone = event.target.value;
    applyFilters();
  });

  document.getElementById("sheet-search").addEventListener("input", (event) => {
    state.filters.search = event.target.value;
    applyFilters();
  });
}

async function loadSheetData(url) {
  const source = getSheetSource(url);
  const response = await fetch(buildSheetApiUrl(source), { cache: "no-store" });

  if (!response.ok) {
    throw new Error(`Google Sheet API returned ${response.status}`);
  }

  const payload = parseGvizResponse(await response.text());
  const mapped = mapPayload(payload);

  if (!mapped.columns.length) {
    throw new Error("No columns found in the selected sheet.");
  }

  state.columns = mapped.columns;
  state.rows = mapped.rows;

  state.filterColumns.system = findColumnKey(mapped.columns, REQUIRED_FILTER_KEYS.system);
  state.filterColumns.milestone = findColumnKey(mapped.columns, REQUIRED_FILTER_KEYS.milestone);
  state.filterColumns.developer = findColumnKey(mapped.columns, ["assigned developer", "developer"]);
  state.filterColumns.manager = findColumnKey(mapped.columns, ["assigned project manager", "project manager"]);

  state.filters.system = "";
  state.filters.milestone = "";

  document.getElementById("system-filter").value = "";
  document.getElementById("milestone-filter").value = "";

  renderFilterOptions();
  applyFilters();
}

async function handleSubmit(event) {
  event.preventDefault();

  const input = document.getElementById("sheet-source-input");

  try {
    const parsed = parseGoogleSheetLink(input.value);
    localStorage.setItem(LOCAL_SOURCE_KEY, parsed.displayUrl);
    input.value = parsed.displayUrl;

    await loadSheetData(parsed.url);
    setFeedback("Sheet loaded successfully.");
  } catch (error) {
    state.columns = [];
    state.rows = [];
    state.filteredRows = [];
    renderTable();
    renderSummary();
    renderFilterOptions();

    setFeedback(error.message || "Unable to load Google Sheets data.", true);
  }
}

function resetView() {
  localStorage.removeItem(LOCAL_SOURCE_KEY);

  document.getElementById("sheet-source-input").value = "";
  document.getElementById("sheet-search").value = "";

  state.columns = [];
  state.rows = [];
  state.filteredRows = [];
  state.filters = { system: "", milestone: "", search: "" };
  state.filterColumns = { system: "", milestone: "", developer: "", manager: "" };

  renderFilterOptions();
  renderTable();
  renderSummary();
  setFeedback("Paste a Google Sheets link and click View Data.");
}

function init() {
  document.getElementById("sheet-source-form").addEventListener("submit", handleSubmit);
  document.getElementById("sheet-source-reset").addEventListener("click", resetView);

  bindFilters();
  resetView();

  const savedLink = localStorage.getItem(LOCAL_SOURCE_KEY);
  if (savedLink) {
    document.getElementById("sheet-source-input").value = savedLink;
    document.getElementById("sheet-source-form").requestSubmit();
  }
}

init();
