const DEFAULT_SOURCE = {
  mode: "sheet",
  appsScriptUrl: "",
  sheetId: "1BHTM5Jx7xAyaRnzuyjD38-Iz5mlQddkA8s2J2ZatIXI",
  sheetTab: "Portfolio Status",
  gid: ""
};

const LOCAL_SOURCE_KEY = "projectStatusSheetSource";

const fallbackIssues = [
  {
    issueId: "PRJ-1001",
    project: "Township Expansion - North",
    owner: "PMO Team A",
    priority: "High",
    status: "On Track",
    blocker: "Waiting for utility permit",
    dueDate: "Feb 20, 2026",
    lastUpdate: "Feb 08, 2026"
  },
  {
    issueId: "PRJ-1027",
    project: "Digital Sales Portal",
    owner: "PMO Team B",
    priority: "Critical",
    status: "At Risk",
    blocker: "UAT defect backlog",
    dueDate: "Mar 03, 2026",
    lastUpdate: "Feb 09, 2026"
  },
  {
    issueId: "PRJ-1064",
    project: "CRM Data Cleansing",
    owner: "PMO Team C",
    priority: "Medium",
    status: "Delayed",
    blocker: "Data quality validation pending",
    dueDate: "Feb 27, 2026",
    lastUpdate: "Feb 07, 2026"
  },
  {
    issueId: "PRJ-1088",
    project: "Client Self-Service Enhancements",
    owner: "PMO Team D",
    priority: "Low",
    status: "Completed",
    blocker: "None",
    dueDate: "Feb 14, 2026",
    lastUpdate: "Feb 10, 2026"
  }
];

function toClassName(status) {
  return String(status)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

function formatMetricCount(count) {
  return String(count).padStart(2, "0");
}

function sanitize(value, fallback = "—") {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function sanitizeOptional(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeHeader(header) {
  return String(header ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseGvizResponse(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("Invalid Google Visualization response body");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function gvizToRows(gvizPayload) {
  const columns = (gvizPayload.table?.cols ?? []).map((column) => normalizeHeader(column.label));
  const rawRows = gvizPayload.table?.rows ?? [];

  return rawRows.map((row) => {
    const cells = row.c ?? [];
    const mapped = {};

    columns.forEach((columnName, index) => {
      const cell = cells[index];
      mapped[columnName] = cell?.f ?? cell?.v ?? "";
    });

    return mapped;
  });
}

function normalizeIssue(raw) {
  return {
    issueId: sanitize(raw["issue id"] ?? raw.issueid ?? raw.id),
    project: sanitize(raw.project),
    owner: sanitize(raw.owner),
    priority: sanitize(raw.priority),
    status: sanitize(raw.status ?? raw.rag),
    blocker: sanitize(raw.blocker ?? raw["key blocker"]),
    dueDate: sanitize(raw["due date"] ?? raw["target date"]),
    lastUpdate: sanitize(raw["last update"])
  };
}

function getKpis(issues) {
  const activeProjects = issues.length;
  const onTrack = issues.filter((issue) => issue.status.toLowerCase() === "on track").length;
  const atRisk = issues.filter((issue) => issue.status.toLowerCase() === "at risk").length;
  const delayed = issues.filter((issue) => issue.status.toLowerCase() === "delayed").length;

  return { activeProjects, onTrack, atRisk, delayed };
}

function renderKpis({ activeProjects, onTrack, atRisk, delayed }) {
  document.getElementById("kpi-active-projects").textContent = formatMetricCount(activeProjects);
  document.getElementById("kpi-on-track").textContent = formatMetricCount(onTrack);
  document.getElementById("kpi-at-risk").textContent = formatMetricCount(atRisk);
  document.getElementById("kpi-delayed").textContent = formatMetricCount(delayed);
}

function renderTable(issues) {
  const tableBody = document.getElementById("status-table-body");

  tableBody.innerHTML = issues
    .map((issue) => {
      const statusClass = toClassName(issue.status);

      return `
        <tr>
          <td class="issue-title">${escapeHtml(issue.issueId)}</td>
          <td>${escapeHtml(issue.project)}</td>
          <td>${escapeHtml(issue.owner)}</td>
          <td>${escapeHtml(issue.priority)}</td>
          <td><span class="status-indicator ${escapeHtml(statusClass)}">${escapeHtml(issue.status)}</span></td>
          <td>${escapeHtml(issue.blocker)}</td>
          <td>${escapeHtml(issue.dueDate)}</td>
          <td>${escapeHtml(issue.lastUpdate)}</td>
        </tr>
      `;
    })
    .join("");
}

function buildSheetViewLink(source) {
  if (!source.sheetId) {
    return "";
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/edit`;
}

function buildStatusApi(source) {
  if (!source.sheetId) {
    throw new Error("Google Sheet ID is missing.");
  }

  const params = new URLSearchParams({ tqx: "out:json" });
  if (source.gid) {
    params.set("gid", source.gid);
  } else if (source.sheetTab) {
    params.set("sheet", source.sheetTab);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/gviz/tq?${params.toString()}`;
}

function isAppsScriptUrl(url) {
  return url.hostname === "script.google.com" || url.hostname.endsWith(".script.google.com");
}

function parseSourceUrl(rawUrl) {
  const url = new URL(rawUrl);

  if (isAppsScriptUrl(url)) {
    return {
      mode: "appsScript",
      appsScriptUrl: url.toString(),
      sheetId: "",
      sheetTab: "",
      gid: ""
    };
  }

  const idMatch = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

  if (!idMatch?.[1]) {
    throw new Error("Provide a valid Google Apps Script Web App URL or Google Sheets URL.");
  }

  return {
    mode: "sheet",
    appsScriptUrl: "",
    sheetId: idMatch[1],
    sheetTab: sanitizeOptional(url.searchParams.get("sheet")),
    gid: sanitizeOptional(url.searchParams.get("gid"))
  };
}

function readSavedSource() {
  try {
    const raw = localStorage.getItem(LOCAL_SOURCE_KEY);
    if (!raw) {
      return { ...DEFAULT_SOURCE };
    }

    const parsed = JSON.parse(raw);
    return {
      mode: parsed.mode === "appsScript" ? "appsScript" : "sheet",
      appsScriptUrl: sanitizeOptional(parsed.appsScriptUrl),
      sheetId: sanitizeOptional(parsed.sheetId),
      sheetTab: sanitizeOptional(parsed.sheetTab),
      gid: sanitizeOptional(parsed.gid)
    };
  } catch {
    return { ...DEFAULT_SOURCE };
  }
}

function saveSource(source) {
  localStorage.setItem(LOCAL_SOURCE_KEY, JSON.stringify(source));
}

function setFeedback(message) {
  document.getElementById("sheet-source-feedback").textContent = message;
}

function updateSourceUi(source) {
  const input = document.getElementById("sheet-source-input");
  const label = document.getElementById("status-source-label");

  if (source.mode === "appsScript" && source.appsScriptUrl) {
    input.value = source.appsScriptUrl;
    label.innerHTML = "Source: <strong>Google Apps Script Web App</strong>";
    return;
  }

  input.value = buildSheetViewLink(source) || "";
  label.innerHTML = "Source: <strong>Google Sheets (gviz)</strong>";
}

async function loadFromAppsScript(source) {
  if (!source.appsScriptUrl) {
    throw new Error("Apps Script URL is missing.");
  }

  const response = await fetch(source.appsScriptUrl, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Apps Script API returned ${response.status}`);
  }

  const payload = await response.json();
  const rows = Array.isArray(payload) ? payload : payload.rows;

  if (!Array.isArray(rows)) {
    throw new Error("Apps Script response must be an array or { rows: [] } JSON payload.");
  }

  const issues = rows.map(normalizeIssue).filter((issue) => issue.project !== "—");

  if (!issues.length) {
    throw new Error("No issue tracking rows found from Apps Script response.");
  }

  return issues;
}

async function loadFromSheet(source) {
  const response = await fetch(buildStatusApi(source), { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Status API returned ${response.status}`);
  }

  const text = await response.text();
  const gvizPayload = parseGvizResponse(text);
  const rows = gvizToRows(gvizPayload);
  const issues = rows.map(normalizeIssue).filter((issue) => issue.project !== "—");

  if (!issues.length) {
    throw new Error("No issue tracking rows found from Google Sheet.");
  }

  return issues;
}

async function loadProjectStatus(source) {
  try {
    const issues = source.mode === "appsScript"
      ? await loadFromAppsScript(source)
      : await loadFromSheet(source);

    renderKpis(getKpis(issues));
    renderTable(issues);
  } catch (error) {
    console.warn("Unable to load live project status. Showing fallback data.", error);
    setFeedback("Could not read your live source yet. Showing built-in issue tracking rows.");
    renderKpis(getKpis(fallbackIssues));
    renderTable(fallbackIssues);
  }
}

function initSourceControls(source) {
  const form = document.getElementById("sheet-source-form");
  const resetButton = document.getElementById("sheet-source-reset");

  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    try {
      const inputValue = document.getElementById("sheet-source-input").value.trim();
      const parsedSource = parseSourceUrl(inputValue);
      saveSource(parsedSource);
      updateSourceUi(parsedSource);
      setFeedback(
        parsedSource.mode === "appsScript"
          ? "Apps Script source saved. Loading issue tracking table..."
          : "Sheet source saved. Loading issue tracking table..."
      );
      await loadProjectStatus(parsedSource);
    } catch (error) {
      setFeedback(error.message || "Please provide a valid Google URL.");
    }
  });

  resetButton.addEventListener("click", async () => {
    saveSource(DEFAULT_SOURCE);
    updateSourceUi(DEFAULT_SOURCE);
    setFeedback("Reset to default sheet source.");
    await loadProjectStatus(DEFAULT_SOURCE);
  });

  updateSourceUi(source);
}

const activeSource = readSavedSource();
initSourceControls(activeSource);
loadProjectStatus(activeSource);
