
const DEFAULT_SOURCE = {
  sheetId: "1BHTM5Jx7xAyaRnzuyjD38-Iz5mlQddkA8s2J2ZatIXI",
  sheetTab: "Portfolio Status",
  gid: ""
};
const LOCAL_SOURCE_KEY = "projectStatusSheetSource";

const fallbackProjects = [
  {
    project: "Township Expansion - North",
    owner: "PMO Team A",
    phase: "Execution",
    progress: 72,
    rag: "On Track",
    milestone: "Site utility handover",
    targetDate: "Feb 20, 2026",
    lastUpdate: "Feb 08, 2026"
  },
  {
    project: "Digital Sales Portal",
    owner: "PMO Team B",
    phase: "UAT",
    progress: 58,
    rag: "At Risk",
    milestone: "UAT signoff",
    targetDate: "Mar 03, 2026",
    lastUpdate: "Feb 09, 2026"
  },
  {
    project: "CRM Data Cleansing",
    owner: "PMO Team C",
    phase: "Planning",
    progress: 32,
    rag: "Delayed",
    milestone: "Baseline approval",
    targetDate: "Feb 27, 2026",
    lastUpdate: "Feb 07, 2026"
  },
  {
    project: "Client Self-Service Enhancements",
    owner: "PMO Team D",
    phase: "Deployment",
    progress: 94,
    rag: "Completed",
    milestone: "Post-launch report",
    targetDate: "Feb 14, 2026",
    lastUpdate: "Feb 10, 2026"
  }
];

function toClassName(rag) {
  return String(rag)
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

function normalizeProject(raw) {
  const progress = Number.parseInt(raw.progress ?? raw["progress %"] ?? "", 10);

  return {
    project: sanitize(raw.project),
    owner: sanitize(raw.owner),
    phase: sanitize(raw.phase),
    progress: Number.isFinite(progress) ? progress : null,
    rag: sanitize(raw.rag ?? raw.status),
    milestone: sanitize(raw["next milestone"] ?? raw.milestone),
    targetDate: sanitize(raw["target date"] ?? raw.targetdate),
    lastUpdate: sanitize(raw["last update"] ?? raw.lastupdate)
  };
}

function getKpis(projects) {
  const activeProjects = projects.length;
  const onTrack = projects.filter((project) => project.rag.toLowerCase() === "on track").length;
  const atRisk = projects.filter((project) => project.rag.toLowerCase() === "at risk").length;
  const delayed = projects.filter((project) => project.rag.toLowerCase() === "delayed").length;

  return { activeProjects, onTrack, atRisk, delayed };
}

function renderKpis({ activeProjects, onTrack, atRisk, delayed }) {
  document.getElementById("kpi-active-projects").textContent = formatMetricCount(activeProjects);
  document.getElementById("kpi-on-track").textContent = formatMetricCount(onTrack);
  document.getElementById("kpi-at-risk").textContent = formatMetricCount(atRisk);
  document.getElementById("kpi-delayed").textContent = formatMetricCount(delayed);
}

function renderTable(projects) {
  const tableBody = document.getElementById("status-table-body");

  tableBody.innerHTML = projects
    .map((project) => {
      const progressValue = Number.isFinite(project.progress) ? `${project.progress}%` : "—";
      const ragClass = toClassName(project.rag);

      return `
        <tr>
          <td class="issue-title">${escapeHtml(project.project)}</td>
          <td>${escapeHtml(project.owner)}</td>
          <td>${escapeHtml(project.phase)}</td>
          <td>${escapeHtml(progressValue)}</td>
          <td><span class="status-indicator ${escapeHtml(ragClass)}">${escapeHtml(project.rag)}</span></td>
          <td>${escapeHtml(project.milestone)}</td>
          <td>${escapeHtml(project.targetDate)}</td>
          <td>${escapeHtml(project.lastUpdate)}</td>
        </tr>
      `;
    })
    .join("");
}

function buildSheetViewLink(source) {
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/edit`;
}

function buildStatusApi(source) {
  const params = new URLSearchParams({ tqx: "out:json" });
  if (source.gid) {
    params.set("gid", source.gid);
  } else {
    params.set("sheet", source.sheetTab);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/gviz/tq?${params.toString()}`;
}

function readSavedSource() {
  try {
    const raw = localStorage.getItem(LOCAL_SOURCE_KEY);
    if (!raw) {
      return { ...DEFAULT_SOURCE };
    }

    const parsed = JSON.parse(raw);
    return {
      sheetId: sanitize(parsed.sheetId, DEFAULT_SOURCE.sheetId),
      sheetTab: sanitize(parsed.sheetTab, DEFAULT_SOURCE.sheetTab),
      gid: sanitize(parsed.gid, "") === "—" ? "" : sanitize(parsed.gid, "")
    };
  } catch {
    return { ...DEFAULT_SOURCE };
  }
}

function saveSource(source) {
  localStorage.setItem(LOCAL_SOURCE_KEY, JSON.stringify(source));
}

function parseGoogleSheetUrl(rawUrl) {
  const url = new URL(rawUrl);
  const idMatch = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

  if (!idMatch?.[1]) {
    throw new Error("Google Sheets link is missing a spreadsheet ID.");
  }

  const gid = url.searchParams.get("gid") ?? "";
  const sheetTab = url.searchParams.get("sheet") ?? DEFAULT_SOURCE.sheetTab;

  return {
    sheetId: idMatch[1],
    sheetTab,
    gid
  };
}

function updateSourceUi(source) {
  const input = document.getElementById("sheet-source-input");
  const masterSheetLink = document.getElementById("master-sheet-link");

  input.value = buildSheetViewLink(source);
  masterSheetLink.href = buildSheetViewLink(source);
}

function setFeedback(message) {
  document.getElementById("sheet-source-feedback").textContent = message;
}

async function loadProjectStatus(source) {
  try {
    const response = await fetch(buildStatusApi(source), { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Status API returned ${response.status}`);
    }

    const text = await response.text();
    const gvizPayload = parseGvizResponse(text);
    const rows = gvizToRows(gvizPayload);
    const projects = rows.map(normalizeProject).filter((project) => project.project !== "—");

    if (!projects.length) {
      throw new Error("No rows from Google Sheet");
    }

    renderKpis(getKpis(projects));
    renderTable(projects);
  } catch (error) {
    console.warn("Unable to load live project status. Showing fallback data.", error);
    renderKpis(getKpis(fallbackProjects));
    renderTable(fallbackProjects);
  }
}

function initSourceControls(source) {
  const form = document.getElementById("sheet-source-form");
  const resetButton = document.getElementById("sheet-source-reset");

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    

    try {
      const inputValue = document.getElementById("sheet-source-input").value.trim();
      const parsedSource = parseGoogleSheetUrl(inputValue);
      saveSource(parsedSource);
      updateSourceUi(parsedSource);
      setFeedback("Sheet source saved. Reloading project status from your provided link...");
      await loadProjectStatus(parsedSource);
    } catch (error) {
      setFeedback(error.message || "Please provide a valid Google Sheets URL.");
    }
  });

  resetButton.addEventListener("click", async () => {
    saveSource(DEFAULT_SOURCE);
    updateSourceUi(DEFAULT_SOURCE);
    setFeedback("Reset to default master status sheet.");
    await loadProjectStatus(DEFAULT_SOURCE);
  });

  updateSourceUi(source);
}

const activeSource = readSavedSource();
initSourceControls(activeSource);
loadProjectStatus(activeSource);