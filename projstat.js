const LOCAL_SOURCE_KEY = "projectStatusOverviewSource";


function sanitize(value, fallback = "—") {
  const text = String(value ?? "").trim();
  return text || fallback;
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
    throw new Error("Invalid Google Sheet response.");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function gvizToRows(payload) {
  const columns = (payload.table?.cols ?? []).map((column, index) => {
    const label = normalizeHeader(column.label);
    return label || `column ${index + 1}`;
  });

  return (payload.table?.rows ?? []).map((row) => {
    const mapped = {};

    columns.forEach((columnName, index) => {
      const cell = row.c?.[index];
      mapped[columnName] = cell?.f ?? cell?.v ?? "";
    });

    return mapped;
  });
}

function pickField(row, keys) {
  for (const key of keys) {
    const value = String(row[key] ?? "").trim();
    if (value) {
      return value;
    }
  }

  return "";
}

function normalizeProject(row) {
  return {
    project: sanitize(pickField(row, ["project", "project name", "column 1"])),
    owner: sanitize(pickField(row, ["owner", "lead", "column 2"])),
    phase: sanitize(pickField(row, ["phase", "stage", "column 3"])),
    progress: sanitize(pickField(row, ["progress", "progress %", "column 4"])),
    rag: sanitize(pickField(row, ["rag", "rag status", "health", "status", "column 5"])),
    milestone: sanitize(pickField(row, ["next milestone", "milestone", "column 6"])),
    targetDate: sanitize(pickField(row, ["target date", "due date", "column 7"])),
    lastUpdate: sanitize(pickField(row, ["last update", "updated", "column 8"]))
  };
}

function isGoogleUrl(url) {
  return [
    "docs.google.com",
    "drive.google.com",
    "script.google.com",
    "forms.google.com",
    "lookerstudio.google.com"
  ].some((host) => url.hostname === host || url.hostname.endsWith(`.${host}`));
}

function isLookerStudioUrl(url) {
  return url.hostname === "lookerstudio.google.com" || url.hostname.endsWith(".lookerstudio.google.com");
}

function parseGoogleLink(rawValue) {
  const url = new URL(rawValue);

  if (!isGoogleUrl(url)) {
    throw new Error("Please enter a valid Google link (Sheets, Docs, Drive, Forms, or Apps Script).");
  }

  return url;
}

function getHashParams(url) {
  return new URLSearchParams(String(url.hash ?? "").replace(/^#/, ""));
}

function getSheetSource(url) {
  const idMatch = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch?.[1]) {
    return null;
  }

  const hashParams = getHashParams(url);
  return {
    sheetId: idMatch[1],
    gid: String(url.searchParams.get("gid") || hashParams.get("gid") || "").trim(),
    sheet: String(url.searchParams.get("sheet") || hashParams.get("sheet") || "").trim()
  };
}

function buildViewerUrl(url) {
  if (url.hostname.includes("docs.google.com") && url.pathname.includes("/spreadsheets/d/")) {
    const source = getSheetSource(url);
    if (!source) return url.toString();

    const gidPart = source.gid ? `#gid=${encodeURIComponent(source.gid)}` : "";
    return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/preview${gidPart}`;
  }

  if (url.hostname.includes("drive.google.com")) {
    const match = url.pathname.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
    if (match?.[1]) {
      return `https://drive.google.com/file/d/${encodeURIComponent(match[1])}/preview`;
    }
  }

  if (isLookerStudioUrl(url)) {
    const parts = url.pathname.split("/").filter(Boolean);
    if (parts[0] === "u" && parts.length > 2) {
      parts.splice(0, 2);
    }

    if (parts[0] === "reporting" && parts[1]) {
      const pageId = parts[3] || "";
      const embedPath = pageId
        ? `/embed/reporting/${encodeURIComponent(parts[1])}/page/${encodeURIComponent(pageId)}`
        : `/embed/reporting/${encodeURIComponent(parts[1])}`;
      return `https://lookerstudio.google.com${embedPath}`;
    }
  }

  return url.toString();
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

function classifyRag(ragValue) {
  const normalized = String(ragValue ?? "").trim().toLowerCase();

  if (!normalized || normalized === "—") {
    return "unknown";
  }

  const onTrackTerms = ["on track", "ontrack", "green", "ok", "healthy", "complete", "completed", "done"];
  const atRiskTerms = ["at risk", "atrisk", "amber", "yellow", "warning"];
  const delayedTerms = ["delay", "delayed", "red", "critical", "blocked", "off track", "offtrack"];

  if (onTrackTerms.some((term) => normalized.includes(term))) {
    return "onTrack";
  }

  if (atRiskTerms.some((term) => normalized.includes(term))) {
    return "atRisk";
  }

  if (delayedTerms.some((term) => normalized.includes(term))) {
    return "delayed";
  }

  if (normalized === "g") return "onTrack";
  if (normalized === "a") return "atRisk";
  if (normalized === "r") return "delayed";

  return "unknown";
}

function getKpis(projects) {
  const totals = {
    active: projects.length,
    onTrack: 0,
    atRisk: 0,
    delayed: 0
  };

  projects.forEach((project) => {
    const bucket = classifyRag(project.rag);
    if (bucket === "onTrack") {
      totals.onTrack += 1;
    } else if (bucket === "atRisk") {
      totals.atRisk += 1;
    } else if (bucket === "delayed") {
      totals.delayed += 1;
    }
  });

  return totals;
}

function renderKpis(kpis) {
  document.getElementById("kpi-active-projects").textContent = String(kpis.active).padStart(2, "0");
  document.getElementById("kpi-on-track").textContent = String(kpis.onTrack).padStart(2, "0");
  document.getElementById("kpi-at-risk").textContent = String(kpis.atRisk).padStart(2, "0");
  document.getElementById("kpi-delayed").textContent = String(kpis.delayed).padStart(2, "0");
}

function renderProjects(projects) {
  const body = document.getElementById("status-table-body");
  body.innerHTML = projects.map((project) => `
    <tr>
      <td>${sanitize(project.project)}</td>
      <td>${sanitize(project.owner)}</td>
      <td>${sanitize(project.phase)}</td>
      <td>${sanitize(project.progress)}</td>
      <td>${sanitize(project.rag)}</td>
      <td>${sanitize(project.milestone)}</td>
      <td>${sanitize(project.targetDate)}</td>
      <td>${sanitize(project.lastUpdate)}</td>
    </tr>
  `).join("");
}

function setFeedback(message) {
  document.getElementById("sheet-source-feedback").textContent = message;
}


function updateScrollCue() {
  const cue = document.getElementById("status-scroll-cue");
  const tablePanel = document.getElementById("project-status-table");

  if (!cue || !tablePanel) {
    return;
  }

  const rect = tablePanel.getBoundingClientRect();
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  const isVisible = rect.top <= viewportHeight * 0.7;
  cue.classList.toggle("hidden", isVisible);
}

function bindScrollCue() {
  const cue = document.getElementById("status-scroll-cue");
  const tablePanel = document.getElementById("project-status-table");

  if (!cue || !tablePanel) {
    return;
  }

  const jumpLink = cue.querySelector(".status-scroll-link");
  if (jumpLink) {
    jumpLink.addEventListener("click", (event) => {
      event.preventDefault();
      tablePanel.scrollIntoView({ behavior: "smooth", block: "start" });
      history.replaceState(null, "", "#project-status-table");
    });
  }

  window.addEventListener("scroll", updateScrollCue, { passive: true });
  window.addEventListener("resize", updateScrollCue);

  updateScrollCue();
}

function setViewer(urlString) {
  const viewer = document.getElementById("google-viewer");
  const viewerWrap = document.getElementById("google-viewer-wrap");
  viewer.src = urlString;
  viewerWrap.hidden = false;
}

function clearViewer() {
  const viewer = document.getElementById("google-viewer");
  const viewerWrap = document.getElementById("google-viewer-wrap");
  viewer.src = "about:blank";
  viewerWrap.hidden = true;
}

async function loadProjectsFromSheet(url) {
  const source = getSheetSource(url);
  if (!source) {
    renderKpis({ active: 0, onTrack: 0, atRisk: 0, delayed: 0 });
    renderProjects([]);

    if (isLookerStudioUrl(url)) {
      setFeedback("Looker Studio dashboard loaded in preview. For KPI counting, use a Google Sheets link with a RAG column.");
    } else {
      setFeedback("Google link opened. To auto-count KPIs by RAG, use a Google Sheets project status link.");
    }

    return;
  }

  try {
    const response = await fetch(buildSheetApiUrl(source), { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Google Sheet API returned ${response.status}`);
    }

    const payload = parseGvizResponse(await response.text());
    const projects = gvizToRows(payload)
      .map(normalizeProject)
      .filter((project) => project.project !== "—");

    if (!projects.length) {
      throw new Error("No rows found.");
    }

    renderKpis(getKpis(projects));
    renderProjects(projects);
    setFeedback("Project status loaded successfully.");
  } catch {
    renderKpis({ active: 0, onTrack: 0, atRisk: 0, delayed: 0 });
    renderProjects([]);
    setFeedback("Could not read sheet rows. Make sure the sheet is shared to Anyone with the link (Viewer) and has a RAG column.");
  }
}

async function handleSubmit(event) {
  event.preventDefault();

  const input = document.getElementById("sheet-source-input");

  try {
    const url = parseGoogleLink(input.value.trim());
    localStorage.setItem(LOCAL_SOURCE_KEY, url.toString());
    setViewer(buildViewerUrl(url));
    await loadProjectsFromSheet(url);
  } catch (error) {
    setFeedback(error.message || "Please enter a valid Google link.");
  }
}

function resetView() {
  localStorage.removeItem(LOCAL_SOURCE_KEY);
  clearViewer();
  renderKpis({ active: 0, onTrack: 0, atRisk: 0, delayed: 0 });
  renderProjects([]);
  document.getElementById("sheet-source-input").value = "";
  setFeedback("Paste a Google link then click View Project.");
}

function init() {
  document.getElementById("sheet-source-form").addEventListener("submit", handleSubmit);
  document.getElementById("sheet-source-reset").addEventListener("click", resetView);

  const saved = localStorage.getItem(LOCAL_SOURCE_KEY);
  if (saved) {
    try {
      document.getElementById("sheet-source-input").value = saved;
      const url = parseGoogleLink(saved);
      setViewer(buildViewerUrl(url));
      loadProjectsFromSheet(url);
    } catch {
      localStorage.removeItem(LOCAL_SOURCE_KEY);
      setFeedback("Paste a Google link then click View Project.");
    }
  } else {
    setFeedback("Paste a Google link then click View Project.");
    renderProjects([]);
  }

  bindScrollCue();
}

init(); 

