/**
 * Shared app state and UI: file upload, manual entry, team parsing, smooth scroll navigation.
 */

const AppState = {
  qual: { teams: [], sourceLabel: null },
  playoff: { teams: [], sourceLabel: null },
};

const HEADER_KEYWORDS = new Set([
  "team", "teams", "name", "names", "contestant", "contestants",
  "player", "players", "school", "schools", "participant", "participants",
]);

function parseTeamsFromWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const teams = [];
  for (let i = 0; i < rows.length; i++) {
    const cell = rows[i][0];
    const name = String(cell ?? "").trim();
    if (!name) continue;

    if (i === 0 && HEADER_KEYWORDS.has(name.toLowerCase())) continue;
    teams.push(name);
  }
  return teams;
}

function parseTeamsFromText(text) {
  const lines = text.split(/\r?\n/);
  const teams = [];

  for (let i = 0; i < lines.length; i++) {
    const name = lines[i].trim();
    if (!name) continue;

    if (i === 0 && HEADER_KEYWORDS.has(name.toLowerCase())) continue;
    teams.push(name);
  }

  return teams;
}

function validateTeamCount(teams, section) {
  if (teams.length < 2) {
    if (section !== "qual") {
      alert("We need at least 2 team names. Check your input and try again.");
    }
    return false;
  }
  return true;
}

function setTeams(section, teams, sourceLabel) {
  AppState[section].teams = teams;
  AppState[section].sourceLabel = sourceLabel;
  updateFileStatus(section);
  document.dispatchEvent(
    new CustomEvent(`${section}-teams-updated`, { detail: { teams, sourceLabel } })
  );
}

function clearTeams(section) {
  AppState[section].teams = [];
  AppState[section].sourceLabel = null;

  const fileInput = document.getElementById(`${section}-team-file`);
  const textarea = document.getElementById(`${section}-team-text`);
  if (fileInput) fileInput.value = "";
  if (textarea) textarea.value = "";

  if (section === "qual") resetQualUploadZone();

  updateFileStatus(section);
  document.dispatchEvent(
    new CustomEvent(`${section}-teams-updated`, { detail: { teams: [], sourceLabel: null } })
  );
}

function updateFileStatus(section) {
  const statusEl = document.getElementById(`${section}-file-status`);
  const textEl = document.getElementById(`${section}-file-status-text`);
  if (!statusEl || !textEl) return;

  const { teams, sourceLabel } = AppState[section];

  if (teams.length === 0) {
    statusEl.hidden = true;
    return;
  }

  statusEl.hidden = false;
  const label = sourceLabel || "Manual entry";
  textEl.textContent = `${label} — ${teams.length} team${teams.length === 1 ? "" : "s"} loaded`;
}

async function handleFile(section, file) {
  if (!file) return;

  if (section === "qual") {
    await stageQualFile(file);
    return;
  }

  const validTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "text/csv",
    "application/vnd.oasis.opendocument.spreadsheet",
  ];
  const ext = file.name.split(".").pop().toLowerCase();
  const validExts = ["xlsx", "xls", "csv", "ods"];

  if (!validExts.includes(ext) && !validTypes.includes(file.type)) {
    alert("Please upload an Excel (.xlsx, .xls), CSV, or ODS file.");
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const teams = parseTeamsFromWorkbook(workbook);

    if (!validateTeamCount(teams, section)) {
      return;
    }

    const textarea = document.getElementById(`${section}-team-text`);
    if (textarea) textarea.value = teams.join("\n");

    setTeams(section, teams, file.name);
  } catch {
    alert("Could not read that file. Check the format and try again.");
  }
}

function handleManualEntry(section) {
  const textarea = document.getElementById(`${section}-team-text`);
  if (!textarea) return;

  const teams = parseTeamsFromText(textarea.value);
  if (!validateTeamCount(teams, section)) return;

  setTeams(section, teams, "Manual entry");
}

function initTeamInput(section) {
  const zone = document.getElementById(`${section}-upload-zone`);
  const input = document.getElementById(`${section}-team-file`);
  const clearBtn = document.getElementById(`${section}-clear-teams`);
  const applyBtn = document.getElementById(`${section}-apply-teams`);

  if (!zone || !input) return;

  zone.addEventListener("click", (e) => {
    if (e.target.closest(".upload-zone-clear")) return;
    input.click();
  });

  input.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleFile(section, file);
  });

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });

  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));

  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file) handleFile(section, file);
  });

  clearBtn?.addEventListener("click", () => clearTeams(section));
  applyBtn?.addEventListener("click", () => handleManualEntry(section));
}

function renderTeamsPreview(section, teams) {
  const preview = document.getElementById(`${section}-teams-preview`);
  const empty = document.getElementById(`${section}-empty-state`);
  const list = document.getElementById(`${section}-team-list`);
  const count = document.getElementById(`${section}-team-count`);

  if (!preview || !empty || !list || !count) return;

  if (teams.length === 0) {
    preview.hidden = true;
    empty.hidden = false;
    return;
  }

  empty.hidden = true;
  preview.hidden = false;
  count.textContent = `(${teams.length})`;
  list.innerHTML = teams
    .map((t) => `<li>${escapeHtml(t)}</li>`)
    .join("");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function initScrollNav() {
  document.querySelectorAll("[data-scroll-to]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-scroll-to");
      const target = document.getElementById(targetId);
      if (target) {
        target.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  });
}

function initPlayoffPreview() {
  document.addEventListener("playoff-teams-updated", (e) => {
    renderTeamsPreview("playoff", e.detail.teams);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initTeamInput("qual");
  initTeamInput("playoff");
  initPlayoffPreview();
  initScrollNav();
});
