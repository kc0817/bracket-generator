/**
 * Shared app state and UI: file upload, team parsing, smooth scroll navigation.
 */

const AppState = {
  teams: [],
  fileName: null,
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

function setTeams(teams, fileName) {
  AppState.teams = teams;
  AppState.fileName = fileName;
  updateFileStatus();
  document.dispatchEvent(new CustomEvent("teams-updated", { detail: { teams, fileName } }));
}

function clearTeams() {
  AppState.teams = [];
  AppState.fileName = null;
  document.getElementById("team-file").value = "";
  updateFileStatus();
  document.dispatchEvent(new CustomEvent("teams-updated", { detail: { teams: [], fileName: null } }));
}

function updateFileStatus() {
  const statusEl = document.getElementById("file-status");
  const textEl = document.getElementById("file-status-text");

  if (AppState.teams.length === 0) {
    statusEl.hidden = true;
    return;
  }

  statusEl.hidden = false;
  textEl.textContent = `${AppState.fileName} — ${AppState.teams.length} team${AppState.teams.length === 1 ? "" : "s"} loaded`;
}

async function handleFile(file) {
  if (!file) return;

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

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const teams = parseTeamsFromWorkbook(workbook);

  if (teams.length < 2) {
    alert("We need at least 2 team names in column A. Check your file and try again.");
    return;
  }

  setTeams(teams, file.name);
}

function initUpload() {
  const zone = document.getElementById("upload-zone");
  const input = document.getElementById("team-file");
  const clearBtn = document.getElementById("clear-file");

  zone.addEventListener("click", () => input.click());

  input.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
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
    if (file) handleFile(file);
  });

  clearBtn.addEventListener("click", clearTeams);
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

document.addEventListener("DOMContentLoaded", () => {
  initUpload();
  initScrollNav();
});
