/**
 * Qual bracket generation: random group assignment + round-robin scheduling.
 */

function shuffle(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function hashTeamList(teams) {
  const key = [...teams].sort().join("\n");
  let hash = 2166136261;
  for (let i = 0; i < key.length; i++) {
    hash ^= key.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function shuffleWithSeed(array, seed) {
  const arr = [...array];
  let state = seed >>> 0;
  for (let i = arr.length - 1; i > 0; i--) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    const j = state % (i + 1);
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function chunkIntoGroups(teams, groupSize, seed) {
  const shuffled = seed !== undefined ? shuffleWithSeed(teams, seed) : shuffle(teams);
  const groups = [];
  for (let i = 0; i < shuffled.length; i += groupSize) {
    groups.push(shuffled.slice(i, i + groupSize));
  }
  return groups;
}

function groupLabel(index) {
  let label = "";
  let n = index;
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

/**
 * Circle-method round robin. Returns array of rounds; each round is an array of [teamA, teamB] pairs.
 * Odd-sized groups get a null BYE slot so one team rests each round.
 */
function generateRoundRobinRounds(teams) {
  const participants = [...teams];
  if (participants.length < 2) return [];

  if (participants.length % 2 === 1) {
    participants.push(null);
  }

  const n = participants.length;
  const roundCount = n - 1;
  const rounds = [];

  let rotation = [...participants];

  for (let r = 0; r < roundCount; r++) {
    const matches = [];
    for (let i = 0; i < n / 2; i++) {
      const a = rotation[i];
      const b = rotation[n - 1 - i];
      if (a !== null && b !== null) {
        matches.push([a, b]);
      }
    }
    rounds.push(matches);

    const fixed = rotation[0];
    const rest = rotation.slice(1);
    rest.unshift(rest.pop());
    rotation = [fixed, ...rest];
  }

  return rounds;
}

/**
 * Build schedule with exactly `gamesPerTeam` matches per team.
 * Uses round-robin rounds first; cycles with rematches if more games are requested.
 */
function buildGroupSchedule(teams, gamesPerTeam) {
  if (teams.length < 2) {
    return { rounds: [], warning: "Groups need at least 2 teams to schedule matches." };
  }

  const maxWithoutRematch = teams.length - 1;
  const baseRounds = generateRoundRobinRounds(teams);
  const selectedRounds = [];
  let roundIndex = 0;
  let cycle = 1;

  while (selectedRounds.length < gamesPerTeam && cycle <= 10) {
    if (baseRounds.length === 0) break;

    const sourceRound = baseRounds[roundIndex % baseRounds.length];
    selectedRounds.push(
      sourceRound.map(([a, b]) => [`${a}`, `${b}`])
    );
    roundIndex++;
    if (roundIndex % baseRounds.length === 0) cycle++;
  }

  const trimmedRounds = selectedRounds.slice(0, gamesPerTeam);

  let warning = null;
  if (gamesPerTeam > maxWithoutRematch) {
    warning = `Requested ${gamesPerTeam} games per team, but a group of ${teams.length} can only play each opponent once (${maxWithoutRematch} games) without rematches. Extra rounds include rematches.`;
  }

  return { rounds: trimmedRounds, warning };
}

function generateQualBracket(teams, groupSize, gamesPerTeam, randomShuffle) {
  if (teams.length < 2) {
    throw new Error("At least 2 teams are required.");
  }
  if (groupSize < 2) {
    throw new Error("Group size must be at least 2.");
  }
  if (gamesPerTeam < 1) {
    throw new Error("Games per team must be at least 1.");
  }

  const groupSeed = randomShuffle ? undefined : hashTeamList(teams);
  const rawGroups = chunkIntoGroups(teams, groupSize, groupSeed);
  const groups = rawGroups.map((groupTeams, i) => {
    const label = groupLabel(i);
    const { rounds, warning } = buildGroupSchedule(groupTeams, gamesPerTeam);
    return { label, teams: groupTeams, rounds, warning };
  });

  return groups;
}

function buildExcelWorkbook(groups) {
  const groupsRows = [["Group", "Slot", "Team"]];
  for (const group of groups) {
    group.teams.forEach((team, idx) => {
      groupsRows.push([group.label, idx + 1, team]);
    });
    groupsRows.push([]);
  }

  const scheduleRows = [["Group", "Round", "Match", "Team 1", "Team 2"]];
  for (const group of groups) {
    group.rounds.forEach((matches, roundIdx) => {
      matches.forEach((match, matchIdx) => {
        scheduleRows.push([
          group.label,
          roundIdx + 1,
          matchIdx + 1,
          match[0],
          match[1],
        ]);
      });
    });
    scheduleRows.push([]);
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(groupsRows), "Groups");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(scheduleRows), "Schedule");
  return wb;
}

function downloadQualBracket(groups) {
  const wb = buildExcelWorkbook(groups);
  XLSX.writeFile(wb, "qual-bracket.xlsx");
}

function updateScheduleNote(groupSize, gamesPerTeam) {
  const noteEl = document.getElementById("schedule-note");
  const maxGames = Math.max(1, groupSize - 1);

  if (gamesPerTeam > maxGames) {
    noteEl.hidden = false;
    noteEl.textContent =
      `With ${groupSize} teams per group, each team can play every opponent once (${maxGames} games). ` +
      `You asked for ${gamesPerTeam}, so some matchups will repeat.`;
  } else {
    noteEl.hidden = true;
  }
}

let qualReadFeedbackTimer = null;
let qualPendingFile = null;
let qualPendingFileTeams = [];
let qualPendingFileValid = false;
let qualPreviewGroups = null;
let qualHasGenerated = false;

const QUAL_READ_SUCCESS_ICON =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M20 6L9 17l-5-5"/></svg>';
const QUAL_READ_ERROR_ICON =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M18 6L6 18M6 6l12 12"/></svg>';

function setQualUploadFilename(fileName) {
  const zone = document.getElementById("qual-upload-zone");
  const prompt = document.getElementById("qual-upload-prompt");
  const fileView = document.getElementById("qual-upload-file");
  const fileNameEl = document.getElementById("qual-upload-filename");

  if (!zone || !prompt || !fileView || !fileNameEl) return;

  fileNameEl.textContent = fileName;
  prompt.hidden = true;
  fileView.hidden = false;
  zone.classList.add("upload-zone--has-file");

  const clearBtn = document.getElementById("qual-upload-clear");
  if (clearBtn) clearBtn.hidden = false;
}

function clearQualUploadFile() {
  const fileInput = document.getElementById("qual-team-file");
  if (fileInput) fileInput.value = "";
  resetQualUploadZone();
  updateQualSaveButtons();
  invalidateQualGeneration();
}

function resetQualPendingFile() {
  qualPendingFile = null;
  qualPendingFileTeams = [];
  qualPendingFileValid = false;
}

function resetQualUploadZone() {
  const zone = document.getElementById("qual-upload-zone");
  const prompt = document.getElementById("qual-upload-prompt");
  const fileView = document.getElementById("qual-upload-file");
  const feedback = document.getElementById("qual-read-feedback");

  if (!zone || !prompt || !fileView) return;

  prompt.hidden = false;
  fileView.hidden = true;
  zone.classList.remove("upload-zone--has-file");
  resetQualPendingFile();

  const clearBtn = document.getElementById("qual-upload-clear");
  if (clearBtn) clearBtn.hidden = true;

  if (feedback) {
    clearTimeout(qualReadFeedbackTimer);
    feedback.hidden = true;
    feedback.className = "upload-read-feedback";
    feedback.innerHTML = "";
  }
}

function showQualReadFeedback(success) {
  const feedback = document.getElementById("qual-read-feedback");
  if (!feedback) return;

  clearTimeout(qualReadFeedbackTimer);

  feedback.className = "upload-read-feedback";
  feedback.classList.add(success ? "upload-read-feedback--success" : "upload-read-feedback--error");
  feedback.innerHTML =
    (success ? QUAL_READ_SUCCESS_ICON : QUAL_READ_ERROR_ICON) +
    (success ? "Read successful" : "Read unsuccessful");
  feedback.hidden = false;

  qualReadFeedbackTimer = window.setTimeout(() => {
    feedback.classList.add("upload-read-feedback--fading");
    qualReadFeedbackTimer = window.setTimeout(() => {
      feedback.hidden = true;
      feedback.classList.remove("upload-read-feedback--fading");
    }, 500);
  }, 2500);
}

async function stageQualFile(file) {
  setQualUploadFilename(file.name);
  qualPendingFile = file;
  qualPendingFileTeams = [];
  qualPendingFileValid = false;

  const validTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "text/csv",
    "application/vnd.oasis.opendocument.spreadsheet",
  ];
  const ext = file.name.split(".").pop().toLowerCase();
  const validExts = ["xlsx", "xls", "csv", "ods"];

  if (!validExts.includes(ext) && !validTypes.includes(file.type)) {
    showQualReadFeedback(false);
    updateQualSaveButtons();
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const teams = parseTeamsFromWorkbook(workbook);

    if (teams.length < 2) {
      showQualReadFeedback(false);
      updateQualSaveButtons();
      return;
    }

    qualPendingFileTeams = teams;
    qualPendingFileValid = true;
    showQualReadFeedback(true);
  } catch {
    showQualReadFeedback(false);
  }

  updateQualSaveButtons();
}

function saveQualFromFile() {
  if (!qualPendingFileValid || qualPendingFileTeams.length < 2) return;

  const textarea = document.getElementById("qual-team-text");
  if (textarea) textarea.value = qualPendingFileTeams.join("\n");

  setTeams("qual", qualPendingFileTeams, qualPendingFile?.name || "Uploaded file");
  updateQualSaveButtons();
  invalidateQualGeneration();
  updateQualFormState();
}

function saveQualFromManual() {
  const textarea = document.getElementById("qual-team-text");
  const teams = textarea ? parseTeamsFromText(textarea.value) : [];
  setTeams("qual", teams, "Manual entry");
  invalidateQualGeneration();
  updateQualFormState();
}

function updateQualSaveButtons() {
  const fileBtn = document.getElementById("qual-save-file-btn");
  if (fileBtn) fileBtn.disabled = !qualPendingFileValid;
}

function invalidateQualGeneration() {
  qualHasGenerated = false;
  qualPreviewGroups = null;

  const preview = document.getElementById("qual-bracket-preview");
  const downloadBlock = document.getElementById("qual-download-block");
  const groupsDisplay = document.getElementById("qual-groups-display");
  const scheduleDisplay = document.getElementById("qual-schedule-display");
  const noteEl = document.getElementById("schedule-note");

  if (preview) preview.hidden = true;
  if (downloadBlock) downloadBlock.hidden = true;
  if (groupsDisplay) groupsDisplay.innerHTML = "";
  if (scheduleDisplay) scheduleDisplay.innerHTML = "";
  if (noteEl) noteEl.hidden = true;
}

function updateQualFormState() {
  const form = document.getElementById("qual-form");
  const empty = document.getElementById("qual-empty-state");
  const generateBtn = document.getElementById("qual-generate-btn");
  const teams = AppState.qual.teams;

  if (!form || !empty) return;

  if (teams.length < 2) {
    form.hidden = true;
    empty.hidden = false;
    invalidateQualGeneration();
    if (generateBtn) generateBtn.disabled = true;
    return;
  }

  empty.hidden = true;
  form.hidden = false;
  if (generateBtn) generateBtn.disabled = false;
}

function runQualGenerate() {
  const teams = AppState.qual.teams;
  const groupSizeInput = document.getElementById("group-size");
  const gamesInput = document.getElementById("games-per-team");
  const errorEl = document.getElementById("qual-error");
  const preview = document.getElementById("qual-bracket-preview");
  const downloadBlock = document.getElementById("qual-download-block");
  const groupSize = parseInt(groupSizeInput?.value, 10) || 4;
  const gamesPerTeam = parseInt(gamesInput?.value, 10) || 3;

  errorEl.hidden = true;

  try {
    if (teams.length < 2) {
      throw new Error("Save at least 2 teams before generating a qual bracket.");
    }

    qualPreviewGroups = generateQualBracket(teams, groupSize, gamesPerTeam, true);
    renderGroupsDisplay(qualPreviewGroups);
    renderScheduleDisplay(qualPreviewGroups);
    updateScheduleNote(groupSize, gamesPerTeam);
    qualHasGenerated = true;
    preview.hidden = false;
    downloadBlock.hidden = false;
  } catch (err) {
    qualHasGenerated = false;
    qualPreviewGroups = null;
    preview.hidden = true;
    downloadBlock.hidden = true;
    errorEl.hidden = false;
    errorEl.textContent = err.message || "Could not generate qual bracket.";
  }
}

function initQualSaveButtons() {
  const fileBtn = document.getElementById("qual-save-file-btn");
  const manualBtn = document.getElementById("qual-save-manual-btn");
  const textarea = document.getElementById("qual-team-text");

  fileBtn?.addEventListener("click", saveQualFromFile);
  manualBtn?.addEventListener("click", saveQualFromManual);
  textarea?.addEventListener("input", updateQualSaveButtons);
  updateQualSaveButtons();
}

function initQualUploadClear() {
  const clearBtn = document.getElementById("qual-upload-clear");
  clearBtn?.addEventListener("click", (e) => {
    e.stopPropagation();
    clearQualUploadFile();
  });
}

function initQualGenerateButton() {
  const generateBtn = document.getElementById("qual-generate-btn");
  generateBtn?.addEventListener("click", runQualGenerate);
}

function renderGroupsDisplay(groups) {
  const container = document.getElementById("qual-groups-display");
  if (!container) return;

  container.innerHTML = groups
    .map(
      (group) => `
        <div class="qual-group-card">
          <h4>Group ${escapeHtml(group.label)}</h4>
          <ol class="qual-group-list">
            ${group.teams.map((team) => `<li>${escapeHtml(team)}</li>`).join("")}
          </ol>
        </div>
      `
    )
    .join("");
}

function renderScheduleDisplay(groups) {
  const container = document.getElementById("qual-schedule-display");
  if (!container) return;

  const rows = [];
  for (const group of groups) {
    group.rounds.forEach((matches, roundIdx) => {
      matches.forEach((match, matchIdx) => {
        rows.push(`
          <tr>
            <td>Group ${escapeHtml(group.label)}</td>
            <td>${roundIdx + 1}</td>
            <td>${matchIdx + 1}</td>
            <td>${escapeHtml(match[0])}</td>
            <td>${escapeHtml(match[1])}</td>
          </tr>
        `);
      });
    });
  }

  if (rows.length === 0) {
    container.innerHTML = '<p class="qual-preview-empty">No matches scheduled.</p>';
    return;
  }

  container.innerHTML = `
    <table class="qual-schedule-table">
      <thead>
        <tr>
          <th>Group</th>
          <th>Round</th>
          <th>Match</th>
          <th>Team 1</th>
          <th>Team 2</th>
        </tr>
      </thead>
      <tbody>${rows.join("")}</tbody>
    </table>
  `;
}

function initQualBracket() {
  const form = document.getElementById("qual-form");
  const groupSizeInput = document.getElementById("group-size");
  const gamesInput = document.getElementById("games-per-team");
  const errorEl = document.getElementById("qual-error");

  document.addEventListener("qual-teams-updated", (e) => {
    updateQualSaveButtons();
    updateQualFormState();
    if (e.detail.teams.length === 0 && !e.detail.sourceLabel) {
      resetQualUploadZone();
    }
  });

  groupSizeInput.addEventListener("input", invalidateQualGeneration);
  gamesInput.addEventListener("input", invalidateQualGeneration);

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    errorEl.hidden = true;

    try {
      if (!qualHasGenerated || !qualPreviewGroups) {
        throw new Error("Click Generate to create groups and a schedule before downloading.");
      }
      downloadQualBracket(qualPreviewGroups);
    } catch (err) {
      errorEl.hidden = false;
      errorEl.textContent = err.message || "Something went wrong generating the bracket.";
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initQualBracket();
  initQualSaveButtons();
  initQualUploadClear();
  initQualGenerateButton();
  updateQualFormState();
});
