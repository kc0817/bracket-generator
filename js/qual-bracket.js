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

function chunkIntoGroups(teams, groupSize) {
  const shuffled = shuffle(teams);
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

function generateQualBracket(teams, groupSize, gamesPerTeam) {
  if (teams.length < 2) {
    throw new Error("At least 2 teams are required.");
  }
  if (groupSize < 2) {
    throw new Error("Group size must be at least 2.");
  }
  if (gamesPerTeam < 1) {
    throw new Error("Games per team must be at least 1.");
  }

  const rawGroups = chunkIntoGroups(teams, groupSize);
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

function updateQualGenerateButton() {
  const btn = document.getElementById("qual-generate-btn");
  const textarea = document.getElementById("qual-team-text");
  if (!btn || !textarea) return;

  const teams = parseTeamsFromText(textarea.value);
  btn.disabled = teams.length < 2;
}

function initQualGenerateButton() {
  const btn = document.getElementById("qual-generate-btn");
  const textarea = document.getElementById("qual-team-text");
  if (!btn || !textarea) return;

  textarea.addEventListener("input", updateQualGenerateButton);
  btn.addEventListener("click", () => handleManualEntry("qual"));
  updateQualGenerateButton();
}

function renderQualTeamsPreview(teams) {
  const preview = document.getElementById("qual-teams-preview");
  const empty = document.getElementById("qual-empty-state");
  const form = document.getElementById("qual-form");
  const errorEl = document.getElementById("qual-error");

  renderTeamsPreview("qual", teams);
  errorEl.hidden = true;

  if (teams.length === 0) {
    form.hidden = true;
    return;
  }

  form.hidden = false;

  const groupSize = parseInt(document.getElementById("group-size").value, 10) || 4;
  const gamesPerTeam = parseInt(document.getElementById("games-per-team").value, 10) || 3;
  updateScheduleNote(groupSize, gamesPerTeam);
}

function initQualBracket() {
  const form = document.getElementById("qual-form");
  const groupSizeInput = document.getElementById("group-size");
  const gamesInput = document.getElementById("games-per-team");
  const errorEl = document.getElementById("qual-error");

  document.addEventListener("qual-teams-updated", (e) => {
    renderQualTeamsPreview(e.detail.teams);
    updateQualGenerateButton();
  });

  function onSettingsChange() {
    const groupSize = parseInt(groupSizeInput.value, 10) || 4;
    const gamesPerTeam = parseInt(gamesInput.value, 10) || 3;
    updateScheduleNote(groupSize, gamesPerTeam);
  }

  groupSizeInput.addEventListener("input", onSettingsChange);
  gamesInput.addEventListener("input", onSettingsChange);

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    errorEl.hidden = true;

    const teams = AppState.qual.teams;
    const groupSize = parseInt(groupSizeInput.value, 10);
    const gamesPerTeam = parseInt(gamesInput.value, 10);

    try {
      if (teams.length < 2) {
        throw new Error("Add at least 2 teams using the upload or manual entry above.");
      }
      const groups = generateQualBracket(teams, groupSize, gamesPerTeam);
      downloadQualBracket(groups);
    } catch (err) {
      errorEl.hidden = false;
      errorEl.textContent = err.message || "Something went wrong generating the bracket.";
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initQualBracket();
  initQualGenerateButton();
});
