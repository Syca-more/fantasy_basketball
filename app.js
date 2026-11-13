// ---------------------------------------------------------------
// SETTINGS
// Player lists live in players.json (edit that file to add players).
// Each player needs a matching data file, e.g. "Ty Lawson" -> ty_lawson.xls
// ---------------------------------------------------------------
const PLAYERS_FILE = "players.json";

const FILE_EXT = ".xls";
const FILE_FOLDER = "";   // use "data/" if you put the player files in a data folder

const COLUMNS = ["Date", "Opp", "Result", "MP", "FG", "FGA", "FG%", "3P", "3PA",
  "3P%", "eFG%", "FT", "FTA", "FT%", "ORB", "DRB", "TRB", "AST", "STL",
  "BLK", "TOV", "PF", "PTS", "GmSc", "+/-"];
const KEY_COLUMNS = ["TRB", "AST", "PTS"];   // highlighted in orange

// TEAM BOOST: if every displayed player's game is for the same team
// (and at least MIN_PLAYERS_FOR_BOOST players are shown), the total gets a bonus.
const TEAM_BOOST = 0.10;            // 10%
const MIN_PLAYERS_FOR_BOOST = 2;    // set to 1 to also boost a single player
const TEAM_COLUMN_NAMES = ["team", "tm"];   // header name(s) of the team column (any capitalization)

// Filled in from players.json when the page loads
let POSITIONS = {};
let SECRET_PLAYER = "";
let SECRET_POSITION = "";

const fieldsEl = document.getElementById("fields");
const result = document.getElementById("result");
const summaryEl = document.getElementById("summary");
const timmyEl = document.getElementById("timmy");
const cache = {};    // so each file is only downloaded once
const combos = {};   // position -> its dropdown

// ---------------------------------------------------------------
// Searchable dropdown: click shows ALL options, typing filters them
// ---------------------------------------------------------------
function createCombo(players, id) {
  const wrap = document.createElement("div");
  wrap.className = "combo";

  const input = document.createElement("input");
  input.id = id;
  input.type = "text";
  input.placeholder = "Select Player";
  input.autocomplete = "off";

  const arrow = document.createElement("span");
  arrow.className = "arrow";
  arrow.textContent = "▾";

  const list = document.createElement("ul");
  list.className = "options";
  list.hidden = true;

  wrap.append(input, arrow, list);

  let selected = "";   // last valid choice
  let shown = [];      // options currently visible
  let active = -1;     // keyboard-highlighted index

  function render(filter) {
    const q = filter.trim().toLowerCase();
    shown = q ? players.filter(p => p.toLowerCase().includes(q)) : players.slice();
    active = -1;
    list.innerHTML = "";
    if (!shown.length) {
      const li = document.createElement("li");
      li.className = "empty";
      li.textContent = players.length ? "No matches" : "No players added yet";
      list.appendChild(li);
      return;
    }
    shown.forEach(p => {
      const li = document.createElement("li");
      li.textContent = p;
      if (p === selected) li.classList.add("selected");
      // mousedown (not click) so it fires before the input loses focus
      li.addEventListener("mousedown", e => { e.preventDefault(); choose(p); });
      list.appendChild(li);
    });
  }

  function setActive(i) {
    const items = list.querySelectorAll("li:not(.empty)");
    items.forEach(li => li.classList.remove("active"));
    active = i;
    if (items[i]) {
      items[i].classList.add("active");
      items[i].scrollIntoView({ block: "nearest" });
    }
  }

  function openAll() { render(""); list.hidden = false; }
  function close() { list.hidden = true; }
  function choose(p) { selected = p; input.value = p; close(); input.blur(); }

  input.addEventListener("focus", () => { input.select(); openAll(); });
  input.addEventListener("click", () => { if (list.hidden) openAll(); });
  input.addEventListener("input", () => { render(input.value); list.hidden = false; });

  input.addEventListener("blur", () => {
    const typed = input.value.trim();
    const exact = players.find(p => p.toLowerCase() === typed.toLowerCase());
    if (exact) selected = exact;            // typed a full valid name
    else if (typed === "") selected = "";   // cleared the box
    input.value = selected;                 // otherwise revert to last valid choice
    close();
  });

  input.addEventListener("keydown", e => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (list.hidden) openAll();
      if (shown.length) setActive(Math.min(active + 1, shown.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (shown.length) setActive(Math.max(active - 1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (!list.hidden) {
        if (active >= 0) choose(shown[active]);
        else if (shown.length === 1) choose(shown[0]);
      }
    } else if (e.key === "Escape") {
      input.value = selected;
      close();
    }
  });

  return { el: wrap, getValue: () => selected };
}

// ---------------------------------------------------------------
// Data loading
// ---------------------------------------------------------------
function fileFor(name) {
  const slug = name.toLowerCase().trim()
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, "_");
  return FILE_FOLDER + slug + FILE_EXT;
}

async function loadRows(name) {
  if (cache[name]) return cache[name];
  const url = fileFor(name);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Could not find ${url} (HTTP ${res.status})`);
  const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const all = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

  // Only rows with a value in MP can be picked (skips Did Not Play / Inactive
  // rows, and repeated header rows where the MP cell just says "MP")
  const rows = all.filter(r => {
    const mp = String(r["MP"] ?? "").trim();
    return mp !== "" && mp !== "MP";
  });

  cache[name] = rows;
  return rows;
}

// Finds the team abbreviation in a row, whatever the team column's capitalization
function getTeam(row) {
  for (const key of Object.keys(row)) {
    if (TEAM_COLUMN_NAMES.includes(key.trim().toLowerCase())) {
      return String(row[key] ?? "").trim().toUpperCase();
    }
  }
  return "";
}

// ---------------------------------------------------------------
// Result cards
// ---------------------------------------------------------------
function cellClass(col, value) {
  if (KEY_COLUMNS.includes(col)) return "key";
  if (col === "Result") {
    if (String(value).trim().startsWith("W")) return "win";
    if (String(value).trim().startsWith("L")) return "loss";
  }
  return "";
}

function errorCard(position, text) {
  return `<div class="card"><h2><span class="pos">${position}</span></h2>
          <div class="message error">${text}</div></div>`;
}

async function buildCard(position, name) {
  try {
    const rows = await loadRows(name);
    if (!rows.length) throw new Error("That file has no rows with minutes played (MP).");
    const row = rows[Math.floor(Math.random() * rows.length)];

    // Pull GmSc out as a number (null if missing or not a number)
    const parsed = parseFloat(String(row["GmSc"] ?? "").replace(/[^0-9.\-]/g, ""));
    const gmsc = isNaN(parsed) ? null : parsed;
    const team = getTeam(row);

    const head = COLUMNS
      .map(c => `<th class="${KEY_COLUMNS.includes(c) ? "key" : ""}">${c}</th>`)
      .join("");
    const cells = COLUMNS
      .map(c => `<td class="${cellClass(c, row[c])}">${row[c] ?? ""}</td>`)
      .join("");

    const html = `
      <div class="card">
        <h2><span class="pos">${position}:</span> ${name}
          <small>${team ? `<span class="team">${team}</span> · ` : ""}${row["Date"] ?? ""} ${row["Opp"] ? "vs " + row["Opp"] : ""}</small>
          ${gmsc !== null ? `<span class="gmsc">GmSc ${row["GmSc"]}</span>` : ""}</h2>
        <div class="wrap"><table><tr>${head}</tr><tr>${cells}</tr></table></div>
      </div>`;
    return { html, gmsc, team, ok: true };
  } catch (err) {
    return { html: errorCard(position, err.message), gmsc: null, team: "", ok: false };
  }
}

// ---------------------------------------------------------------
// Submit
// ---------------------------------------------------------------
async function handleSubmit() {
  const chosen = Object.keys(POSITIONS)
    .map(pos => [pos, combos[pos].getValue()])
    .filter(([, name]) => name !== "");

  summaryEl.innerHTML = "";

  if (!chosen.length) {
    result.innerHTML = `<div class="message">Select at least one player first.</div>`;
    return;
  }

  result.innerHTML = `<div class="message">Loading...</div>`;
  const built = await Promise.all(chosen.map(([pos, name]) => buildCard(pos, name)));

  const scores = built.map(b => b.gmsc).filter(g => g !== null);
  const baseTotal = scores.reduce((sum, g) => sum + g, 0);

  // Team boost: every successfully loaded game must have the same, non-blank team
  const loaded = built.filter(b => b.ok);
  const teams = loaded.map(b => b.team);
  const sameTeam = loaded.length >= MIN_PLAYERS_FOR_BOOST &&
    teams.every(t => t !== "") &&
    teams.every(t => t === teams[0]);
  const boostTeam = sameTeam ? teams[0] : "";
  const total = sameTeam ? baseTotal * (1 + TEAM_BOOST) : baseTotal;

  result.innerHTML = built.map(b => b.html).join("");

  if (scores.length) {
    const pct = Math.round(TEAM_BOOST * 100);
    summaryEl.innerHTML = `
      <div class="summary">
        <div>
          <div class="label">Total Game Score</div>
          <div class="sub">${sameTeam
            ? `${baseTotal.toFixed(1)} base + ${pct}% team boost`
            : `Sum of GmSc across ${scores.length} player${scores.length === 1 ? "" : "s"}`}</div>
        </div>
        ${sameTeam ? `<div class="boost">+${pct}% Team Boost (${boostTeam})</div>` : ""}
        <div class="total">${total.toFixed(1)}</div>
      </div>`;
  }
}

// ---------------------------------------------------------------
// Release the Beast: show Timmy for 3 seconds and unlock Tim Duncan
// ---------------------------------------------------------------
new Image().src = "timmy.png";   // preload so it appears instantly
let timmyTimer;
document.getElementById("beastBtn").addEventListener("click", () => {
  const list = POSITIONS[SECRET_POSITION];
  if (list && !list.includes(SECRET_PLAYER)) list.push(SECRET_PLAYER);

  timmyEl.hidden = false;
  clearTimeout(timmyTimer);
  timmyTimer = setTimeout(() => { timmyEl.hidden = true; }, 3000);
});

// ---------------------------------------------------------------
// Start-up: load players.json, then build the dropdowns
// ---------------------------------------------------------------
async function init() {
  try {
    const res = await fetch(PLAYERS_FILE);
    if (!res.ok) throw new Error(`Could not find ${PLAYERS_FILE} (HTTP ${res.status})`);
    const data = await res.json();
    POSITIONS = data.positions || {};
    SECRET_PLAYER = (data.secret && data.secret.player) || "";
    SECRET_POSITION = (data.secret && data.secret.position) || "";
  } catch (err) {
    result.innerHTML = `<div class="message error">Could not load the player list: ${err.message}.
      If the file exists, check players.json for a missing comma or quote.</div>`;
    return;
  }

  // Build one dropdown per position
  Object.entries(POSITIONS).forEach(([position, players], i) => {
    const field = document.createElement("div");
    field.className = "field";

    const label = document.createElement("label");
    label.htmlFor = "input-" + i;
    label.textContent = position;

    // Passing the array itself (not a copy) lets us add a player to it later
    const combo = createCombo(players, "input-" + i);
    field.append(label, combo.el);
    fieldsEl.appendChild(field);
    combos[position] = combo;
  });

  // Submit button lives in the same row as the dropdowns
  const submitBtn = document.createElement("button");
  submitBtn.id = "submitBtn";
  submitBtn.textContent = "Submit";
  submitBtn.addEventListener("click", handleSubmit);
  fieldsEl.appendChild(submitBtn);
}

init();
