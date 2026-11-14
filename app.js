"use strict";

const STORAGE_KEY = "planner.workbook.v1";

const state = {
  workbook: null,
  sheetIndex: 0,
  sel: { r1: 0, c1: 0, r2: 0, c2: 0 },
  editing: null,
};

const gridEl = document.getElementById("grid");
const tabsEl = document.getElementById("sheet-tabs");
const refEl = document.getElementById("cell-ref");
const formulaInput = document.getElementById("formula-input");
const fileInput = document.getElementById("file-input");

function colName(i) {
  let s = "";
  let n = i + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function colIndex(name) {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function key(r, c) { return r + "," + c; }

function sheet() { return state.workbook.sheets[state.sheetIndex]; }

function cellAt(s, r, c) { return s.cells[key(r, c)] || null; }

function getRaw(s, r, c) { const cell = cellAt(s, r, c); return cell ? cell.v : ""; }

function setCell(s, r, c, props) {
  const k = key(r, c);
  const cur = s.cells[k] || { v: "" };
  const next = Object.assign({}, cur, props);
  if (next.v === "" && !next.b && !next.color && !next.bg) delete s.cells[k];
  else s.cells[k] = next;
}

function defaultWorkbook() {
  return {
    sheets: [
      {
        name: "Schedule",
        rows: 8,
        cols: 6,
        cells: {
          "0,0": { v: "Day", b: true }, "0,1": { v: "Morning", b: true },
          "0,2": { v: "Afternoon", b: true }, "0,3": { v: "Evening", b: true },
          "0,4": { v: "Who", b: true }, "0,5": { v: "Notes", b: true },
          "1,0": { v: "Monday" }, "2,0": { v: "Tuesday" }, "3,0": { v: "Wednesday" },
          "4,0": { v: "Thursday" }, "5,0": { v: "Friday" }, "6,0": { v: "Saturday" },
          "7,0": { v: "Sunday" },
          "1,3": { v: "Dinner with Alex", bg: "#fde68a" },
          "5,2": { v: "Gym", bg: "#bbf7d0" },
          "6,3": { v: "Brunch with family", bg: "#bfdbfe" },
        },
      },
      {
        name: "Tasks",
        rows: 10,
        cols: 5,
        cells: {
          "0,0": { v: "Task", b: true }, "0,1": { v: "Owner", b: true },
          "0,2": { v: "Due", b: true }, "0,3": { v: "Status", b: true },
          "0,4": { v: "Priority", b: true },
          "1,0": { v: "Book flights" }, "1,1": { v: "Partner" }, "1,3": { v: "todo", bg: "#fee2e2" }, "1,4": { v: "high" },
          "2,0": { v: "Buy birthday gift" }, "2,1": { v: "You" }, "2,3": { v: "doing", bg: "#fef3c7" }, "2,4": { v: "medium" },
          "3,0": { v: "Call plumber" }, "3,1": { v: "You" }, "3,3": { v: "done", bg: "#bbf7d0" }, "3,4": { v: "low" },
        },
      },
      {
        name: "Ideas",
        rows: 12,
        cols: 6,
        cells: {
          "0,0": { v: "Plan", b: true }, "0,1": { v: "Budget", b: true },
          "0,2": { v: "Spent", b: true }, "0,3": { v: "Left", b: true },
          "1,0": { v: "Christmas" }, "1,1": { v: "450" }, "1,2": { v: "120" }, "1,3": { v: "=B2-C2" },
          "2,0": { v: "Summer trip" }, "2,1": { v: "2000" }, "2,2": { v: "300" }, "2,3": { v: "=B3-C3" },
        },
      },
    ],
  };
}

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* corrupted storage falls through to default */ }
  return null;
}

function save() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.workbook)); } catch (e) { /* storage full/blocked */ }
}

function evalCell(s, r, c, depth) {
  if (depth === undefined) depth = 0;
  if (depth > 40) return "#CIRC!";
  const raw = getRaw(s, r, c);
  if (typeof raw === "string" && raw.startsWith("=")) return evalFormula(raw.slice(1), s, depth + 1);
  if (raw === "") return "";
  const n = Number(raw);
  return raw.trim() !== "" && Number.isFinite(n) ? n : raw;
}

function tokenizeFormula(s) {
  const tokens = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { i++; continue; }
    const rest = s.slice(i);
    const ref = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})/.exec(rest);
    if (ref) {
      tokens.push({ t: "ref", col: colIndex(ref[1]), row: parseInt(ref[2], 10) - 1 });
      i += ref[0].length;
      continue;
    }
    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < s.length && /[0-9.]/.test(s[j])) j++;
      tokens.push({ t: "num", v: parseFloat(s.slice(i, j)) });
      i = j;
      continue;
    }
    if (/[A-Za-z]/.test(ch)) {
      let j = i;
      while (j < s.length && /[A-Za-z]/.test(s[j])) j++;
      tokens.push({ t: "word", v: s.slice(i, j).toUpperCase() });
      i = j;
      continue;
    }
    if ("+-*/(),:".includes(ch)) { tokens.push({ t: ch }); i++; continue; }
    if (ch === "$") { i++; continue; }
    return { error: "#NAME?" };
  }
  return tokens;
}

function parseFormula(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];

  function parseExpr() {
    let left = parseTerm();
    while (peek() && (peek().t === "+" || peek().t === "-")) {
      const op = next().t;
      left = { op, left, right: parseTerm() };
    }
    return left;
  }

  function parseTerm() {
    let left = parseFactor();
    while (peek() && (peek().t === "*" || peek().t === "/")) {
      const op = next().t;
      left = { op, left, right: parseFactor() };
    }
    return left;
  }

  function parseFactor() {
    const tok = peek();
    if (!tok) return { error: "#SYNTAX!" };
    if (tok.t === "-") { next(); return { op: "neg", left: parseFactor() }; }
    if (tok.t === "+") { next(); return parseFactor(); }
    if (tok.t === "(") {
      next();
      const e = parseExpr();
      if (!peek() || next().t !== ")") return { error: "#SYNTAX!" };
      return e;
    }
    if (tok.t === "num") { next(); return { num: tok.v }; }
    if (tok.t === "ref") {
      next();
      if (peek() && peek().t === ":") {
        next();
        const r2 = next();
        if (!r2 || r2.t !== "ref") return { error: "#SYNTAX!" };
        return { range: [tok.row, tok.col, r2.row, r2.col] };
      }
      return { ref: [tok.row, tok.col] };
    }
    if (tok.t === "word") {
      next();
      if (peek() && peek().t === "(") {
        next();
        const args = [];
        if (peek() && peek().t !== ")") {
          args.push(parseExpr());
          while (peek() && peek().t === ",") { next(); args.push(parseExpr()); }
        }
        if (!peek() || next().t !== ")") return { error: "#SYNTAX!" };
        return { func: tok.v, args };
      }
      return { error: "#NAME?" };
    }
    return { error: "#SYNTAX!" };
  }

  const ast = parseExpr();
  if (pos !== tokens.length) return { error: "#SYNTAX!" };
  return ast;
}

function evalNode(node, s, depth) {
  if (node.error) return node.error;
  if (node.num !== undefined) return node.num;
  if (node.ref) return evalCell(s, node.ref[0], node.ref[1], depth);
  if (node.range) {
    const r1 = Math.min(node.range[0], node.range[2]);
    const r2 = Math.max(node.range[0], node.range[2]);
    const c1 = Math.min(node.range[1], node.range[3]);
    const c2 = Math.max(node.range[1], node.range[3]);
    const vals = [];
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) vals.push(evalCell(s, r, c, depth));
    return vals;
  }
  if (node.op === "neg") {
    const v = evalNode(node.left, s, depth);
    return typeof v === "number" ? -v : "#VALUE!";
  }
  if (node.op) {
    const a = evalNode(node.left, s, depth);
    const b = evalNode(node.right, s, depth);
    if (typeof a !== "number" || typeof b !== "number") return "#VALUE!";
    if (node.op === "+") return a + b;
    if (node.op === "-") return a - b;
    if (node.op === "*") return a * b;
    if (node.op === "/") return b === 0 ? "#DIV/0!" : a / b;
  }
  if (node.func) {
    const args = node.args.map((a) => evalNode(a, s, depth));
    const nums = [];
    for (const a of args) {
      if (Array.isArray(a)) nums.push.apply(nums, a.filter((v) => typeof v === "number"));
      else if (typeof a === "number") nums.push(a);
    }
    if (node.func === "SUM") return nums.reduce((sum, n) => sum + n, 0);
    if (node.func === "AVG") return nums.length ? nums.reduce((sum, n) => sum + n, 0) / nums.length : "#DIV/0!";
    if (node.func === "MIN") return nums.length ? Math.min.apply(Math, nums) : 0;
    if (node.func === "MAX") return nums.length ? Math.max.apply(Math, nums) : 0;
    return "#NAME?";
  }
  return "#VALUE!";
}

function evalFormula(expr, s, depth) {
  const tokens = tokenizeFormula(expr);
  if (tokens.error) return tokens.error;
  const v = evalNode(parseFormula(tokens), s, depth);
  if (typeof v === "number" && !Number.isInteger(v)) return Math.round(v * 1e10) / 1e10;
  return v;
}

function displayValue(s, r, c) {
  const raw = getRaw(s, r, c);
  if (typeof raw === "string" && raw.startsWith("=")) {
    const v = evalCell(s, r, c, 0);
    return v === "" ? "" : String(v);
  }
  return raw;
}

function isSelected(r, c) {
  const sel = state.sel;
  return r >= Math.min(sel.r1, sel.r2) && r <= Math.max(sel.r1, sel.r2) &&
    c >= Math.min(sel.c1, sel.c2) && c <= Math.max(sel.c1, sel.c2);
}

function makeTh(text, cls) {
  const el = document.createElement("th");
  el.textContent = text;
  if (cls) el.className = cls;
  return el;
}

function renderCell(s, r, c) {
  const td = document.createElement("td");
  td.dataset.r = r;
  td.dataset.c = c;
  const cell = cellAt(s, r, c);
  if (cell) {
    if (cell.b) td.style.fontWeight = "bold";
    if (cell.color) td.style.color = cell.color;
    if (cell.bg) td.style.backgroundColor = cell.bg;
  }
  td.textContent = displayValue(s, r, c);
  if (isSelected(r, c)) td.classList.add("selected");
  if (state.sel.r1 === r && state.sel.c1 === c) td.classList.add("active");
  return td;
}

function renderGrid() {
  const s = sheet();
  gridEl.innerHTML = "";
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  hr.appendChild(makeTh("", "corner"));
  for (let c = 0; c < s.cols; c++) {
    const thEl = makeTh(colName(c), "col-head");
    thEl.dataset.c = c;
    hr.appendChild(thEl);
  }
  thead.appendChild(hr);
  gridEl.appendChild(thead);
  const tbody = document.createElement("tbody");
  for (let r = 0; r < s.rows; r++) {
    const tr = document.createElement("tr");
    const rh = makeTh(String(r + 1), "row-head");
    rh.dataset.r = r;
    tr.appendChild(rh);
    for (let c = 0; c < s.cols; c++) tr.appendChild(renderCell(s, r, c));
    tbody.appendChild(tr);
  }
  gridEl.appendChild(tbody);
}

function renderTabs() {
  tabsEl.innerHTML = "";
  state.workbook.sheets.forEach((s, i) => {
    const b = document.createElement("button");
    b.className = "sheet-tab" + (i === state.sheetIndex ? " active" : "");
    b.textContent = s.name;
    b.addEventListener("click", () => {
      state.sheetIndex = i;
      state.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
      render();
    });
    b.addEventListener("dblclick", () => renameSheet(i));
    tabsEl.appendChild(b);
  });
}

function renderFormulaBar() {
  const s = sheet();
  refEl.textContent = colName(state.sel.c1) + (state.sel.r1 + 1);
  if (!state.editing) formulaInput.value = getRaw(s, state.sel.r1, state.sel.c1);
}

function render() {
  renderGrid();
  renderTabs();
  renderFormulaBar();
}

function moveSelTo(r, c) {
  const s = sheet();
  r = Math.max(0, Math.min(s.rows - 1, r));
  c = Math.max(0, Math.min(s.cols - 1, c));
  state.sel = { r1: r, c1: c, r2: r, c2: c };
  render();
}

function clearSel() {
  eachCellInSel((s, r, c) => setCell(s, r, c, { v: "" }));
  save();
  render();
}

function eachCellInSel(fn) {
  const s = sheet();
  const sel = state.sel;
  for (let r = Math.min(sel.r1, sel.r2); r <= Math.max(sel.r1, sel.r2); r++)
    for (let c = Math.min(sel.c1, sel.c2); c <= Math.max(sel.c1, sel.c2); c++) fn(s, r, c);
}

function applyToSel(props) {
  eachCellInSel((s, r, c) => setCell(s, r, c, props));
  save();
  render();
}

function startEdit(r, c) {
  if (state.editing) return;
  const s = sheet();
  const td = gridEl.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
  if (!td) return;
  const input = document.createElement("input");
  input.className = "cell-editor";
  input.value = getRaw(s, r, c);
  td.textContent = "";
  td.appendChild(input);
  state.editing = { r, c };
  input.focus();
  input.select();
  let done = false;
  const finish = (commit, move) => {
    if (done) return;
    done = true;
    if (commit) setCell(s, r, c, { v: input.value });
    state.editing = null;
    save();
    render();
    if (move === "down") moveSelTo(r + 1, c);
    else if (move === "right") moveSelTo(r, c + 1);
    else if (move === "left") moveSelTo(r, c - 1);
  };
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); finish(true, "down"); }
    else if (e.key === "Tab") { e.preventDefault(); e.stopPropagation(); finish(true, e.shiftKey ? "left" : "right"); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(false, null); }
  });
  input.addEventListener("blur", () => finish(true, null));
}

function insertRow(s, at) {
  const cells = {};
  for (const k in s.cells) {
    const parts = k.split(",");
    const r = Number(parts[0]);
    const c = Number(parts[1]);
    cells[key(r >= at ? r + 1 : r, c)] = s.cells[k];
  }
  s.cells = cells;
  s.rows++;
}

function deleteRow(s, at) {
  const cells = {};
  for (const k in s.cells) {
    const parts = k.split(",");
    const r = Number(parts[0]);
    const c = Number(parts[1]);
    if (r === at) continue;
    cells[key(r > at ? r - 1 : r, c)] = s.cells[k];
  }
  s.cells = cells;
  s.rows--;
}

function insertCol(s, at) {
  const cells = {};
  for (const k in s.cells) {
    const parts = k.split(",");
    const r = Number(parts[0]);
    const c = Number(parts[1]);
    cells[key(r, c >= at ? c + 1 : c)] = s.cells[k];
  }
  s.cells = cells;
  s.cols++;
}

function deleteCol(s, at) {
  const cells = {};
  for (const k in s.cells) {
    const parts = k.split(",");
    const r = Number(parts[0]);
    const c = Number(parts[1]);
    if (c === at) continue;
    cells[key(r, c > at ? c - 1 : c)] = s.cells[k];
  }
  s.cells = cells;
  s.cols--;
}

function uniqueSheetName(wb) {
  let i = wb.sheets.length + 1;
  let name = "Sheet " + i;
  while (wb.sheets.some((s) => s.name === name)) { i++; name = "Sheet " + i; }
  return name;
}

function renameSheet(i) {
  const name = prompt("Sheet name", state.workbook.sheets[i].name);
  if (name && name.trim()) {
    state.workbook.sheets[i].name = name.trim();
    save();
    render();
  }
}

function csvEscape(v) {
  v = String(v);
  return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

function sheetToCSV(s) {
  const lines = [];
  for (let r = 0; r < s.rows; r++) {
    const row = [];
    for (let c = 0; c < s.cols; c++) row.push(csvEscape(getRaw(s, r, c)));
    lines.push(row.join(","));
  }
  return lines.join("\n");
}

function parseCSV(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++; } else inQ = false;
      } else cur += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ",") { row.push(cur); cur = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cur);
      cur = "";
      rows.push(row);
      row = [];
    } else cur += ch;
  }
  if (cur !== "" || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function importCSV(text) {
  const rows = parseCSV(text);
  const s = sheet();
  s.cells = {};
  s.rows = Math.max(rows.length, 1);
  s.cols = Math.max.apply(Math, rows.map((r) => r.length).concat([1]));
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== "") setCell(s, r, c, { v }); }));
  save();
  render();
}

function download(filename, text, type) {
  const blob = new Blob([text], { type: type || "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

gridEl.addEventListener("click", (e) => {
  const td = e.target.closest("td");
  const thEl = e.target.closest("th");
  if (td) {
    const r = Number(td.dataset.r);
    const c = Number(td.dataset.c);
    if (e.shiftKey) {
      state.sel.r2 = r;
      state.sel.c2 = c;
    } else {
      state.sel = { r1: r, c1: c, r2: r, c2: c };
    }
    render();
  } else if (thEl && thEl.classList.contains("col-head")) {
    const s = sheet();
    const c = Number(thEl.dataset.c);
    state.sel = { r1: 0, c1: c, r2: s.rows - 1, c2: c };
    render();
  } else if (thEl && thEl.classList.contains("row-head")) {
    const s = sheet();
    const r = Number(thEl.dataset.r);
    state.sel = { r1: r, c1: 0, r2: r, c2: s.cols - 1 };
    render();
  }
});

gridEl.addEventListener("dblclick", (e) => {
  const td = e.target.closest("td");
  if (td) startEdit(Number(td.dataset.r), Number(td.dataset.c));
});

document.addEventListener("keydown", (e) => {
  if (state.editing) return;
  const t = e.target;
  if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "BUTTON" || t.isContentEditable)) return;
  if (document.getElementById("app-planner").hidden) return;
  const s = sheet();
  const r1 = state.sel.r1;
  const c1 = state.sel.c1;
  const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  if (moves[e.key]) {
    e.preventDefault();
    moveSelTo(r1 + moves[e.key][0], c1 + moves[e.key][1]);
  } else if (e.key === "Tab") {
    e.preventDefault();
    moveSelTo(r1, c1 + (e.shiftKey ? -1 : 1));
  } else if (e.key === "Enter" || e.key === "F2") {
    e.preventDefault();
    startEdit(r1, c1);
  } else if (e.key === "Delete" || e.key === "Backspace") {
    e.preventDefault();
    clearSel();
  }
});

formulaInput.addEventListener("change", () => {
  setCell(sheet(), state.sel.r1, state.sel.c1, { v: formulaInput.value });
  save();
  render();
});

document.getElementById("btn-bold").addEventListener("click", () => {
  const cur = !!(cellAt(sheet(), state.sel.r1, state.sel.c1) || {}).b;
  applyToSel({ b: !cur });
});

document.getElementById("inp-color").addEventListener("input", (e) => applyToSel({ color: e.target.value }));
document.getElementById("inp-bg").addEventListener("input", (e) => applyToSel({ bg: e.target.value }));

document.getElementById("btn-add-row").addEventListener("click", () => { insertRow(sheet(), state.sel.r1); save(); render(); });
document.getElementById("btn-del-row").addEventListener("click", () => { const s = sheet(); if (s.rows > 1) { deleteRow(s, state.sel.r1); save(); render(); } });
document.getElementById("btn-add-col").addEventListener("click", () => { insertCol(sheet(), state.sel.c1); save(); render(); });
document.getElementById("btn-del-col").addEventListener("click", () => { const s = sheet(); if (s.cols > 1) { deleteCol(s, state.sel.c1); save(); render(); } });

document.getElementById("btn-new-sheet").addEventListener("click", () => {
  state.workbook.sheets.push({ name: uniqueSheetName(state.workbook), rows: 20, cols: 8, cells: {} });
  state.sheetIndex = state.workbook.sheets.length - 1;
  state.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
  save();
  render();
});

document.getElementById("btn-rename-sheet").addEventListener("click", () => renameSheet(state.sheetIndex));

document.getElementById("btn-del-sheet").addEventListener("click", () => {
  const wb = state.workbook;
  if (wb.sheets.length <= 1) { alert("A workbook needs at least one sheet."); return; }
  if (confirm('Delete sheet "' + wb.sheets[state.sheetIndex].name + '"?')) {
    wb.sheets.splice(state.sheetIndex, 1);
    state.sheetIndex = Math.min(state.sheetIndex, wb.sheets.length - 1);
    save();
    render();
  }
});

document.getElementById("btn-import").addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  const f = fileInput.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    const text = String(reader.result);
    if (f.name.toLowerCase().endsWith(".json")) {
      try {
        const wb = JSON.parse(text);
        if (!wb || !Array.isArray(wb.sheets) || wb.sheets.length === 0) throw new Error("bad workbook");
        state.workbook = wb;
        state.sheetIndex = 0;
        state.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
        save();
        render();
      } catch (err) {
        alert("Could not read that JSON file.");
      }
    } else {
      importCSV(text);
    }
  };
  reader.readAsText(f);
  fileInput.value = "";
});

document.getElementById("btn-export-csv").addEventListener("click", () => {
  const s = sheet();
  download(s.name.replace(/[^\w-]+/g, "_") + ".csv", sheetToCSV(s), "text/csv");
});

document.getElementById("btn-export-json").addEventListener("click", () => {
  download("planner-workbook.json", JSON.stringify(state.workbook, null, 2), "application/json");
});

document.getElementById("btn-reset").addEventListener("click", () => {
  if (confirm("Replace the saved workbook in this browser with the demo data?")) {
    localStorage.removeItem(STORAGE_KEY);
    state.workbook = defaultWorkbook();
    state.sheetIndex = 0;
    state.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
    save();
    render();
  }
});

document.querySelectorAll(".app-pill").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".app-pill").forEach((b) => b.classList.toggle("active", b === btn));
    const app = btn.dataset.app;
    document.getElementById("app-planner").hidden = app !== "planner";
    const ph = document.getElementById("app-placeholder");
    ph.hidden = app === "planner";
    if (app !== "planner") document.getElementById("placeholder-title").textContent = btn.textContent;
  });
});

state.workbook = load() || window.SEED_CALENDAR || defaultWorkbook();
render();
