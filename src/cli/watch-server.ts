import {
  createServer,
  type IncomingMessage,
  type Server,
  type ServerResponse,
} from "node:http";
import { basename } from "node:path";
import { getRegistry } from "../hooks.js";

// ---------------------------------------------------------------------------
// File provider — abstraction over daemon cache vs disk
// ---------------------------------------------------------------------------

export interface WatchFileProvider {
  /** List currently known file absolute paths */
  listFiles(): string[];

  /** Get the xlsx buffer for a file (from memory or disk) */
  getBuffer(absPath: string): Promise<Buffer | null>;

  /** Subscribe to file events. Returns unsubscribe function. */
  subscribe(listener: WatchEventListener): () => void;

  /** Release any provider-owned listeners/resources. */
  stop(): void;
}

export type WatchEvent =
  | { type: "opened"; absPath: string }
  | { type: "changed"; absPath: string };

export type WatchEventListener = (event: WatchEvent) => void;

// ---------------------------------------------------------------------------
// SSE helpers
// ---------------------------------------------------------------------------

function sseConnect(res: ServerResponse, clients: Set<ServerResponse>): void {
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
    "Access-Control-Allow-Origin": "*",
  });
  res.write(": connected\n\n");
  clients.add(res);
  res.on("close", () => clients.delete(res));
}

function sseBroadcast(clients: Set<ServerResponse>, data: string): void {
  const frame = `data: ${data}\n\n`;
  for (const c of clients) {
    try {
      c.write(frame);
    } catch {}
  }
}

// ---------------------------------------------------------------------------
// Hooks summary helper
// ---------------------------------------------------------------------------

function getHookName(fn: unknown): string {
  if (typeof fn === "function" && fn.name) return fn.name;
  return "(anonymous)";
}

async function getHooksSummary(): Promise<{
  hooks: { event: string; name: string }[];
  total: number;
}> {
  const registry = await getRegistry();
  const hooks: { event: string; name: string }[] = [];
  for (const event of [
    "preCommand",
    "postCommand",
    "onOpen",
    "preSave",
    "postSave",
  ] as const) {
    for (const entry of registry[event]) {
      hooks.push({ event, name: getHookName(entry.fn) });
    }
  }
  return { hooks, total: hooks.length };
}

// ---------------------------------------------------------------------------
// Viewer HTML
// ---------------------------------------------------------------------------

function buildViewerHTML(filenames: string[]): string {
  const fileListJSON = JSON.stringify(filenames);
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>hsx watch</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.min.css">
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg-base: #ffffff;
  --bg-surface: #f8f8f8;
  --bg-raised: #f0f0f0;
  --bg-overlay: #e8e8e8;
  --border-subtle: rgba(0,0,0,0.08);
  --border-medium: rgba(0,0,0,0.14);
  --text-primary: #1a1a1a;
  --text-secondary: #6b6b6b;
  --text-tertiary: #999999;
  --accent: #1a8a5c;
  --accent-dim: rgba(26,138,92,0.08);
  --accent-glow: rgba(26,138,92,0.2);
  --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  --font-mono: 'SF Mono', 'JetBrains Mono', Consolas, monospace;
  --radius-sm: 4px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --shadow-popup: 0 8px 32px rgba(0,0,0,0.12), 0 0 0 1px var(--border-subtle);
  --transition-fast: 0.12s cubic-bezier(0.4,0,0.2,1);
  --transition-normal: 0.2s cubic-bezier(0.4,0,0.2,1);
}

* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { height: 100%; }
body { font-family: var(--font-sans); background: var(--bg-base); -webkit-font-smoothing: antialiased; color: var(--text-primary); }

/* ---- Toolbar ---- */
#toolbar {
  display: flex; align-items: stretch;
  background: var(--bg-surface);
  color: var(--text-secondary); font-size: 13px;
  height: 34px;
  border-bottom: 1px solid var(--border-subtle);
  overflow: hidden;
  position: relative;
  z-index: 10;
}

#file-tabs {
  display: flex; align-items: stretch;
  min-width: 0; gap: 1px;
  padding-left: 2px;
  overflow-x: auto; overflow-y: hidden;
  scrollbar-width: none;
}
#file-tabs::-webkit-scrollbar { display: none; }

.file-tab {
  display: flex; align-items: center; gap: 8px;
  padding: 0 18px; cursor: pointer; color: var(--text-tertiary);
  background: transparent;
  white-space: nowrap; user-select: none;
  font-family: var(--font-mono);
  font-size: 12px; font-weight: 400;
  letter-spacing: 0.01em;
  transition: color var(--transition-fast), background var(--transition-fast);
  position: relative;
  outline: none;
  border: none;
  border-bottom: 2px solid transparent;
}
.file-tab::after {
  content: ''; position: absolute; bottom: -1px; left: 18px; right: 18px;
  height: 2px; background: var(--accent); border-radius: 2px 2px 0 0;
  transform: scaleX(0); transition: transform var(--transition-normal);
}
.file-tab:hover { color: var(--text-primary); background: var(--bg-raised); }
.file-tab:focus-visible { color: var(--text-primary); background: var(--bg-raised); }
.file-tab:focus-visible::after { transform: scaleX(1); }
.file-tab.active {
  color: var(--text-primary); background: var(--bg-base);
}
.file-tab.active::after { transform: scaleX(1); }
.file-tab .dot {
  width: 7px; height: 7px; border-radius: 50%;
  background: var(--accent); flex-shrink: 0;
  box-shadow: 0 0 6px var(--accent-glow);
  transition: opacity var(--transition-fast);
}
.file-tab:not(.active) .dot { opacity: 0; }

.file-tab .kbd {
  font-family: var(--font-mono);
  font-size: 9px; letter-spacing: 0.04em;
  color: var(--text-tertiary); background: var(--bg-overlay);
  padding: 2px 6px; border-radius: var(--radius-sm);
  border: 1px solid var(--border-subtle);
  transition: color var(--transition-fast), background var(--transition-fast), border-color var(--transition-fast);
}
.file-tab:hover .kbd { color: var(--text-secondary); background: rgba(255,255,255,0.06); }
.file-tab.active .kbd { color: var(--accent); background: var(--accent-dim); border-color: rgba(86,212,160,0.15); }

.spacer { flex: 1; }

/* ---- Hooks badge ---- */
#hooks-badge {
  display: flex; align-items: center; gap: 6px;
  padding: 0 14px;
  font-family: var(--font-mono);
  font-size: 11px; color: var(--text-tertiary); cursor: pointer;
  border-left: 1px solid var(--border-subtle);
  flex-shrink: 0;
  transition: color var(--transition-fast), background var(--transition-fast);
  position: relative; z-index: 200;
}
#hooks-badge:hover { color: var(--text-secondary); background: var(--bg-raised); }
#hooks-badge .hooks-icon { display: flex; align-items: center; opacity: 0.5; }
#hooks-badge:hover .hooks-icon { opacity: 0.8; }
#hooks-badge .hooks-count {
  font-family: var(--font-mono);
  font-size: 10px; font-weight: 500;
  background: var(--bg-overlay); color: var(--text-secondary);
  padding: 2px 7px; border-radius: 10px;
  border: 1px solid var(--border-subtle);
  min-width: 20px; text-align: center;
}

/* ---- Hooks tooltip ---- */
#hooks-tooltip {
  display: none;
  position: fixed;
  top: 42px; right: 12px;
  background: var(--bg-raised);
  border: 1px solid var(--border-medium);
  border-radius: var(--radius-lg); padding: 6px 0;
  min-width: 260px; z-index: 99999;
  box-shadow: var(--shadow-popup);
  font-size: 12px;
  backdrop-filter: blur(12px);
  animation: tooltipIn 0.15s ease-out;
}
@keyframes tooltipIn {
  from { opacity: 0; transform: translateY(-4px); }
  to { opacity: 1; transform: translateY(0); }
}
#hooks-tooltip.visible { display: block; }
#hooks-tooltip .hook-row {
  display: flex; align-items: center;
  padding: 6px 14px; gap: 10px;
  transition: background var(--transition-fast);
}
#hooks-tooltip .hook-row:hover { background: rgba(255,255,255,0.03); }
#hooks-tooltip .hook-event {
  font-family: var(--font-mono);
  font-size: 10px; font-weight: 500;
  color: var(--accent); background: var(--accent-dim);
  padding: 2px 7px; border-radius: var(--radius-sm); white-space: nowrap;
}
#hooks-tooltip .hook-name {
  font-family: var(--font-mono);
  color: var(--text-secondary); white-space: nowrap; font-size: 12px;
}
#hooks-tooltip .hook-none {
  font-family: var(--font-sans);
  padding: 8px 14px; color: var(--text-tertiary); font-style: italic; font-size: 12px;
}

/* ---- Status ---- */
#status {
  display: flex; align-items: center; gap: 7px;
  padding: 0 16px;
  font-family: var(--font-mono);
  font-size: 11px; color: var(--accent); white-space: nowrap;
  letter-spacing: 0.02em;
  flex-shrink: 0;
}
#status .status-dot {
  width: 6px; height: 6px; border-radius: 50%;
  background: var(--accent);
  box-shadow: 0 0 8px var(--accent-glow);
  animation: pulse 2s ease-in-out infinite;
}
@keyframes pulse {
  0%, 100% { opacity: 1; transform: scale(1); }
  50% { opacity: 0.5; transform: scale(0.85); }
}
#status.error { color: #f47067; }
#status.error .status-dot { background: #f47067; box-shadow: 0 0 8px rgba(244,112,103,0.3); animation: none; }
#status.loading { color: var(--text-tertiary); }
#status.loading .status-dot { background: var(--text-tertiary); box-shadow: none; animation: pulse 1s ease-in-out infinite; }

/* ---- Formula bar ---- */
#formula-bar {
  display: flex; align-items: center;
  background: var(--bg-surface);
  border-bottom: 1px solid var(--border-subtle);
  height: 26px;
}
#cell-ref {
  width: 88px; padding: 4px 10px;
  border-right: 1px solid var(--border-subtle);
  font-size: 12px; font-family: var(--font-mono);
  color: var(--accent); background: var(--bg-raised); text-align: center;
  font-weight: 500; letter-spacing: 0.02em;
}
#fx-label {
  padding: 0 10px; color: var(--text-tertiary);
  font-style: italic; font-size: 12px;
  font-family: var(--font-mono);
}
#formula-input {
  flex: 1; height: 100%; border: none; outline: none;
  padding: 0 10px; font-size: 12px;
  font-family: var(--font-mono);
  color: var(--text-primary);
  background: var(--bg-base);
}

#ss { width: 100%; }

/* ---- Empty state ---- */
#empty-state {
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; height: calc(100vh - 34px);
  color: var(--text-tertiary); font-size: 14px; gap: 6px;
  animation: fadeIn 0.4s ease-out;
}
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}
#empty-state .empty-icon {
  width: 64px; height: 64px; margin-bottom: 12px;
  border-radius: 16px;
  background: var(--bg-raised);
  border: 1px solid var(--border-subtle);
  display: flex; align-items: center; justify-content: center;
}
#empty-state .empty-icon svg { width: 28px; height: 28px; stroke: var(--text-tertiary); }
#empty-state .empty-title {
  font-family: var(--font-mono);
  font-size: 15px; font-weight: 500;
  color: var(--text-secondary);
  margin-top: 4px;
}
#empty-state .empty-desc {
  font-family: var(--font-mono);
  font-size: 13px; color: var(--text-tertiary);
  margin-bottom: 4px;
}
#empty-state code {
  font-family: var(--font-mono);
  color: var(--accent); font-size: 12px;
  background: var(--accent-dim);
  padding: 6px 14px; border-radius: var(--radius-md);
  border: 1px solid rgba(86,212,160,0.12);
  letter-spacing: 0.02em;
}
</style>
</head>
<body>

<div id="toolbar">
  <div id="file-tabs"></div>
  <div class="spacer"></div>
  <div id="hooks-badge">
    <span class="hooks-icon"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg></span>
    <span>hooks</span>&nbsp;
    <span class="hooks-count" id="hooks-count">\u2026</span>
  </div>
  <span id="status">loading\u2026</span>
</div>

<div id="formula-bar" style="display:none">
  <div id="cell-ref">A1</div>
  <span id="fx-label">fx</span>
  <input type="text" id="formula-input" readonly>
</div>

<div id="ss" style="display:none"></div>

<div id="hooks-tooltip"></div>

<div id="empty-state">
  <div class="empty-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 16V12"/><path d="M12 16V8"/><path d="M16 16v-2"/></svg></div>
  <div>Waiting for files\u2026</div>
  <div>Open a workbook in another terminal:</div>
  <code>hsx get myfile.xlsx A1</code>
</div>

<script src="https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/dist/gc.spread.sheets.all.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@grapecity/spread-excelio/dist/gc.spread.excelio.min.js"></script>
<script>
// --- Watermark suppression ---
(function(){
  var orig = CanvasRenderingContext2D.prototype.fillText;
  CanvasRenderingContext2D.prototype.fillText = function(text){
    if (text && typeof text === 'string' && (
      text.includes('GrapeCity') || text.includes('MESCIUS') ||
      text.includes('EVALUATION') || text.includes('Powered') ||
      text.includes('deployment') || text.includes('grapecity.com') ||
      text.includes('mescius.com') || text.includes('Email us') ||
      text === 'Evaluation Version'
    )) return;
    return orig.apply(this, arguments);
  };
})();

var FILES = ${fileListJSON};
var activeFile = FILES[0] || null;
var spread = null;
var viewStates = {};
var statusEl = document.getElementById('status');
var cellRefEl = document.getElementById('cell-ref');
var formulaInputEl = document.getElementById('formula-input');
var tabsEl = document.getElementById('file-tabs');
var formulaBarEl = document.getElementById('formula-bar');
var ssEl = document.getElementById('ss');
var emptyEl = document.getElementById('empty-state');

function showSpread() {
  emptyEl.style.display = 'none';
  formulaBarEl.style.display = 'flex';
  ssEl.style.display = 'block';
  if (!spread) initSpread();
  else { layoutSpread(); spread.refresh(); }
}

function showEmpty() {
  emptyEl.style.display = 'flex';
  formulaBarEl.style.display = 'none';
  ssEl.style.display = 'none';
}

function layoutSpread() {
  ssEl.style.height = (window.innerHeight - 34 - 26) + 'px';
}

// --- File tabs ---
function renderTabs() {
  tabsEl.innerHTML = '';
  FILES.forEach(function(name, i) {
    var tab = document.createElement('div');
    tab.className = 'file-tab' + (name === activeFile ? ' active' : '');
    tab.setAttribute('tabindex', name === activeFile ? '0' : '-1');
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-selected', name === activeFile ? 'true' : 'false');
    tab.setAttribute('data-index', i);
    tab.setAttribute('data-file', name);

    var dot = document.createElement('span');
    dot.className = 'dot';
    tab.appendChild(dot);
    tab.appendChild(document.createTextNode(name));

    // Keyboard shortcut badge
    if (i < 9) {
      var kbd = document.createElement('span');
      kbd.className = 'kbd';
      kbd.textContent = '\\u2303' + (i + 1);
      tab.appendChild(kbd);
    }

    tab.onclick = function() { switchFile(name); };

    // Arrow key navigation within tablist
    tab.onkeydown = function(e) {
      var tabs = tabsEl.querySelectorAll('.file-tab');
      var idx = Number(tab.getAttribute('data-index'));
      if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
        e.preventDefault();
        var next = tabs[(idx + 1) % tabs.length];
        next.focus();
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        e.preventDefault();
        var prev = tabs[(idx - 1 + tabs.length) % tabs.length];
        prev.focus();
      } else if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        switchFile(tab.getAttribute('data-file'));
      } else if (e.key === 'Home') {
        e.preventDefault();
        tabs[0].focus();
      } else if (e.key === 'End') {
        e.preventDefault();
        tabs[tabs.length - 1].focus();
      }
    };

    tabsEl.appendChild(tab);
  });
}

function switchFile(name) {
  if (name === activeFile) return;
  saveViewState();
  activeFile = name;
  renderTabs();
  loadFile(name);
}

// --- Ctrl+N keyboard shortcuts ---
document.addEventListener('keydown', function(e) {
  if ((e.ctrlKey || e.metaKey) && e.key >= '1' && e.key <= '9') {
    var idx = parseInt(e.key, 10) - 1;
    if (idx < FILES.length) {
      e.preventDefault();
      switchFile(FILES[idx]);
      var tabs = tabsEl.querySelectorAll('.file-tab');
      if (tabs[idx]) tabs[idx].focus();
    }
  }
});

// --- Col helper ---
function colLetter(col) {
  var s = '';
  while (col >= 0) { s = String.fromCharCode((col % 26) + 65) + s; col = Math.floor(col / 26) - 1; }
  return s;
}

// --- Formula bar ---
function updateFormulaBar() {
  if (!spread) return;
  var sheet = spread.getActiveSheet();
  var r = sheet.getActiveRowIndex(), c = sheet.getActiveColumnIndex();
  cellRefEl.textContent = colLetter(c) + (r + 1);
  var f = sheet.getFormula(r, c);
  formulaInputEl.value = f ? '=' + f : (sheet.getValue(r, c) ?? '');
}

// --- View state ---
function saveViewState() {
  if (!spread || !activeFile) return;
  var sheet = spread.getActiveSheet();
  viewStates[activeFile] = {
    sheetName: sheet.name(),
    sheetIndex: spread.getActiveSheetIndex(),
    row: sheet.getActiveRowIndex(),
    col: sheet.getActiveColumnIndex(),
    topRow: sheet.getViewportTopRow(1),
    leftCol: sheet.getViewportLeftColumn(1),
    sels: sheet.getSelections().map(function(s){
      return { row: s.row, col: s.col, rowCount: s.rowCount, colCount: s.colCount };
    })
  };
}

function restoreViewState(state) {
  if (!spread || !state) return;
  var idx = spread.getSheetIndex(state.sheetName);
  if (idx === -1) idx = Math.min(state.sheetIndex, spread.getSheetCount() - 1);
  spread.setActiveSheetIndex(idx);
  var sheet = spread.getActiveSheet();
  try { sheet.showRow(state.topRow, GC.Spread.Sheets.VerticalPosition.top); } catch(e){}
  try { sheet.showColumn(state.leftCol, GC.Spread.Sheets.HorizontalPosition.left); } catch(e){}
  if (state.sels && state.sels.length) {
    sheet.clearSelection();
    state.sels.forEach(function(sel, i){
      if (i === 0) sheet.setSelection(sel.row, sel.col, sel.rowCount, sel.colCount);
      else sheet.addSelection(sel.row, sel.col, sel.rowCount, sel.colCount);
    });
  }
  sheet.setActiveCell(state.row, state.col);
}

// --- SpreadJS init ---
function initSpread() {
  layoutSpread();
  spread = new GC.Spread.Sheets.Workbook(ssEl);
  spread.bind(GC.Spread.Sheets.Events.SelectionChanged, updateFormulaBar);
  spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, updateFormulaBar);
  window.addEventListener('resize', function(){
    layoutSpread();
    spread.refresh();
  });
}

// --- Load file ---
function loadFile(name) {
  name = name || activeFile;
  if (!name) return;
  statusEl.textContent = 'loading\u2026';
  saveViewState();

  fetch('/file/' + encodeURIComponent(name) + '?t=' + Date.now())
    .then(function(r){ if (!r.ok) throw new Error(r.status); return r.blob(); })
    .then(function(blob){
      showSpread();
      var io = new GC.Spread.Excel.IO();
      io.open(blob, function(json){
        if (json.sheets) {
          Object.values(json.sheets).forEach(function(s){
            s.rowCount = Math.max(s.rowCount || 0, 10000);
            s.columnCount = Math.max(s.columnCount || 0, 500);
          });
        }
        spread.fromJSON(json);
        restoreViewState(viewStates[name]);
        spread.refresh();
        updateFormulaBar();
        statusEl.textContent = new Date().toLocaleTimeString();
      }, function(err){
        statusEl.textContent = 'error: ' + err;
      });
    })
    .catch(function(e){ statusEl.textContent = 'error: ' + e.message; });
}

// --- Hooks info ---
function loadHooksInfo() {
  fetch('/hooks')
    .then(function(r){ return r.json(); })
    .then(function(data){
      document.getElementById('hooks-count').textContent = data.total;
      var tooltip = document.getElementById('hooks-tooltip');
      if (!data.hooks.length) {
        tooltip.innerHTML = '<div class="hook-none">No hooks registered</div>';
        return;
      }
      tooltip.innerHTML = data.hooks.map(function(h){
        return '<div class="hook-row">' +
          '<span class="hook-event">' + h.event + '</span>' +
          '<span class="hook-name">' + h.name + '</span>' +
          '</div>';
      }).join('');
    })
    .catch(function(){ });
}

// --- Hooks tooltip toggle ---
(function(){
  var badge = document.getElementById('hooks-badge');
  var tooltip = document.getElementById('hooks-tooltip');
  badge.addEventListener('click', function(e){
    e.stopPropagation();
    tooltip.classList.toggle('visible');
  });
  document.addEventListener('click', function(){
    tooltip.classList.remove('visible');
  });
  tooltip.addEventListener('click', function(e){ e.stopPropagation(); });
})();

// --- SSE for live reload ---
function connectSSE() {
  var es = new EventSource('/events');
  es.onopen = function() { statusEl.textContent = FILES.length ? 'watching' : 'waiting\u2026'; };
  es.onmessage = function(e) {
    try {
      var msg = JSON.parse(e.data);
      if (msg.type === 'reload') {
        if (!msg.file || msg.file === activeFile) loadFile();
      }
      if (msg.type === 'files') {
        FILES = msg.files;
        renderTabs();
        // If we had no files and now we do, load the first
        if (!activeFile && FILES.length) {
          activeFile = FILES[0];
          renderTabs();
          loadFile();
        }
      }
    } catch(err) {}
  };
  es.onerror = function() {
    statusEl.textContent = 'reconnecting\u2026';
  };
}

// --- Boot ---
requestAnimationFrame(function(){
  renderTabs();
  loadHooksInfo();
  if (FILES.length) {
    loadFile();
  } else {
    showEmpty();
  }
  connectSSE();
});
</script>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// Watch server
// ---------------------------------------------------------------------------

export class WatchServer {
  private server: Server;
  private sseClients = new Set<ServerResponse>();
  private fileNames: string[] = []; // display names
  private fileMap = new Map<string, string>(); // displayName → absPath
  private reverseMap = new Map<string, string>(); // absPath → displayName
  private provider: WatchFileProvider;
  private unsub: (() => void) | null = null;

  constructor(provider: WatchFileProvider) {
    this.provider = provider;
    this.syncFiles(provider.listFiles());
    this.server = createServer(this.handleHttp.bind(this));
  }

  private displayName(absPath: string): string {
    return basename(absPath);
  }

  private syncFiles(absPaths: string[]): boolean {
    // Deduplicate display names
    const counts = new Map<string, number>();
    for (const p of absPaths) {
      const n = this.displayName(p);
      counts.set(n, (counts.get(n) ?? 0) + 1);
    }

    const newNames: string[] = [];
    const newFileMap = new Map<string, string>();
    const newReverseMap = new Map<string, string>();

    for (const p of absPaths) {
      let name = this.displayName(p);
      if ((counts.get(name) ?? 0) > 1) name = p;
      newNames.push(name);
      newFileMap.set(name, p);
      newReverseMap.set(p, name);
    }

    const changed = JSON.stringify(newNames) !== JSON.stringify(this.fileNames);
    this.fileNames = newNames;
    this.fileMap = newFileMap;
    this.reverseMap = newReverseMap;
    return changed;
  }

  async start(port: number): Promise<number> {
    // Subscribe to file events
    this.unsub = this.provider.subscribe((event) => {
      if (event.type === "opened") {
        const files = this.provider.listFiles();
        if (this.syncFiles(files)) {
          sseBroadcast(
            this.sseClients,
            JSON.stringify({ type: "files", files: this.fileNames }),
          );
        }
      } else if (event.type === "changed") {
        const name = this.reverseMap.get(event.absPath);
        if (name) {
          sseBroadcast(
            this.sseClients,
            JSON.stringify({ type: "reload", file: name }),
          );
        }
      }
    });

    return new Promise<number>((resolve, reject) => {
      this.server.on("error", reject);
      this.server.listen(port, () => {
        const addr = this.server.address();
        const actualPort = typeof addr === "object" && addr ? addr.port : port;
        resolve(actualPort);
      });
    });
  }

  stop(): void {
    if (this.unsub) {
      this.unsub();
      this.unsub = null;
    }
    for (const c of this.sseClients) {
      try {
        c.end();
      } catch {}
    }
    this.sseClients.clear();
    this.provider.stop();
    this.server.close();
  }

  private handleHttp(req: IncomingMessage, res: ServerResponse): void {
    const url = req.url ?? "/";

    if (url === "/events") {
      sseConnect(res, this.sseClients);
      return;
    }

    if (url === "/hooks") {
      getHooksSummary()
        .then((summary) => {
          const body = Buffer.from(JSON.stringify(summary), "utf-8");
          res.writeHead(200, {
            "Content-Type": "application/json",
            "Content-Length": body.length,
            "Cache-Control": "no-cache",
          });
          res.end(body);
        })
        .catch(() => {
          res.writeHead(500);
          res.end("Internal error");
        });
      return;
    }

    if (url === "/" || url === "/index.html") {
      const html = buildViewerHTML(this.fileNames);
      const body = Buffer.from(html, "utf-8");
      res.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Length": body.length,
      });
      res.end(body);
      return;
    }

    if (url.startsWith("/file/")) {
      const reqName = decodeURIComponent(url.slice(6).split("?")[0]);
      const absPath = this.fileMap.get(reqName);
      if (absPath) {
        this.provider
          .getBuffer(absPath)
          .then((buf) => {
            if (buf) {
              res.writeHead(200, {
                "Content-Type":
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Length": buf.length,
                "Cache-Control": "no-cache",
              });
              res.end(buf);
            } else {
              res.writeHead(404);
              res.end("Not found");
            }
          })
          .catch(() => {
            res.writeHead(500);
            res.end("Internal error");
          });
        return;
      }
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    res.writeHead(404);
    res.end("Not found");
  }
}
