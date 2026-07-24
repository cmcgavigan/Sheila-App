// tray.js — Sheila server tray launcher (GOLD icon).
//
// Same approach as the Receipts app: talks directly to the systray2 Go binary
// over stdin/stdout, bypassing the broken systray2 JS wrapper (Node 22+ issue).
// The GOLD icon lets you tell Sheila's server apart from your own Receipts one
// (which uses a GREEN icon), at a glance in the system tray.
// in the system tray at a glance.

import { spawn, execSync } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import readline from 'node:readline';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ─── Config ────────────────────────────────────────────────────────────────
const NODE_EXE = fs.existsSync(path.join(__dirname, 'node', 'node.exe'))
  ? path.join(__dirname, 'node', 'node.exe')
  : process.execPath;

const SERVER_JS  = path.join(__dirname, 'server.js');
const EXCEL_PATH = path.join(__dirname, process.env.EXCEL_PATH || 'her-expenses.xlsx');
const URL_FILE   = path.join(__dirname, 'current-url.txt');
const TRAY_BIN   = path.join(__dirname, 'node_modules', 'systray2', 'traybin', 'tray_windows_release.exe');

// ─── Icon (GOLD) ─────────────────────────────────────────────────────────────
// Pink rounded-square ICO. Written to a temp file then read back as base64,
// exactly as the systray2 JS wrapper would via loadIcon().
const ICO_B64 = 'AAABAAEAICAAAAAAIAD+AAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAgAAAAIAgGAAAAc3p69AAAAMVJREFUeJxjYBhgwIhLgoeH9z81Lfry5TNWuzAEqW0xIYcw0dNybHYw4ZKglyOY8CmkB2BiYKCv72EAZicLOZr3TrPDKeecdYgksxiRXUOJxeQ45MuXz4xEpwFSLCdF/eBIhEPCAaQmLmLVk5QIYYBauYCkREgrMOqAUQeMOmDAHUBWdSynpoFHlrQSk+QQuL0njSJ5ih1AbTA4HICr04ANqLrMokgeBmB2wi2md8MU5gAmdAF6Wo7iAHo5At2OAe+cDjgAAKT5Rmn4MYHmAAAAAElFTkSuQmCC';
const ICON_PATH = path.join(os.tmpdir(), 'sheila-tray.ico');
fs.writeFileSync(ICON_PATH, Buffer.from(ICO_B64, 'base64'));
const ICON_DATA = fs.readFileSync(ICON_PATH).toString('base64');

// ─── Menu ──────────────────────────────────────────────────────────────────
const MENU = {
  icon: ICON_DATA,
  title: '',
  tooltip: 'Sheila Server (gold)',
  items: [
    { title: 'Open Her App',     tooltip: "Open Sheila's app in your browser", checked: false, enabled: true, __id: 1 },
    { title: 'Open Her Excel',   tooltip: 'Open her-expenses.xlsx',             checked: false, enabled: true, __id: 2 },
    { title: 'Setup Page (QRs)', tooltip: 'Open the phone setup QR wall',       checked: false, enabled: true, __id: 3 },
    { title: '<SEPARATOR>',      tooltip: '', checked: false, enabled: true, __id: 4 },
    { title: 'Stop Server',      tooltip: 'Shut down the server and exit',      checked: false, enabled: true, __id: 5 },
  ],
};
const ITEMS = { OPEN_APP: 1, OPEN_XL: 2, SETUP: 3, STOP: 5 };

// ─── Helpers ────────────────────────────────────────────────────────────────
function getAppUrl() {
  try { return fs.readFileSync(URL_FILE, 'utf8').trim(); }
  catch (_) { return 'https://localhost:3001'; }
}
function openUrl(url) { try { execSync(`start "" "${url}"`); } catch (_) {} }
function openExcel()  { try { execSync(`start "" "${EXCEL_PATH}"`); } catch (_) {} }

// ─── Start the server ────────────────────────────────────────────────────────
console.log('[tray] Starting Sheila server...');
const server = spawn(NODE_EXE, [SERVER_JS], {
  cwd: __dirname, detached: false, stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env },
});
server.stdout.on('data', d => process.stdout.write(d));
server.stderr.on('data', d => process.stderr.write(d));
server.on('exit', (code) => { console.log(`[tray] Server exited (code ${code}).`); killTray(); process.exit(code ?? 0); });

// ─── Start tray binary ───────────────────────────────────────────────────────
const tray = spawn(TRAY_BIN, [], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
tray.stderr.on('data', d => console.error('[tray-bin stderr]', d.toString().trim()));
tray.on('exit', (code) => { console.log(`[tray] Binary exited (code ${code}).`); killServer(); process.exit(0); });

const rl = readline.createInterface({ input: tray.stdout, crlfDelay: Infinity });
rl.on('line', (line) => {
  let action;
  try { action = JSON.parse(line); } catch (_) { return; }
  if (action.type === 'ready') {
    tray.stdin.write(JSON.stringify(MENU) + '\n');
    console.log('[tray] Menu sent (pink icon).');
  }
  if (action.type === 'clicked') {
    const id = action.item?.__id;
    switch (id) {
      case ITEMS.OPEN_APP: openUrl(getAppUrl()); break;
      case ITEMS.OPEN_XL:  openExcel(); break;
      case ITEMS.SETUP:    openUrl(getAppUrl() + '/setup'); break;
      case ITEMS.STOP:     killAll(); break;
    }
  }
});

// ─── Shutdown ─────────────────────────────────────────────────────────────────
function killTray()   { try { tray.stdin.write('{"type":"exit"}\n'); } catch (_) {} try { tray.kill(); } catch (_) {} }
function killServer() { try { server.kill('SIGTERM'); } catch (_) {} }
function killAll()    { console.log('[tray] Stopping...'); killTray(); killServer(); setTimeout(() => process.exit(0), 1500); }
process.on('SIGINT',  killAll);
process.on('SIGTERM', killAll);
