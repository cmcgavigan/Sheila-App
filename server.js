// Sheila McGavigan PMU — Receipt PWA backend
//
// Two flows:
//   OUT (expenses): photo -> Groq OCR -> review -> POST /api/save-out
//   IN  (income):   photo -> Groq OCR -> review -> POST /api/save-in
//
// Each save:
//   1. assigns the next sequential serial for that sheet (Out / In each have their own)
//   2. watermarks the serial onto the top-right of the photo (sharp)
//   3. saves the photo to PHOTOS_DIR/{O|I}{serial}.jpg
//   4. converts the total to EUR for the receipt's exact date (Frankfurter.app)
//   5. appends a row to the workbook via a Python/openpyxl helper (append_row.py)
//      so the Dashboard charts + formulas are PRESERVED on every save.
//
// IMPORTANT — Excel engine note:
//   The workbook (her-expenses.xlsx) is built ONCE by create-workbook.py (openpyxl),
//   which is the only library here that can write native charts + live formulas.
//   The server appends rows through append_row.py (openpyxl) — NOT SheetJS — because
//   SheetJS rewrites the whole workbook on save and DESTROYS charts and formulas.
//   This is the agreed v1 architecture; the Excel layer is flagged to revisit later.

import express from 'express';
import dotenv from 'dotenv';
import sharp from 'sharp';
import QRCode from 'qrcode';
import selfsigned from 'selfsigned';
import https from 'node:https';
import { fileURLToPath } from 'node:url';
import { execFileSync, spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = parseInt(process.env.PORT || '3001', 10);
const GROQ_API_KEY = process.env.GROQ_API_KEY || '';
const GROQ_MODEL = process.env.GROQ_MODEL || 'qwen/qwen3.6-27b'; // Groq retired Llama 4 Scout on 2026-06-17
const EXCEL_PATH = path.resolve(__dirname, process.env.EXCEL_PATH || 'her-expenses.xlsx');
const PHOTOS_DIR = path.resolve(__dirname, process.env.PHOTOS_DIR || 'her-photos');
const TREATMENTS_PATH = path.resolve(__dirname, process.env.TREATMENTS_PATH || 'treatments.json');
const TREATMENTS_PIN = String(process.env.TREATMENTS_PIN || '1234');
const CERT_DIR = path.resolve(__dirname, 'cert');
const CERT_KEY = path.join(CERT_DIR, 'key.pem');
const CERT_PEM = path.join(CERT_DIR, 'cert.pem');
const HOSTNAME = process.env.HOSTNAME_LOCAL || 'sheila.local';

// Python interpreter used for the openpyxl workbook helpers.
const PYTHON = process.env.PYTHON_BIN || (process.platform === 'win32' ? 'python' : 'python3');
const APPEND_SCRIPT = path.join(__dirname, 'append_row.py');
const CREATE_SCRIPT = path.join(__dirname, 'create-workbook.py');

// Out-sheet expense categories. Kept in sync with create-workbook.py and index.html.
const ARTICLES = [
  'Fuel', 'Meals (Travel)', 'Client Entertainment', 'Accommodation', 'Parking',
  'Vehicle expenses', 'Phone & Internet', 'Insurance', 'Training & Education',
  'Marketing & Advertising', 'Equipment & Tools', 'General Business Expenses',
  'PMU Pigments & Inks', 'PMU Needles & Cartridges', 'PMU Equipment & Machines',
  'Skincare & Treatment Products', 'Disposables & PPE', 'Salon Supplies',
];

if (!GROQ_API_KEY) {
  console.error('\n[!] GROQ_API_KEY is not set. Paste your key into .env.\n    Free key (no card): https://console.groq.com/keys\n');
}

fs.mkdirSync(PHOTOS_DIR, { recursive: true });

/* ===================================================================
   Groq vision OCR — shared by both Out and In flows.
   =================================================================== */
const OCR_PROMPT_OUT = [
  'You are a receipt data extractor. Read this EXPENSE receipt photo and reply with ONLY a JSON object',
  '(no markdown, no commentary) using exactly these keys:',
  '- merchant: store/company name (string)',
  '- date: transaction date as YYYY-MM-DD (string)',
  '- totalCost: final total paid, a number with no currency symbol',
  '- currency: ISO code (EUR, CHF, USD, GBP, ...) or "" if unclear',
  '- receiptCode: receipt/transaction/invoice reference, or "" if none',
  '- article: best-fit category, EXACTLY one of: ' + ARTICLES.join(', ') + '.',
  '    Choose the closest fit for a permanent-makeup / beauty business. PMU = permanent makeup.',
  '    Use "General Business Expenses" only if nothing else fits. Always pick one; never empty.',
  '- confidence: number 0.0-1.0, your honest confidence that merchant, date and total are correct.',
  '    Use below 0.5 if the photo is blurry, cropped, dark, glare-washed or hard to read.',
  'If a field cannot be read, use "" for strings or 0 for numbers (but article must always be chosen).',
].join('\n');

const OCR_PROMPT_IN = [
  'You are an invoice/receipt data extractor for a permanent-makeup (PMU) artist who is being PAID by a client.',
  'Read this photo and reply with ONLY a JSON object (no markdown, no commentary) using exactly these keys:',
  '- clientName: the client/customer name if shown, else "" (string)',
  '- date: transaction date as YYYY-MM-DD (string)',
  '- totalCost: final total received, a number with no currency symbol',
  '- currency: ISO code (EUR, CHF, USD, GBP, ...) or "" if unclear',
  '- receiptCode: invoice/receipt reference, or "" if none',
  '- treatment: the service/treatment described, or "" if unclear (string)',
  '- confidence: number 0.0-1.0, your honest confidence that the total and date are correct.',
  'If a field cannot be read, use "" for strings or 0 for numbers.',
].join('\n');

async function groqOcr(base64, mimeType, prompt) {
  if (!GROQ_API_KEY) { const e = new Error('GROQ_API_KEY not configured'); e.status = 500; throw e; }
  const resp = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${GROQ_API_KEY}` },
    body: JSON.stringify({
      model: GROQ_MODEL,
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [{
        role: 'user',
        content: [
          { type: 'text', text: prompt },
          { type: 'image_url', image_url: { url: `data:${mimeType};base64,${base64}` } },
        ],
      }],
    }),
  });
  if (!resp.ok) {
    const t = await resp.text().catch(() => '');
    const e = new Error('Groq HTTP ' + resp.status + ' ' + t.slice(0, 200));
    e.status = resp.status;
    throw e;
  }
  const data = await resp.json();
  const text = (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  return JSON.parse(text);
}

/* ===================================================================
   EUR conversion — Frankfurter.app, rate for the receipt's EXACT date.
   - If currency is already EUR (or empty), copy the total as the EUR value.
   - Weekends/holidays: Frankfurter returns the prior business day's rate.
   - On any failure: return { eur: null, flagged: true } so the row is flagged.
   =================================================================== */
async function toEur(total, currency, date) {
  const amount = Number(total) || 0;
  const cur = String(currency || '').toUpperCase().trim();
  if (!amount) return { eur: 0, rate: 1, rateDate: null, flagged: false };
  if (!cur || cur === 'EUR') return { eur: amount, rate: 1, rateDate: date || null, flagged: false };

  // Validate date; fall back to today if missing/garbage (still flag it).
  const isoDate = /^\d{4}-\d{2}-\d{2}$/.test(String(date || '')) ? date : null;
  const useDate = isoDate || new Date().toISOString().slice(0, 10);
  try {
    const url = `https://api.frankfurter.app/${useDate}?from=${encodeURIComponent(cur)}&to=EUR`;
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), 8000);
    const resp = await fetch(url, { signal: ctl.signal });
    clearTimeout(t);
    if (!resp.ok) throw new Error('Frankfurter HTTP ' + resp.status);
    const j = await resp.json();
    const rate = j && j.rates && j.rates.EUR;
    if (typeof rate !== 'number') throw new Error('No EUR rate in response');
    const eur = Math.round(amount * rate * 100) / 100;
    // Flag if we had to substitute a date (original missing/invalid).
    return { eur, rate, rateDate: j.date || useDate, flagged: !isoDate };
  } catch (e) {
    console.warn('[eur] conversion failed for', cur, useDate, '-', e.message);
    return { eur: null, rate: null, rateDate: null, flagged: true };
  }
}

/* ===================================================================
   App setup
   =================================================================== */
const app = express();
app.use(express.json({ limit: '25mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Clean path for the treatments editor (spec + setup QR use /treatments).
app.get('/treatments', (_req, res) => res.sendFile(path.join(__dirname, 'public', 'treatments.html')));

/* ---------- network helpers (LAN / Tailscale) ---------- */
function getLocalIPs() {
  const out = [];
  for (const name of Object.keys(os.networkInterfaces())) {
    for (const iface of os.networkInterfaces()[name] || []) {
      if (iface.family === 'IPv4' && !iface.internal) out.push({ name, address: iface.address });
    }
  }
  return out;
}
function getTailscaleIp() {
  for (const i of getLocalIPs()) {
    const m = /^100\.(\d+)\./.exec(i.address);
    if (m && Number(m[1]) >= 64 && Number(m[1]) <= 127) return i.address;
  }
  return '';
}
function tailscaleBin() {
  const win = 'C:\\Program Files\\Tailscale\\tailscale.exe';
  if (process.platform === 'win32' && fs.existsSync(win)) return win;
  return 'tailscale';
}
function runTailscale(args) {
  try {
    return execFileSync(tailscaleBin(), args, { timeout: 7000, encoding: 'utf8', windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] });
  } catch (e) { return null; }
}
function tailscaleInfo() {
  const out = runTailscale(['status', '--json']);
  if (!out) return null;
  try {
    const j = JSON.parse(out);
    const dns = ((j.Self && j.Self.DNSName) || '').replace(/\.$/, '');
    const httpsEnabled = Array.isArray(j.CertDomains) && j.CertDomains.length > 0;
    return dns ? { dns, httpsEnabled } : null;
  } catch (e) { return null; }
}
// Sheila is served on a SEPARATE Tailscale HTTPS port (default 8443) so it does
// NOT overwrite another app that already owns :443 on this same machine/hostname
// (the Receipts app uses :443). Each `tailscale serve --https=<port>` maps one
// external HTTPS port to one local port; using a distinct port lets both apps
// share the one .ts.net hostname. Override with TS_HTTPS_PORT in .env.
const TS_HTTPS_PORT = parseInt(process.env.TS_HTTPS_PORT || '8443', 10);
function ensureTailscaleServe(port) {
  const info = tailscaleInfo();
  if (!info) return '';
  if (!info.httpsEnabled) {
    console.log(`  Tailscale:   ${info.dns}  — enable HTTPS at https://login.tailscale.com/admin/dns for a no-warning cert`);
    return '';
  }
  runTailscale(['serve', '--bg', `--https=${TS_HTTPS_PORT}`, `https+insecure://127.0.0.1:${port}`]);
  // 443 is implicit in a URL; any other port must be shown explicitly.
  const suffix = TS_HTTPS_PORT === 443 ? '' : `:${TS_HTTPS_PORT}`;
  return `https://${info.dns}${suffix}`;
}

/* ---------- workbook helpers (Python/openpyxl) ---------- */

// Preflight: the whole Excel layer (build template + append rows) runs through
// Python/openpyxl. If either is missing the workbook silently never builds and
// every Save fails. Check once on boot and print a clear, actionable message.
// Returns true if Python+openpyxl are usable, false otherwise (never throws).
let PYTHON_OK = false;
function preflightPython() {
  // 1. Is the Python interpreter itself runnable?
  let pyVersion = '';
  try {
    pyVersion = execFileSync(PYTHON, ['--version'], { encoding: 'utf8', windowsHide: true }).trim();
  } catch (e) {
    console.error('\n  [!] SETUP NEEDED — Python not found.');
    console.error(`      Tried to run "${PYTHON}" but it failed. The Excel filing needs Python 3.`);
    console.error('      Fix: install Python 3, or set PYTHON_BIN in .env to the right command');
    console.error('           (on Windows it is often "py" instead of "python"):');
    console.error('           PYTHON_BIN=py\n');
    return false;
  }
  // 2. Is openpyxl importable?
  try {
    execFileSync(PYTHON, ['-c', 'import openpyxl'], { windowsHide: true, stdio: 'ignore' });
  } catch (e) {
    console.error('\n  [!] SETUP NEEDED — the "openpyxl" Python package is missing.');
    console.error(`      ${pyVersion} is installed, but the Excel filing needs openpyxl.`);
    console.error('      Fix: run this once, then restart the server:');
    console.error(`           ${PYTHON} -m pip install openpyxl\n`);
    return false;
  }
  PYTHON_OK = true;
  return true;
}

// Ensure the workbook exists; if not, build it once from the template script.
function ensureWorkbook() {
  if (fs.existsSync(EXCEL_PATH)) return;
  if (!PYTHON_OK) {
    console.error('  [!] Skipping workbook build — Python/openpyxl not ready (see message above).');
    return;
  }
  console.log('  Workbook not found — building template via create-workbook.py …');
  try {
    execFileSync(PYTHON, [CREATE_SCRIPT], { cwd: __dirname, stdio: 'inherit' });
  } catch (e) {
    console.error('  [!] Failed to build workbook template:', e.message);
    throw e;
  }
}

// Append a row by calling append_row.py. Payload is passed as JSON on stdin.
// Returns a promise resolving to the helper's JSON result ({ ok, serial, rowCount }).
function appendRow(sheet, row) {
  return new Promise((resolve, reject) => {
    ensureWorkbook();
    const child = spawn(PYTHON, [APPEND_SCRIPT, '--sheet', sheet, '--excel', EXCEL_PATH], {
      cwd: __dirname, stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '', stderr = '';
    child.stdout.on('data', d => { stdout += d; });
    child.stderr.on('data', d => { stderr += d; });
    child.on('error', reject);
    child.on('close', (code) => {
      if (code !== 0) {
        const e = new Error(stderr.trim() || ('append_row.py exited ' + code));
        // Surface the Excel-locked case so the client can show a calm message.
        if (/locked|EBUSY|EPERM|EACCES|being used by another process|Permission denied/i.test(stderr)) e.excelOpen = true;
        return reject(e);
      }
      try { resolve(JSON.parse(stdout.trim() || '{}')); }
      catch (e) { reject(new Error('append_row.py returned non-JSON: ' + stdout.slice(0, 200))); }
    });
    child.stdin.write(JSON.stringify(row));
    child.stdin.end();
  });
}

/* ---------- photo watermarking (sharp) ---------- */
async function watermarkAndSave(imageBuffer, label, outputPath) {
  const meta = await sharp(imageBuffer).metadata();
  const width = meta.width || 1080;
  const height = meta.height || 1080;
  const fontSize = Math.max(36, Math.round(Math.min(width, height) * 0.06));
  const padX = Math.round(fontSize * 0.45);
  const padY = Math.round(fontSize * 0.25);
  const textW = Math.round(label.length * fontSize * 0.62);
  const boxW = textW + padX * 2;
  const boxH = fontSize + padY * 2;
  const offset = Math.round(Math.min(width, height) * 0.025);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
    <g transform="translate(${width - offset - boxW}, ${offset})">
      <rect width="${boxW}" height="${boxH}" fill="black" fill-opacity="0.78" rx="${Math.round(padY)}" />
      <text x="${padX}" y="${fontSize + padY * 0.4}" font-family="Arial Black,Helvetica,sans-serif" font-weight="900" font-size="${fontSize}" fill="white" letter-spacing="2">${label}</text>
    </g>
  </svg>`;
  await sharp(imageBuffer).rotate().composite([{ input: Buffer.from(svg), top: 0, left: 0 }]).jpeg({ quality: 90 }).toFile(outputPath);
}

// Build a clean "receipt card" image from typed-in details, for manual entries
// that have no photo (e.g. a bank transfer). Keeps every Excel row visual and
// consistent — the # cell always links to an image. Black & gold to match the app.
function escapeXml(s) {
  return String(s == null ? '' : s).replace(/[<>&'"]/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}
async function generateReceiptImage(label, fields, outputPath) {
  const W = 900, H = 1200;
  const gold = '#c9a24b', goldLt = '#e7c873', ink = '#f4efe6', dim = '#8c8478', bg = '#0c0c0d', line = '#3a342a';
  // fields: ordered array of { k, v }
  const startY = 360, rowH = 92;
  const rowsSvg = fields.filter(f => f && (f.v || f.v === 0)).map((f, i) => {
    const y = startY + i * rowH;
    return `<text x="70" y="${y}" font-family="Helvetica,Arial,sans-serif" font-size="26" fill="${dim}" letter-spacing="2" style="text-transform:uppercase">${escapeXml(String(f.k).toUpperCase())}</text>
      <text x="70" y="${y + 42}" font-family="Helvetica,Arial,sans-serif" font-size="40" font-weight="600" fill="${ink}">${escapeXml(f.v)}</text>
      <line x1="70" y1="${y + 62}" x2="${W - 70}" y2="${y + 62}" stroke="${line}" stroke-width="1"/>`;
  }).join('\n');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="${bg}"/>
    <rect x="24" y="24" width="${W - 48}" height="${H - 48}" rx="36" fill="none" stroke="${line}" stroke-width="2"/>
    <text x="${W / 2}" y="150" text-anchor="middle" font-family="Helvetica,Arial,sans-serif" font-size="32" font-weight="600" fill="${goldLt}" letter-spacing="8">SHEILA McGAVIGAN</text>
    <text x="${W / 2}" y="195" text-anchor="middle" font-family="Helvetica,Arial,sans-serif" font-size="20" fill="${dim}" letter-spacing="8" style="text-transform:uppercase">Permanent Makeup</text>
    <line x1="70" y1="250" x2="${W - 70}" y2="250" stroke="${gold}" stroke-width="2"/>
    <text x="70" y="315" font-family="Helvetica,Arial,sans-serif" font-size="30" font-weight="600" fill="${gold}" letter-spacing="3" style="text-transform:uppercase">Entered manually</text>
    ${rowsSvg}
    <text x="${W - 70}" y="${H - 70}" text-anchor="end" font-family="Arial Black,Helvetica,sans-serif" font-weight="900" font-size="64" fill="${gold}" letter-spacing="4">${escapeXml(label)}</text>
  </svg>`;
  await sharp(Buffer.from(svg)).jpeg({ quality: 92 }).toFile(outputPath);
}

function fmtCoord(n) { return Number(n).toFixed(6); }
function googleMapsUrl(lat, lng) { return `https://www.google.com/maps?q=${lat},${lng}`; }
function nowLocalStamp() {
  const d = new Date(); const p = (n) => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/* ===================================================================
   /api/process-receipt — Groq OCR. mode=out (default) or in.
   =================================================================== */
app.post('/api/process-receipt', async (req, res) => {
  try {
    const { image, mimeType = 'image/jpeg', mode = 'out' } = req.body || {};
    if (!image || typeof image !== 'string') return res.status(400).json({ error: 'Missing "image" (base64 string)' });
    if (!GROQ_API_KEY) return res.status(500).json({ error: 'GROQ_API_KEY not configured on server' });
    const base64 = image.includes(',') ? image.split(',', 2)[1] : image;
    const prompt = mode === 'in' ? OCR_PROMPT_IN : OCR_PROMPT_OUT;

    let parsed;
    try {
      parsed = await groqOcr(base64, mimeType, prompt);
    } catch (err) {
      if (err && err.status === 429) {
        let retryAfter = 0;
        try { const m = /try again in ([\d.]+)s/i.exec(err.message || ''); if (m) retryAfter = Math.ceil(parseFloat(m[1])); } catch (_) {}
        return res.status(429).json({ error: 'quota', quota: true, retryAfter });
      }
      console.error('process-receipt (groq) error:', err.message);
      return res.status(502).json({ error: err.message || 'OCR failed' });
    }

    const num = (v) => (typeof v === 'number' ? v : Number(v) || 0);
    const conf = (v) => (typeof v === 'number' ? v : (v != null ? Number(v) : null));
    if (mode === 'in') {
      return res.json({
        clientName: parsed.clientName || '',
        date: parsed.date || '',
        totalCost: num(parsed.totalCost),
        currency: parsed.currency || '',
        receiptCode: parsed.receiptCode || '',
        treatment: parsed.treatment || '',
        confidence: conf(parsed.confidence),
      });
    }
    const article = ARTICLES.includes(parsed.article) ? parsed.article : 'General Business Expenses';
    res.json({
      merchant: parsed.merchant || '',
      date: parsed.date || '',
      totalCost: num(parsed.totalCost),
      currency: parsed.currency || '',
      receiptCode: parsed.receiptCode || '',
      article,
      confidence: conf(parsed.confidence),
    });
  } catch (err) {
    console.error('process-receipt error:', err);
    res.status(500).json({ error: err.message || 'Unknown error' });
  }
});

/* ===================================================================
   Shared save helper — watermark photo, convert EUR, append row.
   =================================================================== */
async function nextSerialFor(sheet) {
  // Ask the Python helper for the current row count by appending nothing? No —
  // serials are assigned inside append_row.py (atomic with the write). The server
  // does not need to pre-count. We pass the photo prefix and the helper returns the serial.
  return null;
}

async function handleSave(mode, req, res) {
  try {
    const body = req.body || {};
    const { image, lat, lng } = body;
    // Image is OPTIONAL: income (and expenses) can be logged manually with no
    // photo — e.g. a bank transfer with no receipt. hasImage drives whether we
    // watermark/save a photo and hyperlink the row's # cell to it.
    const hasImage = typeof image === 'string' && image.length > 0;

    const date = body.date || '';
    const total = Number(body.totalCost) || 0;
    const currency = String(body.currency || '').toUpperCase().trim();
    const receiptCode = String(body.receiptCode || '');
    const note = String(body.note || '');
    const location = (lat != null && lng != null && lat !== '' && lng !== '' && !Number.isNaN(Number(lat)) && !Number.isNaN(Number(lng)))
      ? `${fmtCoord(lat)}, ${fmtCoord(lng)}` : '';
    const mapUrl = location ? googleMapsUrl(fmtCoord(lat), fmtCoord(lng)) : '';

    if (mode === 'out') {
      if (!body.merchant && !date && !total && !receiptCode) return res.status(400).json({ error: 'Nothing to save' });
    } else {
      if (!body.clientName && !date && !total && !receiptCode) return res.status(400).json({ error: 'Nothing to save' });
    }

    // EUR conversion for the exact date.
    const { eur, flagged } = await toEur(total, currency, date);

    // Build the row payload for append_row.py. The helper assigns the serial and
    // photo filename prefix, and returns them so we know where to write the photo.
    const prefix = mode === 'out' ? 'O' : 'I';
    const common = {
      date, total, currency,
      eur, // may be null -> helper writes blank + flags
      receiptCode, note, location, mapUrl,
      status: flagged ? 'CHECK FX' : 'OK',
      capturedAt: body.capturedAt || '',
      savedAt: nowLocalStamp(),
      photoPrefix: prefix,
      hasPhoto: true,   // every row gets an image (real photo OR a generated card)
    };
    let row;
    if (mode === 'out') {
      row = { ...common, merchant: String(body.merchant || ''), article: String(body.article || 'General Business Expenses'), businessPersonal: (body.businessPersonal === 'Personal' ? 'Personal' : 'Business') };
    } else {
      row = { ...common, clientName: String(body.clientName || ''), treatment: String(body.treatment || '') };
    }

    const result = await appendRow(mode, row);
    const serial = result.serial;
    const photoName = `${prefix}${serial}.jpg`;
    const photoPath = path.join(PHOTOS_DIR, photoName);
    const label = `${prefix}${serial}`;

    if (hasImage) {
      // Real photo → watermark the serial onto it.
      const base64 = image.includes(',') ? image.split(',', 2)[1] : image;
      const imageBuffer = Buffer.from(base64, 'base64');
      await watermarkAndSave(imageBuffer, label, photoPath);
    } else {
      // Manual entry → generate a tidy receipt card from the typed details so
      // the row still has a consistent image to link to.
      const cardFields = mode === 'out'
        ? [
            { k: 'Merchant', v: body.merchant }, { k: 'Category', v: body.article },
            { k: 'Amount', v: (total ? `${total.toFixed(2)} ${currency || ''}`.trim() : '') },
            { k: 'Date', v: date }, { k: 'Type', v: row.businessPersonal },
            { k: 'Receipt', v: receiptCode }, { k: 'Note', v: note },
          ]
        : [
            { k: 'Client', v: body.clientName }, { k: 'Treatment', v: body.treatment },
            { k: 'Amount', v: (total ? `${total.toFixed(2)} ${currency || ''}`.trim() : '') },
            { k: 'Date', v: date }, { k: 'Receipt', v: receiptCode }, { k: 'Note', v: note },
          ];
      await generateReceiptImage(label, cardFields, photoPath);
    }

    res.json({ ok: true, serial, photo: photoName, generated: !hasImage, eur, fxFlagged: flagged, rowCount: result.rowCount, file: EXCEL_PATH });
  } catch (err) {
    console.error(`save-${mode} error:`, err);
    if (err.excelOpen) {
      return res.status(423).json({ error: 'excel-open', message: 'her-expenses.xlsx is open in Excel — close it and this will file automatically.' });
    }
    res.status(500).json({ error: err.message || 'Unknown error' });
  }
}

app.post('/api/save-out', (req, res) => handleSave('out', req, res));
app.post('/api/save-in', (req, res) => handleSave('in', req, res));

/* ===================================================================
   Treatments — GET list, POST update (PIN), POST verify (PIN gate).
   =================================================================== */
function readTreatments() {
  try {
    if (!fs.existsSync(TREATMENTS_PATH)) return [];
    const j = JSON.parse(fs.readFileSync(TREATMENTS_PATH, 'utf8'));
    return Array.isArray(j) ? j : [];
  } catch (e) { console.warn('[treatments] read failed:', e.message); return []; }
}
function writeTreatments(list) {
  const clean = (Array.isArray(list) ? list : [])
    .map(t => ({ name: String(t.name || '').trim(), price: Number(t.price) || 0 }))
    .filter(t => t.name);
  fs.writeFileSync(TREATMENTS_PATH, JSON.stringify(clean, null, 2), 'utf8');
  return clean;
}

app.get('/api/treatments', (_req, res) => res.json({ treatments: readTreatments() }));

app.post('/api/treatments/verify', (req, res) => {
  const pin = String((req.body && req.body.pin) || '');
  res.json({ ok: pin === TREATMENTS_PIN });
});

app.post('/api/treatments', (req, res) => {
  const pin = String((req.body && req.body.pin) || '');
  if (pin !== TREATMENTS_PIN) return res.status(403).json({ error: 'Wrong PIN' });
  const list = (req.body && req.body.treatments) || [];
  if (!Array.isArray(list)) return res.status(400).json({ error: 'treatments must be an array' });
  try {
    const saved = writeTreatments(list);
    res.json({ ok: true, treatments: saved });
  } catch (e) {
    res.status(500).json({ error: e.message || 'Failed to save treatments' });
  }
});

/* ===================================================================
   Misc endpoints
   =================================================================== */
app.get('/api/health', (_req, res) => res.json({ ok: true, excelReady: PYTHON_OK && fs.existsSync(EXCEL_PATH), pythonOk: PYTHON_OK, excelPath: EXCEL_PATH, photosDir: PHOTOS_DIR }));
app.get('/api/articles', (_req, res) => res.json({ articles: ARTICLES }));

/* ---------- In-app auto-update (git installs only) ----------------------------
   /api/update-check : is a newer version on GitHub? (git fetch + compare)
   /api/update       : git pull --ff-only, then restart so the new code loads.
   Branch-agnostic (uses the checked-out branch's origin ref). Installer copies
   without .git get {updateAvailable:false}. Data files (.env, her-expenses.xlsx,
   her-photos/) are gitignored, so pulling never touches them. */
const LOCK_PATH = path.join(__dirname, 'package-lock.json');
function gitCmd(args, timeout = 20000) {
  try {
    return execFileSync('git', ['-C', __dirname, ...args], {
      encoding: 'utf8', timeout, windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
  } catch (_) { return null; }
}
function isGitInstall() { return fs.existsSync(path.join(__dirname, '.git')); }
function currentBranch() { return gitCmd(['rev-parse', '--abbrev-ref', 'HEAD']) || 'master'; }
let lastFetch = 0;
function fetchRemote() {
  const now = Date.now();
  if (now - lastFetch < 60000) return;
  lastFetch = now;
  gitCmd(['fetch', '--quiet', '--tags', 'origin']);
}
function updateStatus() {
  if (!isGitInstall()) return { mode: 'non-git', updateAvailable: false };
  fetchRemote();
  const ref = 'origin/' + currentBranch();
  const behind = parseInt(gitCmd(['rev-list', '--count', 'HEAD..' + ref]) || '0', 10) || 0;
  const current = gitCmd(['describe', '--tags', '--always']) || 'unknown';
  const latest = gitCmd(['describe', '--tags', '--always', ref]) || current;
  return { mode: 'git', current, latest, behind, updateAvailable: behind > 0 };
}
app.get('/api/update-check', (_req, res) => {
  try { res.json(updateStatus()); }
  catch (e) { res.json({ mode: 'error', updateAvailable: false, error: e.message }); }
});
app.post('/api/update', (_req, res) => {
  if (!isGitInstall()) return res.status(400).json({ error: 'Not a git install — update manually.' });
  fetchRemote();
  const branch = currentBranch();
  const ref = 'origin/' + branch;
  const behind = parseInt(gitCmd(['rev-list', '--count', 'HEAD..' + ref]) || '0', 10) || 0;
  if (behind <= 0) return res.json({ ok: true, updated: false, message: 'Already up to date.' });

  const lockBefore = fs.existsSync(LOCK_PATH) ? fs.readFileSync(LOCK_PATH, 'utf8') : '';
  const pulled = gitCmd(['pull', '--ff-only', 'origin', branch]);
  if (pulled === null) {
    return res.status(500).json({ error: 'git pull failed (local edits or conflict). Update manually: git pull.' });
  }
  const lockAfter = fs.existsSync(LOCK_PATH) ? fs.readFileSync(LOCK_PATH, 'utf8') : '';
  const needNpm = lockBefore !== lockAfter;

  res.json({ ok: true, updated: true, restarting: true, npm: needNpm });
  setTimeout(() => {
    try {
      const helper = path.join(__dirname, 'update-restart.cmd');
      if (process.platform === 'win32' && fs.existsSync(helper)) {
        spawn('cmd', ['/c', 'start', '', '/min', helper, needNpm ? '1' : '0'],
          { cwd: __dirname, detached: true, stdio: 'ignore', windowsHide: true }).unref();
      }
    } catch (_) {}
    process.exit(0);
  }, 700);
});

app.get('/cert', (_req, res) => {
  if (!fs.existsSync(CERT_PEM)) return res.status(404).send('Certificate not generated yet.');
  res.setHeader('Content-Type', 'application/x-x509-ca-cert');
  res.setHeader('Content-Disposition', 'attachment; filename="Sheila.crt"');
  fs.createReadStream(CERT_PEM).pipe(res);
});

/* ---------- /setup — three-QR wall ---------- */
app.get('/setup', async (_req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  const lanIps = getLocalIPs();
  const lanIp = (lanIps.find(i => /^(192\.168\.|10\.|172\.(1[6-9]|2\d|3[01])\.)/.test(i.address)) || lanIps[0] || {}).address || '';
  const addr = process.env.TAILSCALE_HOST || getTailscaleIp() || lanIp || HOSTNAME;
  const base = (PUBLIC_URL || `https://${addr}:${PORT}`).replace(/\/+$/, '');
  const onTailscale = /(^https:\/\/[^/]+\.ts\.net)/i.test(base);
  const TS_PLAY = 'https://play.google.com/store/apps/details?id=com.tailscale.ipn';
  const qrData = async (s) => { try { return await QRCode.toDataURL(s, { margin: 1, width: 230 }); } catch (e) { return ''; } };
  const [tsQr, appQr, treatQr] = await Promise.all([qrData(TS_PLAY), qrData(base), qrData(base + '/treatments')]);
  const card = (n, title, body) => `<div class="card"><div class="num">${n}</div><h2>${title}</h2>${body}</div>`;
  const qrImg = (d) => `<img src="${d}" alt="QR" width="230" height="230">`;

  const tsCard = card('1', 'Install Tailscale on the phone',
    `<p>Scan to get Tailscale, then <b>sign in with the SAME account as this laptop</b> and toggle it <b>On</b>. This lets the phone reach the laptop on any network — abroad, mobile data, hotel Wi-Fi.</p>${qrImg(tsQr)}<p class="sub">iPhone: search “Tailscale” in the App Store.</p>`);
  const appCard = card('2', 'Open Sheila’s Receipts app',
    `<p>Scan to open the app${onTailscale ? ' (secure, no warnings)' : ''}. Then Safari/Chrome menu → <b>Add to Home Screen</b>, and tap <b>Allow</b> if it asks about notifications.</p>${qrImg(appQr)}<p class="sub">In/Out toggle is at the top. No key, no typing.</p>`);
  const treatCard = card('3', 'Treatments editor (optional)',
    `<p>Scan to open the treatments list editor. Enter the PIN to add, edit or delete treatments and their prices.</p>${qrImg(treatQr)}<p class="sub">Add this to the Home Screen too if you’ll edit prices often.</p>`);

  const certNote = onTailscale ? '' :
    `<div class="warn">Tailscale isn’t serving a certificate yet, so the phone may show a “not secure” warning. Finish Tailscale setup on the laptop (sign in + turn on HTTPS) and reload — the warning disappears.</div>`;

  res.send(`<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sheila — phone setup</title>
<style>
:root{--ink:#0f1222;--muted:#6b7280;--line:#e5e7eb;--accent:#b6356b}
*{box-sizing:border-box}
body{font:16px/1.55 -apple-system,Segoe UI,Roboto,sans-serif;margin:0;color:var(--ink);background:#fdf2f8}
.wrap{max-width:60rem;margin:0 auto;padding:1.5rem 1.25rem 3rem}
h1{font-size:1.5rem;margin:.2rem 0}.lead{color:var(--muted);margin:.2rem 0 1.4rem}
.grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:1rem}
@media(max-width:820px){.grid{grid-template-columns:1fr}}
.card{position:relative;background:#fff;border:1px solid var(--line);border-radius:16px;padding:1.2rem 1.2rem 1.4rem}
.card .num{position:absolute;top:-12px;left:16px;width:30px;height:30px;border-radius:50%;background:var(--accent);color:#fff;font-weight:700;display:flex;align-items:center;justify-content:center}
.card h2{font-size:1.05rem;margin:.4rem 0 .5rem}
.card p{margin:.4rem 0}.card .sub{color:var(--muted);font-size:.85rem}
.card img{display:block;margin:.6rem auto 0;border-radius:12px;background:#fff;padding:6px;border:1px solid var(--line)}
.warn{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;border-radius:12px;padding:.8rem 1rem;margin:1rem 0}
.done{text-align:center;color:var(--muted);margin-top:1.6rem;font-size:.9rem}
.addr{font:13px ui-monospace,Consolas,monospace;color:var(--muted);word-break:break-all}
</style>
<div class="wrap">
<h1>💅 Sheila — phone setup</h1>
<p class="lead">Scan these codes with the phone. The whole setup takes about a minute.</p>
${certNote}
<div class="grid">${tsCard}${appCard}${treatCard}</div>
<p class="done">App address: <span class="addr">${base}</span><br>After setup, always open from the Home-Screen icon.</p>
</div>`);
});

/* ===================================================================
   TLS cert — generate once, reuse.
   =================================================================== */
function loadOrGenerateCert(ipsForSan, hostnames = []) {
  fs.mkdirSync(CERT_DIR, { recursive: true });
  if (fs.existsSync(CERT_KEY) && fs.existsSync(CERT_PEM)) {
    try {
      const existingCert = fs.readFileSync(CERT_PEM, 'utf8');
      const allCovered = ipsForSan.every(ip => existingCert.includes(ip)) && hostnames.every(h => existingCert.includes(h));
      if (allCovered) return { key: fs.readFileSync(CERT_KEY), cert: fs.readFileSync(CERT_PEM), generated: false };
      console.log('  LAN IPs / hostname changed - regenerating TLS cert...');
    } catch (e) { console.log('  Existing cert unreadable, regenerating...'); }
  }
  const altNames = [
    { type: 2, value: 'localhost' },
    ...hostnames.map(h => ({ type: 2, value: h })),
    { type: 7, ip: '127.0.0.1' },
    ...ipsForSan.map(ip => ({ type: 7, ip })),
  ];
  const pems = selfsigned.generate(
    [{ name: 'commonName', value: HOSTNAME }],
    { keySize: 2048, days: 825, algorithm: 'sha256', extensions: [
      { name: 'subjectAltName', altNames },
      { name: 'basicConstraints', cA: false },
      { name: 'keyUsage', digitalSignature: true, keyEncipherment: true },
      { name: 'extKeyUsage', serverAuth: true },
    ] }
  );
  fs.writeFileSync(CERT_KEY, pems.private);
  fs.writeFileSync(CERT_PEM, pems.cert);
  return { key: pems.private, cert: pems.cert, generated: true };
}

const ips = getLocalIPs();
const ipStrings = ips.map(i => i.address);
const TAILSCALE_HOST = process.env.TAILSCALE_HOST || getTailscaleIp();
const tsIsIp = /^\d+\.\d+\.\d+\.\d+$/.test(TAILSCALE_HOST);
const certIps = (TAILSCALE_HOST && tsIsIp) ? [...ipStrings, TAILSCALE_HOST] : ipStrings;
const certHostnames = (TAILSCALE_HOST && !tsIsIp) ? [HOSTNAME, TAILSCALE_HOST] : [HOSTNAME];
const tls = loadOrGenerateCert(certIps, certHostnames);

const PUBLIC_URL = process.env.PUBLIC_URL || ensureTailscaleServe(PORT);
const SETUP_BASE = PUBLIC_URL || `https://localhost:${PORT}`;
try { fs.writeFileSync(path.join(__dirname, 'current-url.txt'), SETUP_BASE, 'utf8'); } catch (_) {}

preflightPython();
try { ensureWorkbook(); } catch (_) { /* logged inside */ }

https.createServer({ key: tls.key, cert: tls.cert }, app).listen(PORT, '0.0.0.0', () => {
  console.log('\n  Sheila PMU - Receipts PWA running (HTTPS)');
  console.log('  -----------------------------------------');
  console.log(`  Excel file:  ${EXCEL_PATH}`);
  console.log(`  Photos:      ${PHOTOS_DIR}`);
  console.log(`  Treatments:  ${TREATMENTS_PATH}`);
  console.log(`  TLS cert:    ${CERT_PEM}` + (tls.generated ? '  (just generated)' : '  (reused)'));
  console.log(`  On laptop:   https://localhost:${PORT}`);
  console.log(`  On phone:    https://${HOSTNAME}:${PORT}`);
  if (TAILSCALE_HOST) console.log(`  Anywhere:    https://${TAILSCALE_HOST}:${PORT}    (Tailscale)`);
  if (PUBLIC_URL) console.log(`  Public URL:  ${PUBLIC_URL}    (valid cert, used by setup QR)`);
  for (const { name, address } of ips) console.log(`  fallback:    https://${address}:${PORT}    (${name})`);
  console.log(`\n  Phone setup (scan the QR wall):  ${SETUP_BASE}/setup`);
  console.log('  Treatments editor:               ' + SETUP_BASE + '/treatments\n');
});
