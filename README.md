# Sheila-App

Receipt PWA for Sheila McGavigan PMU — captures income (In) and expenses (Out)
from a phone, reads them with Groq vision OCR, watermarks the photo, converts
foreign currency to EUR, and files everything into a local Excel workbook with a
live formula-driven dashboard and charts.

## Run

```
npm install        # already done
node server.js     # or: npm start
```

The server starts on **https://localhost:3001** (self-signed cert). On first boot
it builds `her-expenses.xlsx` from the template script automatically.

Phone setup (QR wall): open **https://localhost:3001/setup** and scan with the phone.
Treatments editor: **https://localhost:3001/treatments** (PIN from `.env`).

## Files

| File | What it does |
|------|--------------|
| `server.js` | Express + HTTPS + Tailscale auto-detect. Groq OCR, sharp watermarking, Frankfurter EUR conversion, all `/api` routes, `/setup`, `/cert`. Appends rows by spawning the Python helper. |
| `create-workbook.py` | **One-time** template builder (openpyxl): Out, In, Combined, Dashboard sheets + formulas + 5 native charts. Refuses to overwrite an existing workbook. |
| `append_row.py` | Appends one row to the Out or In sheet, preserving charts + formulas. Called by `server.js` on every save. |
| `public/index.html` | The phone PWA — Out/In toggle, both capture flows, offline queue, GPS. |
| `public/treatments.html` | PIN-gated treatments editor (served at `/treatments`). |
| `public/sw.js` | Service worker — offline app shell + background-sync queue drain. |
| `public/manifest.webmanifest` | PWA manifest. |
| `treatments.json` | Treatment list + prices (local data; git-ignored). |

## Requirements

- **Node 20+** and **Python 3 with openpyxl** must both be on the laptop.
  Set `PYTHON_BIN` in `.env` if `python`/`python3` isn't on PATH.
  Install openpyxl: `pip install openpyxl`

## Important architecture note (the Excel/charts engine)

The spec called for `create-workbook.js` (Node) and for the server to append rows
with the SheetJS `xlsx` library. **That doesn't work**: SheetJS rebuilds the whole
workbook on every write and **destroys charts and live formulas**. Verified by test.

So the Excel layer uses **Python/openpyxl** instead — the only library here that
both writes native charts and preserves them on append. `server.js` shells out to
`append_row.py`. Everything else (server, OCR, watermarking, EUR, PWA) is Node/JS
exactly as specced.

This Excel/charts layer is the agreed **"build now, polish later"** piece — search
the code for `REVISIT` for the spots flagged for a later pass (see "Known
placeholders" below).

## Known placeholders / things to revisit

- **Income-by-Treatment dashboard table** (`create-workbook.py`, Section 4): the
  treatment-name slots (column L, rows 4–21) are left blank because deriving a
  distinct treatment list without Excel dynamic arrays is awkward. The SUMIFS
  formulas are wired and will compute the moment you type/paste the treatment
  names into those slots (or we wire them to `treatments.json` at build time later).
- **"Cumulative net" chart** is currently a **monthly** net line, not a running
  total — a true cumulative needs a helper column. Flagged `REVISIT`.
- **Chart styling** is default openpyxl. Fine to open and recolour, or we polish later.

## EUR conversion

Frankfurter.app, rate for the receipt's **exact date** (`/{date}?from=X&to=EUR`).
EUR receipts are copied as-is. If the API fails or the date is missing, the EUR
cell is left blank and the row's **Status** column is set to `CHECK FX` so it's
easy to filter and fix.

## Automatic updates

The app checks GitHub when it loads and shows a one-tap **Update now** bar when a newer version is available. Tapping it makes the laptop pull the latest code and restart itself - your data (.env, her-expenses.xlsx, her-photos/) is never touched.

