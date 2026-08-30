# DataVision AI

MVP prototype: upload Excel/CSV → auto column detection & KPIs → chart recommendations →
customizable charts → drag-ordered report builder → print-to-PDF export.

## Run locally

```bash
npm install
npm run dev
```

Then open the URL Vite prints (usually http://localhost:5173).

## Deploy

### 1. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/datavision-ai.git
git push -u origin main
```

### 2. Deploy on Vercel

1. Go to vercel.com → **Add New Project**
2. Import the GitHub repo you just pushed
3. Vercel auto-detects the Vite framework — no config needed
4. Click **Deploy**

You'll get a live URL like `datavision-ai.vercel.app`.

## Notes

- Everything runs client-side: file parsing (PapaParse/SheetJS), charts (Recharts), and PDF
  export (browser print dialog) all happen in the browser — no backend required for the MVP.
- Google Sheets live import is stubbed in the UI (see the "New Analysis" import screen) — wiring
  it up for real requires a small backend endpoint using the Google Sheets API, since Google
  doesn't allow that call directly from the browser for private sheets.
- The in-app AI executive summary feature was removed — insights and KPIs are still fully
  computed from the uploaded data, just without an LLM-written narrative on top.
