# MediaPulse — AI-Powered Media Performance Analytics

A browser-based analytics dashboard for performance marketers. Upload campaign data from Meta, Google, DV360, or Programmatic platforms and get interactive charts, AI-powered insights, and ready-to-share reports.

**No server. No database. 100% client-side.**

## Quick Start

```bash
npm install
npm run dev
```

Open `http://localhost:5173` in your browser.

## Deploy

```bash
npm run build
```

The `dist/` folder is a static site — deploy to Vercel, Netlify, GitHub Pages, or any static host.

### Deploy to Vercel (easiest)

1. Push this repo to GitHub
2. Go to [vercel.com](https://vercel.com) → New Project → Import your repo
3. Framework: Vite → Deploy

### Deploy to GitHub Pages

1. Install: `npm install -D gh-pages`
2. Add to `package.json` scripts: `"deploy": "vite build --base=/REPO_NAME/ && gh-pages -d dist"`
3. Run: `npm run deploy`

## Features

### 1. Data Upload
- Drag & drop CSV or Excel files (XLSX/XLS/TSV)
- Built-in demo data for instant exploration
- Shows data preview with row/column counts

### 2. Column Mapping
- Auto-detects common column names (Spend, Cost, Impressions, etc.)
- Manual override for non-standard headers
- Campaign objective & success metric configuration
- Platform auto-detection

### 3. Performance Dashboard
- KPI cards: Total Spend, Installs, CPI, CTR, ROAS
- Spend & Installs trend chart (dual-axis)
- Device split donut chart
- Platform performance comparison bars
- CPI by platform
- Campaign breakdown table with color-coded CPI

### 4. Analytics Builder
- Custom pivot tables — pick any dimension × metric combination
- Dimensions: Campaign, Platform, Device, City, Placement, Date, Objective
- Metrics: Spend, Impressions, Clicks, Installs, CPI, CTR, CPC, ROAS, CVR
- Sortable results table

### 5. AI Insights
- Automated CPI efficiency analysis per platform
- Device performance comparison
- Budget alert for underperforming platforms
- Objective mismatch detection
- Actionable recommendations

### 6. Report Generator
- Campaign Completion Report
- Weekly Email Template
- Executive Summary
- One-click copy to clipboard
- Auto-populated with your actual KPIs

## Data Validation

MediaPulse validates your data **before** processing and shows clear errors:

| Issue | Handling |
|-------|----------|
| Missing required columns | Error with list of missing columns |
| Currency symbols (₹, $, commas) | Auto-stripped with warning |
| Percentage signs (%) | Auto-stripped with warning |
| Non-numeric values | Set to 0 with warning |
| Wrong date format (DD/MM/YYYY) | Auto-converted to YYYY-MM-DD |
| Blank rows | Silently removed |
| Merged cells | Detected and skipped |
| Total/subtotal rows | Detected and removed |
| Inconsistent names (Meta vs meta) | Warning shown |
| Unsupported file types | Rejected with clear message |

## Required Data Format

Your CSV/Excel **must** have these columns (names are auto-detected):

| Column | Example | Required |
|--------|---------|----------|
| Date | 2025-03-01 | ✅ Yes |
| Campaign | Meta - Installs - Tier1 | ✅ Yes |
| Platform | Meta, Google, DV360 | ✅ Yes |
| Spend | 25000 | ✅ Yes |
| Impressions | 1500000 | ✅ Yes |
| Clicks | 37500 | ✅ Yes |

Recommended columns: Installs, Device, City, Placement, CTR, CPC, CPI, ROAS, Objective

## 10 Data Rules

1. Row 1 = Headers
2. No merged cells
3. No blank rows
4. Plain numbers only (no ₹, $, commas)
5. Percentages as decimals (1.25, not 1.25%)
6. Dates as YYYY-MM-DD
7. One row per data point (no totals)
8. UTF-8 encoding
9. No formulas (paste as values)
10. Consistent names (Meta ≠ meta)

## Tech Stack

| Library | Purpose |
|---------|---------|
| React 18 | UI framework |
| Vite 5 | Build tool |
| Recharts | Charts (Line, Bar, Pie) |
| PapaParse | CSV parsing |
| SheetJS (xlsx) | Excel file parsing |
| Lucide React | Icons |

## Sample Data

A `sample-media-plan.csv` is included in `public/` for testing. It contains 24 rows across Meta, Google, DV360, and Programmatic platforms with Indian city data.

## Privacy

All processing happens in your browser. No data is sent to any server. Nothing is stored.

## License

MIT
