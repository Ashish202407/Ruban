# The VC Corner — Project Context

## What is The VC Corner?

The VC Corner is a **100% client-side investor dashboard generator**. A founder uploads their financial Excel template (`.xlsx`) in the browser and instantly gets a premium, investor-grade dashboard — no server, no data leaves the device.

**Live site:** https://ashish202407.github.io/Ruban/

---

## How It Works

1. **Template** — An Excel workbook (`VCCorner_Template_v2.xlsx`) with:
   - **Setup** sheet: company name, business type (AI-SaaS / D2C / Healthcare), currency, FY start
   - **Checklist** sheet: toggle sections on/off (column D = "Yes"/"No", column E = section ID). Original 29 sections default "Yes", 29 new sections default "No"
   - **Sector sheets** (AI-SaaS, D2C, Healthcare): financial data with 5-year projections, section anchors in column H, row types in column I

2. **Upload** — User drops the `.xlsx` file on the upload screen

3. **Parse** — `parser.js` reads Setup, Checklist, and sector data using SheetJS. Includes a client-side formula evaluator for SUM, IF, and arithmetic (handles openpyxl files that don't cache formula results)

4. **Render** — `renderer.js` builds the DOM: header, VC Readiness Radar scorecard, KPI strip (up to 8 cards), chart grid (3-col landscape) grouped by category with sticky headers, YoY badges, HTML legends, data tables, and privacy badge

5. **Charts** — `charts.js` creates Chart.js instances + custom Canvas 2D charts (waterfall, gauge, radar, stacked area, combo, donut) with dynamic color palettes and theme-aware styling

---

## Architecture

```
dashboard/docs-v2/       ← Source files (v2)
├── index.html            ← Single page: upload screen + dashboard screen
├── css/dashboard.css     ← All styles, CSS variables, light/dark themes, print mode
└── js/
    ├── registry.js       ← 58 section definitions (chart type, KPI rules, row mappings, categories)
    ├── parser.js         ← SheetJS Excel parser + formula evaluator
    ├── theme.js          ← Theme/palette state, controls, re-render logic
    ├── charts.js         ← Chart.js factory + custom chart types (waterfall, gauge, radar, etc.)
    ├── renderer.js       ← DOM builder (header, radar scorecard, KPIs, category groups, charts, tables, PDF export)
    └── app.js            ← Orchestrator (file upload, parse, render, theme init)

docs-v2/                  ← Mirror of dashboard/docs-v2/ for GitHub Pages (gh-pages-v2 branch)
```

**Script load order:** `registry.js` → `parser.js` → `theme.js` → `charts.js` → `renderer.js` → `app.js`

No build step, no bundler, no frameworks — plain vanilla JS modules as IIFEs on `window`.

---

## Supported Sectors & Sections (58 total)

### AI-SaaS (20 sections)
**Original 10:** MRR & ARR, Revenue by Plan, Churn & Retention, NDR, CAC & LTV, Rule of 40 & Burn Multiple, Sales Efficiency, P&L, Cash & Runway, Fundraising & Cap Table

**New 10:** Cohort MRR Retention, Feature Adoption by Plan, API Usage & Overage Revenue, NPS Trend, Headcount & Rev per Employee, Infrastructure Cost vs Revenue, ARR Bridge, Geographic Revenue Split, Support & CSAT, Token/Compute Cost per Customer

### D2C (20 sections)
**Original 10:** GMV & AOV, Channel Revenue, Return Rates, Customer Funnel, Cohort Retention, Unit Economics, P&L, Balance Sheet, Cash Flow, Fundraising & Cap Table

**New 10:** SKU-level Revenue, Ad Spend & ROAS by Channel, Seasonal Revenue Index, Refund & Dispute Rate, Loyalty Program, Inventory Turnover, Contribution Margin by Channel, LTV by Acquisition Channel, WhatsApp & Email Marketing, Influencer & Affiliate Revenue

### Healthcare (18 sections)
**Original 9:** Bed Occupancy, ARPOB, Department Revenue, Payer Mix, OPD & IPD Volumes, Cost per Bed Day, P&L, Cash Flow & Capex, Fundraising & Cap Table

**New 9:** Procedure Volume by Department, Insurance Claim Settlement, Doctor Productivity, Pharmacy Margin, ICU Utilization, Readmission Rate, Patient Satisfaction Score, Revenue Cycle Metrics, Capex & Asset Schedule

---

## Chart Types

| Type | Implementation | Used By |
|------|---------------|---------|
| Line | Chart.js | MRR/ARR, Churn, NPS, Cohort Retention, etc. |
| Bar | Chart.js | CAC/LTV, NDR, Cash & Runway, etc. |
| Stacked Bar | Chart.js (stacked) | Revenue by Plan, Channel Revenue, Headcount, etc. |
| Doughnut | Chart.js | Payer Mix |
| Waterfall | Chart.js stacked bar + transparent bases | P&L Summary (SaaS) |
| Gauge | Custom Canvas 2D semicircular arc | Rule of 40 & Burn Multiple |
| Stacked Area | Chart.js line with fill + stacking | Fundraising & Cap Table (SaaS) |
| Combo | Chart.js bar + stepped line, dual Y axes | Fundraising valuations |
| Donut KPI | Chart.js doughnut + center text plugin | Revenue by Plan snapshot |
| Radar | Chart.js radar + benchmark overlay | VC Readiness Scorecard |

---

## Dashboard Features

- **VC Readiness Radar** — 6-axis spider chart with benchmark overlay, per-sector axis definitions, overall score /100
- **Category Grouping** — sections sorted by: Revenue, Retention, Customers, Unit Economics, Operations, Efficiency, Marketing, Cost, Quality, Financials, Fundraising — with sticky headers
- **YoY Badges** — auto-computed Year 4→Year 5 delta on every section card title
- **Accent Stripes** — left border color per card via `--section-accent` CSS custom property
- **Dark/Light Theme** — toggle in nav bar with smooth transitions
- **10 Color Palettes** — 5 per theme (Dark: Gold, Ocean, Sage, Lavender, Copper; Light: Slate, Indigo, Teal, Rose, Amber)
- **Export PDF** — one-click print-friendly layout (hides nav, 2-col grid, white background)
- **Landscape-first** — 3-col grid optimized for 100% zoom on standard monitors, 4-col at 1920px+

---

## Theme System

- **Two base themes:** Light (default) and Dark — toggle in nav bar
- **10 monochromatic palettes** (5 per theme, 8 shades each)
- **Persistence:** `localStorage` (`ruban-theme`, `ruban-palette`)
- **FOUC prevention:** Inline `<script>` in `<head>` sets `data-theme` before paint
- **Re-render:** Theme changes trigger `Renderer.cleanup()` + `Renderer.render()` with `.no-animate` class to skip entrance animations
- **Chart theming:** Per-theme colors for grid, tooltips, ticks, point borders, legend labels

---

## Key Design Decisions

- **No server:** Privacy-first, everything runs in the browser
- **Column anchors:** Section IDs in hidden column H allow flexible row ordering in Excel
- **Formula evaluator:** Handles SUM, IF, arithmetic, comparison operators — needed because openpyxl (Python) doesn't evaluate formulas on save
- **Currency formatting:** Indian notation (Cr, L, K) as default, supports INR/USD/EUR/GBP
- **HTML legends:** Placed outside canvas to avoid overlap with y-axis labels
- **DPR capped at 2:** Canvas rendering uses `Math.min(devicePixelRatio, 2)` for sharp charts without excessive memory
- **Deferred gauge drawing:** Custom Canvas 2D gauge uses `requestAnimationFrame` retry loop since card may not be in DOM when chart is created
- **Compact layout:** Base font 14px, nav 44px, chart height 200px — optimized for landscape at 100% zoom

---

## Tooling

- **Template generator:** `generate_template_v2.py` — creates the Excel template with 58 sections across 3 sectors
- **Trial data filler:** `fill_trial1_v2.py` — populates template with sample AI-SaaS data (all 20 sections) for testing
- **v1 generators:** `generate_template.py` and `fill_trial1.py` still available for original 29-section template
- **CDN dependencies:** SheetJS (xlsx@0.18.5), Chart.js (4.4.1)
- **Fonts:** DM Sans (UI), JetBrains Mono (data/numbers)
- **Deployment:** GitHub Pages via `gh-pages-v2` branch (v2 files at root `/`)

---

## Commit History

| Hash | Description |
|------|-------------|
| `671a65a` | Fix zoom/layout for landscape at 100% — compact dashboard |
| `8c5df9f` | Fix empty gauge chart + update README for v2 |
| `29292cb` | The VC Corner v2: rebrand + 29 new sections, new chart types, category grouping, radar scorecard |
| `891df2c` | Add theme + color palette system with light/dark toggle |
| `9212933` | Move chart legends to HTML outside canvas to fix overlap |
| `7b6059c` | Fix $0 KPIs: formula cells now evaluate correctly |
| `df0c31f` | Add README.md |
| `4634f5f` | Initial commit: template generator + dark investor dashboard |
