# Dashboard Generator

**Fill an Excel template. Get an investor-grade dashboard.**

Upload your 5-year financial projections and instantly generate a polished, interactive dashboard - 100% client-side, your data never leaves your browser.

**[Try it live →](https://ashish202407.github.io/Ruban/)**

---

## How It Works

1. **Download** the `DashGen_Template_v2.xlsx` from `dashboard/`
2. **Fill** the Setup sheet (company name, sector, currency)
3. **Toggle** sections on the Checklist sheet (Yes/No)
4. **Enter** your numbers in the relevant sector sheet (blue cells = your inputs, black cells = auto-calculated formulas)
5. **Upload** the filled file at the link above
6. **Done** - dashboard renders instantly with KPIs, charts, and data tables

## Supported Sectors

| Sector | Sections | Highlights |
|--------|----------|------------|
| **AI-SaaS** | 20 | MRR/ARR, Revenue by Plan, Churn, NDR, CAC/LTV, Rule of 40 (gauge), Sales Efficiency, P&L (waterfall), Cash & Runway, Cap Table (stacked area), Cohort Retention, Feature Adoption, API Usage, NPS, Headcount, Infra Cost, ARR Bridge, Geo Revenue, Support/CSAT, Token Cost |
| **D2C** | 20 | GMV/AOV, Channel Revenue, Return Rates, Customer Funnel, Cohort Retention, Unit Economics, P&L, Balance Sheet, Cash Flow, Cap Table, SKU Revenue, Ad Spend/ROAS, Seasonal Index, Refund/Dispute, Loyalty, Inventory Turnover, Contribution Margin, LTV by Channel, Email/WhatsApp Marketing, Influencer/Affiliate |
| **Healthcare** | 18 | Bed Occupancy, ARPOB, Department Revenue, Payer Mix, OPD/IPD Volumes, Cost per Bed Day, P&L, Cash Flow & Capex, Cap Table, Procedure Volume, Insurance Claims, Doctor Productivity, Pharmacy Margin, ICU Utilization, Readmission Rate, Patient Satisfaction, Revenue Cycle, Capex & Assets |

## Dashboard Features

- **Readiness Radar** - 6-axis spider chart scoring your startup vs benchmarks (per sector)
- **Category Grouping** - sections organized by Revenue, Retention, Operations, etc. with sticky headers
- **YoY Badges** - auto-computed Year 4→5 deltas on every section card
- **Chart Types** - line, bar, stacked bar, doughnut, waterfall (P&L), gauge (Rule of 40), stacked area (cap table), combo (fundraising), radar
- **Dark/Light Theme** - toggle with smooth transitions
- **Export PDF** - one-click print-friendly layout
- **100% Private** - all processing in-browser, no server calls

## Project Structure

```
Ruban/
├── dashboard/
│   ├── generate_template_v2.py    # Generates DashGen_Template_v2.xlsx (58 sections)
│   ├── fill_trial1_v2.py          # Sample AI-SaaS data filler (20 sections)
│   ├── DashGen_Template_v2.xlsx   # Blank template
│   ├── Trial_1_v2_AI_SaaS.xlsx   # Pre-filled sample
│   ├── generate_template.py      # v1 template generator (29 sections)
│   ├── fill_trial1.py             # v1 sample filler
│   └── docs-v2/                   # Web app source (v2)
│       ├── index.html
│       ├── css/dashboard.css
│       └── js/
│           ├── registry.js        # 58 section configs + category order
│           ├── parser.js          # SheetJS parser + formula evaluator
│           ├── charts.js          # Chart.js factory + custom charts (waterfall, gauge, radar, etc.)
│           ├── renderer.js        # DOM builder + radar scorecard + PDF export
│           ├── theme.js           # Dark/light theme system
│           └── app.js             # Upload orchestrator
├── legacy-model/                  # Original D2C financial model
│   ├── generate_model.py
│   └── *.xlsx
└── docs-v2/                       # GitHub Pages deployment (mirrors dashboard/docs-v2)
```

## Tech Stack

- **Template:** Python + openpyxl (generates `.xlsx` with formulas, dropdowns, hidden anchors)
- **Dashboard:** Vanilla JS, no build step
- **Parsing:** [SheetJS](https://sheetjs.com/) + custom client-side formula evaluator
- **Charts:** [Chart.js](https://www.chartjs.org/) + custom Canvas 2D (gauge, waterfall)
- **Hosting:** GitHub Pages - static, no backend

## Privacy

All processing happens in your browser. No data is uploaded to any server. The Excel file is read using SheetJS in JavaScript - nothing leaves your device.

## Generating the Template

```bash
cd dashboard
pip install openpyxl
python3 generate_template_v2.py   # → DashGen_Template_v2.xlsx
```

## Generating a Sample Filled File

```bash
cd dashboard
python3 fill_trial1_v2.py         # → Trial_1_v2_AI_SaaS.xlsx
```

## License

MIT
