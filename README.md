# Ruban

**Fill an Excel template. Get an investor-grade dashboard.**

Upload your 5-year financial projections and instantly generate a dark-themed, interactive dashboard — 100% client-side, your data never leaves your browser.

**[Try it live →](https://ashish202407.github.io/Ruban/)**

---

## How It Works

1. **Download** the `Ruban_Template.xlsx` from `dashboard/`
2. **Fill** the Setup sheet (company name, sector, currency)
3. **Toggle** sections on the Checklist sheet (Yes/No)
4. **Enter** your numbers in the relevant sector sheet (blue cells = your inputs, black cells = auto-calculated formulas)
5. **Upload** the filled file at the link above
6. **Done** — dashboard renders instantly with KPIs, charts, and data tables

## Supported Sectors

| Sector | Sections | Metrics |
|--------|----------|---------|
| **AI-SaaS** | 10 | MRR/ARR, Revenue by Plan, Churn, NDR, CAC/LTV, Rule of 40, Sales Efficiency, P&L, Cash & Runway, Cap Table |
| **D2C** | 10 | GMV/AOV, Channel Revenue, Return Rates, Customer Funnel, Cohort Retention, Unit Economics, P&L, Balance Sheet, Cash Flow, Cap Table |
| **Healthcare** | 9 | Bed Occupancy, ARPOB, Department Revenue, Payer Mix, OPD/IPD Volumes, Cost per Bed Day, P&L, Cash Flow & Capex, Cap Table |

## Project Structure

```
Ruban/
├── dashboard/                  # Multi-sector template + web dashboard
│   ├── generate_template.py    # Generates Ruban_Template.xlsx
│   ├── fill_trial1.py          # Sample AI-SaaS data filler
│   ├── Ruban_Template.xlsx     # Blank template
│   ├── Trial_1_AI_SaaS.xlsx    # Pre-filled sample
│   └── docs/                   # Web app source
│       ├── index.html
│       ├── css/dashboard.css
│       └── js/
│           ├── registry.js     # 29 section configs
│           ├── parser.js       # SheetJS parser + formula evaluator
│           ├── charts.js       # Chart.js factory (dark theme)
│           ├── renderer.js     # DOM builder
│           └── app.js          # Upload orchestrator
├── legacy-model/               # Original D2C financial model
│   ├── generate_model.py
│   └── *.xlsx
└── docs/                       # GitHub Pages deployment (mirrors dashboard/docs)
```

## Tech Stack

- **Template:** Python + openpyxl (generates `.xlsx` with formulas, dropdowns, hidden anchors)
- **Dashboard:** Vanilla JS, no build step
- **Parsing:** [SheetJS](https://sheetjs.com/) + custom client-side formula evaluator
- **Charts:** [Chart.js](https://www.chartjs.org/) with dark theme, 3x DPR rendering
- **Hosting:** GitHub Pages — static, no backend

## Privacy

All processing happens in your browser. No data is uploaded to any server. The Excel file is read using SheetJS in JavaScript — nothing leaves your device.

## Generating the Template

```bash
cd dashboard
pip install openpyxl
python3 generate_template.py   # → Ruban_Template.xlsx
```

## Generating a Sample Filled File

```bash
cd dashboard
python3 fill_trial1.py         # → Trial_1_AI_SaaS.xlsx
```

## License

MIT
