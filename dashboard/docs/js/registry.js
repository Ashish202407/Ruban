/**
 * Ruban Section Registry
 * Maps 29 section IDs → chart types, KPI rules, display config
 */
const REGISTRY = {
  // ─── AI-SaaS ───
  saas_mrr_arr: {
    title: "MRR & ARR Progression",
    sector: "AI-SaaS",
    chartType: "line",
    kpis: [
      { label: "Ending MRR", row: "Ending MRR", format: "currency" },
      { label: "ARR", row: "ARR", format: "currency" },
      { label: "MRR Growth", row: "Ending MRR", format: "growth" }
    ],
    chartRows: ["Beginning MRR", "Ending MRR", "ARR"],
    tableRows: "all"
  },
  saas_revenue_plan: {
    title: "Revenue by Plan/Tier",
    sector: "AI-SaaS",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total Revenue", row: "Total Revenue", format: "currency" },
      { label: "Revenue Growth", row: "Total Revenue", format: "growth" }
    ],
    chartRows: ["Starter Revenue", "Pro Revenue", "Enterprise Revenue"],
    tableRows: "all"
  },
  saas_churn_retention: {
    title: "Churn & Retention",
    sector: "AI-SaaS",
    chartType: "line",
    kpis: [
      { label: "Logo Churn %", row: "Logo Churn %", format: "percent" },
      { label: "Ending Customers", row: "Ending Customers", format: "number" }
    ],
    chartRows: ["Starting Customers", "Ending Customers"],
    chartRowsSecondary: ["Logo Churn %"],
    tableRows: "all"
  },
  saas_ndr: {
    title: "Net Dollar Retention",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "NDR %", row: "NDR %", format: "percent" },
      { label: "Ending ARR", row: "Ending ARR", format: "currency" }
    ],
    chartRows: ["Beginning ARR", "Expansion ARR", "Churned ARR", "Ending ARR"],
    tableRows: "all"
  },
  saas_cac_ltv: {
    title: "CAC & LTV",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "CAC", row: "CAC", format: "currency" },
      { label: "LTV", row: "LTV", format: "currency" },
      { label: "LTV/CAC", row: "LTV / CAC", format: "ratio" },
      { label: "Payback", row: "Payback (months)", format: "months" }
    ],
    chartRows: ["CAC", "LTV"],
    tableRows: "all"
  },
  saas_rule_of_40: {
    title: "Rule of 40 & Burn Multiple",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "Rule of 40", row: "Rule of 40 Score", format: "percent" },
      { label: "Burn Multiple", row: "Burn Multiple", format: "ratio" }
    ],
    chartRows: ["Revenue Growth %", "EBITDA Margin %", "Rule of 40 Score"],
    tableRows: "all"
  },
  saas_sales_efficiency: {
    title: "Sales Efficiency",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "Magic Number", row: "Magic Number", format: "ratio" }
    ],
    chartRows: ["Quarterly Net New ARR", "Previous Quarter S&M Spend", "Magic Number"],
    tableRows: "all"
  },
  saas_pnl: {
    title: "P&L Summary",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "Revenue", row: "Revenue", format: "currency" },
      { label: "Gross Profit", row: "Gross Profit", format: "currency" },
      { label: "EBITDA", row: "EBITDA", format: "currency" },
      { label: "PAT", row: "PAT (Profit After Tax)", format: "currency" }
    ],
    chartRows: ["Revenue", "Gross Profit", "EBITDA", "PAT (Profit After Tax)"],
    tableRows: "all"
  },
  saas_cash_runway: {
    title: "Cash & Runway",
    sector: "AI-SaaS",
    chartType: "bar",
    kpis: [
      { label: "Closing Cash", row: "Closing Cash", format: "currency" },
      { label: "Runway", row: "Runway (months)", format: "months" }
    ],
    chartRows: ["Opening Cash", "Closing Cash"],
    tableRows: "all"
  },
  saas_fundraising: {
    title: "Fundraising & Cap Table",
    sector: "AI-SaaS",
    chartType: "stacked_bar",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  },

  // ─── D2C ───
  d2c_gmv_aov: {
    title: "GMV & AOV",
    sector: "D2C",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total GMV", row: "Total GMV", format: "currency" },
      { label: "Blended AOV", row: "Blended AOV", format: "currency" }
    ],
    chartRows: ["GMV — Website", "GMV — Retail", "GMV — Marketplace"],
    tableRows: "all"
  },
  d2c_channel_revenue: {
    title: "Channel Revenue Breakup",
    sector: "D2C",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total Net Revenue", row: "Total Net Revenue", format: "currency" },
      { label: "Revenue Growth", row: "Total Net Revenue", format: "growth" }
    ],
    chartRows: ["Website — Net Revenue", "Retail — Net Revenue", "Marketplace — Net Revenue"],
    tableRows: "all"
  },
  d2c_return_rates: {
    title: "Return Rates & Net Revenue",
    sector: "D2C",
    chartType: "bar",
    kpis: [],
    chartRows: ["Website — Net Revenue", "Retail — Net Revenue", "Marketplace — Net Revenue"],
    tableRows: "all"
  },
  d2c_customer_funnel: {
    title: "Customer Funnel",
    sector: "D2C",
    chartType: "bar",
    kpis: [
      { label: "Visitors", row: "Website Visitors", format: "number" },
      { label: "Customers", row: "Customers Converted", format: "number" },
      { label: "Lead → Cust %", row: "Lead → Customer %", format: "percent" }
    ],
    chartRows: ["Website Visitors", "Leads Generated", "Customers Converted", "Repeat Customers"],
    tableRows: "all"
  },
  d2c_cohort_retention: {
    title: "Cohort Retention & Churn",
    sector: "D2C",
    chartType: "line",
    kpis: [
      { label: "Blended Churn", row: "Blended Churn Rate", format: "percent" },
      { label: "Repeat Rate", row: "Repeat Purchase Rate", format: "percent" }
    ],
    chartRows: ["Website — Churn Rate", "Retail — Churn Rate", "Marketplace — Churn Rate", "Blended Churn Rate"],
    tableRows: "all"
  },
  d2c_unit_economics: {
    title: "Unit Economics",
    sector: "D2C",
    chartType: "bar",
    kpis: [
      { label: "CAC", row: "CAC (Customer Acquisition Cost)", format: "currency" },
      { label: "LTV", row: "LTV (Lifetime Value)", format: "currency" },
      { label: "LTV/CAC", row: "LTV / CAC", format: "ratio" }
    ],
    chartRows: ["CAC (Customer Acquisition Cost)", "LTV (Lifetime Value)"],
    tableRows: "all"
  },
  d2c_pnl: {
    title: "P&L Summary",
    sector: "D2C",
    chartType: "bar",
    kpis: [
      { label: "Net Revenue", row: "Net Revenue", format: "currency" },
      { label: "Gross Profit", row: "Gross Profit", format: "currency" },
      { label: "EBITDA", row: "EBITDA", format: "currency" },
      { label: "PAT", row: "PAT (Profit After Tax)", format: "currency" }
    ],
    chartRows: ["Net Revenue", "Gross Profit", "EBITDA", "PAT (Profit After Tax)"],
    tableRows: "all"
  },
  d2c_balance_sheet: {
    title: "Balance Sheet",
    sector: "D2C",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total Assets", row: "Total Assets", format: "currency" },
      { label: "Total Equity", row: "Total Equity", format: "currency" }
    ],
    chartRows: ["Cash & Equivalents", "Inventory", "Accounts Receivable", "Fixed Assets (Net)"],
    tableRows: "all"
  },
  d2c_cash_flow: {
    title: "Cash Flow",
    sector: "D2C",
    chartType: "bar",
    kpis: [
      { label: "CF from Ops", row: "CF from Operations", format: "currency" },
      { label: "Closing Cash", row: "Closing Cash Balance", format: "currency" }
    ],
    chartRows: ["CF from Operations", "CF from Investing", "CF from Financing", "Closing Cash Balance"],
    tableRows: "all"
  },
  d2c_fundraising: {
    title: "Fundraising & Cap Table",
    sector: "D2C",
    chartType: "stacked_bar",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  },

  // ─── Healthcare ───
  hc_bed_occupancy: {
    title: "Bed Occupancy & Utilization",
    sector: "Healthcare",
    chartType: "bar",
    kpis: [
      { label: "Occupancy %", row: "Occupancy Rate %", format: "percent" },
      { label: "Total Beds", row: "Total Beds Available", format: "number" },
      { label: "Discharges", row: "Total Discharges", format: "number" }
    ],
    chartRows: ["Total Beds Available", "Total Discharges"],
    chartRowsSecondary: ["Occupancy Rate %"],
    tableRows: "all"
  },
  hc_arpob: {
    title: "ARPOB & Revenue per Bed",
    sector: "Healthcare",
    chartType: "bar",
    kpis: [
      { label: "ARPOB", row: "ARPOB (Avg Rev per Occupied Bed)", format: "currency" },
      { label: "Rev/Bed", row: "Revenue per Available Bed", format: "currency" }
    ],
    chartRows: ["Inpatient Revenue", "ARPOB (Avg Rev per Occupied Bed)", "Revenue per Available Bed"],
    tableRows: "all"
  },
  hc_dept_revenue: {
    title: "Department Revenue",
    sector: "Healthcare",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total Revenue", row: "Total Revenue", format: "currency" },
      { label: "Revenue Growth", row: "Total Revenue", format: "growth" }
    ],
    chartRows: ["IPD Revenue", "OPD Revenue", "Pharmacy Revenue", "Diagnostics Revenue", "Surgical Revenue", "Emergency Revenue"],
    tableRows: "all"
  },
  hc_payer_mix: {
    title: "Payer Mix",
    sector: "Healthcare",
    chartType: "stacked_bar",
    kpis: [
      { label: "Total Revenue", row: "Total Revenue", format: "currency" }
    ],
    chartRows: ["Insurance Revenue", "Cash/Self-Pay Revenue", "Government Revenue", "Corporate/TPA Revenue"],
    tableRows: "all"
  },
  hc_opd_ipd: {
    title: "OPD & IPD Volumes",
    sector: "Healthcare",
    chartType: "bar",
    kpis: [
      { label: "Annual OPD", row: "Annual OPD Visits", format: "number" },
      { label: "IPD Admissions", row: "IPD Admissions", format: "number" },
      { label: "Conversion %", row: "OPD → IPD Conversion %", format: "percent" }
    ],
    chartRows: ["Annual OPD Visits", "IPD Admissions"],
    tableRows: "all"
  },
  hc_cost_per_bed: {
    title: "Cost per Bed Day",
    sector: "Healthcare",
    chartType: "stacked_bar",
    kpis: [
      { label: "Cost/Bed Day", row: "Cost per Occupied Bed Day", format: "currency" },
      { label: "Contribution", row: "Contribution", format: "currency" }
    ],
    chartRows: ["Medical Supplies Cost", "Staff Cost (allocated to beds)", "Pharmacy COGS"],
    tableRows: "all"
  },
  hc_pnl: {
    title: "P&L Summary",
    sector: "Healthcare",
    chartType: "bar",
    kpis: [
      { label: "Revenue", row: "Total Revenue", format: "currency" },
      { label: "Gross Profit", row: "Gross Profit", format: "currency" },
      { label: "EBITDA", row: "EBITDA", format: "currency" },
      { label: "PAT", row: "PAT (Profit After Tax)", format: "currency" }
    ],
    chartRows: ["Total Revenue", "Gross Profit", "EBITDA", "PAT (Profit After Tax)"],
    tableRows: "all"
  },
  hc_cash_flow: {
    title: "Cash Flow & Capex",
    sector: "Healthcare",
    chartType: "bar",
    kpis: [
      { label: "CF from Ops", row: "CF from Operations", format: "currency" },
      { label: "Closing Cash", row: "Closing Cash Balance", format: "currency" }
    ],
    chartRows: ["CF from Operations", "CF from Investing", "CF from Financing", "Closing Cash Balance"],
    tableRows: "all"
  },
  hc_fundraising: {
    title: "Fundraising & Cap Table",
    sector: "Healthcare",
    chartType: "stacked_bar",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  }
};
