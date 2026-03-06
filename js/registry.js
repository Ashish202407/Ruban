/**
 * The VC Corner Section Registry
 * Maps 58 section IDs → chart types, KPI rules, display config
 *
 * yFormat: "currency" (default) | "percent" | "number" | "ratio"
 *   Controls how the chart y-axis is formatted
 */
const REGISTRY = {
  // ─── AI-SaaS (Original 10) ───
  saas_mrr_arr: {
    title: "MRR & ARR Progression",
    sector: "AI-SaaS",
    chartType: "line",
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "number",
    category: "Retention",
    kpis: [
      { label: "Logo Churn %", row: "Logo Churn %", format: "percent" },
      { label: "Ending Customers", row: "Ending Customers", format: "number" }
    ],
    chartRows: ["Starting Customers", "Ending Customers"],
    tableRows: "all"
  },
  saas_ndr: {
    title: "Net Dollar Retention",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "currency",
    category: "Retention",
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
    yFormat: "currency",
    category: "Unit Economics",
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
    chartType: "gauge",
    yFormat: "percent",
    category: "Efficiency",
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
    yFormat: "currency",
    category: "Efficiency",
    kpis: [
      { label: "Magic Number", row: "Magic Number", format: "ratio" }
    ],
    chartRows: ["Quarterly Net New ARR", "Previous Quarter S&M Spend", "Magic Number"],
    tableRows: "all"
  },
  saas_pnl: {
    title: "P&L Summary",
    sector: "AI-SaaS",
    chartType: "waterfall",
    yFormat: "currency",
    category: "Financials",
    kpis: [
      { label: "Revenue", row: "Revenue", format: "currency" },
      { label: "Gross Profit", row: "Gross Profit", format: "currency" },
      { label: "EBITDA", row: "EBITDA", format: "currency" },
      { label: "PAT", row: "PAT (Profit After Tax)", format: "currency" }
    ],
    chartRows: ["Revenue", "COGS", "Gross Profit", "R&D Expense", "S&M Expense", "G&A Expense", "EBITDA", "PAT (Profit After Tax)"],
    tableRows: "all"
  },
  saas_cash_runway: {
    title: "Cash & Runway",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "currency",
    category: "Financials",
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
    chartType: "stacked_area",
    yFormat: "percent_whole",
    category: "Fundraising",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  },

  // ─── AI-SaaS (New 10) ───
  saas_cohort_mrr: {
    title: "Monthly Cohort MRR Retention",
    sector: "AI-SaaS",
    chartType: "line",
    yFormat: "percent",
    category: "Operations",
    kpis: [
      { label: "Cohort 1 Y5", row: "Cohort 1 Retention %", format: "percent" }
    ],
    chartRows: ["Cohort 1 Retention %", "Cohort 2 Retention %", "Cohort 3 Retention %", "Cohort 4 Retention %", "Cohort 5 Retention %"],
    tableRows: "all"
  },
  saas_feature_adoption: {
    title: "Feature Adoption by Plan",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "percent",
    category: "Operations",
    kpis: [
      { label: "Enterprise", row: "Enterprise Adoption %", format: "percent" }
    ],
    chartRows: ["Free Tier Adoption %", "Starter Adoption %", "Pro Adoption %", "Enterprise Adoption %"],
    tableRows: "all"
  },
  saas_api_usage: {
    title: "API Usage & Overage Revenue",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "number",
    category: "Revenue",
    kpis: [
      { label: "API Calls", row: "API Calls (thousands)", format: "number" },
      { label: "Overage Rev", row: "Overage Revenue", format: "currency" }
    ],
    chartRows: ["API Calls (thousands)", "Included API Calls (thousands)", "Overage Revenue"],
    tableRows: "all"
  },
  saas_nps: {
    title: "NPS Trend",
    sector: "AI-SaaS",
    chartType: "line",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "NPS", row: "NPS Score", format: "number" }
    ],
    chartRows: ["NPS Score"],
    tableRows: "all"
  },
  saas_headcount: {
    title: "Headcount & Rev per Employee",
    sector: "AI-SaaS",
    chartType: "stacked_bar",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "Headcount", row: "Total Headcount", format: "number" },
      { label: "Rev/Emp", row: "Revenue per Employee", format: "currency" }
    ],
    chartRows: ["R&D Headcount", "S&M Headcount", "G&A Headcount"],
    tableRows: "all"
  },
  saas_infra_cost: {
    title: "Infrastructure Cost vs Revenue",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "currency",
    category: "Efficiency",
    kpis: [
      { label: "Infra %", row: "Infra as % of Revenue", format: "percent" }
    ],
    chartRows: ["Hosting & Cloud Cost", "Total Revenue"],
    tableRows: "all"
  },
  saas_arr_bridge: {
    title: "ARR Bridge",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "currency",
    category: "Revenue",
    kpis: [
      { label: "Net New ARR", row: "Net New ARR", format: "currency" },
      { label: "Ending ARR", row: "Ending ARR", format: "currency" }
    ],
    chartRows: ["New ARR", "Expansion ARR", "Contraction ARR", "Churned ARR", "Net New ARR"],
    tableRows: "all"
  },
  saas_geo_revenue: {
    title: "Geographic Revenue Split",
    sector: "AI-SaaS",
    chartType: "stacked_bar",
    yFormat: "currency",
    category: "Revenue",
    kpis: [
      { label: "Total Rev", row: "Total Revenue", format: "currency" }
    ],
    chartRows: ["India Revenue", "US Revenue", "Europe Revenue", "RoW Revenue"],
    tableRows: "all"
  },
  saas_support_csat: {
    title: "Support & CSAT",
    sector: "AI-SaaS",
    chartType: "line",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "CSAT", row: "CSAT Score %", format: "percent" },
      { label: "Tickets", row: "Total Tickets", format: "number" }
    ],
    chartRows: ["Total Tickets", "Tickets Resolved"],
    tableRows: "all"
  },
  saas_token_cost: {
    title: "Token/Compute Cost per Customer",
    sector: "AI-SaaS",
    chartType: "bar",
    yFormat: "currency",
    category: "Efficiency",
    kpis: [
      { label: "Cost/Cust", row: "Cost per Customer", format: "currency" }
    ],
    chartRows: ["Total AI Inference Cost", "Cost per Customer"],
    tableRows: "all"
  },

  // ─── D2C (Original 10) ───
  d2c_gmv_aov: {
    title: "GMV & AOV",
    sector: "D2C",
    chartType: "stacked_bar",
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "currency",
    category: "Revenue",
    kpis: [],
    chartRows: ["Website — Net Revenue", "Retail — Net Revenue", "Marketplace — Net Revenue"],
    tableRows: "all"
  },
  d2c_customer_funnel: {
    title: "Customer Funnel",
    sector: "D2C",
    chartType: "bar",
    yFormat: "number",
    category: "Customers",
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
    yFormat: "percent",
    category: "Retention",
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
    yFormat: "currency",
    category: "Unit Economics",
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
    yFormat: "currency",
    category: "Financials",
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
    yFormat: "currency",
    category: "Financials",
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
    yFormat: "currency",
    category: "Financials",
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
    yFormat: "percent_whole",
    category: "Fundraising",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  },

  // ─── D2C (New 10) ───
  d2c_sku_revenue: {
    title: "SKU-level Revenue",
    sector: "D2C",
    chartType: "stacked_bar",
    yFormat: "currency",
    category: "Revenue",
    kpis: [
      { label: "Total SKU Rev", row: "Total SKU Revenue", format: "currency" }
    ],
    chartRows: ["SKU Category 1 Revenue", "SKU Category 2 Revenue", "SKU Category 3 Revenue", "SKU Category 4 Revenue"],
    tableRows: "all"
  },
  d2c_ad_roas: {
    title: "Ad Spend & ROAS by Channel",
    sector: "D2C",
    chartType: "bar",
    yFormat: "currency",
    category: "Marketing",
    kpis: [
      { label: "Blended ROAS", row: "Blended ROAS", format: "ratio" }
    ],
    chartRows: ["Meta Ad Spend", "Google Ad Spend", "Other Ad Spend", "Total Ad Revenue"],
    tableRows: "all"
  },
  d2c_seasonal: {
    title: "Seasonal Revenue Index",
    sector: "D2C",
    chartType: "stacked_bar",
    yFormat: "currency",
    category: "Revenue",
    kpis: [
      { label: "Annual Rev", row: "Annual Revenue", format: "currency" }
    ],
    chartRows: ["Q1 Revenue", "Q2 Revenue", "Q3 Revenue", "Q4 Revenue"],
    tableRows: "all"
  },
  d2c_refund_dispute: {
    title: "Refund & Dispute Rate",
    sector: "D2C",
    chartType: "line",
    yFormat: "percent",
    category: "Operations",
    kpis: [
      { label: "Refund Rate", row: "Refund Rate %", format: "percent" },
      { label: "Dispute Rate", row: "Dispute Rate %", format: "percent" }
    ],
    chartRows: ["Refund Rate %", "Dispute Rate %"],
    tableRows: "all"
  },
  d2c_loyalty: {
    title: "Loyalty Program Metrics",
    sector: "D2C",
    chartType: "bar",
    yFormat: "number",
    category: "Retention",
    kpis: [
      { label: "Members", row: "Loyalty Members", format: "number" },
      { label: "Repeat %", row: "Repeat Purchase Rate %", format: "percent" }
    ],
    chartRows: ["Loyalty Members", "Active Members"],
    tableRows: "all"
  },
  d2c_inventory_turnover: {
    title: "Inventory Turnover",
    sector: "D2C",
    chartType: "bar",
    yFormat: "ratio",
    category: "Operations",
    kpis: [
      { label: "Turnover", row: "Inventory Turnover Ratio", format: "ratio" },
      { label: "DIO", row: "Days Inventory Outstanding", format: "number" }
    ],
    chartRows: ["Inventory Turnover Ratio", "Days Inventory Outstanding"],
    tableRows: "all"
  },
  d2c_contribution_margin: {
    title: "Contribution Margin by Channel",
    sector: "D2C",
    chartType: "bar",
    yFormat: "currency",
    category: "Unit Economics",
    kpis: [],
    chartRows: ["Website — Contribution Margin", "Retail — Contribution Margin", "Marketplace — Contribution Margin"],
    tableRows: "all"
  },
  d2c_ltv_channel: {
    title: "LTV by Acquisition Channel",
    sector: "D2C",
    chartType: "bar",
    yFormat: "currency",
    category: "Unit Economics",
    kpis: [],
    chartRows: ["Organic LTV", "Paid Social LTV", "Paid Search LTV", "Referral LTV"],
    tableRows: "all"
  },
  d2c_email_marketing: {
    title: "WhatsApp & Email Marketing",
    sector: "D2C",
    chartType: "bar",
    yFormat: "currency",
    category: "Marketing",
    kpis: [
      { label: "Email Rev", row: "Email Revenue", format: "currency" },
      { label: "WA Rev", row: "WhatsApp Revenue", format: "currency" }
    ],
    chartRows: ["Email Revenue", "WhatsApp Revenue"],
    tableRows: "all"
  },
  d2c_influencer: {
    title: "Influencer & Affiliate Revenue",
    sector: "D2C",
    chartType: "bar",
    yFormat: "currency",
    category: "Marketing",
    kpis: [
      { label: "Inf ROAS", row: "Influencer ROAS", format: "ratio" },
      { label: "Aff ROAS", row: "Affiliate ROAS", format: "ratio" }
    ],
    chartRows: ["Influencer Spend", "Influencer Revenue", "Affiliate Spend", "Affiliate Revenue"],
    tableRows: "all"
  },

  // ─── Healthcare (Original 9) ───
  hc_bed_occupancy: {
    title: "Bed Occupancy & Utilization",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "Occupancy %", row: "Occupancy Rate %", format: "percent" },
      { label: "Total Beds", row: "Total Beds Available", format: "number" },
      { label: "Discharges", row: "Total Discharges", format: "number" }
    ],
    chartRows: ["Total Beds Available", "Total Discharges"],
    tableRows: "all"
  },
  hc_arpob: {
    title: "ARPOB & Revenue per Bed",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "currency",
    category: "Revenue",
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
    yFormat: "number",
    category: "Operations",
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
    yFormat: "currency",
    category: "Cost",
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
    yFormat: "currency",
    category: "Financials",
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
    yFormat: "currency",
    category: "Financials",
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
    yFormat: "percent_whole",
    category: "Fundraising",
    kpis: [
      { label: "Amount Raised", row: "Amount Raised", format: "currency" }
    ],
    chartRows: ["Founders %", "ESOP %", "Seed Investors %", "Series A %", "Series B %", "Others %"],
    tableRows: "all"
  },

  // ─── Healthcare (New 9) ───
  hc_procedure_volume: {
    title: "Procedure Volume by Department",
    sector: "Healthcare",
    chartType: "stacked_bar",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "Total Procedures", row: "Total Procedures", format: "number" }
    ],
    chartRows: ["General Surgery Procedures", "Orthopaedic Procedures", "Cardiology Procedures", "Oncology Procedures", "Other Procedures"],
    tableRows: "all"
  },
  hc_insurance_claims: {
    title: "Insurance Claim Settlement",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "Approval %", row: "Approval Rate %", format: "percent" }
    ],
    chartRows: ["Claims Submitted", "Claims Approved", "Claims Rejected"],
    tableRows: "all"
  },
  hc_doctor_productivity: {
    title: "Doctor Productivity",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "currency",
    category: "Efficiency",
    kpis: [
      { label: "Rev/Doctor", row: "Revenue per Doctor", format: "currency" }
    ],
    chartRows: ["Total Doctors", "Revenue per Doctor"],
    tableRows: "all"
  },
  hc_pharmacy_margin: {
    title: "Pharmacy Margin",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "currency",
    category: "Revenue",
    kpis: [
      { label: "Margin %", row: "Pharmacy Margin %", format: "percent" }
    ],
    chartRows: ["Pharmacy Revenue", "Pharmacy COGS", "Pharmacy Gross Profit"],
    tableRows: "all"
  },
  hc_icu_utilization: {
    title: "ICU Utilization",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "number",
    category: "Operations",
    kpis: [
      { label: "ICU Occ %", row: "ICU Occupancy %", format: "percent" }
    ],
    chartRows: ["ICU Beds Available", "ICU Occupied Bed Days"],
    tableRows: "all"
  },
  hc_readmission: {
    title: "Readmission Rate",
    sector: "Healthcare",
    chartType: "line",
    yFormat: "percent",
    category: "Quality",
    kpis: [
      { label: "Readmission %", row: "Readmission Rate %", format: "percent" }
    ],
    chartRows: ["Readmission Rate %"],
    tableRows: "all"
  },
  hc_patient_satisfaction: {
    title: "Patient Satisfaction Score",
    sector: "Healthcare",
    chartType: "line",
    yFormat: "number",
    category: "Quality",
    kpis: [
      { label: "Overall Score", row: "Overall Satisfaction Score", format: "number" }
    ],
    chartRows: ["Overall Satisfaction Score", "IPD Satisfaction Score", "OPD Satisfaction Score"],
    tableRows: "all"
  },
  hc_revenue_cycle: {
    title: "Revenue Cycle Metrics",
    sector: "Healthcare",
    chartType: "bar",
    yFormat: "currency",
    category: "Financials",
    kpis: [
      { label: "Collection %", row: "Collection Rate %", format: "percent" }
    ],
    chartRows: ["Gross Revenue", "Net Revenue"],
    tableRows: "all"
  },
  hc_capex_assets: {
    title: "Capex & Asset Schedule",
    sector: "Healthcare",
    chartType: "stacked_bar",
    yFormat: "currency",
    category: "Financials",
    kpis: [
      { label: "Total Capex", row: "Total Capex", format: "currency" },
      { label: "Net FA", row: "Net Fixed Assets", format: "currency" }
    ],
    chartRows: ["Equipment Capex", "Building/Expansion Capex", "IT & Technology Capex"],
    tableRows: "all"
  }
};

/** Category display order for grouping */
const CATEGORY_ORDER = [
  "Revenue", "Retention", "Customers", "Unit Economics", "Operations",
  "Efficiency", "Marketing", "Cost", "Quality", "Financials", "Fundraising"
];
