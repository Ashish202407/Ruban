#!/usr/bin/env python3
"""
Dashboard Generator - User Guide PDF Generator
Generates a polished, branded user guide for founders.
"""

from fpdf import FPDF
import os

# ── Brand Colors (white background, black/dark grey text) ─────
PAGE_BG      = (255, 255, 255)    # White background
CARD_BG      = (247, 247, 247)    # #F7F7F7 - light grey card bg
BLACK        = (35, 35, 35)       # #232323 - primary text (matches logo)
DARK_GREY    = (60, 60, 60)       # #3C3C3C - secondary text
MID_GREY     = (120, 120, 120)    # #787878 - tertiary/muted text
LIGHT_GREY   = (200, 200, 200)    # #C8C8C8 - borders, dividers
ACCENT       = (35, 35, 35)       # Black accent (matching logo)
TABLE_HEADER = (35, 35, 35)       # Black header
TABLE_ROW_1  = (255, 255, 255)    # White
TABLE_ROW_2  = (245, 245, 245)    # Very light grey
BLUE_INPUT   = (37, 102, 176)     # #2566B0 - matches template blue
GREEN_OK     = (40, 140, 70)      # Darker green for white bg
RED_WARN     = (200, 50, 50)      # Darker red for white bg
NOTE_BLUE    = (50, 80, 160)      # Blue for note boxes

LOGO_PATH = os.path.join(os.path.dirname(__file__), "VC - Logo.jpeg")


class VCGuide(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_auto_page_break(auto=True, margin=20)
        self.set_margins(18, 18, 18)
        # Add fonts
        self.add_page()  # triggers header

    # ── Page chrome ───────────────────────────────────────────
    def header(self):
        self.set_fill_color(*PAGE_BG)
        self.rect(0, 0, 210, 297, "F")
        # Black accent line at top
        self.set_fill_color(*ACCENT)
        self.rect(0, 0, 210, 1.5, "F")

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*MID_GREY)
        self.cell(0, 10, f"Dashboard Generator  |  User Guide  |  Page {self.page_no()}/{{nb}}", align="C")

    # ── Helpers ───────────────────────────────────────────────
    def accent_line(self, y=None):
        if y is None:
            y = self.get_y()
        self.set_fill_color(*ACCENT)
        self.rect(18, y, 174, 0.5, "F")
        self.set_y(y + 2)

    def section_title(self, num, title):
        self.ln(6)
        self.set_font("Helvetica", "B", 18)
        self.set_text_color(*BLACK)
        self.cell(0, 10, f"{num}.  {title}", ln=True)
        self.accent_line()
        self.ln(2)

    def sub_title(self, title):
        self.ln(3)
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*BLACK)
        self.cell(0, 8, title, ln=True)
        self.ln(1)

    def sub_sub_title(self, title):
        self.ln(2)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*DARK_GREY)
        self.cell(0, 7, title, ln=True)
        self.ln(1)

    def body(self, text):
        self.set_font("Helvetica", "", 9.5)
        self.set_text_color(*DARK_GREY)
        self.multi_cell(0, 5, text)
        self.ln(1)

    def body_light(self, text):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*MID_GREY)
        self.multi_cell(0, 5, text)
        self.ln(1)

    def bullet(self, text, indent=6):
        x = self.get_x()
        self.set_x(x + indent)
        self.set_font("Helvetica", "", 9.5)
        self.set_text_color(*BLACK)
        self.cell(4, 5, ">")
        self.set_text_color(*DARK_GREY)
        self.multi_cell(0, 5, text)
        self.set_x(x)

    def bullet_light(self, text, indent=6):
        x = self.get_x()
        self.set_x(x + indent)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*BLACK)
        self.cell(4, 5, ">")
        self.set_text_color(*MID_GREY)
        self.multi_cell(0, 5, text)
        self.set_x(x)

    def key_value(self, key, value, indent=6):
        x = self.get_x()
        self.set_x(x + indent)
        self.set_font("Helvetica", "B", 9.5)
        self.set_text_color(*BLACK)
        kw = self.get_string_width(key + ": ") + 2
        self.cell(kw, 5, key + ": ")
        self.set_font("Helvetica", "", 9.5)
        self.set_text_color(*DARK_GREY)
        self.multi_cell(0, 5, value)
        self.set_x(x)

    def tip_box(self, text):
        self.ln(2)
        y = self.get_y()
        self.set_fill_color(240, 250, 240)  # Very light green bg
        self.set_font("Helvetica", "", 9)
        lines = self.multi_cell(164, 5, text, dry_run=True, output="LINES")
        h = len(lines) * 5 + 8
        if y + h > 280:
            self.add_page()
            y = self.get_y()
        self.rect(20, y, 170, h, "F")
        self.set_fill_color(*GREEN_OK)
        self.rect(20, y, 1.5, h, "F")
        self.set_xy(25, y + 3)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*GREEN_OK)
        self.cell(10, 5, "TIP  ")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*DARK_GREY)
        self.multi_cell(155, 5, text)
        self.set_y(y + h + 2)

    def warning_box(self, text):
        self.ln(2)
        y = self.get_y()
        self.set_fill_color(255, 242, 240)  # Very light red bg
        self.set_font("Helvetica", "", 9)
        lines = self.multi_cell(164, 5, text, dry_run=True, output="LINES")
        h = len(lines) * 5 + 8
        if y + h > 280:
            self.add_page()
            y = self.get_y()
        self.rect(20, y, 170, h, "F")
        self.set_fill_color(*RED_WARN)
        self.rect(20, y, 1.5, h, "F")
        self.set_xy(25, y + 3)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*RED_WARN)
        self.cell(18, 5, "WARNING  ")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*DARK_GREY)
        self.multi_cell(147, 5, text)
        self.set_y(y + h + 2)

    def note_box(self, text):
        self.ln(2)
        y = self.get_y()
        self.set_fill_color(238, 242, 255)  # Very light blue bg
        self.set_font("Helvetica", "", 9)
        lines = self.multi_cell(164, 5, text, dry_run=True, output="LINES")
        h = len(lines) * 5 + 8
        if y + h > 280:
            self.add_page()
            y = self.get_y()
        self.rect(20, y, 170, h, "F")
        self.set_fill_color(*NOTE_BLUE)
        self.rect(20, y, 1.5, h, "F")
        self.set_xy(25, y + 3)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*NOTE_BLUE)
        self.cell(12, 5, "NOTE  ")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*DARK_GREY)
        self.multi_cell(153, 5, text)
        self.set_y(y + h + 2)

    def simple_table(self, headers, rows, col_widths=None):
        """Draw a styled table."""
        if col_widths is None:
            w = 174 / len(headers)
            col_widths = [w] * len(headers)

        # Check if table fits, otherwise new page
        needed = 7 + len(rows) * 6 + 4
        if self.get_y() + needed > 275:
            self.add_page()

        # Header row
        self.set_font("Helvetica", "B", 8.5)
        self.set_fill_color(*TABLE_HEADER)
        self.set_text_color(255, 255, 255)  # White text on black header
        x0 = 18
        self.set_x(x0)
        for i, h in enumerate(headers):
            self.cell(col_widths[i], 7, f" {h}", border=0, fill=True)
        self.ln()

        # Data rows
        self.set_font("Helvetica", "", 8.5)
        for ri, row in enumerate(rows):
            bg = TABLE_ROW_1 if ri % 2 == 0 else TABLE_ROW_2
            self.set_fill_color(*bg)
            self.set_text_color(*DARK_GREY)
            self.set_x(x0)
            for i, val in enumerate(row):
                self.cell(col_widths[i], 6, f" {val}", border=0, fill=True)
            self.ln()
        self.ln(2)

    def check_space(self, needed=30):
        """Add page if less than `needed` mm remain."""
        if self.get_y() + needed > 275:
            self.add_page()


def build_guide():
    pdf = VCGuide()
    pdf.alias_nb_pages()

    # ══════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════
    pdf.set_fill_color(*PAGE_BG)
    pdf.rect(0, 0, 210, 297, "F")
    pdf.set_fill_color(*ACCENT)
    pdf.rect(0, 0, 210, 1.5, "F")

    # Logo
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=70, y=45, w=70)

    # Title
    pdf.set_y(125)
    pdf.set_font("Helvetica", "B", 32)
    pdf.set_text_color(*BLACK)
    pdf.cell(0, 14, "User Guide", align="C", ln=True)

    pdf.ln(4)
    pdf.set_font("Helvetica", "", 13)
    pdf.set_text_color(*DARK_GREY)
    pdf.cell(0, 8, "Your complete guide to building investor-grade dashboards", align="C", ln=True)

    pdf.ln(8)
    pdf.set_draw_color(*BLACK)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y())
    pdf.ln(8)

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(*MID_GREY)
    pdf.cell(0, 7, "Supports:  AI-SaaS  |  D2C  |  Healthcare", align="C", ln=True)
    pdf.cell(0, 7, "58 Financial Sections  |  10 Chart Types  |  5-Year Projections", align="C", ln=True)

    pdf.ln(20)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*MID_GREY)
    pdf.cell(0, 6, "100% Client-Side  -  Your Data Never Leaves Your Browser", align="C", ln=True)
    pdf.cell(0, 6, "https://ashish202407.github.io/Ruban/", align="C", ln=True)

    # Version
    pdf.set_y(265)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*MID_GREY)
    pdf.cell(0, 5, "Version 2.0  |  March 2026", align="C", ln=True)

    # ══════════════════════════════════════════════════════════
    # TABLE OF CONTENTS
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.set_y(25)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*BLACK)
    pdf.cell(0, 12, "Table of Contents", ln=True)
    pdf.accent_line()
    pdf.ln(6)

    toc_items = [
        ("1", "Getting Started", "What you need and how the process works"),
        ("2", "The Excel Template", "Setup, Checklist, and sector data sheets explained"),
        ("3", "Filling the Template - Setup Sheet", "Company name, business type, currency, FY start"),
        ("4", "Filling the Template - Checklist Sheet", "Toggling sections on and off"),
        ("5", "Filling the Template - Sector Data", "Blue cells, formulas, and 5-year projections"),
        ("6", "AI-SaaS Sections (20)", "Every metric and what to enter"),
        ("7", "D2C Sections (20)", "Every metric and what to enter"),
        ("8", "Healthcare Sections (18)", "Every metric and what to enter"),
        ("9", "Uploading Your File", "Drag-and-drop, validation, and what to expect"),
        ("10", "The Dashboard", "Header, KPIs, charts, data tables, and categories"),
        ("11", "Chart Types", "Line, bar, waterfall, gauge, radar, and more"),
        ("12", "Readiness Radar", "Your startup scored against benchmarks"),
        ("13", "Theme & Palettes", "Dark/light mode and 10 color palettes"),
        ("14", "Exporting to PDF", "One-click print-friendly export"),
        ("15", "Currency & Number Formatting", "INR, USD, EUR, GBP and display conventions"),
        ("16", "Common Mistakes & Troubleshooting", "Fixing the most frequent issues"),
        ("17", "FAQ", "Quick answers to common questions"),
    ]

    for num, title, desc in toc_items:
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(10, 6, num + ".")
        pdf.set_text_color(*DARK_GREY)
        pdf.cell(70, 6, title)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*MID_GREY)
        pdf.cell(0, 6, desc, ln=True)
        pdf.ln(1)

    # ══════════════════════════════════════════════════════════
    # 1. GETTING STARTED
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("1", "Getting Started")

    pdf.body(
        "Dashboard Generator turns your financial projections into a polished, investor-grade dashboard "
        "in seconds. There is no sign-up, no server, and no data sharing - everything runs in your "
        "browser."
    )

    pdf.sub_title("What You Need")
    pdf.bullet("A modern web browser (Chrome, Edge, Firefox, or Safari)")
    pdf.bullet("Microsoft Excel, Google Sheets, or any spreadsheet app that can edit .xlsx files")
    pdf.bullet("Your 5-year financial projections")

    pdf.sub_title("The 3-Step Process")
    pdf.ln(1)
    pdf.key_value("Step 1", "Download the blank template (DashGen_Template_v2.xlsx)")
    pdf.key_value("Step 2", "Fill in your company details and financial data")
    pdf.key_value("Step 3", "Upload the filled file at the website - your dashboard appears instantly")

    pdf.tip_box(
        "Your data never leaves your device. The Excel file is read entirely in JavaScript "
        "inside your browser. No information is sent to any server."
    )

    pdf.sub_title("Supported Business Types")
    pdf.simple_table(
        ["Sector", "Sections", "Focus Areas"],
        [
            ["AI-SaaS", "20", "MRR/ARR, Churn, NDR, CAC/LTV, Rule of 40, P&L, Cap Table, and more"],
            ["D2C", "20", "GMV/AOV, Channel Revenue, Unit Economics, Balance Sheet, Cash Flow, and more"],
            ["Healthcare", "18", "Bed Occupancy, ARPOB, Payer Mix, OPD/IPD, Pharmacy, P&L, and more"],
        ],
        [30, 20, 124]
    )

    # ══════════════════════════════════════════════════════════
    # 2. THE EXCEL TEMPLATE
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("2", "The Excel Template")

    pdf.body(
        "The template is a single .xlsx workbook with 5 sheets. You only need to interact with "
        "the first 3 (Setup, Checklist, and your sector's data sheet)."
    )

    pdf.simple_table(
        ["Sheet", "Purpose", "Your Action"],
        [
            ["Setup", "Company metadata", "Fill all 4 fields"],
            ["Checklist", "Toggle sections on/off", "Set Yes or No for each section"],
            ["AI-SaaS", "AI-SaaS financial data", "Fill blue cells (if your sector)"],
            ["D2C", "D2C financial data", "Fill blue cells (if your sector)"],
            ["Healthcare", "Healthcare financial data", "Fill blue cells (if your sector)"],
        ],
        [30, 60, 84]
    )

    pdf.note_box(
        "You only need to fill the data sheet that matches your business type. For example, "
        "if you selected 'AI-SaaS' on the Setup sheet, fill only the AI-SaaS sheet. "
        "The other sector sheets can be left blank."
    )

    pdf.sub_title("How the Template is Organized")
    pdf.body(
        "Each sector data sheet contains rows organized into sections. Each section has:"
    )
    pdf.bullet("A section header row (bold, with the section name)")
    pdf.bullet("Input rows (blue text) - these are the cells you fill with your data")
    pdf.bullet("Formula rows (black text) - these calculate automatically from your inputs")
    pdf.bullet("Percentage/ratio rows (grey italic) - derived metrics computed from your data")
    pdf.bullet("5 columns of data: Year 1 through Year 5")

    pdf.warning_box(
        "Do not modify formula cells (black text). They contain formulas that auto-calculate. "
        "Overwriting them will produce incorrect dashboard results."
    )

    # ══════════════════════════════════════════════════════════
    # 3. SETUP SHEET
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("3", "Filling the Template - Setup Sheet")

    pdf.body(
        "The Setup sheet is the first thing you fill. It has 4 fields that configure your "
        "entire dashboard."
    )

    pdf.simple_table(
        ["Cell", "Field", "Options", "What It Does"],
        [
            ["B3", "Company Name", "Any text", "Displayed in the dashboard header"],
            ["B4", "Business Type", "AI-SaaS / D2C / Healthcare", "Determines which data sheet to read"],
            ["B5", "Currency", "INR / USD / EUR / GBP", "Sets currency symbol and formatting"],
            ["B6", "FY Start Month", "January - December", "Shown in dashboard metadata"],
        ],
        [14, 30, 56, 74]
    )

    pdf.tip_box(
        "Business Type is a dropdown - click the cell and select from the list. "
        "This must match the sector sheet you filled with data."
    )

    pdf.warning_box(
        "If you select 'AI-SaaS' as your business type but fill data in the D2C sheet, "
        "the dashboard will show no data. Always match the business type to the sheet you filled."
    )

    pdf.sub_title("Currency Options")
    pdf.body("The dashboard formats all monetary values using the currency you select:")
    pdf.simple_table(
        ["Currency", "Symbol", "Large Number Format"],
        [
            ["INR", "Rs.", "L (Lakhs), Cr (Crores)"],
            ["USD", "$", "K (Thousands), M (Millions)"],
            ["EUR", "E", "K (Thousands), M (Millions)"],
            ["GBP", "GBP", "K (Thousands), M (Millions)"],
        ],
        [40, 30, 104]
    )

    # ══════════════════════════════════════════════════════════
    # 4. CHECKLIST SHEET
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("4", "Filling the Template - Checklist Sheet")

    pdf.body(
        "The Checklist sheet controls which sections appear on your dashboard. Each row "
        "represents one section. Column D has a Yes/No dropdown - set it to 'Yes' for sections "
        "you want to include, 'No' for sections to hide."
    )

    pdf.sub_title("Default Settings")
    pdf.body(
        "When you open a fresh template:"
    )
    pdf.bullet("Original sections (first 10 for SaaS/D2C, first 9 for Healthcare) default to 'Yes'")
    pdf.bullet("New/advanced sections (the remaining 10/9) default to 'No'")
    pdf.body(
        "This means you get a solid dashboard with just the core sections. Turn on additional "
        "sections as you have the data for them."
    )

    pdf.tip_box(
        "Start with the defaults. You can always re-upload with more sections toggled on later. "
        "Only enable a section if you have actual data to fill it - empty sections will show "
        "blank charts."
    )

    pdf.sub_title("Checklist Structure")
    pdf.simple_table(
        ["Column", "Content", "Your Action"],
        [
            ["A", "Sector name", "Read only"],
            ["B", "Category (Revenue, Retention, etc.)", "Read only"],
            ["C", "Section name", "Read only"],
            ["D", "Include? (Yes/No)", "Select from dropdown"],
            ["E", "Section ID (hidden)", "Do not modify"],
        ],
        [25, 65, 84]
    )

    pdf.warning_box(
        "Do not edit column E (Section ID). It contains hidden identifiers that the dashboard "
        "uses to match your data. Changing these will break the dashboard."
    )

    # ══════════════════════════════════════════════════════════
    # 5. SECTOR DATA
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("5", "Filling the Template - Sector Data")

    pdf.body(
        "This is where you enter your actual financial numbers. Navigate to the sheet matching "
        "your business type (AI-SaaS, D2C, or Healthcare)."
    )

    pdf.sub_title("Understanding Cell Colors")
    pdf.simple_table(
        ["Cell Color", "Row Type", "What To Do"],
        [
            ["Blue text", "Input (I)", "Enter your data here"],
            ["Black text", "Formula (F)", "Do not edit - auto-calculated"],
            ["Grey italic", "Percent/Ratio (P)", "Do not edit - derived metric"],
            ["Bold black", "Total/Subtotal (T)", "Do not edit - summed automatically"],
            ["Bold section name", "Header (H/S)", "Section label - do not edit"],
        ],
        [30, 40, 104]
    )

    pdf.sub_title("Data Entry Rules")
    pdf.bullet("Enter numbers only - no currency symbols, no commas, no text in numeric cells")
    pdf.bullet("Use positive numbers. For losses or outflows, the formulas handle the signs")
    pdf.bullet("Percentages: enter as whole numbers (e.g., enter 15 for 15%, not 0.15)")
    pdf.bullet("Fill all 5 years (Year 1 through Year 5) for each input row")
    pdf.bullet("Leave a cell as 0 if you have no data for that year - do not leave it blank")

    pdf.tip_box(
        "Year 1 is your earliest year (could be historical), Year 5 is your furthest projection. "
        "Many investors focus on the Year 4 to Year 5 trend, which the dashboard highlights "
        "as a YoY badge on each section."
    )

    pdf.sub_title("Hidden Columns")
    pdf.body(
        "Columns H and I are hidden in the template. They contain:"
    )
    pdf.bullet("Column H: Section ID anchors (e.g., 'saas_mrr_arr') - used by the parser to locate sections")
    pdf.bullet("Column I: Row type markers (I, F, P, T, H, S) - used to style and process rows")
    pdf.warning_box(
        "Never unhide, edit, or delete columns H and I. They are critical for the dashboard parser. "
        "Modifying them will cause sections to not render or render incorrectly."
    )

    # ══════════════════════════════════════════════════════════
    # 6. AI-SaaS SECTIONS
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("6", "AI-SaaS Sections (20)")

    pdf.body(
        "If you selected AI-SaaS as your business type, these are the 20 sections available. "
        "The first 10 are enabled by default; the remaining 10 are optional."
    )

    # --- Original 10 ---
    pdf.sub_title("Core Sections (Enabled by Default)")

    saas_core = [
        ("MRR & ARR Progression", "Revenue", "Line",
         "Tracks your Monthly Recurring Revenue growth over 5 years.",
         ["Beginning MRR", "New MRR", "Expansion MRR", "Churned MRR"],
         ["Ending MRR", "ARR (Annual Recurring Revenue)"],
         "Ending MRR, ARR, MRR Growth %"),

        ("Revenue by Plan/Tier", "Revenue", "Stacked Bar",
         "Shows revenue breakdown across your pricing tiers.",
         ["Starter Revenue", "Pro Revenue", "Enterprise Revenue"],
         ["Total Revenue"],
         "Total Revenue, Revenue Growth %"),

        ("Churn & Retention", "Retention", "Line",
         "Measures customer logo churn and retention over time.",
         ["Starting Customers", "New Customers", "Churned Customers"],
         ["Ending Customers", "Logo Churn %"],
         "Logo Churn %, Ending Customers"),

        ("Net Dollar Retention (NDR)", "Retention", "Bar",
         "NDR above 100% means existing customers spend more each year - a key VC metric.",
         ["Beginning ARR", "Expansion ARR", "Churned ARR"],
         ["Ending ARR", "NDR %"],
         "NDR %, Ending ARR"),

        ("CAC & LTV", "Unit Economics", "Bar",
         "Customer Acquisition Cost vs Lifetime Value - shows unit economics health.",
         ["S&M Spend", "New Customers Acquired", "ARPA", "Gross Margin %", "Annual Churn Rate %"],
         ["CAC", "LTV", "LTV/CAC", "Payback (months)"],
         "CAC, LTV, LTV/CAC, Payback"),

        ("Rule of 40 & Burn Multiple", "Efficiency", "Gauge",
         "Rule of 40 = Revenue Growth % + EBITDA Margin %. Score above 40 is excellent.",
         ["Revenue Growth %", "EBITDA Margin %", "Net Burn", "Net New ARR"],
         ["Rule of 40 Score", "Burn Multiple"],
         "Rule of 40 Score, Burn Multiple"),

        ("Sales Efficiency", "Efficiency", "Bar",
         "The Magic Number measures how efficiently S&M spend converts to ARR.",
         ["Quarterly Net New ARR", "Previous Quarter S&M Spend"],
         ["Magic Number"],
         "Magic Number"),

        ("P&L Summary", "Financials", "Waterfall",
         "Full Profit & Loss statement rendered as a waterfall chart.",
         ["Revenue", "COGS", "R&D Expense", "S&M Expense", "G&A Expense", "D&A", "Interest", "Tax"],
         ["Gross Profit", "Total OPEX", "EBITDA", "PAT"],
         "Revenue, Gross Profit, EBITDA, PAT"),

        ("Cash & Runway", "Financials", "Bar",
         "Tracks your cash position and how many months of runway you have.",
         ["Opening Cash", "Cash Flow from Operations", "Equity Raised", "Monthly Burn Rate"],
         ["Closing Cash", "Runway (months)"],
         "Closing Cash, Runway (months)"),

        ("Fundraising & Cap Table", "Fundraising", "Stacked Area",
         "Shows your funding rounds and ownership dilution over time.",
         ["Round Name", "Amount Raised", "Pre-Money Valuation", "Post-Money Valuation", "Dilution %",
          "Founders %", "ESOP %", "Seed %", "Series A %", "Series B %", "Others %"],
         [],
         "Amount Raised"),
    ]

    for name, category, chart, desc, inputs, formulas, kpis in saas_core:
        pdf.check_space(45)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(37, 102, 176)
        pdf.cell(0, 5, "  Your inputs:", ln=True)
        for inp in inputs:
            pdf.bullet_light(inp, indent=10)
        if formulas:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*MID_GREY)
            pdf.cell(0, 5, "  Auto-calculated:", ln=True)
            for f in formulas:
                pdf.bullet_light(f, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    # --- New 10 ---
    pdf.add_page()
    pdf.sub_title("Advanced Sections (Off by Default)")

    saas_advanced = [
        ("Monthly Cohort MRR Retention", "Operations", "Line",
         "Tracks how well each customer cohort retains MRR over time.",
         ["Cohort 1-5 Retention %"],
         "Cohort 1 Year 5 Retention"),

        ("Feature Adoption by Plan", "Operations", "Bar",
         "Measures feature adoption rates across pricing tiers.",
         ["Free Tier, Starter, Pro, Enterprise Adoption %"],
         "Enterprise Adoption %"),

        ("API Usage & Overage Revenue", "Revenue", "Bar",
         "For API-first products: track usage volumes and overage billing.",
         ["API Calls (thousands), Included API Calls, Overage Rate"],
         "API Calls, Overage Revenue"),

        ("NPS Trend", "Operations", "Line",
         "Net Promoter Score trend with promoter/detractor breakdown.",
         ["NPS Score, Promoters %, Passives %, Detractors %"],
         "NPS Score"),

        ("Headcount & Rev per Employee", "Operations", "Stacked Bar",
         "Team size by function and revenue efficiency per employee.",
         ["R&D / S&M / G&A Headcount, Total Revenue"],
         "Total Headcount, Revenue per Employee"),

        ("Infrastructure Cost vs Revenue", "Efficiency", "Bar",
         "Cloud/hosting costs as a percentage of revenue.",
         ["Hosting & Cloud Cost, Total Revenue"],
         "Infra as % of Revenue"),

        ("ARR Bridge", "Revenue", "Bar",
         "Waterfall-style ARR movements: new, expansion, contraction, churn.",
         ["Beginning ARR, New ARR, Expansion ARR, Contraction ARR, Churned ARR"],
         "Net New ARR, Ending ARR"),

        ("Geographic Revenue Split", "Revenue", "Stacked Bar",
         "Revenue breakdown by region (India, US, Europe, Rest of World).",
         ["India, US, Europe, RoW Revenue"],
         "Total Revenue"),

        ("Support & CSAT", "Operations", "Line",
         "Support ticket volumes and customer satisfaction scores.",
         ["Total Tickets, Tickets Resolved, CSAT Score %, Avg Resolution Time"],
         "CSAT %, Total Tickets"),

        ("Token/Compute Cost per Customer", "Efficiency", "Bar",
         "For AI products: tracks inference cost efficiency per customer.",
         ["Total AI Inference Cost, Total Customers"],
         "Cost per Customer"),
    ]

    for name, category, chart, desc, inputs, kpis in saas_advanced:
        pdf.check_space(30)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLUE_INPUT)
        pdf.cell(0, 5, "  Your inputs: ", new_x="LMARGIN", new_y="NEXT")
        for inp in inputs:
            pdf.bullet_light(inp, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    # ══════════════════════════════════════════════════════════
    # 7. D2C SECTIONS
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("7", "D2C Sections (20)")

    pdf.body(
        "If you selected D2C as your business type, these are the 20 sections available. "
        "The first 10 are enabled by default."
    )

    pdf.sub_title("Core Sections (Enabled by Default)")

    d2c_core = [
        ("GMV & AOV", "Revenue", "Stacked Bar",
         "Gross Merchandise Value and Average Order Value across channels (Website, Retail, Marketplace).",
         ["Customers, Orders/Customer, AOV per channel"],
         "Total GMV, Blended AOV"),

        ("Channel Revenue Breakup", "Revenue", "Stacked Bar",
         "Net revenue per channel after returns and commissions.",
         ["GMV, Return Rate %, Commission % per channel"],
         "Total Net Revenue, Revenue Growth %"),

        ("Return Rates & Net Revenue", "Revenue", "Bar",
         "Gross-to-net revenue bridge showing impact of returns and commissions.",
         ["Gross Revenue, Returns, Commission per channel"],
         "Net Revenue per channel"),

        ("Customer Funnel", "Customers", "Bar",
         "Top-of-funnel visitors through to repeat customers.",
         ["Website Visitors, Leads Generated, Customers Converted, Repeat Customers"],
         "Visitors, Customers, Lead-to-Customer %"),

        ("Cohort Retention & Churn", "Retention", "Line",
         "Churn and repeat purchase rates by acquisition channel.",
         ["Churn Rate % per channel, Repeat Purchase Rate %"],
         "Blended Churn, Repeat Rate %"),

        ("Unit Economics", "Unit Economics", "Bar",
         "CAC, LTV, and payback economics per customer.",
         ["CAC, LTV, Payback Period, Contribution Margin per Order"],
         "CAC, LTV, LTV/CAC"),

        ("P&L Summary", "Financials", "Bar",
         "Full profit and loss statement with margins.",
         ["Net Revenue, COGS, R&D, S&M, G&A, D&A, Interest, Tax"],
         "Net Revenue, Gross Profit, EBITDA, PAT"),

        ("Balance Sheet", "Financials", "Stacked Bar",
         "Assets, liabilities, and equity snapshot.",
         ["Cash, Inventory, Receivables, Fixed Assets, Payables, Debt, Equity components"],
         "Total Assets, Total Equity"),

        ("Cash Flow", "Financials", "Bar",
         "Operating, investing, and financing cash flows.",
         ["EBITDA, Working Capital, Tax, Capex, Equity Raised, Debt"],
         "CF from Operations, Closing Cash"),

        ("Fundraising & Cap Table", "Fundraising", "Stacked Bar",
         "Funding rounds and ownership structure over time.",
         ["Round details, ownership percentages"],
         "Amount Raised"),
    ]

    for name, category, chart, desc, inputs, kpis in d2c_core:
        pdf.check_space(30)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        for inp in inputs:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*BLUE_INPUT)
            pdf.bullet_light(inp, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    pdf.add_page()
    pdf.sub_title("Advanced Sections (Off by Default)")

    d2c_advanced = [
        ("SKU-level Revenue", "Revenue", "Stacked Bar",
         "Revenue broken down by product category or SKU group.",
         ["SKU Category 1-4 Revenue"],
         "Total SKU Revenue"),

        ("Ad Spend & ROAS by Channel", "Marketing", "Bar",
         "Ad spend and return on ad spend across Meta, Google, and other channels.",
         ["Ad Spend and Revenue Attributed per channel"],
         "Blended ROAS"),

        ("Seasonal Revenue Index", "Revenue", "Stacked Bar",
         "Quarterly revenue patterns to highlight seasonality.",
         ["Q1-Q4 Revenue"],
         "Annual Revenue"),

        ("Refund & Dispute Rate", "Operations", "Line",
         "Track refund and payment dispute trends.",
         ["Total Orders, Refunds Issued, Disputes Filed"],
         "Refund Rate %, Dispute Rate %"),

        ("Loyalty Program Metrics", "Retention", "Bar",
         "Loyalty member base, activity, and point redemption.",
         ["Loyalty Members, Active Members, Repeat Rate %, Points Redeemed Value"],
         "Members, Repeat %"),

        ("Inventory Turnover", "Operations", "Bar",
         "How efficiently inventory is sold and replaced.",
         ["COGS, Average Inventory"],
         "Turnover Ratio, Days Inventory Outstanding"),

        ("Contribution Margin by Channel", "Unit Economics", "Bar",
         "Per-channel profitability after variable costs.",
         ["Revenue, Variable Cost per channel"],
         "Contribution Margin per channel"),

        ("LTV by Acquisition Channel", "Unit Economics", "Bar",
         "Customer lifetime value segmented by how they were acquired.",
         ["Organic, Paid Social, Paid Search, Referral LTV"],
         "LTV per channel"),

        ("WhatsApp & Email Marketing", "Marketing", "Bar",
         "Campaign volumes and revenue from email and WhatsApp channels.",
         ["Email Campaigns Sent, Open/Click Rate %, Email Revenue, WhatsApp Messages, WhatsApp Revenue"],
         "Email Revenue, WhatsApp Revenue"),

        ("Influencer & Affiliate Revenue", "Marketing", "Bar",
         "ROI from influencer and affiliate marketing spend.",
         ["Influencer/Affiliate Spend & Revenue"],
         "Influencer ROAS, Affiliate ROAS"),
    ]

    for name, category, chart, desc, inputs, kpis in d2c_advanced:
        pdf.check_space(28)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        for inp in inputs:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*BLUE_INPUT)
            pdf.bullet_light(inp, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    # ══════════════════════════════════════════════════════════
    # 8. HEALTHCARE SECTIONS
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("8", "Healthcare Sections (18)")

    pdf.body(
        "If you selected Healthcare as your business type, these are the 18 sections available. "
        "The first 9 are enabled by default."
    )

    pdf.sub_title("Core Sections (Enabled by Default)")

    hc_core = [
        ("Bed Occupancy & Utilization", "Operations", "Bar",
         "Tracks bed usage, occupancy rate, average length of stay, and discharges.",
         ["Total Beds Available, Occupied Bed Days, Total Available Bed Days, ALOS"],
         "Occupancy %, Total Beds, Discharges"),

        ("ARPOB & Revenue per Bed", "Revenue", "Bar",
         "Average Revenue Per Occupied Bed - a key hospital efficiency metric.",
         ["Inpatient Revenue, Occupied Bed Days"],
         "ARPOB, Revenue per Available Bed"),

        ("Department Revenue", "Revenue", "Stacked Bar",
         "Revenue split across hospital departments.",
         ["IPD, OPD, Pharmacy, Diagnostics, Surgical, Emergency Revenue"],
         "Total Revenue, Revenue Growth %"),

        ("Payer Mix", "Revenue", "Stacked Bar",
         "Revenue breakdown by payment source (insurance, cash, government, corporate).",
         ["Insurance, Cash/Self-Pay, Government, Corporate/TPA Revenue"],
         "Total Revenue"),

        ("OPD & IPD Volumes", "Operations", "Bar",
         "Outpatient and inpatient visit volumes with conversion tracking.",
         ["Annual OPD Visits, IPD Admissions"],
         "Annual OPD, IPD Admissions, OPD-to-IPD Conversion %"),

        ("Cost per Bed Day", "Cost", "Stacked Bar",
         "Operating cost breakdown per occupied bed day.",
         ["Medical Supplies Cost, Staff Cost, Pharmacy COGS, Inpatient Revenue"],
         "Cost per Occupied Bed Day, Contribution"),

        ("P&L Summary", "Financials", "Bar",
         "Profit and loss statement for the hospital.",
         ["Total Revenue, COGS, Operating Expenses, D&A, Interest, Tax"],
         "Revenue, Gross Profit, EBITDA, PAT"),

        ("Cash Flow & Capex", "Financials", "Bar",
         "Cash flow from operations, investing, and financing activities.",
         ["CF from Operations, CF from Investing, CF from Financing, Opening Cash"],
         "CF from Operations, Closing Cash"),

        ("Fundraising & Cap Table", "Fundraising", "Stacked Bar",
         "Funding rounds and ownership structure.",
         ["Round details, ownership percentages"],
         "Amount Raised"),
    ]

    for name, category, chart, desc, inputs, kpis in hc_core:
        pdf.check_space(30)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        for inp in inputs:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*BLUE_INPUT)
            pdf.bullet_light(inp, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    pdf.add_page()
    pdf.sub_title("Advanced Sections (Off by Default)")

    hc_advanced = [
        ("Procedure Volume by Department", "Operations", "Stacked Bar",
         "Surgical and procedure counts by department.",
         ["General Surgery, Orthopaedic, Cardiology, Oncology, Other Procedures"],
         "Total Procedures"),

        ("Insurance Claim Settlement", "Operations", "Bar",
         "Claims submitted vs approved, with settlement rate tracking.",
         ["Claims Submitted, Claims Approved, Claims Rejected"],
         "Approval Rate %"),

        ("Doctor Productivity", "Efficiency", "Bar",
         "Revenue generated per doctor on staff.",
         ["Total Doctors, Total Revenue"],
         "Revenue per Doctor"),

        ("Pharmacy Margin", "Revenue", "Bar",
         "Pharmacy gross profit and margin percentage.",
         ["Pharmacy Revenue, Pharmacy COGS"],
         "Pharmacy Margin %"),

        ("ICU Utilization", "Operations", "Bar",
         "ICU bed occupancy and utilization rates.",
         ["ICU Beds Available, ICU Occupied Bed Days"],
         "ICU Occupancy %"),

        ("Readmission Rate", "Quality", "Line",
         "Percentage of patients readmitted - a quality of care indicator.",
         ["Readmission Rate %"],
         "Readmission Rate %"),

        ("Patient Satisfaction Score", "Quality", "Line",
         "Overall, IPD, and OPD patient satisfaction trends.",
         ["Overall Score, IPD Score, OPD Score"],
         "Overall Satisfaction Score"),

        ("Revenue Cycle Metrics", "Financials", "Bar",
         "Gross-to-net revenue collection efficiency.",
         ["Gross Revenue, Net Revenue"],
         "Collection Rate %"),

        ("Capex & Asset Schedule", "Financials", "Stacked Bar",
         "Capital expenditure by category and net fixed asset position.",
         ["Equipment Capex, Building/Expansion Capex, IT Capex, Net Fixed Assets"],
         "Total Capex, Net Fixed Assets"),
    ]

    for name, category, chart, desc, inputs, kpis in hc_advanced:
        pdf.check_space(28)
        pdf.sub_sub_title(f"{name}  [{category}]  -  {chart} Chart")
        pdf.body_light(desc)
        for inp in inputs:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(*BLUE_INPUT)
            pdf.bullet_light(inp, indent=10)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*BLACK)
        pdf.cell(0, 5, f"  Dashboard KPIs: {kpis}", ln=True)
        pdf.ln(1)

    # ══════════════════════════════════════════════════════════
    # 9. UPLOADING YOUR FILE
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("9", "Uploading Your File")

    pdf.body(
        "Once your template is filled, head to the website to generate your dashboard."
    )

    pdf.sub_title("How to Upload")
    pdf.bullet("Go to https://ashish202407.github.io/Ruban/")
    pdf.bullet("You will see the upload screen with a drop zone")
    pdf.bullet("Either drag and drop your .xlsx file onto the drop zone, or click it to browse")
    pdf.bullet("The file is read instantly in your browser - no upload to any server")
    pdf.bullet("Your dashboard appears within seconds")

    pdf.sub_title("File Requirements")
    pdf.simple_table(
        ["Requirement", "Details"],
        [
            ["Format", ".xlsx or .xls (Excel format)"],
            ["Setup sheet", "Must have Company Name, Business Type, Currency filled"],
            ["Checklist", "At least one section must be set to 'Yes'"],
            ["Data", "At least some numeric data in the sector sheet"],
            ["File size", "No strict limit - processed locally"],
        ],
        [40, 134]
    )

    pdf.sub_title("What Happens After Upload")
    pdf.body("The dashboard processes your file in this order:")
    pdf.key_value("1. Parse Setup", "Reads company name, business type, currency, FY start")
    pdf.key_value("2. Parse Checklist", "Identifies which sections are toggled 'Yes'")
    pdf.key_value("3. Parse Data", "Reads your sector sheet, evaluates formulas, extracts 5-year values")
    pdf.key_value("4. Render", "Builds the dashboard: header, radar, KPIs, charts, and data tables")

    pdf.sub_title("Error Messages")
    pdf.body("If something goes wrong, you may see one of these messages:")
    pdf.simple_table(
        ["Error", "What It Means", "How to Fix"],
        [
            ["Please upload an Excel file", "Wrong file format", "Use .xlsx or .xls only"],
            ["No sections marked Yes", "Checklist is all 'No'", "Set at least one section to 'Yes'"],
            ["No data found", "Sector sheet is empty", "Fill in data on the correct sector sheet"],
            ["Error parsing file", "Corrupted or incompatible file", "Re-download template and re-fill"],
        ],
        [45, 50, 79]
    )

    # ══════════════════════════════════════════════════════════
    # 10. THE DASHBOARD
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("10", "The Dashboard")

    pdf.body(
        "After uploading, your dashboard renders with several components. Here is what you see, "
        "from top to bottom."
    )

    pdf.sub_title("Dashboard Header")
    pdf.body(
        "Displays your company name, business type, currency, FY start month, and "
        "'5-Year Projection' label. This information comes directly from your Setup sheet."
    )

    pdf.sub_title("Readiness Radar")
    pdf.body(
        "A 6-axis spider/radar chart that scores your startup against benchmarks specific to "
        "your sector. Each axis represents a key metric, and the chart shows how you compare "
        "to typical VC expectations. An overall score out of 100 is displayed."
    )

    pdf.sub_title("KPI Strip")
    pdf.body(
        "Up to 8 key performance indicators displayed as cards at the top of the dashboard. "
        "Each KPI card shows the metric name, its Year 5 value, and a trend arrow (up/down) "
        "comparing Year 4 to Year 5."
    )

    pdf.sub_title("Chart Grid")
    pdf.body(
        "The main body of the dashboard. Your enabled sections are displayed as cards in a "
        "3-column grid (4 columns on wide screens). Each card contains:"
    )
    pdf.bullet("Section title with a colored accent stripe on the left")
    pdf.bullet("YoY badge showing the Year 4 to Year 5 percentage change")
    pdf.bullet("A chart visualization (line, bar, waterfall, gauge, etc.)")
    pdf.bullet("An HTML legend below the chart")
    pdf.bullet("A 'View Data' button to toggle the raw data table")

    pdf.sub_title("Category Grouping")
    pdf.body("Sections are organized into categories with sticky headers:")
    categories = [
        "Revenue", "Retention", "Customers", "Unit Economics", "Operations",
        "Efficiency", "Marketing", "Cost", "Quality", "Financials", "Fundraising"
    ]
    pdf.body("  |  ".join(categories))
    pdf.body_light(
        "Categories appear in the order listed above. Only categories with active sections are shown."
    )

    pdf.sub_title("Data Tables")
    pdf.body(
        "Each chart card has a toggleable data table showing the raw numbers behind the chart. "
        "Click 'View Data' to expand or 'Hide Data' to collapse. Tables show all rows with "
        "Year 1-5 values, formatted with your selected currency."
    )

    # ══════════════════════════════════════════════════════════
    # 11. CHART TYPES
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("11", "Chart Types")

    pdf.body(
        "The dashboard uses 10 different chart types, automatically selected based on the section. "
        "You do not choose the chart type - it is determined by the data."
    )

    pdf.simple_table(
        ["Chart Type", "Used For", "How to Read It"],
        [
            ["Line", "Trends over time (MRR, Churn, NPS)", "Track the direction and slope of lines"],
            ["Bar", "Comparisons (CAC, LTV, Cash)", "Compare heights across years"],
            ["Stacked Bar", "Composition (Revenue by Plan)", "Each color is a component of the total"],
            ["Waterfall", "Flow analysis (P&L)", "Shows how values build up or break down"],
            ["Gauge", "Single score (Rule of 40)", "Needle position on semicircular scale"],
            ["Stacked Area", "Composition over time (Cap Table)", "Filled areas show ownership %"],
            ["Doughnut", "Proportions (Payer Mix)", "Slices show relative share"],
            ["Radar", "Multi-axis scoring (Readiness)", "6 axes, further from center = better"],
            ["Combo", "Dual metrics (Valuations)", "Bars + line on two Y-axes"],
            ["Donut KPI", "Single metric highlight", "Donut with center value display"],
        ],
        [28, 55, 91]
    )

    pdf.tip_box(
        "Hover over any chart to see detailed tooltips with exact values. "
        "All charts use interactive hover in index mode, so you see all data series at once."
    )

    # ══════════════════════════════════════════════════════════
    # 12. READINESS RADAR
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("12", "Readiness Radar")

    pdf.body(
        "The Readiness Radar is a 6-axis spider chart that automatically scores your "
        "startup against sector-specific benchmarks. It appears at the top of your dashboard "
        "and provides an overall score out of 100."
    )

    pdf.sub_title("How Scoring Works")
    pdf.body(
        "Each axis has a benchmark value considered 'good' for your sector. Your actual "
        "Year 5 metric is compared to this benchmark. The closer you are (or exceed), "
        "the higher your score on that axis. The overall score is the average of all 6 axes."
    )

    pdf.sub_title("Axes by Sector")
    pdf.simple_table(
        ["Axis", "AI-SaaS", "D2C", "Healthcare"],
        [
            ["1", "Revenue Growth", "GMV Growth", "Occupancy Rate"],
            ["2", "Gross Margin", "Gross Margin", "ARPOB Growth"],
            ["3", "NDR", "Repeat Rate", "OPD-IPD Conversion"],
            ["4", "LTV/CAC", "LTV/CAC", "Revenue Growth"],
            ["5", "Rule of 40", "EBITDA Margin", "EBITDA Margin"],
            ["6", "Runway", "Cash Position", "Cash Position"],
        ],
        [15, 50, 50, 59]
    )

    pdf.note_box(
        "The radar only uses data from your enabled sections. If a section needed for "
        "a radar axis is turned off in the Checklist, that axis may use a default or zero value."
    )

    # ══════════════════════════════════════════════════════════
    # 13. THEME & PALETTES
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("13", "Theme & Palettes")

    pdf.body(
        "The dashboard has a full theme system that lets you customize the look and feel."
    )

    pdf.sub_title("Dark / Light Mode")
    pdf.body(
        "Toggle between dark and light mode using the sun/moon icon in the navigation bar. "
        "Dark mode (default) uses a dark background with light text. Light mode uses a white "
        "background with dark text."
    )

    pdf.sub_title("Color Palettes")
    pdf.body("Each theme has 5 color palettes to choose from:")

    pdf.simple_table(
        ["Dark Theme Palettes", "Light Theme Palettes"],
        [
            ["Gold (default)", "Slate (default)"],
            ["Ocean", "Indigo"],
            ["Sage", "Teal"],
            ["Lavender", "Rose"],
            ["Copper", "Amber"],
        ],
        [87, 87]
    )

    pdf.body(
        "Select a palette from the dropdown in the navigation bar. Each palette provides "
        "8 color shades used for chart colors, accents, and card borders."
    )

    pdf.sub_title("Persistence")
    pdf.body(
        "Your theme and palette choices are saved in your browser's local storage. "
        "When you revisit the site, your last selected theme and palette are restored automatically."
    )

    pdf.tip_box(
        "Try different palettes to find one that matches your brand or presentation context. "
        "If presenting to investors, the default dark + gold palette gives a premium, "
        "investor-deck aesthetic."
    )

    # ══════════════════════════════════════════════════════════
    # 14. EXPORTING TO PDF
    # ══════════════════════════════════════════════════════════
    pdf.section_title("14", "Exporting to PDF")

    pdf.body(
        "The dashboard has a one-click PDF export feature."
    )

    pdf.sub_title("How to Export")
    pdf.bullet("Click the 'Export PDF' button in the navigation bar")
    pdf.bullet("The dashboard reformats for print: white background, 2-column layout, navigation hidden")
    pdf.bullet("Your browser's print dialog opens")
    pdf.bullet("Select 'Save as PDF' or your preferred printer")
    pdf.bullet("Recommended: Landscape orientation, 100% scale")

    pdf.tip_box(
        "For the best PDF output, use Chrome or Edge. The print layout is optimized for "
        "landscape orientation at 100% zoom. Expand all data tables you want included before exporting."
    )

    # ══════════════════════════════════════════════════════════
    # 15. CURRENCY & NUMBER FORMATTING
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("15", "Currency & Number Formatting")

    pdf.body(
        "The dashboard automatically formats numbers based on your currency selection and "
        "the type of metric being displayed."
    )

    pdf.sub_title("Number Display Formats")
    pdf.simple_table(
        ["Metric Type", "Format", "Example"],
        [
            ["Currency (INR)", "Rs. with L/Cr suffixes", "Rs. 1.5 Cr, Rs. 50 L"],
            ["Currency (USD/EUR/GBP)", "Symbol with K/M suffixes", "$1.5M, $500K"],
            ["Percentage", "1 decimal place", "15.2%"],
            ["Ratio", "1 decimal with 'x'", "3.5x"],
            ["Months", "0 decimals with 'mo'", "18 mo"],
            ["Count/Number", "0 decimals", "1,250"],
        ],
        [45, 55, 74]
    )

    pdf.sub_title("Data Entry vs Display")
    pdf.body(
        "In the Excel template, always enter raw numbers without formatting. The dashboard "
        "handles all display formatting automatically."
    )
    pdf.simple_table(
        ["Enter This", "Not This", "Dashboard Shows"],
        [
            ["1500000", "15,00,000 or 1.5M", "Rs. 15 L or $1.5M"],
            ["15", "15% or 0.15", "15.0%"],
            ["3.5", "3.5x", "3.5x"],
        ],
        [50, 62, 62]
    )

    # ══════════════════════════════════════════════════════════
    # 16. COMMON MISTAKES & TROUBLESHOOTING
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("16", "Common Mistakes & Troubleshooting")

    pdf.sub_title("Template Errors")

    mistakes = [
        ("Wrong business type selected",
         "You filled the D2C sheet but set Business Type to 'AI-SaaS' on the Setup sheet.",
         "Change the Business Type dropdown on the Setup sheet to match the sheet you filled."),

        ("Editing formula cells",
         "Black-text cells contain formulas. Overwriting them removes the calculation.",
         "Re-download a fresh template and only fill blue-text cells."),

        ("Editing hidden columns H or I",
         "These columns contain section anchors and row types that the dashboard depends on.",
         "Never unhide or modify columns H and I. If corrupted, re-download the template."),

        ("Leaving input cells blank",
         "Empty cells can cause formulas to return errors or $0 values.",
         "Enter 0 for any metric you do not have. Do not leave cells empty."),

        ("Entering formatted numbers",
         "Typing '$1,500' or '15%' in cells instead of plain numbers.",
         "Enter raw numbers only: 1500, 15. The dashboard formats them for display."),

        ("Changing section names or row labels",
         "The dashboard matches data by hidden column anchors, not by visible labels. However, "
         "changing labels may confuse the formula references.",
         "Do not rename rows or sections. The labels are there for your reference while filling."),

        ("Toggling sections without data",
         "Setting a section to 'Yes' on the Checklist but leaving its data empty.",
         "Only enable sections you have data for. Empty sections produce blank charts."),

        ("Using Google Sheets export",
         "Google Sheets may not preserve all Excel features (dropdowns, hidden columns) when "
         "exporting back to .xlsx.",
         "Best results with Microsoft Excel or LibreOffice. If using Google Sheets, "
         "verify dropdowns and hidden columns are intact after export."),
    ]

    for title, problem, fix in mistakes:
        pdf.check_space(25)
        pdf.sub_sub_title(title)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(200, 50, 50)
        pdf.cell(0, 5, "  Problem:", ln=True)
        pdf.body_light(f"    {problem}")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(40, 140, 70)
        pdf.cell(0, 5, "  Fix:", ln=True)
        pdf.body_light(f"    {fix}")

    pdf.add_page()
    pdf.sub_title("Dashboard Issues")

    dash_issues = [
        ("Dashboard shows $0 or blank values",
         "Formula cells may not have cached values if the file was saved without Excel recalculating.",
         "Open the file in Excel, press Ctrl+Shift+F9 to recalculate all formulas, then save and re-upload. "
         "The dashboard also has a built-in formula evaluator for common cases (SUM, IF, arithmetic)."),

        ("Charts are missing or not rendering",
         "The section may be toggled 'No' on the Checklist, or the data sheet may be empty for those rows.",
         "Check the Checklist sheet and ensure the section is set to 'Yes'. Verify the data is in the "
         "correct sector sheet."),

        ("Gauge chart appears empty",
         "The Rule of 40 gauge needs valid Revenue Growth % and EBITDA Margin % inputs.",
         "Ensure both inputs are filled with numeric values. The gauge renders via custom Canvas 2D "
         "and may need a moment to appear."),

        ("Theme or palette not saving",
         "Browser privacy settings may block localStorage.",
         "Check that your browser allows local storage for the site. Try a different browser if needed."),

        ("PDF export looks different",
         "Print mode uses a white background and 2-column layout regardless of your theme.",
         "This is expected. The print layout is optimized for readability on paper."),

        ("Very large numbers look wrong",
         "Extremely large or small values may not format correctly with the default suffixes.",
         "Ensure your data is in the same unit scale (e.g., all in actual values, not millions). "
         "The formatter handles up to Cr (10M) for INR and M (millions) for USD."),
    ]

    for title, problem, fix in dash_issues:
        pdf.check_space(25)
        pdf.sub_sub_title(title)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(200, 50, 50)
        pdf.cell(0, 5, "  Problem:", ln=True)
        pdf.body_light(f"    {problem}")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(40, 140, 70)
        pdf.cell(0, 5, "  Fix:", ln=True)
        pdf.body_light(f"    {fix}")

    # ══════════════════════════════════════════════════════════
    # 17. FAQ
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.section_title("17", "Frequently Asked Questions")

    faqs = [
        ("Is my data safe?",
         "Yes. 100% of the processing happens in your browser using JavaScript. Your Excel file "
         "is never uploaded to any server. No data leaves your device."),

        ("Can I use Google Sheets to fill the template?",
         "You can, but Microsoft Excel gives the best results. If using Google Sheets, download "
         "as .xlsx and verify that dropdowns and hidden columns are preserved."),

        ("What if I do not have 5 years of data?",
         "Enter 0 for years you do not have projections for. The dashboard will still render, "
         "but charts will show zero values for those years."),

        ("Can I customize which sections appear?",
         "Yes. Use the Checklist sheet to toggle individual sections on or off. Set column D "
         "to 'Yes' or 'No' for each section."),

        ("What is the YoY badge on each chart?",
         "It shows the percentage change from Year 4 to Year 5 for the primary metric in that "
         "section. Green with an up arrow means growth; red with a down arrow means decline."),

        ("Can I upload multiple files or compare companies?",
         "Currently, the dashboard supports one file at a time. To compare, export each as PDF "
         "and review side by side."),

        ("Does it work offline?",
         "Once the page is loaded, the core functionality works offline since it is all client-side. "
         "However, the initial page load requires an internet connection to fetch the JavaScript "
         "libraries (SheetJS, Chart.js) from CDN."),

        ("Which browsers are supported?",
         "Chrome, Edge, Firefox, and Safari (latest versions). Chrome and Edge give the best "
         "performance and PDF export quality."),

        ("Can I add my own sections?",
         "The current version supports the 58 predefined sections. Custom sections would require "
         "modifications to both the template generator and the JavaScript registry."),

        ("What currencies are supported?",
         "INR (Indian Rupees), USD (US Dollars), EUR (Euros), and GBP (British Pounds). "
         "Select your currency on the Setup sheet."),

        ("Why do some formulas show $0?",
         "If your spreadsheet app did not evaluate formulas before saving, the cached values "
         "may be missing. The dashboard has a built-in evaluator for SUM, IF, and basic "
         "arithmetic, but complex formulas may need Excel to recalculate first (Ctrl+Shift+F9)."),

        ("How do I get the template?",
         "Download DashGen_Template_v2.xlsx from the upload page on the website. "
         "There is a download link on the page itself."),
    ]

    for q, a in faqs:
        pdf.check_space(20)
        pdf.set_x(18)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*BLACK)
        pdf.multi_cell(174, 5.5, f"Q: {q}")
        pdf.set_x(18)
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(*DARK_GREY)
        pdf.multi_cell(174, 5, f"A: {a}")
        pdf.ln(3)

    # ══════════════════════════════════════════════════════════
    # BACK COVER
    # ══════════════════════════════════════════════════════════
    pdf.add_page()
    pdf.set_fill_color(*PAGE_BG)
    pdf.rect(0, 0, 210, 297, "F")

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=75, y=75, w=60)

    pdf.set_y(150)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(*BLACK)
    pdf.cell(0, 10, "Dashboard Generator", align="C", ln=True)

    pdf.ln(4)
    pdf.set_draw_color(*BLACK)
    pdf.line(70, pdf.get_y(), 140, pdf.get_y())
    pdf.ln(8)

    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(*DARK_GREY)
    pdf.cell(0, 7, "Fill an Excel template. Get an investor-grade dashboard.", align="C", ln=True)

    pdf.ln(12)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(*MID_GREY)
    pdf.cell(0, 7, "https://ashish202407.github.io/Ruban/", align="C", ln=True)

    pdf.ln(30)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*MID_GREY)
    pdf.cell(0, 6, "100% Private  |  No Server  |  No Sign-up", align="C", ln=True)
    pdf.cell(0, 6, "Your data never leaves your browser.", align="C", ln=True)

    # ── Save ──────────────────────────────────────────────────
    output_path = os.path.join(os.path.dirname(__file__), "DashGen_UserGuide.pdf")
    pdf.output(output_path)
    print(f"\nUser guide generated: {output_path}")
    print(f"Pages: {pdf.pages_count}")


if __name__ == "__main__":
    build_guide()
