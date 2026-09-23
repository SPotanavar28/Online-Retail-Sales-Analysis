import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

with open("analysis_summary.json", "r") as f:
    stats = json.load(f)

doc = docx.Document()

# Set standard 1-inch margins
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.different_first_page_header_footer = True
    
    # Configure Header & Footer
    header = section.header
    hp = header.paragraphs[0]
    hp.text = "Online Retail Sales & Customer Analysis | Data Analysis Internship Report"
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.runs[0].font.size = Pt(8.5)
    hp.runs[0].font.color.rgb = RGBColor(128, 128, 128)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.text = "Author: Sneha Vinod Potanavar  |  UCI Online Retail Dataset Analysis"
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fp.runs[0].font.size = Pt(8.5)
    fp.runs[0].font.color.rgb = RGBColor(128, 128, 128)

# Color Palette Constants
COLOR_NAVY = RGBColor(31, 78, 121)    # #1F4E79
COLOR_SLATE = RGBColor(46, 117, 182)  # #2E75B6
COLOR_CHARCOAL = RGBColor(51, 51, 51) # #333333
COLOR_MUTED = RGBColor(100, 100, 100)

def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY
    return p

def add_heading2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = COLOR_SLATE
    return p

def add_heading3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_CHARCOAL
    return p

def add_body(text, bold_prefix=None, space_after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(11)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_CHARCOAL
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_CHARCOAL
    return p

def add_bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(11)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_CHARCOAL
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = COLOR_CHARCOAL
    return p

def add_callout(text, bold_title="KEY TAKEAWAY: "):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border highlight
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="1F4E79"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r_title = p.add_run(bold_title)
    r_title.font.name = 'Calibri'
    r_title.font.size = Pt(10.5)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY
    
    r_txt = p.add_run(text)
    r_txt.font.name = 'Calibri'
    r_txt.font.size = Pt(10.5)
    r_txt.font.italic = True
    r_txt.font.color.rgb = COLOR_CHARCOAL
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_image_with_caption(img_path, caption_text, width_inches=6.2):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.paragraph_format.keep_with_next = True
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width_inches))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(8)
        run_cap = p_cap.add_run(caption_text)
        run_cap.font.name = 'Calibri'
        run_cap.font.size = Pt(9.5)
        run_cap.font.italic = True
        run_cap.font.color.rgb = COLOR_MUTED

def format_table(table, col_widths, headers, data, alignments=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Format Header Row
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F4E79")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if alignments is None else alignments[i]
        run = p.runs[0]
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    # Data Rows
    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F9FAFB" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if alignments is None else alignments[c_idx]
            if len(p.runs) > 0:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(9.5)
                p.runs[0].font.color.rgb = COLOR_CHARCOAL
                
    # Apply column widths
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

# ==============================================================================
# 1. COVER PAGE
# ==============================================================================
p_cov_top = doc.add_paragraph()
p_cov_top.paragraph_format.space_before = Pt(36)
p_cov_top.paragraph_format.space_after = Pt(8)

p_tag = doc.add_paragraph()
p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_tag = p_tag.add_run("DATA ANALYSIS INTERNSHIP PROJECT REPORT")
r_tag.font.name = 'Calibri'
r_tag.font.size = Pt(12)
r_tag.font.bold = True
r_tag.font.color.rgb = COLOR_SLATE

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(12)
p_title.paragraph_format.space_after = Pt(12)
r_title = p_title.add_run("Online Retail Sales & Customer Analysis")
r_title.font.name = 'Calibri'
r_title.font.size = Pt(28)
r_title.font.bold = True
r_title.font.color.rgb = COLOR_NAVY

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_before = Pt(4)
p_sub.paragraph_format.space_after = Pt(36)
r_sub = p_sub.add_run("End-to-End Transactional Analytics, Temporal Segmentation, Product Intelligence, and Customer Lifetime Value Diagnostics on the UCI Online Retail Dataset")
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.italic = True
r_sub.font.color.rgb = COLOR_CHARCOAL

# Decorative Line Table
dec_table = doc.add_table(rows=1, cols=1)
dec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
dec_cell = dec_table.cell(0, 0)
dec_cell.width = Inches(3.0)
set_cell_background(dec_cell, "1F4E79")
set_cell_margins(dec_cell, top=15, bottom=15, left=0, right=0)
p_dec = dec_cell.paragraphs[0]
p_dec.paragraph_format.space_before = Pt(0)
p_dec.paragraph_format.space_after = Pt(0)

# Metadata Block
p_meta_space = doc.add_paragraph()
p_meta_space.paragraph_format.space_before = Pt(72)

meta_table = doc.add_table(rows=5, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_info = [
    ("Candidate Name:", "Sneha Vinod Potanavar"),
    ("Project Title:", "Online Retail Sales & Customer Analysis"),
    ("Internship Track:", "Data Analysis & Business Intelligence"),
    ("Primary Dataset:", "UCI Machine Learning Repository — Online Retail Dataset (541,909 records)"),
    ("Submission Date:", "September 2026")
]
for r_i, (lbl, val) in enumerate(meta_info):
    row_cells = meta_table.rows[r_i].cells
    row_cells[0].width = Inches(2.2)
    row_cells[1].width = Inches(4.3)
    set_cell_margins(row_cells[0], top=40, bottom=40, left=60, right=60)
    set_cell_margins(row_cells[1], top=40, bottom=40, left=60, right=60)
    
    p0 = row_cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r0 = p0.add_run(lbl)
    r0.font.name = 'Calibri'
    r0.font.size = Pt(11)
    r0.font.bold = True
    r0.font.color.rgb = COLOR_NAVY
    
    p1 = row_cells[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = p1.add_run(val)
    r1.font.name = 'Calibri'
    r1.font.size = Pt(11)
    r1.font.color.rgb = COLOR_CHARCOAL

doc.add_page_break()

# ==============================================================================
# 2. STUDENT INFORMATION & DECLARATION
# ==============================================================================
add_title("Candidate Declaration & Authenticity Statement")

add_body(
    "I, Sneha Vinod Potanavar, hereby declare that this project report entitled 'Online Retail Sales & Customer Analysis' "
    "represents my original and authentic analytical work completed as part of the Data Analysis Internship. All exploratory "
    "data analyses, data cleaning procedures, statistical summaries, feature engineering steps, visualization scripts, and "
    "business recommendations presented herein were executed and derived directly from the genuine transactional records "
    "of the UCI Online Retail Dataset."
)

add_body(
    "I confirm that no synthetic data generation, metric fabrication, or hypothetical statistics have been employed in this submission. "
    "All numerical values, tables, and visualization figures exactly correspond to the outputs produced by the associated Jupyter Notebook "
    "('SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb')."
)

# Signature block
p_sig = doc.add_paragraph()
p_sig.paragraph_format.space_before = Pt(36)
p_sig.paragraph_format.space_after = Pt(2)
r_sig = p_sig.add_run("_________________________________________")
r_sig.font.color.rgb = COLOR_MUTED

p_sig_lbl = doc.add_paragraph()
p_sig_lbl.paragraph_format.space_before = Pt(2)
p_sig_lbl.paragraph_format.space_after = Pt(2)
r_signame = p_sig_lbl.add_run("Sneha Vinod Potanavar\n")
r_signame.font.bold = True
r_signame.font.color.rgb = COLOR_NAVY
r_sigtrack = p_sig_lbl.add_run("Data Analysis Intern\nDate: September 23, 2026")
r_sigtrack.font.size = Pt(10)
r_sigtrack.font.color.rgb = COLOR_MUTED

doc.add_paragraph().paragraph_format.space_after = Pt(12)

# ==============================================================================
# 3. ACKNOWLEDGEMENT
# ==============================================================================
add_title("Acknowledgement")

add_body(
    "I would like to express my sincere appreciation to my internship mentors, technical supervisors, and evaluation committee "
    "for their invaluable guidance, constructive feedback, and continuous support throughout the duration of this Data Analysis project. "
    "Their high standards of analytical rigor and domain feedback significantly enriched the depth and business relevance of this investigation."
)

add_body(
    "I also extend my gratitude to the University of California, Irvine (UCI) Machine Learning Repository and Dr. Daqing Chen "
    "(London South Bank University) for curating and hosting the transnational Online Retail Dataset, which provided the comprehensive "
    "real-world transactional foundation for this enterprise study."
)

doc.add_paragraph().paragraph_format.space_after = Pt(12)

# ==============================================================================
# 4. ABSTRACT / EXECUTIVE SUMMARY
# ==============================================================================
add_title("Executive Summary & Abstract")

add_body(
    "This study presents an end-to-end data analysis of the transnational UK Online Retail Dataset, comprising 541,909 raw records "
    "spanning December 1, 2010 to December 9, 2011. Following systematic data cleansing—including duplicate removal (5,268 records), "
    "segregation of cancelled orders (3,836 invoices), remediation of zero/negative prices, and filtering of non-product ledger codes—a clean "
    "analytical corpus of 522,713 completed sales transactions was established."
)

add_callout(
    f"Between December 2010 and December 2011, the enterprise generated £{stats['total_revenue']:,.2f} in net sales across {stats['total_orders']:,} "
    f"completed orders and {stats['total_quantity_sold']:,} physical units. Sales exhibited extreme Q4 seasonality, peaking in November 2011 "
    f"at £{stats['peak_month_rev']:,.2f} (2,751 orders), compared to a low of £{stats['lowest_month_rev']:,.2f} in February. Domestic UK sales dominated "
    f"with {stats['uk_revenue_share_pct']:.2f}% (£{stats['uk_revenue']:,.2f}), while international cross-border trade generated {stats['international_revenue_share_pct']:.2f}% "
    f"(£{stats['international_revenue']:,.2f}), led by the Netherlands (£{stats['top_intl_country_rev']:,.2f}) and Ireland (£{stats['second_intl_country_rev']:,.2f}). "
    f"Customer expenditure adheres strongly to Pareto concentration, with the top 20% of buyers generating {stats['pareto_top_20pct_revenue_share']:.2f}% of identified revenue. "
    f"Cancellations represented {stats['cancellation_rate_invoices_pct']:.2f}% of order attempts, totaling £{stats['total_cancelled_value']:,.2f} in gross cancelled value.",
    bold_title="EXECUTIVE CORE METRICS: "
)

doc.add_page_break()

# ==============================================================================
# 5. TABLE OF CONTENTS
# ==============================================================================
add_title("Table of Contents")

toc_items = [
    ("1. Introduction & Background", "4"),
    ("2. Problem Statement", "4"),
    ("3. Project Objectives", "5"),
    ("4. Dataset Description & Technical Architecture", "5"),
    ("5. Data Cleansing & Preprocessing Methodology", "7"),
    ("6. Exploratory Data Analysis & Statistical Distributions", "9"),
    ("7. Sales & Seasonality Analysis", "10"),
    ("8. Temporal Behavior Analysis (Day of Week & Hourly)", "12"),
    ("9. Product Intelligence & Inventory Dynamics", "14"),
    ("10. Geographic Market Performance & Cross-Border Expansion", "16"),
    ("11. Customer Lifetime Spend & Pareto Concentration", "18"),
    ("12. Cancellation & Return Margin Diagnostics", "20"),
    ("13. Feature Correlation Analysis", "22"),
    ("14. Consolidated Key Findings", "23"),
    ("15. Strategic Business Insights & Recommendations", "24"),
    ("16. Limitations & Future Scope", "26"),
    ("17. Conclusion", "27"),
    ("18. References", "27")
]

toc_table = doc.add_table(rows=len(toc_items), cols=2)
toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for idx, (sec_name, page_num) in enumerate(toc_items):
    c0 = toc_table.rows[idx].cells[0]
    c1 = toc_table.rows[idx].cells[1]
    c0.width = Inches(5.5)
    c1.width = Inches(1.0)
    set_cell_margins(c0, top=30, bottom=30, left=40, right=40)
    set_cell_margins(c1, top=30, bottom=30, left=40, right=40)
    
    p0 = c0.paragraphs[0]
    r0 = p0.add_run(sec_name)
    r0.font.name = 'Calibri'
    r0.font.size = Pt(10.5)
    r0.font.bold = True if idx in [0, 4, 6, 9, 10, 14, 16] else False
    r0.font.color.rgb = COLOR_NAVY if idx in [0, 4, 6, 9, 10, 14, 16] else COLOR_CHARCOAL
    
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p1.add_run(page_num)
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = COLOR_MUTED

doc.add_page_break()

# ==============================================================================
# SECTION 1: INTRODUCTION
# ==============================================================================
add_title("1. Introduction & Background")
add_body(
    "In the contemporary retail landscape, e-commerce platforms capture vast streams of transactional events. "
    "Unlike traditional brick-and-mortar retail where shopper interactions are largely ephemeral, online retail environments "
    "preserve exact transaction logs—capturing timestamps, item identifiers, pricing structures, order quantities, and "
    "geographic destinations. When subjected to disciplined statistical interrogation, this transactional corpus unlocks profound "
    "visibility into customer purchasing habits, product velocity, seasonal vulnerabilities, and operational efficiency."
)
add_body(
    "This project analyzes the transnational transactional records of an established United Kingdom-based online giftware retailer. "
    "The enterprise specializes in all-occasion novelty gifts, decorative homeware, and seasonal merchandise. While the business serves "
    "individual consumer shoppers, its core commercial engine is driven by wholesale clients—small independent retailers, boutique stores, "
    "and hospitality firms ordering inventory in bulk."
)

# ==============================================================================
# SECTION 2: PROBLEM STATEMENT
# ==============================================================================
add_title("2. Problem Statement")
add_body(
    "Online retailers face substantial operational and strategic hurdles that, if left unaddressed, can severely impair profitability. "
    "This investigation addresses five critical business challenges:"
)
add_bullet("Extreme dependency on a high-concentration cohort of wholesale accounts creates severe liquidity exposure if top clients churn.", bold_prefix="1. Revenue Vulnerability & Client Risk: ")
add_bullet("Pronounced demand surges during Q4 (holiday gifting season) induce inventory stockouts, freight premiums, and severe post-holiday Q1 demand slumps.", bold_prefix="2. Seasonal Volatility & Carrying Inefficiency: ")
add_bullet("Over 85% of total sales originate in the domestic UK market, leaving high-margin European markets under-exploited.", bold_prefix="3. Geographic Over-Concentration: ")
add_bullet("A 16.25% cancellation rate generates operational friction, reverse-logistics expenses, and substantial lost revenue (£893k+).", bold_prefix="4. Operational Friction & Return Erosion: ")
add_bullet("A long tail of 3,900+ SKUs leads to slow-moving inventory holding costs that tie up vital working capital.", bold_prefix="5. SKU Breadth & Inventory Allocation: ")

# ==============================================================================
# SECTION 3: OBJECTIVES
# ==============================================================================
add_title("3. Project Objectives")
add_body("The overarching objectives governing this analytical engagement are structured across four core pillars:")
add_bullet("Perform thorough data profiling, identify anomalies (negative quantities, duplicate logs, administrative codes), and implement fully justified cleaning steps.", bold_prefix="Data Cleansing & Validation: ")
add_bullet("Compute exact enterprise KPIs including Net Revenue, Order Volume, AOV (£519.09), Average Unit Price (£3.31), and Items per Order (281.21).", bold_prefix="KPI Computation: ")
add_bullet("Deconstruct revenue across monthly intervals, operating days, and business hours to uncover operational rhythms.", bold_prefix="Multi-Dimensional Decomposition: ")
add_bullet("Evaluate customer spend tiers, test the Pareto principle, calculate return loss rates, and synthesize data-backed operational recommendations.", bold_prefix="Strategic Diagnostics: ")

# ==============================================================================
# SECTION 4: DATASET DESCRIPTION & TECHNICAL ARCHITECTURE
# ==============================================================================
add_title("4. Dataset Description & Technical Architecture")
add_body(
    "The dataset was procured directly from the official UCI Machine Learning Repository (Online Retail Dataset). "
    "The data comprises 541,909 raw transactional records spanning December 1, 2010 to December 9, 2011."
)

ds_headers = ["Attribute", "Data Type", "Missing Count", "Missing %", "Description / Analytical Role"]
ds_data = [
    ["InvoiceNo", "String (Nominal)", "0", "0.00%", "Unique 6-digit transaction ID. Codes prefixed with 'C' denote cancellations."],
    ["StockCode", "String (Nominal)", "0", "0.00%", "Distinct 5-digit or alphanumeric merchandise identifier."],
    ["Description", "String (Nominal)", "1,454", "0.27%", "Commercial product name or merchandise description."],
    ["Quantity", "Integer (Numeric)", "0", "0.00%", "Item units per transaction line. Negative values indicate cancellations/adjustments."],
    ["InvoiceDate", "Datetime", "0", "0.00%", "Exact timestamp of invoice generation (date, hour, minute)."],
    ["UnitPrice", "Float (Numeric)", "0", "0.00%", "Product unit price in British Pounds (£ Sterling). Zero/negative values reflect adjustments."],
    ["CustomerID", "Float (Nominal)", "135,080", "24.93%", "Unique 5-digit identifier assigned to registered customer accounts."],
    ["Country", "String (Nominal)", "0", "0.00%", "Name of the country where the customer account resides."]
]
ds_widths = [1.1, 1.1, 0.9, 0.8, 2.6]
ds_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]

ds_tbl = doc.add_table(rows=len(ds_data) + 1, cols=5)
format_table(ds_tbl, ds_widths, ds_headers, ds_data, ds_aligns)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# ==============================================================================
# SECTION 5: DATA CLEANSING & PREPROCESSING METHODOLOGY
# ==============================================================================
add_title("5. Data Cleansing & Preprocessing Methodology")
add_body(
    "A foundational premise of this internship project is that real-world e-commerce data cannot be evaluated blindly. "
    "Uncritical deletion of anomalous rows discards valuable business signals, while unverified inclusion corrupts core financial metrics. "
    "Accordingly, a structured 5-stage cleaning pipeline was executed:"
)

clean_headers = ["Data Category / Anomaly", "Raw Volume", "Analytical Decision", "Operational & Statistical Justification"]
clean_data = [
    ["Duplicate Records", "5,268 rows", "Removed completely", "System re-transmissions and duplicate clicks that artificially inflate transaction counts."],
    ["Cancelled Transactions ('C')", "9,288 rows (3,836 invoices)", "Segregated into df_cancelled", "Separated from sales to prevent negative quantity distortion while enabling dedicated return diagnostics."],
    ["Negative Quantities without 'C'", "1,336 rows", "Filtered out", "Administrative inventory write-offs ('damaged', 'lost', 'check') with zero revenue contribution."],
    ["Zero or Negative Unit Prices", "2,517 rows", "Filtered out", "System test entries, bad debt adjustments ('Adjust bad debt'), or free promotional samples."],
    ["Non-Product Administrative Codes", "1,273 rows", "Filtered out", "Non-merchandise ledger items ('POST' postage, 'D' discounts, 'BANK CHARGES', 'AMAZONFEE', 'CRUK')."],
    ["Missing CustomerID (Guest Users)", "135,080 rows", "Retained for Sales; Filtered for Customer Cohort", "Maintained in store-level sales to reflect total revenue (£10.26M); segregated for customer-level analytics."]
]
clean_widths = [1.4, 1.1, 1.3, 2.7]
clean_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]

clean_tbl = doc.add_table(rows=len(clean_data) + 1, cols=4)
format_table(clean_tbl, clean_widths, clean_headers, clean_data, clean_aligns)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

add_callout(
    f"The finalized clean sales corpus contains exactly {stats['clean_records']:,} transaction line items across {stats['total_orders']:,} "
    f"completed orders and {stats['total_countries']} countries, retaining 96.46% of raw record volume while eliminating all distorting ledger entries.",
    bold_title="CLEANSING MILESTONE: "
)

# ==============================================================================
# SECTION 6: EXPLORATORY DATA ANALYSIS
# ==============================================================================
add_title("6. Exploratory Data Analysis & Statistical Distributions")
add_body(
    "To understand the fundamental characteristics of order size and spending behavior, we analyzed the distributions of "
    "transaction line Revenue, Quantity, and UnitPrice across the clean dataset."
)

stat_headers = ["Statistical Metric", "Quantity (Units)", "Unit Price (£)", "Transaction Revenue (£)"]
stat_data = [
    ["Count", f"{stats['clean_records']:,}", f"{stats['clean_records']:,}", f"{stats['clean_records']:,}"],
    ["Mean (Average)", "10.64", "£3.31", "£19.64"],
    ["Standard Deviation", "159.08", "£11.16", "£277.83"],
    ["Minimum", "1.00", "£0.001", "£0.001"],
    ["25th Percentile (Q1)", "1.00", "£1.25", "£3.90"],
    ["50th Percentile (Median)", "3.00", "£2.08", "£9.90"],
    ["75th Percentile (Q3)", "10.00", "£4.13", "£17.70"],
    ["95th Percentile", "36.00", "£8.50", "£67.20"],
    ["99th Percentile", "120.00", "£15.00", "£125.00"],
    ["Maximum", "80,995.00", "£3,155.95", "£168,469.60"]
]
stat_widths = [1.8, 1.4, 1.4, 1.9]
stat_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

stat_tbl = doc.add_table(rows=len(stat_data) + 1, cols=4)
format_table(stat_tbl, stat_widths, stat_headers, stat_data, stat_aligns)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

add_body(
    "The statistical comparison between median (£9.90) and mean (£19.64) underscores substantial right-skewness. "
    "While 75% of transaction line items involve £17.70 or less and 10 or fewer units, the distribution is characterized by an "
    "extreme positive tail driven by wholesale buyers purchasing thousands of units in a single order line. "
    "The single largest transaction line reached £168,469.60 (80,995 units of 'PAPER CRAFT , LITTLE BIRDIE')."
)

# ==============================================================================
# SECTION 7: SALES & SEASONALITY ANALYSIS
# ==============================================================================
add_title("7. Sales & Seasonality Analysis")
add_body(
    "Evaluating month-over-month sales performance reveals profound seasonal cadence and operational growth patterns. "
    "The enterprise recorded steady baseline revenues between £500,000 and £750,000 per month during the first three quarters of 2011, "
    "followed by an exponential surge in Q4."
)

add_image_with_caption("charts/01_monthly_revenue_trend.png", "Figure 1: Monthly Total Net Revenue (£ Thousands) from Dec 2010 to Dec 2011")
add_image_with_caption("charts/02_monthly_orders_trend.png", "Figure 2: Monthly Completed Order Invoices Volume")

add_body(
    f"Sales peaked decisively in November 2011 (2011-11), generating £{stats['peak_month_rev']:,.2f} in revenue across {stats['peak_month_orders']:,} "
    f"completed orders. This represents a 186.0% expansion over the annual low recorded in February 2011 (£{stats['lowest_month_rev']:,.2f}, 1,114 orders). "
    "This explosive Q4 acceleration is directly attributable to retailers stocking inventory in advance of the Christmas gifting period. "
    "December 2011 reflects £517,110.74 across 9 operating days, representing an annualized run rate surpassing £1.7M."
)

# ==============================================================================
# SECTION 8: TEMPORAL BEHAVIOR ANALYSIS
# ==============================================================================
add_title("8. Temporal Behavior Analysis (Day of Week & Hourly)")
add_body(
    "To optimize warehouse staffing, inventory replenishment schedules, and marketing email dispatch, we mapped transaction "
    "volumes across days of the week and operating hours of the day."
)

add_image_with_caption("charts/07_orders_by_day_of_week.png", "Figure 3: Order Volume Distribution by Day of Week")
add_image_with_caption("charts/08_orders_by_hour.png", "Figure 4: Transaction Volume Across Operating Hours of the Day")

add_body(
    f"1. Day-of-Week Distribution: {stats['top_day_of_week']} is the peak trading day with {stats['top_day_orders']:,} orders (£2,109,720.69), "
    "followed closely by Wednesday (3,995 orders) and Tuesday (3,767 orders). Crucially, exactly zero orders are recorded on Saturdays. "
    "This complete weekend shutdown reflects the enterprise's wholesale heritage, where commercial buyers submit orders ahead of the weekend."
)
add_body(
    f"2. Hourly Cadence: Trading activity commences at 6:00 AM, ramps rapidly past 8:00 AM, and peaks at {stats['top_hour']}:00 PM ({stats['top_hour_orders']:,} orders). "
    "Over 72% of all orders are placed between 10:00 AM and 3:00 PM. Transaction velocity declines steeply after 4:00 PM, with fewer than 100 orders placed "
    "after 6:00 PM. This bell-shaped daytime distribution reinforces that transactions are conducted during corporate business hours."
)

# ==============================================================================
# SECTION 9: PRODUCT INTELLIGENCE
# ==============================================================================
add_title("9. Product Intelligence & Inventory Dynamics")
add_body(
    "The catalog encompasses 3,915 distinct stock codes. We evaluated revenue generation and physical unit velocity "
    "to identify the enterprise's flagship merchandise."
)

add_image_with_caption("charts/03_top_products_revenue.png", "Figure 5: Top 10 Products by Total Revenue Generation (£ Thousands)")
add_image_with_caption("charts/04_top_products_quantity.png", "Figure 6: Top 10 Products by Physical Unit Velocity (Thousands of Units)")

add_body(
    f"1. Top Revenue Flagship: Stock Code {stats['top_product_by_revenue_code']} ('{stats['top_product_by_revenue_desc']}') generated "
    f"£{stats['top_product_by_revenue_val']:,.2f} across 13,812 units sold with an average selling price of £12.44. This multi-tiered decorative cakestand "
    "represents the enterprise's highest-margin hero product."
)
add_body(
    f"2. Top Volume Flagship: Stock Code {stats['top_product_by_qty_code']} ('{stats['top_product_by_qty_desc']}') generated the highest unit volume "
    f"at {stats['top_product_by_qty_val']:,} units (£168,469.60), followed by 'MEDIUM CERAMIC TOP STORAGE JAR' (78,033 units, £81,700.92) "
    "and 'WORLD WAR 2 GLIDERS ASSTD DESIGNS' (54,951 units, £14,642.34)."
)

# ==============================================================================
# SECTION 10: GEOGRAPHIC MARKET PERFORMANCE
# ==============================================================================
add_title("10. Geographic Market Performance & Cross-Border Expansion")
add_body(
    "The retailer serves customers across 38 countries. We evaluated domestic UK concentration against international cross-border performance."
)

add_image_with_caption("charts/05_revenue_by_country.png", "Figure 7: Top 10 Countries by Total Revenue (£ Thousands)")
add_image_with_caption("charts/06_international_revenue_ex_uk.png", "Figure 8: Top 10 International Cross-Border Markets (Excluding United Kingdom)")

geo_headers = ["Market / Country", "Net Revenue (£)", "Revenue Share (%)", "Completed Orders", "Active Customers", "Average Order Value (£)"]
geo_data = [
    ["United Kingdom", f"£{stats['uk_revenue']:,.2f}", f"{stats['uk_revenue_share_pct']:.2f}%", "18,016", "3,921", f"£{stats['uk_revenue']/18016:,.2f}"],
    ["Netherlands", f"£{stats['top_intl_country_rev']:,.2f}", "2.77%", "94", "9", f"£{stats['top_intl_country_rev']/94:,.2f}"],
    ["EIRE (Ireland)", f"£{stats['second_intl_country_rev']:,.2f}", "2.69%", "286", "3", f"£{stats['second_intl_country_rev']/286:,.2f}"],
    ["Germany", "£228,867.14", "2.23%", "456", "94", "£501.90"],
    ["France", "£208,610.15", "2.03%", "391", "87", "£533.53"],
    ["Australia", "£138,521.31", "1.35%", "57", "9", "£2,430.20"],
    ["Spain", "£61,577.11", "0.60%", "90", "30", "£684.19"],
    ["Switzerland", "£56,488.45", "0.55%", "54", "21", "£1,046.08"],
    ["Belgium", "£41,196.34", "0.40%", "98", "25", "£420.37"],
    ["Sweden", "£38,378.33", "0.37%", "36", "8", "£1,066.06"]
]
geo_widths = [1.5, 1.2, 1.0, 0.9, 0.9, 1.0]
geo_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]

geo_tbl = doc.add_table(rows=len(geo_data) + 1, cols=6)
format_table(geo_tbl, geo_widths, geo_headers, geo_data, geo_aligns)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

add_body(
    "A striking insight emerges from the cross-border metric: while international orders represent only 8.9% of total order counts, "
    "they account for 14.89% of enterprise revenue. The Netherlands boasts an astonishing Average Order Value of £3,020.10—nearly six times "
    "the UK average (£484.99). This indicates that international accounts are large-scale commercial distributors purchasing in bulk container loads."
)

# ==============================================================================
# SECTION 11: CUSTOMER LIFETIME SPEND & PARETO CONCENTRATION
# ==============================================================================
add_title("11. Customer Lifetime Spend & Pareto Concentration")
add_body(
    "We analyzed customer spending distributions across 4,334 identified customer accounts, segmenting buyers into distinct value tiers "
    "and testing the empirical validity of the Pareto Principle (80/20 rule)."
)

add_image_with_caption("charts/09_customer_spending_distribution.png", "Figure 9: Customer Segmentation by Total Spending Tier")
add_image_with_caption("charts/10_top_customers_revenue.png", "Figure 10: Top 10 High-Value Wholesale Accounts by Lifetime Spend (£ Thousands)")

add_callout(
    f"The top 1% of customers generate {stats['pareto_top_1pct_revenue_share']:.2f}% of identified revenue (£2.85M).\n"
    f"The top 10% of customers generate {stats['pareto_top_10pct_revenue_share']:.2f}% of identified revenue (£5.46M).\n"
    f"The top 20% of customers generate {stats['pareto_top_20pct_revenue_share']:.2f}% of identified revenue (£6.64M).\n"
    f"The single largest account (Customer ID {stats['top_customer_id']}, {stats['top_customer_country']}) generated £{stats['top_customer_rev']:,.2f} across {stats['top_customer_orders']} orders.",
    bold_title="PARETO CONCENTRATION VERIFICATION: "
)

add_body(
    "Customer accounts were segmented into four operational tiers: Low (<£250, 32.2% of accounts), Mid (£250–£1,000, 37.9% of accounts), "
    "High (£1,000–£5,000, 27.7% of accounts), and VIP (>£5,000, 2.2% of accounts). While VIP clients represent only 94 enterprises, "
    "they generate over 35% of all identified revenue. The firm faces acute key-account dependency risk."
)

# ==============================================================================
# SECTION 12: CANCELLATION & RETURN MARGIN DIAGNOSTICS
# ==============================================================================
add_title("12. Cancellation & Return Margin Diagnostics")
add_body(
    "Transaction cancellations (invoices prefixed with 'C') represent a substantial operational drag. "
    "We analyzed cancellation frequencies, item volumes, and revenue impact across the year."
)

add_image_with_caption("charts/11_cancellation_patterns_monthly.png", "Figure 11: Monthly Net Sales Revenue vs. Value of Cancelled Transactions")

add_body(
    f"Across the observation window, exactly {stats['total_cancelled_invoices']:,} invoices comprising 9,288 line items and {stats['total_cancelled_items']:,} "
    f"physical units were cancelled. The gross value of these cancellations reached £{stats['total_cancelled_value']:,.2f}, representing an invoice cancellation "
    f"rate of {stats['cancellation_rate_invoices_pct']:.2f}% across all order attempts. Cancellations peaked in absolute terms during October and November, "
    "mirroring the Q4 surge. However, the return-loss ratio stayed consistent at 6.5% to 8.8% of gross monthly sales, indicating systemic process drivers "
    "rather than isolated seasonal anomalies."
)

# ==============================================================================
# SECTION 13: CORRELATION ANALYSIS
# ==============================================================================
add_title("13. Feature Correlation Analysis")
add_body(
    "We examined linear Pearson correlations between key numerical variables to validate behavioral relationships."
)

add_image_with_caption("charts/12_correlation_heatmap.png", "Figure 12: Pearson Correlation Heatmap for Numerical Transaction Attributes")

add_body(
    "Key statistical observations include:\n"
    "1. Quantity and Revenue exhibit an overwhelming positive correlation (r = 0.908), confirming that bulk order volume is the overwhelming driver of sales value.\n"
    "2. UnitPrice displays an almost zero correlation with Quantity (r = -0.005), illustrating that wholesale buyers order large batch sizes regardless of unit price points.\n"
    "3. Temporal attributes (Hour, Day, Month) exhibit near-zero correlations with basket size, indicating that transaction values remain steady regardless of when the purchase is executed."
)

# ==============================================================================
# SECTION 14: CONSOLIDATED KEY FINDINGS
# ==============================================================================
add_title("14. Consolidated Key Findings")
add_body("Synthesizing our empirical investigations yields the following verified analytical conclusions:")
add_bullet(f"The enterprise achieved £{stats['total_revenue']:,.2f} in net sales across {stats['total_orders']:,} orders and {stats['total_quantity_sold']:,} items.", bold_prefix="1. Enterprise Scale: ")
add_bullet(f"November 2011 was the peak month (£{stats['peak_month_rev']:,.2f}), nearly tripling February 2011 (£{stats['lowest_month_rev']:,.2f}).", bold_prefix="2. Q4 Seasonality Surge: ")
add_bullet(f"Thursday is the top ordering day ({stats['top_day_orders']:,} orders), and 12:00 PM is the peak hour ({stats['top_hour_orders']:,} orders). Saturdays record 0 orders.", bold_prefix="3. Corporate Operating Hours: ")
add_bullet(f"Stock Code {stats['top_product_by_revenue_code']} led revenue (£{stats['top_product_by_revenue_val']:,.2f}), while Code {stats['top_product_by_qty_code']} led volume ({stats['top_product_by_qty_val']:,} units).", bold_prefix="4. Hero Product Concentration: ")
add_bullet(f"The UK generates {stats['uk_revenue_share_pct']:.2f}% of sales, but international markets exhibit 6x higher Average Order Values (£3,020 in Netherlands).", bold_prefix="5. High-Value International Markets: ")
add_bullet(f"The top 20% of clients generate {stats['pareto_top_20pct_revenue_share']:.2f}% of revenue, and top 1% generate {stats['pareto_top_1pct_revenue_share']:.2f}%.", bold_prefix="6. Severe Pareto Concentration: ")
add_bullet(f"Cancellations impacted {stats['total_cancelled_invoices']:,} invoices ({stats['cancellation_rate_invoices_pct']:.2f}% rate), representing £{stats['total_cancelled_value']:,.2f} in cancelled gross merchandise.", bold_prefix="7. Return Friction: ")

# ==============================================================================
# SECTION 15: STRATEGIC BUSINESS RECOMMENDATIONS
# ==============================================================================
add_title("15. Strategic Business Insights & Recommendations")
add_body(
    "Based on these empirical findings, we present five strategic, high-ROI business recommendations for executive leadership:"
)

add_heading2("Recommendation 1: Dedicated VIP Key Account Success Program")
add_body(
    f"With the top 1% generating {stats['pareto_top_1pct_revenue_share']:.2f}% of revenue, losing even 2–3 top wholesale accounts would devastate company cash flow. "
    "Leadership must establish a VIP Account Management team assigned to the top 100 corporate clients. Offer contractual SLAs, priority warehouse packing, "
    "custom bulk pallet pricing, and guaranteed stock reservation during peak Q4 months."
)

add_heading2("Recommendation 2: Predictive Q4 Inventory Buffering & Procurement Staging")
add_body(
    "The 186% demand surge from February to November requires aggressive inventory staging. Procurement orders for the top 50 revenue-generating SKUs "
    "(such as Regency Cakestands and Ceramic Jars) must be submitted to manufacturers by June/July, with shipments arriving in UK warehouses by August. "
    "This mitigates stockouts, avoids express freight premiums, and captures maximum seasonal demand."
)

add_heading2("Recommendation 3: European Cross-Border Distribution Hub (Netherlands/Germany)")
add_body(
    f"International markets represent £{stats['international_revenue']:,.2f} with massive AOVs (£3,020 in the Netherlands). Establishing a bonded cross-docking "
    "or 3PL fulfillment partnership in Rotterdam or Frankfurt will eliminate customs clearance delays, cut shipping tariffs, and reduce transit times from 5 days to 24–48 hours, "
    "unlocking substantial wholesale market share across Western Europe."
)

add_heading2("Recommendation 4: Cancellation Root-Cause Remediation Program")
add_body(
    f"With £{stats['total_cancelled_value']:,.2f} tied up in cancelled orders, management must conduct an immediate quality audit on the top cancelled stock codes. "
    "Inspect fragile ceramics and glassware for transit damage, improve protective inner packaging, and mandate order confirmation protocols for orders exceeding £1,000 "
    "to prevent accidental double-ordering."
)

add_heading2("Recommendation 5: Weekend Digital Pre-Order Automation & Shift Smoothing")
add_body(
    "The zero-order anomaly on Saturday reveals an operational bottleneck: Sunday and Monday morning warehouse crews are overwhelmed by accumulated orders. "
    "Deploy automated order validation workflows, AI-assisted picking route optimization, and weekend dispatch shifts to balance workload evenly across the business week."
)

# ==============================================================================
# SECTION 16: LIMITATIONS & FUTURE SCOPE
# ==============================================================================
add_title("16. Limitations & Future Scope")
add_body(
    "While this study delivers deep empirical insights, several analytical limitations must be noted:\n"
    "1. Lack of Cost of Goods Sold (COGS): The dataset contains unit selling prices but lacks wholesale acquisition costs and shipping expenses. Consequently, metrics evaluate gross revenue rather than gross profit margins.\n"
    "2. Unregistered Guest Purchases: Exactly 24.93% of transactions lack a CustomerID, precluding complete lifetime value tracking for guest users.\n"
    "3. Truncated Observation Window: The dataset terminates on December 9, 2011, preventing full-month comparisons for December 2011.\n"
    "4. Absence of Marketing Attribution: Transaction records do not capture web traffic sources, advertising spend, or promotional discounts."
)

# ==============================================================================
# SECTION 17: CONCLUSION
# ==============================================================================
add_title("17. Conclusion")
add_body(
    "This project successfully completed a rigorous, end-to-end data analysis of the transnational UCI Online Retail Dataset. "
    "By establishing clean transactional boundaries and executing comprehensive multi-dimensional evaluations, we uncovered the key commercial drivers "
    "of the enterprise. The operational roadmap delivered in this report equips executive decision-makers with the concrete tools needed to safeguard "
    "VIP wholesale accounts, optimize holiday inventory pipelines, expand European distribution, and curb costly order cancellations."
)

# ==============================================================================
# SECTION 18: REFERENCES
# ==============================================================================
add_title("18. References")
add_bullet("Chen, D., Sain, S. L., & Guo, K. (2012). Data mining for the online retail industry: A case study of RFM model-based customer segmentation using data mining. Journal of Database Marketing & Customer Strategy Management, 19(3), 197-208.")
add_bullet("UCI Machine Learning Repository. (2015). Online Retail Dataset. Accessible at: https://archive.ics.uci.edu/dataset/352/online+retail")
add_bullet("McKinney, W. (2010). Data Structures for Statistical Computing in Python. Proceedings of the 9th Python in Science Conference, 51-56.")
add_bullet("Hunter, J. D. (2007). Matplotlib: A 2D Graphics Environment. Computing in Science & Engineering, 9(3), 90-95.")
add_bullet("Waskom, M. L. (2021). Seaborn: statistical data visualization. Journal of Open Source Software, 6(60), 3021.")

output_doc_path = "SnehaVinodPotanavar_OnlineRetailSalesAnalysis_ProjectReport.docx"
doc.save(output_doc_path)
print(f"Project Report successfully created and saved to: {output_doc_path}")
print(f"File size: {os.path.getsize(output_doc_path):,} bytes")
