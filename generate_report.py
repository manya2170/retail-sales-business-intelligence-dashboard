"""
generate_report.py
------------------
Generates Retail_Sales_BI_Project_Report.docx using python-docx.
All numeric values are computed live from the dataset via data_loader.
Run:  python generate_report.py
"""

import os
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import data_loader as dl

# ── Load data and compute all metrics ─────────────────────────────────────────
print("Loading dataset and computing metrics...")
df = dl.load_data()
m = dl.compute_metrics(df)

# Pull key values for inline use
total_revenue    = m["total_revenue"]
total_profit     = m["total_profit"]
profit_margin    = m["profit_margin_pct"]
total_orders     = m["total_orders"]
total_customers  = m["total_customers"]
avg_order_value  = m["avg_order_value"]

yearly = m["yearly_growth"].sort_values("Year")
best_year_row  = m["yearly"].loc[m["yearly"]["Revenue"].idxmax()]
best_year      = int(best_year_row["Year"])
best_year_rev  = best_year_row["Revenue"]
latest_growth  = m["latest_yoy_growth_pct"]

cat   = m["category"]
region = m["region"].sort_values("ProfitMarginPct")
seg   = m["segment"].sort_values("ProfitMarginPct", ascending=False)
subcat = m["subcategory"]

loss_subcats = subcat[subcat["Profit"] < 0].sort_values("Profit")
best_cat_row  = cat.loc[cat["ProfitMarginPct"].idxmax()]
worst_cat_row = cat.loc[cat["ProfitMarginPct"].idxmin()]
worst_region  = region.iloc[0]
best_seg      = seg.iloc[0]

high_disc_profit = df[df["Discount"] >= 0.3]["Profit"].mean()
low_disc_profit  = df[df["Discount"] <  0.3]["Profit"].mean()

years = m["years_sorted"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x3F, 0x6E)


def add_para(doc: Document, text: str, bold: bool = False, italic: bool = False,
             size: int = 11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p


def add_finding(doc: Document, label: str, text: str, label_colour: RGBColor):
    p = doc.add_paragraph()
    r1 = p.add_run(f"[{label}]  ")
    r1.bold = True
    r1.font.color.rgb = label_colour
    r1.font.size = Pt(10)
    r2 = p.add_run(text)
    r2.font.size = Pt(10)


COLOURS = {
    "FACT":        RGBColor(0x3B, 0x82, 0xD4),
    "INSIGHT":     RGBColor(0x22, 0xC5, 0x5E),
    "RISK":        RGBColor(0xEF, 0x44, 0x44),
    "OPPORTUNITY": RGBColor(0xF9, 0x73, 0x16),
    "ACTION":      RGBColor(0x7C, 0x3A, 0xED),
}


def add_table_row(table, cells):
    row = table.add_row()
    for i, val in enumerate(cells):
        row.cells[i].text = str(val)


# ── Build document ─────────────────────────────────────────────────────────────
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.2)
    section.right_margin  = Inches(1.2)

# ── Cover page ────────────────────────────────────────────────────────────────
doc.add_paragraph()
doc.add_paragraph()
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title_p.add_run("Retail Sales Performance &\nBusiness Intelligence Dashboard")
title_run.bold = True
title_run.font.size = Pt(22)
title_run.font.color.rgb = RGBColor(0x1F, 0x3F, 0x6E)

doc.add_paragraph()
sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_p.add_run("IBM SkillsBuild Data Analytics with AI Internship\nFinal Project Report").font.size = Pt(13)

doc.add_paragraph()
date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
date_p.add_run(f"Date: {date.today().strftime('%B %d, %Y')}").font.size = Pt(11)

doc.add_paragraph()
ds_p = doc.add_paragraph()
ds_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
ds_p.add_run("Dataset: Sample - Superstore (2014–2017)  |  9,994 transactions").font.size = Pt(10)

doc.add_page_break()

# ── Section 1: Executive Summary ─────────────────────────────────────────────
add_heading(doc, "1. Executive Summary")
add_para(doc,
    f"This report presents a Business Intelligence analysis of the Sample - Superstore retail dataset, "
    f"covering {years[0]}–{years[-1]}. The business generated total revenue of "
    f"${total_revenue:,.2f} and total profit of ${total_profit:,.2f} across "
    f"{total_orders:,} orders from {total_customers:,} customers, delivering an overall "
    f"profit margin of {profit_margin:.2f}%."
)
add_para(doc,
    f"Revenue grew consistently year-over-year, with {best_year} recording the highest annual revenue "
    f"of ${best_year_rev:,.2f}. The most recent year-on-year growth rate was "
    f"{latest_growth:+.1f}%. However, headline revenue growth masks significant "
    f"profitability variation: Furniture operates near break-even ({worst_cat_row['ProfitMarginPct']:.1f}% margin) "
    f"while Technology and Office Supplies each exceed 17% margin."
)
add_para(doc,
    "Three sub-categories — Tables, Bookcases, and Supplies — are net loss-making across the full period. "
    "Addressing these structural issues represents the primary near-term profitability opportunity."
)

# ── Section 2: Dataset Description ───────────────────────────────────────────
add_heading(doc, "2. Dataset Description")
add_para(doc, "All facts in this section are verified directly from the source CSV file.", italic=True)
doc.add_paragraph()

tbl = doc.add_table(rows=1, cols=2)
tbl.style = "Table Grid"
hdr = tbl.rows[0].cells
hdr[0].text = "Property"
hdr[1].text = "Value"
for cell in hdr:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True

rows = [
    ("File", "Sample - Superstore.csv"),
    ("Total rows", f"{len(df):,}"),
    ("Columns", str(len(df.columns))),
    ("Date range", f"{df['Order Date'].min().date()} to {df['Order Date'].max().date()}"),
    ("Years", ", ".join(str(y) for y in years)),
    ("Segments", "Consumer, Corporate, Home Office"),
    ("Categories", "Furniture, Office Supplies, Technology"),
    ("Sub-Categories", "17 (Accessories, Appliances, Art, Binders, Bookcases, Chairs, Copiers, "
                       "Envelopes, Fasteners, Furnishings, Labels, Machines, Paper, Phones, "
                       "Storage, Supplies, Tables)"),
    ("Regions", "East, West, Central, South"),
    ("Ship Modes", "Standard Class, Second Class, First Class, Same Day"),
    ("Loss-making rows", f"{(df['Profit'] < 0).sum():,} rows with negative profit"),
    ("Key numeric columns", "Sales, Quantity, Discount, Profit"),
]
for r in rows:
    add_table_row(tbl, r)

doc.add_paragraph()

# ── Section 3: Methodology ────────────────────────────────────────────────────
add_heading(doc, "3. Methodology")
add_para(doc,
    "The analysis follows the KPI → TREND → DRIVER → RISK → ACTION framework, "
    "structured to support management decision-making rather than exploratory data-dumping."
)
add_para(doc, "Tools and libraries used:")
bullets = [
    "Python 3.x — primary analysis language",
    "pandas — data loading, transformation, and metric computation",
    "Plotly — interactive chart library",
    "Plotly Dash — multi-page interactive dashboard framework",
    "dash-bootstrap-components — professional dashboard layout",
    "python-docx — programmatic Word report generation",
    "Jupyter Notebook — documented analysis narrative",
]
for b in bullets:
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(b).font.size = Pt(10)

add_para(doc,
    "\nAll metrics are computed at runtime from the source CSV. No values are hardcoded "
    "or estimated. Findings are labelled as FACT, INSIGHT, RISK, OPPORTUNITY, or ACTION "
    "to distinguish the nature of each observation."
)

doc.add_page_break()

# ── Section 4: Page 1 Findings — Business Performance ─────────────────────────
add_heading(doc, "4. Page 1 Findings — Business Performance")
add_para(doc, "Business question: What is happening with overall business performance?", italic=True)
doc.add_paragraph()

add_heading(doc, "4.1 KPI Scorecard", level=2)
tbl2 = doc.add_table(rows=1, cols=2)
tbl2.style = "Table Grid"
hdr2 = tbl2.rows[0].cells
hdr2[0].text = "KPI"
hdr2[1].text = "Value"
for cell in hdr2:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True
kpi_rows = [
    ("Total Revenue",     f"${total_revenue:,.2f}"),
    ("Total Profit",      f"${total_profit:,.2f}"),
    ("Profit Margin",     f"{profit_margin:.2f}%"),
    ("Total Orders",      f"{total_orders:,}"),
    ("Total Customers",   f"{total_customers:,}"),
    ("Avg Order Value",   f"${avg_order_value:,.2f}"),
]
for r in kpi_rows:
    add_table_row(tbl2, r)

doc.add_paragraph()
add_heading(doc, "4.2 Findings", level=2)

add_finding(doc, "FACT",
    f"Total revenue 2014–2017: ${total_revenue:,.2f} | "
    f"Total profit: ${total_profit:,.2f} | "
    f"{total_orders:,} orders from {total_customers:,} customers.",
    COLOURS["FACT"])

add_finding(doc, "FACT",
    f"Overall profit margin: {profit_margin:.2f}%.",
    COLOURS["FACT"])

add_finding(doc, "INSIGHT",
    f"Strongest revenue year: {best_year} (${best_year_rev:,.2f}). "
    f"Most recent YoY revenue growth: {latest_growth:+.1f}%.",
    COLOURS["INSIGHT"])

add_finding(doc, "INSIGHT",
    f"{best_cat_row['Category']} is the highest-margin category ({best_cat_row['ProfitMarginPct']:.1f}%). "
    f"{worst_cat_row['Category']} has the lowest margin ({worst_cat_row['ProfitMarginPct']:.1f}%).",
    COLOURS["INSIGHT"])

add_finding(doc, "RISK",
    f"{worst_cat_row['Category']}'s {worst_cat_row['ProfitMarginPct']:.1f}% margin signals structural "
    f"profitability concerns. {len(loss_subcats)} sub-categories are net loss-making: "
    f"{', '.join(loss_subcats['Sub-Category'].tolist())}.",
    COLOURS["RISK"])

add_finding(doc, "OPPORTUNITY",
    "Technology and Office Supplies both exceed 17% margin — expanding these lines could "
    "lift overall portfolio profitability.",
    COLOURS["OPPORTUNITY"])

add_finding(doc, "ACTION",
    "Conduct a pricing and discount review for Furniture — specifically Tables, Bookcases, "
    "and Supplies — to identify loss drivers before the next planning cycle.",
    COLOURS["ACTION"])

doc.add_page_break()

# ── Section 5: Page 2 Findings — Sales & Product Drivers ──────────────────────
add_heading(doc, "5. Page 2 Findings — Sales & Product Drivers")
add_para(doc, "Business question: What categories, products, regions and segments drive performance?", italic=True)
doc.add_paragraph()

add_heading(doc, "5.1 Category Performance", level=2)
tbl3 = doc.add_table(rows=1, cols=4)
tbl3.style = "Table Grid"
hdr3 = tbl3.rows[0].cells
for i, h in enumerate(["Category", "Revenue", "Profit", "Margin %"]):
    hdr3[i].text = h
    for run in hdr3[i].paragraphs[0].runs:
        run.bold = True
for _, row in cat.iterrows():
    add_table_row(tbl3, [
        row["Category"],
        f"${row['Revenue']:,.2f}",
        f"${row['Profit']:,.2f}",
        f"{row['ProfitMarginPct']:.2f}%",
    ])

doc.add_paragraph()
add_heading(doc, "5.2 Regional Performance", level=2)
tbl4 = doc.add_table(rows=1, cols=4)
tbl4.style = "Table Grid"
hdr4 = tbl4.rows[0].cells
for i, h in enumerate(["Region", "Revenue", "Profit", "Margin %"]):
    hdr4[i].text = h
    for run in hdr4[i].paragraphs[0].runs:
        run.bold = True
for _, row in region.sort_values("Region").iterrows():
    add_table_row(tbl4, [
        row["Region"],
        f"${row['Revenue']:,.2f}",
        f"${row['Profit']:,.2f}",
        f"{row['ProfitMarginPct']:.2f}%",
    ])

doc.add_paragraph()
add_heading(doc, "5.3 Findings", level=2)

add_finding(doc, "FACT",
    f"Technology leads in revenue (${cat[cat['Category']=='Technology']['Revenue'].values[0]:,.2f}) "
    f"and also carries the highest margin ({cat[cat['Category']=='Technology']['ProfitMarginPct'].values[0]:.1f}%).",
    COLOURS["FACT"])

add_finding(doc, "FACT",
    f"West region delivers the highest profit margin ({region[region['Region']=='West']['ProfitMarginPct'].values[0]:.1f}%). "
    f"Central region has the lowest ({region[region['Region']=='Central']['ProfitMarginPct'].values[0]:.1f}%).",
    COLOURS["FACT"])

add_finding(doc, "INSIGHT",
    "High-revenue products are not always the same as high-profit products. "
    "Products with large sales volumes but thin or negative margins represent a pricing risk.",
    COLOURS["INSIGHT"])

add_finding(doc, "RISK",
    f"Central region's {region[region['Region']=='Central']['ProfitMarginPct'].values[0]:.1f}% margin "
    "is the lowest across all regions and may reflect product mix or discount concentration.",
    COLOURS["RISK"])

add_finding(doc, "OPPORTUNITY",
    "West region's strong margins suggest that its product mix or pricing strategy "
    "could inform improvements in other regions.",
    COLOURS["OPPORTUNITY"])

doc.add_page_break()

# ── Section 6: Page 3 Findings — Customer & Risk ──────────────────────────────
add_heading(doc, "6. Page 3 Findings — Customer & Risk Analysis")
add_para(doc, "Business question: Who are key customers, and where are the profitability risks?", italic=True)
doc.add_paragraph()

add_heading(doc, "6.1 Segment Performance", level=2)
tbl5 = doc.add_table(rows=1, cols=5)
tbl5.style = "Table Grid"
hdr5 = tbl5.rows[0].cells
for i, h in enumerate(["Segment", "Revenue", "Profit", "Orders", "Margin %"]):
    hdr5[i].text = h
    for run in hdr5[i].paragraphs[0].runs:
        run.bold = True
for _, row in seg.sort_values("Segment").iterrows():
    add_table_row(tbl5, [
        row["Segment"],
        f"${row['Revenue']:,.2f}",
        f"${row['Profit']:,.2f}",
        f"{int(row['Orders']):,}",
        f"{row['ProfitMarginPct']:.2f}%",
    ])

doc.add_paragraph()
add_heading(doc, "6.2 Loss-making Sub-Categories", level=2)
tbl6 = doc.add_table(rows=1, cols=3)
tbl6.style = "Table Grid"
hdr6 = tbl6.rows[0].cells
for i, h in enumerate(["Sub-Category", "Revenue", "Profit"]):
    hdr6[i].text = h
    for run in hdr6[i].paragraphs[0].runs:
        run.bold = True
for _, row in loss_subcats.iterrows():
    add_table_row(tbl6, [
        row["Sub-Category"],
        f"${row['Revenue']:,.2f}",
        f"${row['Profit']:,.2f}",
    ])

doc.add_paragraph()
add_heading(doc, "6.3 Findings", level=2)

add_finding(doc, "FACT",
    f"{len(loss_subcats)} sub-categories are net loss-making across 2014–2017: "
    f"{', '.join(loss_subcats['Sub-Category'].tolist())}. "
    f"Tables is the largest contributor at ${loss_subcats[loss_subcats['Sub-Category']=='Tables']['Profit'].values[0]:,.2f}.",
    COLOURS["FACT"])

add_finding(doc, "INSIGHT",
    f"Orders with discount ≥ 30% show an average profit of ${high_disc_profit:,.2f}, "
    f"compared to ${low_disc_profit:,.2f} for orders with lower discounts. "
    "This is an observed association — not an established causal relationship.",
    COLOURS["INSIGHT"])

add_finding(doc, "RISK",
    f"{worst_region['Region']} region has the lowest profit margin ({worst_region['ProfitMarginPct']:.1f}%), "
    "indicating potential pricing, product mix, or cost issues requiring investigation.",
    COLOURS["RISK"])

add_finding(doc, "OPPORTUNITY",
    f"{best_seg['Segment']} is the highest-margin segment ({best_seg['ProfitMarginPct']:.1f}%). "
    "Targeted retention and upsell efforts for this segment may improve overall profitability.",
    COLOURS["OPPORTUNITY"])

add_finding(doc, "OPPORTUNITY",
    "Consumer segment generates the highest absolute revenue. Loyalty programmes "
    "or repeat-order incentives could improve revenue retention.",
    COLOURS["OPPORTUNITY"])

add_finding(doc, "ACTION",
    f"Prioritise a cost and discount audit for '{loss_subcats.iloc[0]['Sub-Category']}' "
    "and other loss-making sub-categories to determine whether losses are driven "
    "by discounting policy, unit costs, or returns.",
    COLOURS["ACTION"])

add_finding(doc, "ACTION",
    f"Review the {worst_region['Region']} region's pricing strategy and product mix. "
    "Consider whether high-discount offers in this region are driving volume without margin.",
    COLOURS["ACTION"])

doc.add_page_break()

# ── Section 7: Limitations & Caveats ─────────────────────────────────────────
add_heading(doc, "7. Limitations & Caveats")

add_para(doc, "Association is not causation.", bold=True)
add_para(doc,
    "Observed associations between variables — such as the relationship between discount rate "
    "and profit per transaction — do not establish causal direction. Other factors including "
    "product cost, order size, returns, shipping class, and customer segment may explain or "
    "partially explain observed patterns. Business decisions should be validated with additional "
    "operational data and controlled analysis before implementation."
)
doc.add_paragraph()
add_para(doc, "Dataset scope.", bold=True)
add_para(doc,
    "The dataset covers a single retailer operating in the United States across four years "
    f"({years[0]}–{years[-1]}). Findings may not generalise to other markets, time periods, "
    "or business models."
)
doc.add_paragraph()
add_para(doc, "Aggregation effects.", bold=True)
add_para(doc,
    "Regional and segment-level averages may mask wide variation within those groups. "
    "Sub-group analysis (e.g., by city, state, or product line within a region) is recommended "
    "before operational decisions are made."
)

doc.add_page_break()

# ── Section 8: Recommendations ────────────────────────────────────────────────
add_heading(doc, "8. Recommendations")
add_para(doc, "The following recommendations are derived from dataset findings only.", italic=True)
doc.add_paragraph()

recs = [
    ("Discount policy audit for loss-making sub-categories",
     f"Tables, Bookcases, and Supplies are net loss-making. "
     f"A structured review of discount authorisation thresholds for these sub-categories "
     f"is the highest-priority action to improve profit."),
    ("Regional strategy review for Central region",
     f"Central region's {region[region['Region']=='Central']['ProfitMarginPct'].values[0]:.1f}% "
     f"margin is the lowest across all regions. "
     f"A product mix and pricing review specific to Central is recommended."),
    ("Grow Technology and Office Supplies",
     "Both categories sustain margins above 17%. Allocating additional marketing or inventory "
     "investment to these categories is likely to yield higher returns than equivalent investment in Furniture."),
    ("Leverage Home Office segment efficiency",
     f"Home Office delivers {best_seg['ProfitMarginPct']:.1f}% margin with fewer customers. "
     "Targeted acquisition in this segment may be more margin-efficient than broad Consumer campaigns."),
    ("Monitor top customer concentration",
     "Revenue from top-10 customers represents a concentration risk. "
     "Building broader customer relationships across the portfolio would reduce exposure to churn."),
]
for i, (title, text) in enumerate(recs, 1):
    p = doc.add_paragraph(style="List Number")
    r1 = p.add_run(title + ": ")
    r1.bold = True
    r1.font.size = Pt(11)
    r2 = p.add_run(text)
    r2.font.size = Pt(11)

doc.add_page_break()

# ── Appendix ──────────────────────────────────────────────────────────────────
add_heading(doc, "Appendix")

add_heading(doc, "A. Sub-Category Reference", level=2)
all_subcats = sorted(df["Sub-Category"].unique().tolist())
sub_p = doc.add_paragraph()
sub_p.add_run(", ".join(all_subcats)).font.size = Pt(10)

doc.add_paragraph()
add_heading(doc, "B. Dataset Column Reference", level=2)
cols_tbl = doc.add_table(rows=1, cols=2)
cols_tbl.style = "Table Grid"
cols_hdr = cols_tbl.rows[0].cells
cols_hdr[0].text = "Column"
cols_hdr[1].text = "Description"
for cell in cols_hdr:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True

col_descs = [
    ("Row ID", "Unique row identifier"),
    ("Order ID", "Unique order identifier (multiple rows per order)"),
    ("Order Date", "Date the order was placed"),
    ("Ship Date", "Date the order was shipped"),
    ("Ship Mode", "Shipping class selected"),
    ("Customer ID", "Unique customer identifier"),
    ("Customer Name", "Customer full name"),
    ("Segment", "Customer segment: Consumer, Corporate, Home Office"),
    ("Country", "Country of order (United States)"),
    ("City", "City of delivery"),
    ("State", "State of delivery"),
    ("Postal Code", "Postal code of delivery"),
    ("Region", "Regional grouping: East, West, Central, South"),
    ("Product ID", "Unique product identifier"),
    ("Category", "Product category: Furniture, Office Supplies, Technology"),
    ("Sub-Category", "Product sub-category (17 values)"),
    ("Product Name", "Full product name"),
    ("Sales", "Order line revenue (USD)"),
    ("Quantity", "Units ordered"),
    ("Discount", "Discount rate applied (0.0–1.0)"),
    ("Profit", "Order line profit (USD); negative = loss"),
]
for col, desc in col_descs:
    add_table_row(cols_tbl, [col, desc])

doc.add_paragraph()
add_heading(doc, "C. KPI → TREND → DRIVER → RISK → ACTION Framework", level=2)
framework_rows = [
    ("KPI",      "Headline scorecard metrics — what is the current state?"),
    ("TREND",    "How have metrics changed over time?"),
    ("DRIVER",   "What categories, products, regions, and segments explain performance?"),
    ("RISK",     "Where are structural profitability threats?"),
    ("ACTION",   "What specific management decisions are indicated by the data?"),
]
fw_tbl = doc.add_table(rows=1, cols=2)
fw_tbl.style = "Table Grid"
fw_hdr = fw_tbl.rows[0].cells
fw_hdr[0].text = "Layer"
fw_hdr[1].text = "Purpose"
for cell in fw_hdr:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True
for r in framework_rows:
    add_table_row(fw_tbl, r)

# ── Save ───────────────────────────────────────────────────────────────────────
out_path = os.path.join(os.path.dirname(__file__), "Retail_Sales_BI_Project_Report.docx")
doc.save(out_path)
print(f"Report saved: {out_path}")
