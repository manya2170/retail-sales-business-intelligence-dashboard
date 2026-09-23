# Retail Sales Performance & Business Intelligence Dashboard

**IBM SkillsBuild Data Analytics with AI Internship — Final Project**

---

## Project Description

A decision-oriented, multi-page Business Intelligence dashboard built with Python and Plotly Dash.  
The dashboard answers five business questions structured around the **KPI → TREND → DRIVER → RISK → ACTION** framework:

1. What is happening with overall business performance?
2. How is revenue and profit changing over time?
3. What categories, products, regions and customer segments drive performance?
4. What are the major profitability risks?
5. What opportunities and management actions can be identified?

---

## Dataset

| Property | Value |
|---|---|
| File | `Sample - Superstore.csv` |
| Rows | 9,994 transactions |
| Date range | January 2014 – December 2017 |
| Segments | Consumer, Corporate, Home Office |
| Categories | Furniture, Office Supplies, Technology |
| Regions | East, West, Central, South |

All metrics are computed directly from the CSV. No values are hardcoded or estimated.

---

## Prerequisites

- Python 3.9 or higher
- pip

---

## Setup

```bash
# 1. Clone or unzip the project folder
cd Retail-Sales-BI-Dashboard

# 2. (Recommended) Create a virtual environment
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS / Linux

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Run the Dashboard

```bash
python app.py
```

Open your browser at: **http://127.0.0.1:8050**

The dashboard has three pages:
- **Executive Overview** — KPI scorecard, trends, category performance, findings
- **Sales & Products** — Category, region, and product analysis with Year filter
- **Customer & Risk** — Segment analysis, loss-making products, risk findings with Year filter

---

## Run the Jupyter Notebook

```bash
jupyter notebook Retail_Sales_BI_Dashboard.ipynb
```

Run all cells from top to bottom. The notebook reproduces the full analysis with commentary.

---

## Generate the Project Report (Word .docx)

```bash
python generate_report.py
```

This creates `Retail_Sales_BI_Project_Report.docx` in the project folder.  
All values in the report are calculated live from the dataset — run this after setup.

---

## File Structure

```
Retail-Sales-BI-Dashboard/
├── Sample - Superstore.csv              Source dataset (do not modify)
├── app.py                               Dash dashboard entry point
├── data_loader.py                       Data loading and metric computation
├── generate_report.py                   Generates the Word project report
├── requirements.txt                     Python dependencies
├── README.md                            This file
├── Retail_Sales_BI_Dashboard.ipynb      Jupyter Notebook analysis
├── Retail_Sales_BI_Project_Report.docx  Word project report (generated)
└── pages/
    ├── page1_executive.py               Page 1: Executive Overview
    ├── page2_sales_product.py           Page 2: Sales & Product Analysis
    └── page3_customer_risk.py           Page 3: Customer & Risk Analysis
```

---

## Mandatory Submission Files

| File | Description |
|---|---|
| `Retail_Sales_BI_Dashboard.ipynb` | Jupyter Notebook with full analysis |
| `requirements.txt` | Python dependencies |
| `README.md` | Setup and run instructions |
| `Retail_Sales_BI_Project_Report.docx` | Microsoft Word project report |
| `app.py` | Dash dashboard entry point |
| `data_loader.py` | Data loader and metric calculator |
| `Sample - Superstore.csv` | Source dataset |

---

## Notes

- All findings in the dashboard and report are derived from the dataset — no LLM or API calls are used.
- Associations between variables (e.g. discount and profit) are stated as associations, not causal relationships.
- The dashboard runs entirely offline after `pip install -r requirements.txt`.
