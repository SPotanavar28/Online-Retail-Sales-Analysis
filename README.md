# Online Retail Sales & Customer Analysis

**Author:** Sneha Vinod Potanavar  
**Role:** Data Analyst Intern  
**Project Title:** Online Retail Sales & Customer Analysis  
**Repository:** [Online-Retail-Sales-Analysis](https://github.com/SnehaVinodPotanavar/Online-Retail-Sales-Analysis)  
**Dataset Source:** [UCI Machine Learning Repository — Online Retail Dataset](https://archive.ics.uci.edu/dataset/352/online+retail)  

---

## Table of Contents

- [Project Overview](#project-overview)
- [Business Problem Statement](#business-problem-statement)
- [Project Objectives](#project-objectives)
- [Dataset Description](#dataset-description)
- [Technologies & Environment](#technologies--environment)
- [Project Structure](#project-structure)
- [Installation & Setup](#installation--setup)
- [How to Run the Project](#how-to-run-the-project)
- [Data Cleansing & Methodology](#data-cleansing--methodology)
- [Actual Key Findings & Metrics](#actual-key-findings--metrics)
- [Strategic Business Recommendations](#strategic-business-recommendations)
- [Project Limitations](#project-limitations)
- [Conclusion](#conclusion)
- [License & Acknowledgements](#license--acknowledgements)

---

## Project Overview

In the modern e-commerce landscape, transactional logs represent the single most granular and reliable measure of operational health, customer demand, and product performance. This project delivers an end-to-end data analysis of the transnational **Online Retail Dataset** from the **UCI Machine Learning Repository**.

The enterprise under examination is an established United Kingdom-based online giftware retailer operating primarily in the business-to-business (B2B) wholesale sector alongside consumer retail accounts. The dataset documents all transactions occurring between **December 1, 2010 and December 9, 2011**.

By combining rigorous statistical data cleansing, feature engineering, exploratory data analysis, temporal decomposition, geographic profiling, and customer concentration analysis, this project translates raw transactional data into actionable business intelligence and operational strategies.

---

## Business Problem Statement

Online retail businesses face intricate operational and financial challenges:
1. **Extreme Revenue Concentration:** High reliance on a small cluster of wholesale buyers exposes the enterprise to severe liquidity risks if top clients churn.
2. **Volatile Demand Seasonality:** Massive quarterly spikes during Q4 (holiday gifting season) induce inventory stockouts, freight premiums, and severe post-holiday Q1 slumps.
3. **Geographic Imbalance:** Over 85% of total sales originate in the domestic UK market, while high-value European cross-border markets remain under-penetrated.
4. **Order Cancellation & Return Friction:** A high volume of cancelled invoices creates reverse-logistics expenses, ties up inventory, and erodes operating margins.
5. **Inventory Long-Tail Inefficiencies:** A wide catalog of 3,900+ SKUs leads to slow-moving inventory holding costs that tie up vital working capital.

---

## Project Objectives

* **Data Profiling & Cleansing:** Systematically detect and resolve missing values, duplicate entries, negative quantities, administrative non-merchandise codes, and pricing anomalies with full analytical justification.
* **Exact KPI Computation:** Calculate real business metrics: Gross Revenue, Net Revenue, Order Volume, AOV, Average Unit Price, and Cancellation Rates without synthetic estimates.
* **Sales & Seasonality Analysis:** Track month-over-month sales trends, compound monthly growth, and identify peak trading windows.
* **Temporal Behavior Diagnostics:** Profile orders across days of the week and operating hours of the day to identify purchasing habits.
* **Product Performance Evaluation:** Pinpoint top-performing merchandise by revenue and physical unit velocity.
* **Geographic Market Segmentation:** Compare domestic UK sales against international export markets (Netherlands, Ireland, Germany, France).
* **Customer Value & Pareto Testing:** Segment customer accounts by spend tiers, evaluate the Pareto 80/20 distribution, and quantify VIP client concentration.
* **Cancellation & Return Analysis:** Quantify return rates, examine seasonal return dynamics, and calculate the financial impact of cancellations.
* **Strategic Business Recommendations:** Formulate prioritized, high-ROI operational roadmaps for executive leadership.

---

## Dataset Description

The dataset was obtained from the **UCI Machine Learning Repository**:
- **Official Dataset Link:** [https://archive.ics.uci.edu/dataset/352/online+retail](https://archive.ics.uci.edu/dataset/352/online+retail)
- **Raw Dimensions:** 541,909 rows × 8 attributes
- **Timeframe:** December 1, 2010 – December 9, 2011

### Attribute Breakdown

| Attribute | Data Type | Description |
| :--- | :--- | :--- |
| `InvoiceNo` | String (Nominal) | Unique 6-digit transaction ID. Codes prefixed with **'C'** indicate cancellations. |
| `StockCode` | String (Nominal) | Unique 5-digit or alphanumeric merchandise identifier. |
| `Description` | String (Nominal) | Commercial product name or item description. |
| `Quantity` | Integer (Numeric) | Quantity ordered per line item. Negative values indicate cancellations or stock adjustments. |
| `InvoiceDate` | Datetime | Exact timestamp when the invoice was generated. |
| `UnitPrice` | Float (Numeric) | Product unit price in British Pounds (£ Sterling). |
| `CustomerID` | Float (Nominal) | Unique 5-digit identifier assigned to registered customer accounts. |
| `Country` | String (Nominal) | Country of customer residence. |

---

## Technologies & Environment

* **Language:** Python 3.14.3
* **Core Libraries:**
  - `pandas` — High-performance data manipulation, filtering, and aggregation
  - `numpy` — Numerical computations and vectorized array operations
  - `matplotlib` — Publication-ready data visualization
  - `seaborn` — Statistical distributions and correlation heatmaps
  - `openpyxl` — Excel workbook parsing
  - `python-docx` — Programmatic Microsoft Word project report generation
  - `nbclient` & `nbformat` — Headless programmatic Jupyter notebook execution
* **Development Interface:** Jupyter Notebook / Visual Studio Code

---

## Project Structure

```text
Online-Retail-Sales-Analysis/
│
├── SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb      # Complete 28-section executed Jupyter Notebook
├── SnehaVinodPotanavar_OnlineRetailSalesAnalysis_ProjectReport.docx  # Formal Word Project Report (~2.0 MB)
├── requirements.txt                                         # Python package dependencies
├── README.md                                                # Repository documentation & guide
├── analysis_summary.json                                    # Verified JSON ground truth metrics
├── download_data.py                                         # Automated UCI dataset downloader
├── run_full_analysis.py                                     # End-to-end data pipeline & chart generator
├── generate_and_run_notebook.py                             # Notebook builder script
├── execute_notebook.py                                      # Headless notebook execution runner
├── build_report_docx.py                                     # Word report document compiler
│
├── data/
│   └── Online Retail.xlsx                                   # Official raw UCI dataset (23.7 MB)
│
└── charts/                                                  # 12 Publication-quality figures (300 DPI)
    ├── 01_monthly_revenue_trend.png
    ├── 02_monthly_orders_trend.png
    ├── 03_top_products_revenue.png
    ├── 04_top_products_quantity.png
    ├── 05_revenue_by_country.png
    ├── 06_international_revenue_ex_uk.png
    ├── 07_orders_by_day_of_week.png
    ├── 08_orders_by_hour.png
    ├── 09_customer_spending_distribution.png
    ├── 10_top_customers_revenue.png
    ├── 11_cancellation_patterns_monthly.png
    └── 12_correlation_heatmap.png
```

---

## Installation & Setup

### 1. Clone or Open the Workspace
```bash
git clone https://github.com/SnehaVinodPotanavar/Online-Retail-Sales-Analysis.git
cd Online-Retail-Sales-Analysis
```

### 2. Set Up a Virtual Environment (Recommended)
```bash
python -m venv venv
# On Windows:
.\venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate
```

### 3. Install Required Dependencies
```bash
pip install -r requirements.txt
```

### 4. Download the Official UCI Dataset
Run the automated downloader to fetch and extract the official dataset directly from the UCI repository:
```bash
python download_data.py
```

---

## How to Run the Project

### Option A: Execute the Jupyter Notebook
Open and interact with the notebook:
```bash
jupyter notebook SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb
```
Or re-run the notebook headlessly:
```bash
python execute_notebook.py
```

### Option B: Run the Full Analytical Pipeline & Regenerate Charts
To recalculate all metrics and export high-resolution charts into `charts/`:
```bash
python run_full_analysis.py
```

### Option C: Rebuild the Formal Word Project Report
To regenerate the Word document with fresh data and figures:
```bash
python build_report_docx.py
```

---

## Data Cleansing & Methodology

To ensure analytical integrity without blind record deletion, a 5-step data preprocessing framework was implemented:

1. **Deduplication:** Exactly **5,268 duplicate rows** (0.97% of total records) resulting from re-transmissions were eliminated.
2. **Cancellations Segregation:** Invoices starting with **'C'** (**9,288 line items across 3,836 invoices**) were separated into `df_cancelled` for specialized return rate diagnostics, preventing negative quantities from distorting gross sales.
3. **Inventory Write-off Filtering:** **1,336 records** with negative quantities but no 'C' prefix (e.g., damaged items, inventory loss adjustments) were removed.
4. **Price Remediation:** **2,517 records** with `UnitPrice <= 0` (e.g., bad debt adjustments, administrative write-offs) were filtered out.
5. **Administrative Code Filtering:** **1,273 records** corresponding to non-merchandise ledger items (`POST` postage, `D` discount, `M` manual, `BANK CHARGES`, `AMAZONFEE`, `CRUK`) were excluded.
6. **Customer ID Handling:** Maintained the full **522,713 clean sales transactions** for store-level and geographic analytics, while segregating a clean cohort (**397,884 records, 4,334 unique customers**) for customer-level analytics.

---

## Actual Key Findings & Metrics

All numerical statistics below represent actual values computed from the real UCI Online Retail Dataset:

### 1. Executive Key Performance Indicators (KPIs)
* **Total Gross Revenue:** **£10,266,018.19**
* **Total Physical Units Sold:** **5,561,563 units**
* **Total Completed Orders (Invoices):** **19,777 invoices**
* **Unique Merchandise Items (SKUs):** **3,915 stock codes**
* **Unique Identified Customer Accounts:** **4,334 customers**
* **Active Export Countries:** **38 countries**
* **Average Order Value (AOV):** **£519.09**
* **Average Selling Unit Price:** **£3.31**
* **Average Items per Order:** **281.21 units**

### 2. Seasonality & Sales Trends
* **Peak Revenue Month:** **November 2011 (2011-11)** reached **£1,453,265.98** across **2,751 orders** due to Q4 holiday retail stocking.
* **Trough Month:** **February 2011 (2011-02)** reached **£508,081.54** across **1,114 orders** (a 186.0% gap compared to November).
* **Final Window:** December 2011 recorded **£517,110.74** across 9 operating days (projected run-rate >£1.7M).

### 3. Temporal Purchasing Cadence
* **Peak Trading Day:** **Thursday** is the top ordering day with **4,209 orders** (£2,109,720.69), followed by Wednesday (3,995 orders).
* **Weekend Shutdown:** **0 orders are recorded on Saturdays**, reflecting a wholesale business model where buyers order ahead of weekends.
* **Peak Trading Hour:** **12:00 PM (noon)** represents the global peak (**3,205 orders**), with 72% of orders placed between 10:00 AM and 3:00 PM.

### 4. Product Intelligence
* **Top Product by Revenue:** Stock Code `22423` — **`REGENCY CAKESTAND 3 TIER`** generated **£174,156.54** across 13,812 units sold (average price ~£12.44).
* **Top Product by Quantity:** Stock Code `23843` — **`PAPER CRAFT , LITTLE BIRDIE`** achieved the highest unit sales with **80,995 units** (£168,469.60), followed by `MEDIUM CERAMIC TOP STORAGE JAR` (**78,033 units**, £81,700.92).

### 5. Geographic Market Analysis
* **Domestic Dominance:** The **United Kingdom** accounts for **85.11% (£8,737,666.54)** of total revenue and 18,016 orders.
* **International Export:** Cross-border sales account for **14.89% (£1,528,351.65)** across 37 foreign destinations.
* **Top International Markets:**
  1. **Netherlands:** £283,889.34 (94 orders, AOV = £3,020.10)
  2. **EIRE (Ireland):** £276,090.86 (286 orders, AOV = £965.35)
  3. **Germany:** £228,867.14 (456 orders, AOV = £501.90)
  4. **France:** £208,610.15 (391 orders, AOV = £533.53)
  5. **Australia:** £138,521.31 (57 orders, AOV = £2,430.20)

### 6. Customer Spend & Pareto Concentration
* **Pareto Concentration:**
  - **Top 1% of customers** generate **32.02%** of identified revenue (£2.85M).
  - **Top 10% of customers** generate **61.39%** of identified revenue (£5.46M).
  - **Top 20% of customers** generate **74.59%** of identified revenue (£6.64M).
* **Top Client:** Customer ID `14646` (Netherlands wholesale buyer) spent **£279,138.02** across 72 orders (196,915 units).
* **Spend Tiers:**
  - Low (<£250): 1,396 accounts (32.2%)
  - Mid (£250–£1,000): 1,643 accounts (37.9%)
  - High (£1,000–£5,000): 1,201 accounts (27.7%)
  - VIP (>£5,000): 94 accounts (2.2%) — generating over 35% of identified revenue.

### 7. Cancellation & Return Analysis
* **Cancelled Invoices:** **3,836 invoices** (9,288 lines) were cancelled.
* **Cancelled Unit Volume:** **275,560 units**.
* **Gross Monetary Value:** **£893,979.73**.
* **Cancellation Invoice Rate:** **16.25%** across all order attempts.
* **Return Loss Ratio:** Remained steady between 6.5% and 8.8% of monthly gross sales throughout the year.

---

## Strategic Business Recommendations

1. **Dedicated VIP Key Account Success Team:** Implement dedicated account management for the top 100 wholesale clients who generate the vast majority of profits. Provide custom bulk pallet pricing, contracted SLAs, and inventory reservations during Q4.
2. **Predictive Q4 Inventory Staging:** Initiate manufacturing procurement for top-tier revenue drivers (e.g., Regency Cakestands, Ceramic Storage Jars) by June/July, ensuring UK warehouse arrival by August to avoid Q4 stockouts.
3. **European Cross-Border Logistics Hub:** Establish a bonded 3PL fulfillment partnership in the Netherlands or Germany to cut delivery transit times from 5 days to 24–48 hours for European wholesale buyers.
4. **Cancellation Quality Audits:** Perform physical packaging inspections on fragile glassware and ceramic merchandise with high return rates to reduce transit breakage and lower the £893k cancellation volume.
5. **Weekend Digital Pre-Order Automation:** Implement automated order intake and intelligent picking wave pre-scheduling on Saturdays to eliminate the Monday morning warehouse dispatch backlog.

---

## Project Limitations

1. **Lack of Cost Data (COGS):** The dataset records gross selling prices without product acquisition costs or shipping overhead, limiting financial analysis to gross revenues rather than net margin profitability.
2. **Unregistered Guest Transactions:** 24.93% of transactions lack a CustomerID, precluding complete lifetime value analysis for guest users.
3. **Truncated Final Month:** Data collection concluded on December 9, 2011, preventing full-month comparisons for December 2011.
4. **Absence of Marketing Attribution:** Channel attribution, advertising costs, and device categories were not tracked.

---

## Conclusion

This project delivers an end-to-end, scientifically validated data analysis of the UCI Online Retail Dataset. By establishing clean transaction boundaries and conducting multi-dimensional explorations, this work quantifies key sales drivers across temporal, product, geographic, and customer domains. The resulting insights and recommendations equip leadership with concrete strategies to protect wholesale revenue, streamline inventory, expand European markets, and reduce return liabilities.

---

## License & Acknowledgements

* **Dataset Attribution:** Daqing Chen, London South Bank University; hosted by the University of California, Irvine (UCI) Machine Learning Repository.
* **Author:** Sneha Vinod Potanavar (September 2026).
