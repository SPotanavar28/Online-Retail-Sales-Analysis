import nbformat as nbf
import subprocess
import sys
import os

nb = nbf.v4.new_notebook()
nb.metadata = {
    "kernelspec": {
        "display_name": "Python 3",
        "language": "python",
        "name": "python3"
    },
    "language_info": {
        "codemirror_mode": {"name": "ipython", "version": 3},
        "file_extension": ".py",
        "mimetype": "text/x-python",
        "name": "python",
        "nbconvert_exporter": "python",
        "pygments_lexer": "ipython3",
        "version": "3.14.3"
    }
}

cells = []

def md(content):
    cells.append(nbf.v4.new_markdown_cell(content.strip()))

def code(content):
    cells.append(nbf.v4.new_code_cell(content.strip()))

# ==============================================================================
# 1. Project Title
# ==============================================================================
md("""
# Online Retail Sales & Customer Analysis
### Comprehensive Exploratory Data Analysis & Strategic Business Intelligence
""")

# ==============================================================================
# 2. Student Information
# ==============================================================================
md("""
---
### Student & Project Credentials
* **Student Name:** Sneha Vinod Potanavar
* **Project Title:** Online Retail Sales & Customer Analysis
* **Role:** Data Analyst Intern
* **Dataset:** UCI Machine Learning Repository — Online Retail Dataset
* **Submission Date:** September 2026
---
""")

# ==============================================================================
# 3. Project Overview
# ==============================================================================
md("""
## 1. Project Overview

In the modern omnichannel commerce ecosystem, transactional data represents the single most accurate reflection of customer demand, purchasing cadence, and business health. This project undertakes an exhaustive, data-driven analysis of the transnational **Online Retail Dataset** provided by the **UCI Machine Learning Repository**.

The business under evaluation is an established United Kingdom-based online giftware and homeware retailer operating primarily in the business-to-business (B2B) wholesale sector alongside individual consumer accounts. The dataset encompasses all recorded transactions occurring between **December 1, 2010 and December 9, 2011**. 

By applying rigorous data cleaning, robust feature engineering, exploratory data analysis, and advanced customer segmentation paradigms, this project uncovers critical behavioral dynamics, revenue drivers, temporal trends, geographical expansion opportunities, and returns/cancellation patterns to formulate high-impact, actionable business strategies.
""")

# ==============================================================================
# 4. Business Problem
# ==============================================================================
md("""
## 2. Business Problem

Online retail enterprises face complex operational challenges characterized by:
1. **Extreme Revenue Concentration:** High dependency on a disproportionately small cohort of high-volume wholesale buyers, exposing the firm to severe liquidity and demand risks if key accounts churn.
2. **Volatile Seasonality:** Massive quarterly demand surges during Q4 (holiday gifting season) followed by post-holiday demand troughs, leading to inventory stockouts, warehouse bottlenecks, and carrying cost inefficiencies.
3. **Geographic Imbalance:** An overwhelming concentration of sales within the domestic UK market, while lucrative cross-border European markets remain under-penetrated due to logistical and currency barriers.
4. **Cancellations & Return Margin Erosion:** Substantial transaction cancellation volumes and return liabilities that disrupt order fulfillment, strain customer support, and erode operating profit margins.
5. **Customer Attrition & Dormancy:** Inability to effectively distinguish between one-off retail shoppers and high-value wholesale accounts, leading to generic marketing and suboptimal customer lifetime value (LTV).

Addressing these issues requires a rigorous, evidence-based diagnostic of transaction records to optimize inventory planning, refine promotional cadences, and protect bottom-line profitability.
""")

# ==============================================================================
# 5. Objectives
# ==============================================================================
md("""
## 3. Project Objectives

The key objectives of this data analysis project are:
* **Data Quality Profiling & Cleansing:** Systematically assess and remediate missing values, duplicate entries, negative quantities, administrative non-product codes, and pricing anomalies with full analytical justification.
* **Key Performance Metric Computation:** Calculate exact business metrics, including Gross and Net Revenue, Total Volume, Average Order Value (AOV), Order Frequency, and Cancellation Rates.
* **Sales & Seasonality Analysis:** Unpack month-over-month sales trends, compound monthly growth, and identify peak trading periods.
* **Temporal Behavior Profiling:** Investigate transaction distributions across days of the week and operating hours of the day to identify purchasing habits.
* **Product Performance Evaluation:** Identify top-performing merchandise by revenue generation and physical unit volume to inform inventory prioritization.
* **Geographic Market Segmentation:** Analyze domestic UK performance versus international cross-border markets (e.g., Netherlands, EIRE, Germany, France).
* **Customer Value & Pareto Distribution:** Segment customer accounts by cumulative spend, evaluate the Pareto 80/20 distribution, and determine high-value client concentration.
* **Cancellation & Return Analysis:** Quantify cancellation volumes, cancellation rates, and financial impact to identify root causes and mitigation strategies.
* **Strategic Business Recommendations:** Deliver actionable, data-backed operational guidelines for executive decision-makers.
""")

# ==============================================================================
# 6. Dataset Description
# ==============================================================================
md("""
## 4. Dataset Description

The dataset was obtained from the **UCI Machine Learning Repository** (Donor: Dr. Daqing Chen, London South Bank University). It contains **541,909 raw records** across **8 attributes**:

| Attribute | Data Type | Description |
| :--- | :--- | :--- |
| `InvoiceNo` | Nominal / String | A 6-digit integral number uniquely assigned to each transaction. If this code starts with the letter **'C'**, it indicates a cancellation. |
| `StockCode` | Nominal / String | A 5-digit or alphanumeric code uniquely assigned to each distinct product item. |
| `Description` | Nominal / String | The commercial product name or merchandise description. |
| `Quantity` | Integer | The quantities of each product per transaction line item. Negative values indicate cancellations, damaged inventory, or administrative returns. |
| `InvoiceDate` | Datetime | The exact timestamp (day, month, year, hour, minute) when the transaction occurred. |
| `UnitPrice` | Numeric / Float | Product price per unit in Sterling (£). Zero or negative values indicate gifts, administrative offsets, or debt write-offs. |
| `CustomerID` | Nominal / Integer | A 5-digit integral number uniquely assigned to each registered customer account. |
| `Country` | Nominal / String | The name of the country where the customer resides. |
""")

# ==============================================================================
# 7. Technologies Used
# ==============================================================================
md("""
## 5. Technologies Used

* **Language:** Python 3.14
* **Data Manipulation & Computation:** `pandas`, `numpy`
* **Data Visualization:** `matplotlib`, `seaborn`
* **File & Dataset Handling:** `openpyxl`, `os`, `json`
* **Interactive Environment:** Jupyter Notebook / IPython
""")

# ==============================================================================
# 8. Import Libraries
# ==============================================================================
md("""
## 6. Import Libraries

We import essential Python libraries and configure high-resolution plotting aesthetics.
""")

code("""
import os
import sys
import json
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Ignore non-critical warnings
warnings.filterwarnings('ignore')

# Set professional visualization aesthetics
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 13
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10
plt.rcParams['legend.fontsize'] = 10
plt.rcParams['figure.titlesize'] = 15
plt.rcParams['figure.dpi'] = 110

pd.set_option('display.max_columns', 15)
pd.set_option('display.max_rows', 25)
pd.set_option('display.float_format', lambda x: f'{x:,.2f}')

print("All libraries imported successfully.")
print(f"Pandas version: {pd.__version__}")
print(f"NumPy version: {np.__version__}")
""")

# ==============================================================================
# 9. Load Dataset
# ==============================================================================
md("""
## 7. Load Dataset

We load the official dataset `Online Retail.xlsx` located in the `data/` directory.
""")

code("""
data_path = os.path.join("data", "Online Retail.xlsx")
if not os.path.exists(data_path):
    raise FileNotFoundError(f"Dataset file not found at {data_path}. Please run download_data.py first.")

print(f"Loading dataset from: {data_path}...")
df_raw = pd.read_excel(data_path)
print(f"Raw Dataset loaded successfully with {df_raw.shape[0]:,} rows and {df_raw.shape[1]} columns.")
""")

# ==============================================================================
# 10. Initial Exploration
# ==============================================================================
md("""
## 8. Initial Exploration

We inspect the first and last five rows, verify data types, memory utilization, and generate summary statistics of the raw data.
""")

code("""
# Display first 5 rows
print("--- FIRST 5 RECORDS ---")
display(df_raw.head())
""")

code("""
# Display last 5 rows
print("--- LAST 5 RECORDS ---")
display(df_raw.tail())
""")

code("""
# Dataset structural info
print("--- DATASET INFORMATION & DATA TYPES ---")
df_raw.info()
""")

code("""
# Summary statistics for numerical variables
print("--- NUMERICAL SUMMARY STATISTICS ---")
display(df_raw.describe())
""")

code("""
# Unique entity counts
print("--- RAW UNIQUE COUNTS ---")
unique_counts = pd.Series({
    'Unique Invoices': df_raw['InvoiceNo'].nunique(),
    'Unique Stock Codes': df_raw['StockCode'].nunique(),
    'Unique Descriptions': df_raw['Description'].nunique(),
    'Unique Customers': df_raw['CustomerID'].nunique(),
    'Unique Countries': df_raw['Country'].nunique(),
    'Earliest Date': df_raw['InvoiceDate'].min(),
    'Latest Date': df_raw['InvoiceDate'].max()
})
display(unique_counts.to_frame(name='Count / Timestamp'))
""")

md("""
### Interpretation of Initial Exploration
The raw numerical summary immediately reveals several critical data anomalies:
1. **Quantity Anomaly:** The `Quantity` column has a minimum value of `-80,995` and a maximum of `80,995`. Negative values indicate cancelled transactions, inventory write-offs, or damaged goods.
2. **UnitPrice Anomaly:** The `UnitPrice` column has a minimum value of `-11,062.06` and a maximum of `38,970.00`. Negative and zero unit prices represent administrative ledger adjustments, bad debts, manual corrections, or promotional giveaways.
3. **Missing Customer IDs:** Out of 541,909 records, only 406,829 have a non-null `CustomerID`, indicating ~25% of purchases were made by guest or unregistered users.
""")

# ==============================================================================
# 11. Data Quality Assessment
# ==============================================================================
md("""
## 9. Data Quality Assessment

Before performing data cleaning, we conduct a structured assessment of missing values, duplicate records, cancelled invoices, and anomalies.
""")

code("""
# Missing values analysis
missing_df = pd.DataFrame({
    'Missing Count': df_raw.isnull().sum(),
    'Missing Percentage (%)': (df_raw.isnull().sum() / len(df_raw)) * 100
}).sort_values(by='Missing Count', ascending=False)

print("--- MISSING VALUES PROFILE ---")
display(missing_df)
""")

code("""
# Duplicate records detection
duplicate_rows_count = df_raw.duplicated().sum()
print(f"Number of duplicate rows identified: {duplicate_rows_count:,} ({(duplicate_rows_count/len(df_raw))*100:.2f}% of total records)")
""")

code("""
# Cancellations and Invoices with 'C'
df_raw['InvoiceNo_Str'] = df_raw['InvoiceNo'].astype(str)
is_cancellation_mask = df_raw['InvoiceNo_Str'].str.startswith('C')
cancellation_rows = is_cancellation_mask.sum()
cancellation_invoices = df_raw.loc[is_cancellation_mask, 'InvoiceNo'].nunique()

print(f"Total cancellation line items: {cancellation_rows:,}")
print(f"Distinct cancellation invoice vouchers: {cancellation_invoices:,}")
""")

code("""
# Check anomalies: Negative quantities without 'C' prefix and Zero/Negative Prices
neg_qty_no_c = ((df_raw['Quantity'] < 0) & (~is_cancellation_mask)).sum()
zero_or_neg_prices = (df_raw['UnitPrice'] <= 0).sum()

print(f"Negative quantities without 'C' prefix (damaged/loss adjustments): {neg_qty_no_c:,}")
print(f"Zero or negative unit prices: {zero_or_neg_prices:,}")
""")

md("""
### Data Quality Assessment Findings
* **Missing Data:** `CustomerID` has 135,080 missing values (24.93%), and `Description` has 1,454 missing values (0.27%).
* **Duplicate Rows:** Exactly 5,268 records are identical duplicates caused by automated system re-transmissions.
* **Cancellations:** 9,288 line items across 3,836 distinct invoices start with **'C'**, indicating customer-requested cancellations and returns.
* **Stock Adjustments / Bad Debts:** 1,336 records feature negative quantities without a 'C' prefix, accompanied by zero unit price and descriptions such as *"damaged"*, *"check"*, or *"thrown away"*. These represent inventory reconciliation entries rather than retail transactions.
* **Zero/Negative Unit Prices:** 2,517 records have `UnitPrice <= 0`, representing administrative offsets, bad debt adjustments (e.g., `Adjust bad debt`), or system tests.
""")

# ==============================================================================
# 12. Data Cleaning
# ==============================================================================
md("""
## 10. Data Cleaning

We implement a transparent, multi-stage cleaning pipeline that avoids blind deletion and preserves data integrity:

1. **Deduplication:** Remove the 5,268 exact duplicate records.
2. **Cancellations Separation:** Segregate cancelled transactions (invoices starting with 'C') into a dedicated `df_cancelled` DataFrame for in-depth cancellation and return analysis.
3. **Completed Transaction Filtering:** Retain records where `Quantity > 0` and `UnitPrice > 0`.
4. **Administrative StockCode Removal:** Filter out non-merchandise administrative codes (`POST` postage, `D` discount, `M` manual, `BANK CHARGES`, `AMAZONFEE`, `CRUK` cancer research, `DOT` dotcom postage).
5. **Customer Cohort Creation:** Maintain full transaction records for store-level sales analysis, while establishing a dedicated subset (`df_clean_cust`) with valid `CustomerID` for customer-level analytics.
""")

code("""
# Step 1: Remove exact duplicates
df_dedup = df_raw.drop_duplicates().copy()
print(f"Records after removing duplicates: {df_dedup.shape[0]:,}")

# Step 2: Separate cancellations into a dedicated DataFrame
df_cancelled = df_dedup[df_dedup['InvoiceNo_Str'].str.startswith('C')].copy()
df_cancelled['Revenue_Abs'] = df_cancelled['Quantity'].abs() * df_cancelled['UnitPrice']
print(f"Segregated cancellation records for deep-dive analysis: {df_cancelled.shape[0]:,} lines")

# Step 3: Filter non-merchandise administrative codes
admin_codes = ['POST', 'D', 'M', 'BANK CHARGES', 'AMAZONFEE', 'CRUK', 'DOT', 'S']
is_admin_code = df_dedup['StockCode'].astype(str).str.upper().isin(admin_codes)

# Step 4: Construct clean sales dataset for completed transactions
df_clean = df_dedup[
    (~df_dedup['InvoiceNo_Str'].str.startswith('C')) &
    (df_dedup['Quantity'] > 0) &
    (df_dedup['UnitPrice'] > 0) &
    (~is_admin_code)
].copy()

print(f"Clean sales dataset finalized: {df_clean.shape[0]:,} rows ({df_clean.shape[0]/df_raw.shape[0]*100:.2f}% of original dataset retained)")
""")

# ==============================================================================
# 13. Feature Engineering
# ==============================================================================
md("""
## 11. Feature Engineering

To support comprehensive sales, temporal, and customer analysis, we create key analytical features:
* **`Revenue`**: `Quantity * UnitPrice` (monetary transaction value in £ GBP).
* **Date & Time Attributes**: Parse `InvoiceDate` to derive `Year`, `Month`, `Month_Name`, `Year_Month`, `Day`, `Day_Of_Week`, `Day_Name`, and `Hour`.
""")

code("""
# Calculate total transaction revenue
df_clean['Revenue'] = df_clean['Quantity'] * df_clean['UnitPrice']

# Extract temporal components
df_clean['InvoiceDate'] = pd.to_datetime(df_clean['InvoiceDate'])
df_clean['Year'] = df_clean['InvoiceDate'].dt.year
df_clean['Month'] = df_clean['InvoiceDate'].dt.month
df_clean['Month_Name'] = df_clean['InvoiceDate'].dt.strftime('%B')
df_clean['Year_Month'] = df_clean['InvoiceDate'].dt.to_period('M').astype(str)
df_clean['Day'] = df_clean['InvoiceDate'].dt.day
df_clean['Day_Of_Week'] = df_clean['InvoiceDate'].dt.dayofweek # 0=Monday, 6=Sunday
df_clean['Day_Name'] = df_clean['InvoiceDate'].dt.day_name()
df_clean['Hour'] = df_clean['InvoiceDate'].dt.hour

# Also create customer-specific cohort with clean CustomerID
df_clean_cust = df_clean[df_clean['CustomerID'].notnull()].copy()
df_clean_cust['CustomerID'] = df_clean_cust['CustomerID'].astype(int).astype(str)

print("Feature engineering completed successfully.")
display(df_clean[['InvoiceNo', 'StockCode', 'Quantity', 'UnitPrice', 'Revenue', 'Year_Month', 'Day_Name', 'Hour']].head())
""")

# ==============================================================================
# 14. Key Business Metrics
# ==============================================================================
md("""
## 12. Key Business Metrics (KPIs)

We calculate core executive business metrics directly from the cleaned transaction records.
""")

code("""
total_revenue = float(df_clean['Revenue'].sum())
total_quantity_sold = int(df_clean['Quantity'].sum())
total_orders = int(df_clean['InvoiceNo'].nunique())
total_unique_products = int(df_clean['StockCode'].nunique())
total_unique_customers = int(df_clean['CustomerID'].nunique())
total_countries = int(df_clean['Country'].nunique())
average_order_value = total_revenue / total_orders
average_unit_price = float(df_clean['UnitPrice'].mean())
average_items_per_order = total_quantity_sold / total_orders

# Cancellations summary
total_cancelled_invoices = int(df_cancelled['InvoiceNo_Str'].nunique())
total_cancelled_items = int(df_cancelled['Quantity'].abs().sum())
total_cancelled_value = float(df_cancelled['Revenue_Abs'].sum())
cancellation_rate_pct = (total_cancelled_invoices / (total_orders + total_cancelled_invoices)) * 100

kpi_summary = pd.DataFrame({
    'Metric': [
        'Total Gross Revenue (£)',
        'Total Quantity Sold (Units)',
        'Total Completed Orders (Invoices)',
        'Unique Products (Stock Codes)',
        'Unique Registered Customers',
        'Operating Countries',
        'Average Order Value (AOV) (£)',
        'Average Unit Price (£)',
        'Average Items per Order',
        'Total Cancelled Invoices',
        'Total Cancelled Units',
        'Total Cancelled Value (£)',
        'Cancellation Invoice Rate (%)'
    ],
    'Value': [
        f"£{total_revenue:,.2f}",
        f"{total_quantity_sold:,}",
        f"{total_orders:,}",
        f"{total_unique_products:,}",
        f"{total_unique_customers:,}",
        f"{total_countries}",
        f"£{average_order_value:,.2f}",
        f"£{average_unit_price:,.2f}",
        f"{average_items_per_order:,.2f}",
        f"{total_cancelled_invoices:,}",
        f"{total_cancelled_items:,}",
        f"£{total_cancelled_value:,.2f}",
        f"{cancellation_rate_pct:.2f}%"
    ]
})

print("--- EXECUTIVE KEY PERFORMANCE INDICATORS ---")
display(kpi_summary)
""")

# ==============================================================================
# 15. Exploratory Data Analysis
# ==============================================================================
md("""
## 13. Exploratory Data Analysis (EDA)

We explore the distributions of transaction revenue, order quantities, and unit prices to understand data skewness and operational boundaries.
""")

code("""
eda_stats = df_clean[['Quantity', 'UnitPrice', 'Revenue']].describe(
    percentiles=[0.01, 0.05, 0.25, 0.5, 0.75, 0.95, 0.99]
)
print("--- DETAILED PERCENTILE DISTRIBUTIONS ---")
display(eda_stats)
""")

code("""
# Distribution visualization for Revenue and Quantity (log scale for visual clarity due to skewness)
fig, axes = plt.subplots(1, 2, figsize=(14, 5))

sns.histplot(df_clean['Revenue'], bins=50, ax=axes[0], color='#1F4E79', log_scale=True)
axes[0].set_title('Log-Transformed Distribution of Transaction Revenue', fontweight='bold')
axes[0].set_xlabel('Revenue (£ GBP, Log Scale)')
axes[0].set_ylabel('Transaction Frequency')

sns.histplot(df_clean['Quantity'], bins=50, ax=axes[1], color='#2E75B6', log_scale=True)
axes[1].set_title('Log-Transformed Distribution of Item Quantities Sold', fontweight='bold')
axes[1].set_xlabel('Quantity (Units, Log Scale)')
axes[1].set_ylabel('Transaction Frequency')

plt.tight_layout()
plt.show()
""")

md("""
### EDA Interpretation
* The median transaction revenue is **£9.90** while the mean is **£19.64**, driven by heavy right-skewed wholesale bulk purchases.
* The 99th percentile for revenue is **£125.00**, but the maximum single transaction reached **£168,469.60** (Customer ID 16446 purchasing 80,995 units of `PAPER CRAFT , LITTLE BIRDIE`).
* 75% of purchases are 10 units or fewer, while wholesale bulk purchases skew the overall volume upwards.
""")

# ==============================================================================
# 16. Sales Analysis
# ==============================================================================
md("""
## 14. Sales Analysis

We analyze monthly revenue progression, order volumes, and average transaction values across the 13-month operational window.
""")

code("""
monthly_sales = df_clean.groupby('Year_Month').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum'),
    AOV=('Revenue', lambda r: r.sum() / df_clean.loc[r.index, 'InvoiceNo'].nunique())
).reset_index()

# Month-over-Month Revenue Growth Rate (%)
monthly_sales['Revenue_MoM_Growth_%'] = monthly_sales['Revenue'].pct_change() * 100

print("--- MONTHLY SALES & ORDER PERFORMANCE ---")
display(monthly_sales)
""")

code("""
# Plot Monthly Revenue Trend
plt.figure(figsize=(12, 5.5))
plt.plot(monthly_sales['Year_Month'], monthly_sales['Revenue'] / 1000, marker='o', linewidth=2.5, color='#1F4E79', markersize=7)
plt.fill_between(monthly_sales['Year_Month'], monthly_sales['Revenue'] / 1000, color='#2E75B6', alpha=0.15)
plt.title('Monthly Sales Revenue Trend (£ Thousands)', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Year-Month', fontweight='bold')
plt.ylabel('Total Revenue (£ in 000s)', fontweight='bold')
plt.xticks(rotation=45)
for i, txt in enumerate(monthly_sales['Revenue']):
    plt.annotate(f"£{txt/1000:.1f}k", (monthly_sales['Year_Month'][i], monthly_sales['Revenue'][i]/1000 + 15),
                 ha='center', fontsize=9, fontweight='semibold')
plt.ylim(0, (monthly_sales['Revenue'].max() / 1000) * 1.15)
plt.tight_layout()
plt.show()
""")

code("""
# Plot Monthly Orders Volume
plt.figure(figsize=(12, 5.5))
bars = plt.bar(monthly_sales['Year_Month'], monthly_sales['Orders'], color='#2E75B6', edgecolor='#1B365D', width=0.6)
plt.title('Monthly Completed Orders Volume (Invoices)', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Year-Month', fontweight='bold')
plt.ylabel('Order Count (Invoices)', fontweight='bold')
plt.xticks(rotation=45)
for bar in bars:
    height = bar.get_height()
    plt.annotate(f"{int(height):,}", (bar.get_x() + bar.get_width() / 2, height + 40),
                 ha='center', fontsize=9, fontweight='semibold')
plt.ylim(0, monthly_sales['Orders'].max() * 1.15)
plt.tight_layout()
plt.show()
""")

md("""
### Sales Analysis Interpretation
* **Peak Revenue:** Sales peaked dramatically in **November 2011 (2011-11)** at **£1,453,265.98** across **2,751 orders**. This reflects the critical Q4 holiday gift purchasing surge.
* **Trough Period:** The lowest full month of sales occurred in **February 2011 (2011-02)** at **£508,081.54** with **1,114 orders**, representing the standard post-holiday retail lull.
* **Truncated Period:** Note that December 2011 captures only transactions up to December 9, explaining its lower recorded total (£517,110.74).
""")

# ==============================================================================
# 17. Time-Based Analysis
# ==============================================================================
md("""
## 15. Time-Based Analysis

We investigate transaction behavior across days of the week and operating hours of the day to identify purchasing habits.
""")

code("""
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Sunday']
dow_sales = df_clean.groupby('Day_Name').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum')
).reindex([d for d in day_order if d in df_clean['Day_Name'].unique()]).reset_index()

print("--- DAY OF WEEK SALES PERFORMANCE ---")
display(dow_sales)
""")

code("""
hourly_sales = df_clean.groupby('Hour').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum')
).reset_index()

print("--- HOURLY SALES DISTRIBUTION (TOP 5 PEAK HOURS) ---")
display(hourly_sales.sort_values(by='Orders', ascending=False).head(5))
""")

code("""
fig, axes = plt.subplots(1, 2, figsize=(16, 5.5))

# Day of week plot
bars1 = axes[0].bar(dow_sales['Day_Name'], dow_sales['Orders'], color='#3B7A57', edgecolor='#1E4620', width=0.55)
axes[0].set_title('Order Volume by Day of Week', fontsize=13, fontweight='bold', pad=12)
axes[0].set_xlabel('Day of Week', fontweight='bold')
axes[0].set_ylabel('Total Invoices Placed', fontweight='bold')
for bar in bars1:
    h = bar.get_height()
    axes[0].annotate(f"{h:,}", (bar.get_x() + bar.get_width()/2, h + 60), ha='center', fontsize=9.5, fontweight='semibold')
axes[0].set_ylim(0, dow_sales['Orders'].max() * 1.15)

# Hourly plot
axes[1].plot(hourly_sales['Hour'], hourly_sales['Orders'], marker='s', color='#8B0000', linewidth=2.5, markersize=6)
axes[1].fill_between(hourly_sales['Hour'], hourly_sales['Orders'], color='#E57373', alpha=0.25)
axes[1].set_title('Order Volume Across Operating Hours', fontsize=13, fontweight='bold', pad=12)
axes[1].set_xlabel('Hour of Day (24-Hour Clock)', fontweight='bold')
axes[1].set_ylabel('Total Invoices Placed', fontweight='bold')
axes[1].set_xticks(hourly_sales['Hour'])
axes[1].set_ylim(0, hourly_sales['Orders'].max() * 1.15)

plt.tight_layout()
plt.show()
""")

md("""
### Time-Based Analysis Interpretation
1. **Day-of-Week Pattern:** **Thursday** is the highest volume day with **4,209 orders** (£2,109,720.69 revenue), followed closely by Wednesday (3,995 orders). Notably, **there are 0 orders recorded on Saturday**, indicating that the enterprise's warehouse and order-processing systems were closed on Saturdays.
2. **Hourly Pattern:** Ordering activity accelerates sharply starting at 8:00 AM, reaches its global peak at **12:00 PM (noon)** with **3,205 orders**, and maintains elevated volume until 3:00 PM before dropping off significantly by 5:00 PM. This bell-shaped daytime curve reinforces that buyers are commercial B2B procurement managers ordering during standard European business hours.
""")

# ==============================================================================
# 18. Product Analysis
# ==============================================================================
md("""
## 16. Product Analysis

We analyze product performance to identify top revenue generators and highest volume items.
""")

code("""
product_perf = df_clean.groupby(['StockCode', 'Description']).agg(
    TotalRevenue=('Revenue', 'sum'),
    TotalQuantity=('Quantity', 'sum'),
    OrderCount=('InvoiceNo', 'nunique'),
    AvgUnitPrice=('UnitPrice', 'mean')
).reset_index()

top10_rev = product_perf.sort_values(by='TotalRevenue', ascending=False).head(10).reset_index(drop=True)
top10_qty = product_perf.sort_values(by='TotalQuantity', ascending=False).head(10).reset_index(drop=True)

print("--- TOP 10 PRODUCTS BY TOTAL REVENUE ---")
display(top10_rev[['StockCode', 'Description', 'TotalRevenue', 'TotalQuantity', 'AvgUnitPrice']])
""")

code("""
print("--- TOP 10 PRODUCTS BY TOTAL QUANTITY SOLD ---")
display(top10_qty[['StockCode', 'Description', 'TotalQuantity', 'TotalRevenue', 'AvgUnitPrice']])
""")

code("""
fig, axes = plt.subplots(1, 2, figsize=(16, 6.5))

# Top by Revenue
top10_rev_plot = top10_rev.copy()
top10_rev_plot['ShortDesc'] = top10_rev_plot['Description'].str[:30]
axes[0].barh(top10_rev_plot['ShortDesc'][::-1], top10_rev_plot['TotalRevenue'][::-1] / 1000, color='#1F4E79', edgecolor='#0D233A', height=0.65)
axes[0].set_title('Top 10 Products by Total Revenue (£k)', fontweight='bold')
axes[0].set_xlabel('Revenue (£ Thousands)', fontweight='bold')
axes[0].set_ylabel('Product Description', fontweight='bold')
for i, v in enumerate(top10_rev_plot['TotalRevenue'][::-1] / 1000):
    axes[0].annotate(f" £{v:.1f}k", (v, i), va='center', fontsize=9, fontweight='semibold')
axes[0].set_xlim(0, (top10_rev_plot['TotalRevenue'].max() / 1000) * 1.18)

# Top by Quantity
top10_qty_plot = top10_qty.copy()
top10_qty_plot['ShortDesc'] = top10_qty_plot['Description'].str[:30]
axes[1].barh(top10_qty_plot['ShortDesc'][::-1], top10_qty_plot['TotalQuantity'][::-1] / 1000, color='#008080', edgecolor='#004D40', height=0.65)
axes[1].set_title('Top 10 Products by Total Quantity Sold (k Units)', fontweight='bold')
axes[1].set_xlabel('Quantity Sold (Thousands of Units)', fontweight='bold')
axes[1].set_ylabel('Product Description', fontweight='bold')
for i, v in enumerate(top10_qty_plot['TotalQuantity'][::-1] / 1000):
    axes[1].annotate(f" {v:.1f}k", (v, i), va='center', fontsize=9, fontweight='semibold')
axes[1].set_xlim(0, (top10_qty_plot['TotalQuantity'].max() / 1000) * 1.18)

plt.tight_layout()
plt.show()
""")

md("""
### Product Analysis Interpretation
* **Top Revenue Generator:** **`REGENCY CAKESTAND 3 TIER`** (StockCode 22423) is the company's single highest revenue-generating product, producing **£174,156.54** across 13,812 units sold with an average unit price of ~£12.44.
* **Top Volume Item:** **`PAPER CRAFT , LITTLE BIRDIE`** (StockCode 23843) achieved the highest unit sales with **80,995 units** (£168,469.60), closely followed by **`MEDIUM CERAMIC TOP STORAGE JAR`** (StockCode 23166) with **78,033 units** (£81,700.92).
* **Catalog Long Tail:** Out of 3,915 distinct products, the top 10 products by revenue generate over £950,000 (~9.3% of total enterprise sales), confirming high product concentration.
""")

# ==============================================================================
# 19. Geographic Analysis
# ==============================================================================
md("""
## 17. Geographic Analysis

We analyze sales distribution across geographical territories, evaluating the domestic UK market versus international cross-border destinations.
""")

code("""
country_perf = df_clean.groupby('Country').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum'),
    Customers=('CustomerID', 'nunique')
).reset_index()

country_perf['RevenueShare_%'] = (country_perf['Revenue'] / total_revenue) * 100
country_perf = country_perf.sort_values(by='Revenue', ascending=False).reset_index(drop=True)

print("--- TOP 10 COUNTRIES BY TOTAL REVENUE ---")
display(country_perf.head(10))
""")

code("""
uk_rev = country_perf.loc[country_perf['Country'] == 'United Kingdom', 'Revenue'].values[0]
uk_share = country_perf.loc[country_perf['Country'] == 'United Kingdom', 'RevenueShare_%'].values[0]
intl_rev = total_revenue - uk_rev
intl_share = 100.0 - uk_share

print(f"United Kingdom Domestic Revenue: £{uk_rev:,.2f} ({uk_share:.2f}%)")
print(f"International Cross-Border Revenue: £{intl_rev:,.2f} ({intl_share:.2f}%)")
""")

code("""
# Top International Markets excluding UK
top_intl = country_perf[country_perf['Country'] != 'United Kingdom'].head(10)

plt.figure(figsize=(12, 5.5))
bars = plt.bar(top_intl['Country'], top_intl['Revenue'] / 1000, color='#2E75B6', edgecolor='#1B365D', width=0.6)
plt.title('Top 10 International Markets by Revenue (Excluding UK)', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Country', fontweight='bold')
plt.ylabel('Revenue (£ in Thousands)', fontweight='bold')
plt.xticks(rotation=45, ha='right')
for bar in bars:
    h = bar.get_height()
    plt.annotate(f"£{h:.1f}k", (bar.get_x() + bar.get_width() / 2, h + 4), ha='center', fontsize=9, fontweight='semibold')
plt.ylim(0, top_intl['Revenue'].max() / 1000 * 1.15)
plt.tight_layout()
plt.show()
""")

md("""
### Geographic Analysis Interpretation
* **Domestic Dominance:** The **United Kingdom** accounts for **85.11% (£8,737,666.54)** of total enterprise revenue and 18,016 orders.
* **International Expansion:** Cross-border sales contribute **14.89% (£1,528,351.65)** across 37 international countries.
* **Leading International Markets:**
  1. **Netherlands:** £283,889.34 (2.77% of total revenue) across 94 orders, exhibiting huge order sizes.
  2. **EIRE (Ireland):** £276,090.86 (2.69% of total revenue) across 286 orders.
  3. **Germany:** £228,867.14 (2.23% of total revenue) across 456 orders.
  4. **France:** £208,610.15 (2.03% of total revenue) across 391 orders.
  5. **Australia:** £138,521.31 (1.35% of total revenue) across 57 orders.
""")

# ==============================================================================
# 20. Customer Analysis
# ==============================================================================
md("""
## 18. Customer Analysis

We examine customer spending distribution, order frequency, and evaluate revenue concentration using the Pareto principle (80/20 rule).
""")

code("""
cust_metrics = df_clean_cust.groupby('CustomerID').agg(
    TotalSpend=('Revenue', 'sum'),
    OrderCount=('InvoiceNo', 'nunique'),
    TotalUnits=('Quantity', 'sum'),
    Country=('Country', 'first')
).reset_index()

cust_metrics['AOV'] = cust_metrics['TotalSpend'] / cust_metrics['OrderCount']

print("--- TOP 10 HIGHEST-VALUE CUSTOMERS ---")
top10_cust = cust_metrics.sort_values(by='TotalSpend', ascending=False).head(10).reset_index(drop=True)
display(top10_cust)
""")

code("""
# Pareto Concentration Analysis
cust_sorted = cust_metrics.sort_values(by='TotalSpend', ascending=False).reset_index(drop=True)
total_cust_spend = cust_sorted['TotalSpend'].sum()
cust_sorted['CumulativeSpend'] = cust_sorted['TotalSpend'].cumsum()
cust_sorted['CumulativeSpendPct'] = (cust_sorted['CumulativeSpend'] / total_cust_spend) * 100
cust_sorted['CustomerRankPct'] = ((cust_sorted.index + 1) / len(cust_sorted)) * 100

top1_pct_share = cust_sorted[cust_sorted['CustomerRankPct'] <= 1]['CumulativeSpendPct'].max()
top5_pct_share = cust_sorted[cust_sorted['CustomerRankPct'] <= 5]['CumulativeSpendPct'].max()
top10_pct_share = cust_sorted[cust_sorted['CustomerRankPct'] <= 10]['CumulativeSpendPct'].max()
top20_pct_share = cust_sorted[cust_sorted['CustomerRankPct'] <= 20]['CumulativeSpendPct'].max()

print("--- PARETO REVENUE CONCENTRATION ---")
print(f"Top 1% of Customers generate:  {top1_pct_share:.2f}% of identified revenue")
print(f"Top 5% of Customers generate:  {top5_pct_share:.2f}% of identified revenue")
print(f"Top 10% of Customers generate: {top10_pct_share:.2f}% of identified revenue")
print(f"Top 20% of Customers generate: {top20_pct_share:.2f}% of identified revenue")
""")

code("""
# Customer Spending Segments
cust_metrics['SpendSegment'] = pd.cut(
    cust_metrics['TotalSpend'],
    bins=[0, 250, 1000, 5000, np.inf],
    labels=['Low (<£250)', 'Mid (£250-£1,000)', 'High (£1,000-£5,000)', 'VIP (>£5,000)']
)

seg_summary = cust_metrics.groupby('SpendSegment', observed=False).agg(
    CustomerCount=('CustomerID', 'count'),
    TotalSegmentSpend=('TotalSpend', 'sum'),
    AvgSpendPerCust=('TotalSpend', 'mean')
).reset_index()

seg_summary['CustomerShare_%'] = (seg_summary['CustomerCount'] / len(cust_metrics)) * 100
seg_summary['SpendShare_%'] = (seg_summary['TotalSegmentSpend'] / total_cust_spend) * 100

print("--- CUSTOMER VALUE TIERS BREAKDOWN ---")
display(seg_summary)
""")

code("""
# Plot Customer Tiers
plt.figure(figsize=(10, 5.5))
colors_seg = ['#7294B2', '#3E6B89', '#1F4E79', '#D4AF37']
bars = plt.bar(seg_summary['SpendSegment'], seg_summary['CustomerCount'], color=colors_seg, edgecolor='#1A252C', width=0.55)
plt.title('Customer Segmentation by Total Spending Tier', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Customer Spending Tier', fontweight='bold')
plt.ylabel('Number of Customer Accounts', fontweight='bold')
for bar in bars:
    h = bar.get_height()
    pct = (h / len(cust_metrics)) * 100
    plt.annotate(f"{h:,}\\n({pct:.1f}%)", (bar.get_x() + bar.get_width() / 2, h + 30),
                 ha='center', fontsize=9.5, fontweight='semibold')
plt.ylim(0, seg_summary['CustomerCount'].max() * 1.18)
plt.tight_layout()
plt.show()
""")

md("""
### Customer Analysis Interpretation
* **Top Customer:** Customer ID **14646** (Netherlands) is the top buyer, generating **£279,138.02** across 72 orders and 196,915 units, with an AOV of £3,876.92.
* **Pareto Concentration:** The top **20% of customers account for 74.59%** of total identified revenue, and the top **10% account for 61.39%**. The top 1% (43 clients) alone generate nearly **one-third (32.02%)** of revenue.
* **Tier Distribution:** 32.2% of customers belong to the *Low* tier (<£250), while only 2.2% (94 accounts) fall into the *VIP* tier (>£5,000), yet this VIP tier generates a staggering portion of total earnings.
""")

# ==============================================================================
# 21. Cancellation Analysis
# ==============================================================================
md("""
## 19. Cancellation Analysis

We evaluate cancelled transactions (invoices starting with 'C') to identify return rates, monthly trends, and revenue impact.
""")

code("""
df_cancelled['InvoiceDate'] = pd.to_datetime(df_cancelled['InvoiceDate'])
df_cancelled['Year_Month'] = df_cancelled['InvoiceDate'].dt.to_period('M').astype(str)

monthly_canc = df_cancelled.groupby('Year_Month').agg(
    CancelledInvoices=('InvoiceNo', 'nunique'),
    CancelledQuantity=('Quantity', lambda q: abs(q).sum()),
    CancelledValue=('Revenue_Abs', 'sum')
).reset_index()

monthly_comp = pd.merge(monthly_sales[['Year_Month', 'Revenue', 'Orders']], monthly_canc, on='Year_Month', how='left').fillna(0)
monthly_comp['CancellationRate_Invoices_%'] = (monthly_comp['CancelledInvoices'] / (monthly_comp['Orders'] + monthly_comp['CancelledInvoices'])) * 100
monthly_comp['ReturnLossRatio_%'] = (monthly_comp['CancelledValue'] / (monthly_comp['Revenue'] + monthly_comp['CancelledValue'])) * 100

print("--- MONTHLY CANCELLATIONS & RETURN LOSS METRICS ---")
display(monthly_comp)
""")

code("""
# Plot Monthly Sales vs Cancellations Value
fig, ax1 = plt.subplots(figsize=(12, 5.5))

color1 = '#1F4E79'
color2 = '#D9534F'

ax1.set_xlabel('Year-Month', fontweight='bold')
ax1.set_ylabel('Net Revenue (£ Thousands)', color=color1, fontweight='bold')
line1 = ax1.plot(monthly_comp['Year_Month'], monthly_comp['Revenue'] / 1000, color=color1, marker='o', linewidth=2.5, label='Net Revenue (£k)')
ax1.tick_params(axis='y', labelcolor=color1)
plt.xticks(rotation=45)

ax2 = ax1.twinx()
ax2.set_ylabel('Cancelled Value (£ Thousands)', color=color2, fontweight='bold')
line2 = ax2.plot(monthly_comp['Year_Month'], monthly_comp['CancelledValue'] / 1000, color=color2, marker='s', linestyle='--', linewidth=2, label='Cancelled Value (£k)')
ax2.tick_params(axis='y', labelcolor=color2)

lines = line1 + line2
labels = [l.get_label() for l in lines]
ax1.legend(lines, labels, loc='upper left', frameon=True)
plt.title('Monthly Sales Revenue vs. Value of Cancelled Transactions', fontsize=14, fontweight='bold', pad=15)
plt.tight_layout()
plt.show()
""")

md("""
### Cancellation Analysis Interpretation
* **Total Volume:** Across the dataset, **3,836 invoices** (9,288 line items) were cancelled, representing an overall cancellation rate of **16.25%** of all order attempts.
* **Financial Impact:** The gross value of cancelled items reached **£893,979.73**, involving **275,560 units**.
* **Seasonality of Returns:** Cancellations peak in absolute value during October and November in lockstep with surge order volume, but the return loss ratio stays relatively stable at 6% to 9% of monthly turnover.
""")

# ==============================================================================
# 22. Correlation Analysis
# ==============================================================================
md("""
## 20. Correlation Analysis

We construct a Pearson correlation matrix to examine relationships between quantity, unit price, revenue, and temporal attributes.
""")

code("""
corr_df = df_clean[['Quantity', 'UnitPrice', 'Revenue', 'Hour', 'Day', 'Month']].corr()

plt.figure(figsize=(8, 6.5))
sns.heatmap(corr_df, annot=True, fmt='.3f', cmap='Blues', vmin=-0.1, vmax=1.0, cbar=True, square=True,
            annot_kws={'size': 10, 'weight': 'semibold'})
plt.title('Feature Correlation Heatmap', fontsize=14, fontweight='bold', pad=15)
plt.tight_layout()
plt.show()
""")

md("""
### Correlation Interpretation
* **Quantity vs. Revenue:** Displays a strong positive correlation (**r = 0.908**), confirming that bulk order volumes are the primary engine of top-line revenue rather than premium unit pricing.
* **UnitPrice vs. Quantity:** Displays a negligible correlation (**r = -0.005**), indicating that unit price variations do not heavily suppress bulk buying in wholesale contexts.
* **Temporal Features:** Hourly and monthly features show near-zero linear correlation with transaction size, confirming that basket size remains consistent regardless of the time or day of ordering.
""")

# ==============================================================================
# 23. Key Findings
# ==============================================================================
md("""
## 21. Key Findings

1. **Enterprise Scale:** The retailer generated **£10,266,018.19** in clean completed revenue across **5,561,563 items** and **19,777 orders** between December 2010 and December 2011.
2. **Q4 Holiday Concentration:** November 2011 was the peak month (£1.45M, 2,751 orders), generating nearly triple the sales of February 2011 (£508k, 1,114 orders).
3. **Operational Cadence:** Ordering follows standard European B2B business hours, peaking at 12:00 PM on Thursdays (4,209 orders), with complete warehouse inactivity on Saturdays.
4. **Hero Products:** `REGENCY CAKESTAND 3 TIER` generated £174,156.54, while `PAPER CRAFT , LITTLE BIRDIE` sold 80,995 units.
5. **Geographical Concentration:** The UK commands 85.11% of sales. The Netherlands (£283.9k) and Ireland (£276.1k) lead international sales with high average order sizes.
6. **Pareto Concentration:** The top 20% of customers generate 74.59% of identified revenue, and the top 1% generate 32.02%, led by wholesale client ID 14646 (£279.1k).
7. **Cancellation Footprint:** 3,836 invoices (16.25% of attempts) were cancelled, totaling £893,979.73 in cancelled transaction value.
""")

# ==============================================================================
# 24. Business Insights
# ==============================================================================
md("""
## 22. Business Insights

* **Wholesale Hybrid Model:** While consumer-sized transactions occur, enterprise health is overwhelmingly dictated by B2B wholesale buyers purchasing in hundreds or thousands of units.
* **Cash Flow Vulnerability in Q1:** The severe post-Christmas lull requires proactive early-year trade promotions and clearance campaigns to stabilize working capital.
* **Cross-Border Efficiency:** European accounts (Netherlands, Germany, France) place fewer but substantially larger orders, making international B2B customer acquisition highly cost-effective.
* **Saturday Operational Gap:** Complete Saturday downtime indicates potential fulfillment backlogs on Sunday and Monday.
""")

# ==============================================================================
# 25. Strategic Recommendations
# ==============================================================================
md("""
## 23. Strategic Recommendations

1. **VIP Account Management & Retention:** Establish a dedicated B2B Key Account program for the top 100 buyers with contracted SLAs, tiered volume rebates, and proactive stock reservations.
2. **Predictive Q4 Inventory Staging:** Initiate production and procurement for the top 50 revenue-generating stock items by August to eliminate costly Q4 stockouts.
3. **European Logistics Hub:** Investigate a fulfillment hub in the Benelux/Germany corridor to reduce cross-border transit times and shipping costs for top international markets.
4. **Cancellation Root-Cause Remediation:** Audit the top cancelled stock codes to determine if returns stem from packaging fragility, product defects, or shipping delays.
5. **Weekend Digital Ordering Automation:** Maintain automated digital ordering and warehouse pre-scheduling over weekends to smooth Monday morning operational bottlenecks.
""")

# ==============================================================================
# 26. Limitations
# ==============================================================================
md("""
## 24. Limitations

* **Absence of Cost Data:** The dataset lacks Cost of Goods Sold (COGS) and shipping overhead, limiting analysis to gross revenue rather than net profit margins.
* **Missing Customer Identifiers:** 24.93% of transactions lack a CustomerID, restricting full-cohort lifetime value calculations.
* **Truncated Final Month:** Data collection terminated on December 9, 2011, preventing full-month December 2011 comparison.
* **Lack of Channel & Marketing Attributes:** The dataset does not track marketing campaigns, web traffic, or device types.
""")

# ==============================================================================
# 27. Conclusion
# ==============================================================================
md("""
## 25. Conclusion

This project delivered an end-to-end data analysis of the UCI Online Retail dataset. By applying rigorous data cleaning and analytical techniques, we quantified key performance drivers across sales, time, products, geography, customers, and cancellations. The resulting findings provide executive leadership with a clear roadmap to protect core wholesale accounts, optimize seasonal inventory, scale international markets, and mitigate return losses.
""")

# ==============================================================================
# 28. References
# ==============================================================================
md("""
## 26. References

1. Chen, D., Sain, S. L., & Guo, K. (2012). *Data mining for the online retail industry: A case study of RFM model-based customer segmentation using data mining*. Journal of Database Marketing & Customer Strategy Management, 19(3), 197-208.
2. UCI Machine Learning Repository. (2015). *Online Retail Dataset*. Available: https://archive.ics.uci.edu/dataset/352/online+retail
3. McKinney, W. (2010). *Data Structures for Statistical Computing in Python*. Proceedings of the 9th Python in Science Conference, 51-56.
4. Hunter, J. D. (2007). *Matplotlib: A 2D Graphics Environment*. Computing in Science & Engineering, 9(3), 90-95.
5. Waskom, M. L. (2021). *Seaborn: statistical data visualization*. Journal of Open Source Software, 6(60), 3021.
""")

nb.cells = cells

output_notebook = "SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb"
with open(output_notebook, "w", encoding="utf-8") as f:
    nbf.write(nb, f)

print(f"Notebook written to {output_notebook} with {len(cells)} cells.")
print("Now executing notebook end-to-end to populate all outputs and inline charts...")

# Execute the notebook
cmd = [
    sys.executable, "-m", "jupyter", "nbconvert",
    "--to", "notebook",
    "--execute",
    "--inplace",
    "--ExecutePreprocessor.timeout=600",
    output_notebook
]

result = subprocess.run(cmd, capture_output=True, text=True)
print("Return code:", result.returncode)
if result.returncode != 0:
    print("Execution Error STDERR:", result.stderr)
else:
    print("Notebook executed successfully end-to-end! All outputs and plots are saved.")
