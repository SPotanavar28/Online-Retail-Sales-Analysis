import os
import json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Set style for professional publication-ready figures
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['axes.labelsize'] = 12
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10
plt.rcParams['legend.fontsize'] = 10
plt.rcParams['figure.titlesize'] = 16

os.makedirs("charts", exist_ok=True)
data_file = os.path.join("data", "Online Retail.xlsx")

print("1. Loading raw dataset...")
df_raw = pd.read_excel(data_file)
raw_shape = df_raw.shape
print(f"Raw dataset shape: {raw_shape}")

# 2. Raw Inspection Metrics
raw_cols = list(df_raw.columns)
raw_nulls = df_raw.isnull().sum().to_dict()
raw_duplicates = int(df_raw.duplicated().sum())

# Date range raw
min_date_raw = str(df_raw['InvoiceDate'].min())
max_date_raw = str(df_raw['InvoiceDate'].max())

# Unique entities raw
raw_unique_invoices = int(df_raw['InvoiceNo'].nunique())
raw_unique_stockcodes = int(df_raw['StockCode'].nunique())
raw_unique_customers = int(df_raw['CustomerID'].nunique())
raw_unique_countries = int(df_raw['Country'].nunique())

# Convert InvoiceNo to string to handle 'C' prefixes properly
df_raw['InvoiceNo_Str'] = df_raw['InvoiceNo'].astype(str)

# Cancellations in raw
is_cancellation = df_raw['InvoiceNo_Str'].str.startswith('C')
num_cancellations = int(is_cancellation.sum())
cancellation_invoices = int(df_raw.loc[is_cancellation, 'InvoiceNo'].nunique())
cancellation_qty_sum = int(df_raw.loc[is_cancellation, 'Quantity'].sum())

# Negative prices / Zero prices
zero_or_neg_price_count = int((df_raw['UnitPrice'] <= 0).sum())
negative_qty_not_cancelled = int(((df_raw['Quantity'] < 0) & (~is_cancellation)).sum())

print(f"Raw duplicates: {raw_duplicates:,}")
print(f"Cancellations count: {num_cancellations:,} lines across {cancellation_invoices:,} invoices")
print(f"Negative quantities not starting with C: {negative_qty_not_cancelled:,}")
print(f"Zero or negative unit prices: {zero_or_neg_price_count:,}")

# 3. Clean Dataset Construction
# Step A: Remove exact duplicates
df_no_dups = df_raw.drop_duplicates()

# Step B: Separate cancellations for cancellation analysis
df_cancelled = df_no_dups[df_no_dups['InvoiceNo_Str'].str.startswith('C')].copy()
df_cancelled['Revenue_Abs'] = df_cancelled['Quantity'].abs() * df_cancelled['UnitPrice']

# Step C: Valid completed transactions
# Criteria for completed sales transactions:
# - Not cancelled (InvoiceNo does not start with 'C')
# - Quantity > 0
# - UnitPrice > 0
# - StockCode does not contain non-merchandise codes (POST, D, M, BANK CHARGES, AMAZONFEE, CRUK, DOT, SAMPLES)
non_product_codes = ['POST', 'D', 'M', 'BANK CHARGES', 'AMAZONFEE', 'CRUK', 'DOT', 'S']
is_non_product = df_no_dups['StockCode'].astype(str).str.upper().isin(non_product_codes)

df_clean = df_no_dups[
    (~df_no_dups['InvoiceNo_Str'].str.startswith('C')) &
    (df_no_dups['Quantity'] > 0) &
    (df_no_dups['UnitPrice'] > 0) &
    (~is_non_product)
].copy()

# Feature Engineering on clean dataset
df_clean['Revenue'] = df_clean['Quantity'] * df_clean['UnitPrice']
df_clean['InvoiceDate'] = pd.to_datetime(df_clean['InvoiceDate'])
df_clean['Year'] = df_clean['InvoiceDate'].dt.year
df_clean['Month'] = df_clean['InvoiceDate'].dt.month
df_clean['Month_Name'] = df_clean['InvoiceDate'].dt.strftime('%B')
df_clean['Year_Month'] = df_clean['InvoiceDate'].dt.to_period('M').astype(str)
df_clean['Day'] = df_clean['InvoiceDate'].dt.day
df_clean['Day_Of_Week'] = df_clean['InvoiceDate'].dt.dayofweek # 0=Monday, 6=Sunday
df_clean['Day_Name'] = df_clean['InvoiceDate'].dt.day_name()
df_clean['Hour'] = df_clean['InvoiceDate'].dt.hour

clean_shape = df_clean.shape
print(f"Clean sales dataset shape: {clean_shape}")

# Also customer-specific clean cohort (clean records with valid CustomerID)
df_clean_cust = df_clean[df_clean['CustomerID'].notnull()].copy()
df_clean_cust['CustomerID'] = df_clean_cust['CustomerID'].astype(int).astype(str)

# 4. Global KPIs & Metrics
total_revenue = float(df_clean['Revenue'].sum())
total_quantity_sold = int(df_clean['Quantity'].sum())
total_orders = int(df_clean['InvoiceNo'].nunique())
total_unique_products = int(df_clean['StockCode'].nunique())
total_unique_customers_overall = int(df_clean['CustomerID'].nunique())
total_countries = int(df_clean['Country'].nunique())
average_order_value = float(total_revenue / total_orders)
average_unit_price = float(df_clean['UnitPrice'].mean())
average_items_per_order = float(total_quantity_sold / total_orders)

print(f"Total Revenue: £{total_revenue:,.2f}")
print(f"Total Quantity Sold: {total_quantity_sold:,}")
print(f"Total Orders: {total_orders:,}")
print(f"Total Unique Products: {total_unique_products:,}")
print(f"Total Unique Customers (with ID): {total_unique_customers_overall:,}")
print(f"Average Order Value: £{average_order_value:,.2f}")

# 5. Monthly Sales Analysis
monthly_sales = df_clean.groupby('Year_Month').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum'),
    AOV=('Revenue', lambda r: r.sum() / df_clean.loc[r.index, 'InvoiceNo'].nunique())
).reset_index()

# Peak and lowest months
peak_month_rev = monthly_sales.sort_values(by='Revenue', ascending=False).iloc[0]
lowest_month_rev = monthly_sales.sort_values(by='Revenue', ascending=True).iloc[0]

# 6. Time Analysis: Day of Week & Hour
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Sunday'] # Note: Saturday has no transactions in this dataset
dow_sales = df_clean.groupby('Day_Name').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum')
).reindex([d for d in day_order if d in df_clean['Day_Name'].unique()]).reset_index()

hourly_sales = df_clean.groupby('Hour').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum')
).reset_index()

# 7. Product Analysis
product_sales = df_clean.groupby(['StockCode', 'Description']).agg(
    TotalRevenue=('Revenue', 'sum'),
    TotalQuantity=('Quantity', 'sum'),
    OrderCount=('InvoiceNo', 'nunique'),
    AvgUnitPrice=('UnitPrice', 'mean')
).reset_index()

top10_products_revenue = product_sales.sort_values(by='TotalRevenue', ascending=False).head(10)
top10_products_quantity = product_sales.sort_values(by='TotalQuantity', ascending=False).head(10)

# 8. Geographic Analysis
country_sales = df_clean.groupby('Country').agg(
    Revenue=('Revenue', 'sum'),
    Orders=('InvoiceNo', 'nunique'),
    Quantity=('Quantity', 'sum'),
    Customers=('CustomerID', 'nunique')
).reset_index()
country_sales['RevenueShare'] = (country_sales['Revenue'] / total_revenue) * 100
country_sales = country_sales.sort_values(by='Revenue', ascending=False)

uk_revenue = float(country_sales[country_sales['Country'] == 'United Kingdom']['Revenue'].values[0])
uk_share = float(country_sales[country_sales['Country'] == 'United Kingdom']['RevenueShare'].values[0])
non_uk_sales = country_sales[country_sales['Country'] != 'United Kingdom'].copy()
non_uk_revenue = float(non_uk_sales['Revenue'].sum())
top5_international = non_uk_sales.head(5)

# 9. Customer Analysis
customer_metrics = df_clean_cust.groupby('CustomerID').agg(
    TotalSpend=('Revenue', 'sum'),
    OrderCount=('InvoiceNo', 'nunique'),
    TotalUnits=('Quantity', 'sum'),
    FirstPurchase=('InvoiceDate', 'min'),
    LastPurchase=('InvoiceDate', 'max'),
    Country=('Country', 'first')
).reset_index()
customer_metrics['AOV'] = customer_metrics['TotalSpend'] / customer_metrics['OrderCount']

top10_customers_revenue = customer_metrics.sort_values(by='TotalSpend', ascending=False).head(10)

# Pareto / Concentration
customer_metrics_sorted = customer_metrics.sort_values(by='TotalSpend', ascending=False).copy()
customer_metrics_sorted['CumulativeSpend'] = customer_metrics_sorted['TotalSpend'].cumsum()
customer_metrics_sorted['CumulativePct'] = customer_metrics_sorted['CumulativeSpend'] / customer_metrics['TotalSpend'].sum() * 100
customer_metrics_sorted['RankPct'] = np.arange(1, len(customer_metrics_sorted) + 1) / len(customer_metrics_sorted) * 100

top_20pct_customers_rev_share = float(customer_metrics_sorted[customer_metrics_sorted['RankPct'] <= 20]['CumulativePct'].max())
top_10pct_customers_rev_share = float(customer_metrics_sorted[customer_metrics_sorted['RankPct'] <= 10]['CumulativePct'].max())
top_1pct_customers_rev_share = float(customer_metrics_sorted[customer_metrics_sorted['RankPct'] <= 1]['CumulativePct'].max())

# Spending segments
customer_metrics['SpendSegment'] = pd.cut(
    customer_metrics['TotalSpend'],
    bins=[0, 250, 1000, 5000, np.inf],
    labels=['Low (<£250)', 'Mid (£250-£1,000)', 'High (£1,000-£5,000)', 'VIP (>£5,000)']
)
spend_segment_dist = customer_metrics['SpendSegment'].value_counts().reset_index()

# 10. Cancellation Analysis
total_cancelled_invoices = int(df_cancelled['InvoiceNo_Str'].nunique())
total_cancelled_items_abs = int(df_cancelled['Quantity'].abs().sum())
total_cancelled_value = float(df_cancelled['Revenue_Abs'].sum())
cancellation_rate_invoices = float((total_cancelled_invoices / (total_orders + total_cancelled_invoices)) * 100)

df_cancelled['InvoiceDate'] = pd.to_datetime(df_cancelled['InvoiceDate'])
df_cancelled['Year_Month'] = df_cancelled['InvoiceDate'].dt.to_period('M').astype(str)
monthly_cancellations = df_cancelled.groupby('Year_Month').agg(
    CancelledInvoices=('InvoiceNo', 'nunique'),
    CancelledQuantity=('Quantity', lambda x: abs(x).sum()),
    CancelledValue=('Revenue_Abs', 'sum')
).reset_index()

# Merge monthly sales with monthly cancellations for return rate
monthly_combined = pd.merge(monthly_sales, monthly_cancellations, on='Year_Month', how='left').fillna(0)
monthly_combined['ReturnRateValuePct'] = (monthly_combined['CancelledValue'] / (monthly_combined['Revenue'] + monthly_combined['CancelledValue'])) * 100

# Top cancelled products
top_cancelled_products = df_cancelled.groupby(['StockCode', 'Description']).agg(
    CancelledQty=('Quantity', lambda q: abs(q).sum()),
    CancelledValue=('Revenue_Abs', 'sum'),
    CancelCount=('InvoiceNo', 'count')
).reset_index().sort_values(by='CancelledQty', ascending=False).head(10)

# 11. GENERATE PROFESSIONAL VISUALIZATIONS

# Palette: Slate, Navy, Royal Blue, Coral, Emerald
PRIMARY = '#1F4E79'
SECONDARY = '#2E75B6'
ACCENT = '#D9534F'
TEAL = '#008080'
GOLD = '#D4AF37'
SLATE = '#595959'

# Chart 1: Monthly Revenue Trend
fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(monthly_sales['Year_Month'], monthly_sales['Revenue'] / 1000, marker='o', linewidth=2.5, color=PRIMARY, markersize=7)
ax.fill_between(monthly_sales['Year_Month'], monthly_sales['Revenue'] / 1000, color=SECONDARY, alpha=0.15)
ax.set_title('Monthly Revenue Trend (£ Thousands)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Year-Month', fontweight='bold')
ax.set_ylabel('Revenue (£ in 000s)', fontweight='bold')
plt.xticks(rotation=45)
for i, txt in enumerate(monthly_sales['Revenue']):
    ax.annotate(f"£{txt/1000:.1f}k", (monthly_sales['Year_Month'][i], monthly_sales['Revenue'][i]/1000 + 15),
                ha='center', fontsize=9, fontweight='semibold', color='#333333')
ax.set_ylim(0, (monthly_sales['Revenue'].max() / 1000) * 1.15)
plt.tight_layout()
plt.savefig('charts/01_monthly_revenue_trend.png', dpi=300)
plt.close()

# Chart 2: Monthly Orders Trend
fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.bar(monthly_sales['Year_Month'], monthly_sales['Orders'], color=SECONDARY, edgecolor='#1B365D', width=0.6, alpha=0.9)
ax.set_title('Monthly Total Order Volume (Invoices)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Year-Month', fontweight='bold')
ax.set_ylabel('Number of Invoices', fontweight='bold')
plt.xticks(rotation=45)
for bar in bars:
    height = bar.get_height()
    ax.annotate(f"{int(height):,}", (bar.get_x() + bar.get_width() / 2, height + 40),
                ha='center', fontsize=9, fontweight='semibold', color='#333333')
ax.set_ylim(0, monthly_sales['Orders'].max() * 1.15)
plt.tight_layout()
plt.savefig('charts/02_monthly_orders_trend.png', dpi=300)
plt.close()

# Chart 3: Top 10 Products by Revenue
fig, ax = plt.subplots(figsize=(12, 7))
# Shorten description
top10_rev = top10_products_revenue.copy()
top10_rev['ShortDesc'] = top10_rev['Description'].str[:35]
bars = ax.barh(top10_rev['ShortDesc'][::-1], top10_rev['TotalRevenue'][::-1] / 1000, color=PRIMARY, edgecolor='#0D233A', height=0.65)
ax.set_title('Top 10 Products by Total Revenue (£ Thousands)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Total Revenue (£ in 000s)', fontweight='bold')
ax.set_ylabel('Product Description', fontweight='bold')
for bar in bars:
    width = bar.get_width()
    ax.annotate(f" £{width:.1f}k", (width, bar.get_y() + bar.get_height() / 2),
                va='center', fontsize=9.5, fontweight='semibold')
ax.set_xlim(0, (top10_rev['TotalRevenue'].max() / 1000) * 1.18)
plt.tight_layout()
plt.savefig('charts/03_top_products_revenue.png', dpi=300)
plt.close()

# Chart 4: Top 10 Products by Quantity Sold
fig, ax = plt.subplots(figsize=(12, 7))
top10_qty = top10_products_quantity.copy()
top10_qty['ShortDesc'] = top10_qty['Description'].str[:35]
bars = ax.barh(top10_qty['ShortDesc'][::-1], top10_qty['TotalQuantity'][::-1] / 1000, color=TEAL, edgecolor='#004D40', height=0.65)
ax.set_title('Top 10 Products by Total Quantity Sold (Thousands of Units)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Total Quantity Sold (in 000s)', fontweight='bold')
ax.set_ylabel('Product Description', fontweight='bold')
for bar in bars:
    width = bar.get_width()
    ax.annotate(f" {width:.1f}k units", (width, bar.get_y() + bar.get_height() / 2),
                va='center', fontsize=9.5, fontweight='semibold')
ax.set_xlim(0, (top10_qty['TotalQuantity'].max() / 1000) * 1.18)
plt.tight_layout()
plt.savefig('charts/04_top_products_quantity.png', dpi=300)
plt.close()

# Chart 5: Revenue by Top 10 Countries
top10_countries = country_sales.head(10).copy()
fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.bar(top10_countries['Country'], top10_countries['Revenue'] / 1000, color=PRIMARY, edgecolor='#0D233A', width=0.6)
ax.set_title('Top 10 Countries by Total Revenue (£ Thousands)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Country', fontweight='bold')
ax.set_ylabel('Revenue (£ in 000s)', fontweight='bold')
plt.xticks(rotation=45, ha='right')
for bar in bars:
    height = bar.get_height()
    ax.annotate(f"£{height:.0f}k", (bar.get_x() + bar.get_width() / 2, height + 80),
                ha='center', fontsize=8.5, fontweight='semibold', rotation=45)
ax.set_ylim(0, (top10_countries['Revenue'].max() / 1000) * 1.15)
plt.tight_layout()
plt.savefig('charts/05_revenue_by_country.png', dpi=300)
plt.close()

# Chart 6: International Sales (Excluding United Kingdom)
top10_intl = non_uk_sales.head(10).copy()
fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.bar(top10_intl['Country'], top10_intl['Revenue'] / 1000, color=SECONDARY, edgecolor='#1B365D', width=0.6)
ax.set_title('Top 10 International Markets by Revenue (Excluding UK)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Country', fontweight='bold')
ax.set_ylabel('Revenue (£ in 000s)', fontweight='bold')
plt.xticks(rotation=45, ha='right')
for bar in bars:
    height = bar.get_height()
    ax.annotate(f"£{height:.1f}k", (bar.get_x() + bar.get_width() / 2, height + 4),
                ha='center', fontsize=9, fontweight='semibold')
ax.set_ylim(0, (top10_intl['Revenue'].max() / 1000) * 1.15)
plt.tight_layout()
plt.savefig('charts/06_international_revenue_ex_uk.png', dpi=300)
plt.close()

# Chart 7: Orders by Day of Week
fig, ax = plt.subplots(figsize=(10, 5.5))
bars = ax.bar(dow_sales['Day_Name'], dow_sales['Orders'], color='#3B7A57', edgecolor='#1E4620', width=0.55)
ax.set_title('Order Volume by Day of Week', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Day of Week', fontweight='bold')
ax.set_ylabel('Total Invoices Placed', fontweight='bold')
for bar in bars:
    height = bar.get_height()
    ax.annotate(f"{height:,}", (bar.get_x() + bar.get_width() / 2, height + 80),
                ha='center', fontsize=9.5, fontweight='semibold')
ax.set_ylim(0, dow_sales['Orders'].max() * 1.15)
plt.tight_layout()
plt.savefig('charts/07_orders_by_day_of_week.png', dpi=300)
plt.close()

# Chart 8: Orders by Hour of Day
fig, ax = plt.subplots(figsize=(12, 5.5))
ax.plot(hourly_sales['Hour'], hourly_sales['Orders'], marker='s', color='#8B0000', linewidth=2.5, markersize=7)
ax.fill_between(hourly_sales['Hour'], hourly_sales['Orders'], color='#E57373', alpha=0.25)
ax.set_title('Transaction Volume Across Operating Hours of the Day', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Hour of Day (24-Hour Clock)', fontweight='bold')
ax.set_ylabel('Number of Invoices Placed', fontweight='bold')
ax.set_xticks(hourly_sales['Hour'])
for i, txt in enumerate(hourly_sales['Orders']):
    ax.annotate(f"{txt:,}", (hourly_sales['Hour'][i], hourly_sales['Orders'][i] + 70),
                ha='center', fontsize=8.5, fontweight='semibold')
ax.set_ylim(0, hourly_sales['Orders'].max() * 1.15)
plt.tight_layout()
plt.savefig('charts/08_orders_by_hour.png', dpi=300)
plt.close()

# Chart 9: Customer Spending Distribution (Segments)
fig, ax = plt.subplots(figsize=(10, 6))
seg_counts = customer_metrics['SpendSegment'].value_counts()[['Low (<£250)', 'Mid (£250-£1,000)', 'High (£1,000-£5,000)', 'VIP (>£5,000)']]
colors_seg = ['#7294B2', '#3E6B89', '#1F4E79', '#D4AF37']
bars = ax.bar(seg_counts.index, seg_counts.values, color=colors_seg, edgecolor='#1A252C', width=0.55)
ax.set_title('Customer Segmentation by Total Spending Tier', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Customer Spending Tier', fontweight='bold')
ax.set_ylabel('Number of Customers', fontweight='bold')
for bar in bars:
    height = bar.get_height()
    pct = (height / len(customer_metrics)) * 100
    ax.annotate(f"{height:,}\n({pct:.1f}%)", (bar.get_x() + bar.get_width() / 2, height + 40),
                ha='center', fontsize=9.5, fontweight='semibold')
ax.set_ylim(0, seg_counts.max() * 1.18)
plt.tight_layout()
plt.savefig('charts/09_customer_spending_distribution.png', dpi=300)
plt.close()

# Chart 10: Top 10 Customers by Revenue
fig, ax = plt.subplots(figsize=(12, 6))
top10_cust = top10_customers_revenue.copy()
top10_cust['Label'] = 'ID ' + top10_cust['CustomerID'] + ' (' + top10_cust['Country'] + ')'
bars = ax.barh(top10_cust['Label'][::-1], top10_cust['TotalSpend'][::-1] / 1000, color=PRIMARY, edgecolor='#0D233A', height=0.65)
ax.set_title('Top 10 High-Value Customers by Lifetime Spend (£ Thousands)', fontsize=15, fontweight='bold', pad=15)
ax.set_xlabel('Total Spend (£ in 000s)', fontweight='bold')
ax.set_ylabel('Customer ID (Country)', fontweight='bold')
for bar in bars:
    width = bar.get_width()
    ax.annotate(f" £{width:.1f}k", (width, bar.get_y() + bar.get_height() / 2),
                va='center', fontsize=9.5, fontweight='semibold')
ax.set_xlim(0, (top10_cust['TotalSpend'].max() / 1000) * 1.18)
plt.tight_layout()
plt.savefig('charts/10_top_customers_revenue.png', dpi=300)
plt.close()

# Chart 11: Monthly Cancellations vs Revenue
fig, ax1 = plt.subplots(figsize=(12, 6))
color1 = PRIMARY
color2 = ACCENT

ax1.set_xlabel('Year-Month', fontweight='bold')
ax1.set_ylabel('Net Revenue (£ Thousands)', color=color1, fontweight='bold')
line1 = ax1.plot(monthly_combined['Year_Month'], monthly_combined['Revenue'] / 1000, color=color1, marker='o', linewidth=2.5, label='Net Revenue (£k)')
ax1.tick_params(axis='y', labelcolor=color1)
plt.xticks(rotation=45)

ax2 = ax1.twinx()
ax2.set_ylabel('Cancelled Value (£ Thousands)', color=color2, fontweight='bold')
line2 = ax2.plot(monthly_combined['Year_Month'], monthly_combined['CancelledValue'] / 1000, color=color2, marker='s', linestyle='--', linewidth=2, label='Cancelled Value (£k)')
ax2.tick_params(axis='y', labelcolor=color2)

lines = line1 + line2
labels = [l.get_label() for l in lines]
ax1.legend(lines, labels, loc='upper left', frameon=True)
plt.title('Monthly Sales Revenue vs. Value of Cancelled Transactions', fontsize=15, fontweight='bold', pad=15)
plt.tight_layout()
plt.savefig('charts/11_cancellation_patterns_monthly.png', dpi=300)
plt.close()

# Chart 12: Correlation Heatmap
corr_data = df_clean[['Quantity', 'UnitPrice', 'Revenue', 'Hour', 'Day', 'Month']].corr()
fig, ax = plt.subplots(figsize=(8, 6.5))
sns.heatmap(corr_data, annot=True, fmt='.3f', cmap='Blues', vmin=-0.1, vmax=1.0, cbar=True, square=True, ax=ax,
            annot_kws={'size': 10, 'weight': 'semibold'})
ax.set_title('Feature Correlation Heatmap', fontsize=14, fontweight='bold', pad=15)
plt.tight_layout()
plt.savefig('charts/12_correlation_heatmap.png', dpi=300)
plt.close()

print("All charts successfully generated and saved to charts/")

# 12. Save all exact calculated metrics to JSON for report & notebook insertion
summary_stats = {
    "student_name": "Sneha Vinod Potanavar",
    "project_title": "Online Retail Sales & Customer Analysis",
    "raw_records": int(raw_shape[0]),
    "raw_columns": int(raw_shape[1]),
    "raw_duplicates": raw_duplicates,
    "raw_null_customer_ids": int(raw_nulls.get('CustomerID', 0)),
    "raw_null_descriptions": int(raw_nulls.get('Description', 0)),
    "date_range_start": min_date_raw,
    "date_range_end": max_date_raw,
    "raw_cancellation_lines": num_cancellations,
    "raw_cancellation_invoices": cancellation_invoices,
    "clean_records": int(clean_shape[0]),
    "clean_columns": int(clean_shape[1]),
    "total_revenue": round(total_revenue, 2),
    "total_quantity_sold": total_quantity_sold,
    "total_orders": total_orders,
    "total_unique_products": total_unique_products,
    "total_unique_customers": total_unique_customers_overall,
    "total_countries": total_countries,
    "average_order_value": round(average_order_value, 2),
    "average_unit_price": round(average_unit_price, 2),
    "average_items_per_order": round(average_items_per_order, 2),
    "uk_revenue": round(uk_revenue, 2),
    "uk_revenue_share_pct": round(uk_share, 2),
    "international_revenue": round(non_uk_revenue, 2),
    "international_revenue_share_pct": round(100 - uk_share, 2),
    "top_intl_country": str(top5_international.iloc[0]['Country']),
    "top_intl_country_rev": round(float(top5_international.iloc[0]['Revenue']), 2),
    "second_intl_country": str(top5_international.iloc[1]['Country']),
    "second_intl_country_rev": round(float(top5_international.iloc[1]['Revenue']), 2),
    "peak_month": str(peak_month_rev['Year_Month']),
    "peak_month_rev": round(float(peak_month_rev['Revenue']), 2),
    "peak_month_orders": int(peak_month_rev['Orders']),
    "lowest_month": str(lowest_month_rev['Year_Month']),
    "lowest_month_rev": round(float(lowest_month_rev['Revenue']), 2),
    "top_day_of_week": str(dow_sales.sort_values(by='Orders', ascending=False).iloc[0]['Day_Name']),
    "top_day_orders": int(dow_sales.sort_values(by='Orders', ascending=False).iloc[0]['Orders']),
    "top_hour": int(hourly_sales.sort_values(by='Orders', ascending=False).iloc[0]['Hour']),
    "top_hour_orders": int(hourly_sales.sort_values(by='Orders', ascending=False).iloc[0]['Orders']),
    "top_product_by_revenue_code": str(top10_products_revenue.iloc[0]['StockCode']),
    "top_product_by_revenue_desc": str(top10_products_revenue.iloc[0]['Description']),
    "top_product_by_revenue_val": round(float(top10_products_revenue.iloc[0]['TotalRevenue']), 2),
    "top_product_by_qty_code": str(top10_products_quantity.iloc[0]['StockCode']),
    "top_product_by_qty_desc": str(top10_products_quantity.iloc[0]['Description']),
    "top_product_by_qty_val": int(top10_products_quantity.iloc[0]['TotalQuantity']),
    "top_customer_id": str(top10_customers_revenue.iloc[0]['CustomerID']),
    "top_customer_rev": round(float(top10_customers_revenue.iloc[0]['TotalSpend']), 2),
    "top_customer_orders": int(top10_customers_revenue.iloc[0]['OrderCount']),
    "top_customer_country": str(top10_customers_revenue.iloc[0]['Country']),
    "pareto_top_20pct_revenue_share": round(top_20pct_customers_rev_share, 2),
    "pareto_top_10pct_revenue_share": round(top_10pct_customers_rev_share, 2),
    "pareto_top_1pct_revenue_share": round(top_1pct_customers_rev_share, 2),
    "total_cancelled_invoices": total_cancelled_invoices,
    "total_cancelled_items": total_cancelled_items_abs,
    "total_cancelled_value": round(total_cancelled_value, 2),
    "cancellation_rate_invoices_pct": round(cancellation_rate_invoices, 2),
}

with open("analysis_summary.json", "w") as f:
    json.dump(summary_stats, f, indent=4)

print("\n--- EXACT SUMMARY METRICS ---")
for k, v in summary_stats.items():
    print(f"{k}: {v}")

print("\nAnalysis execution completed successfully!")
