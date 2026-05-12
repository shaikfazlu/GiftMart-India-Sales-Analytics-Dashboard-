# 🎁 GiftMart India — Sales & Customer Analytics Dashboard

> **End-to-end Excel analytics project** on a real-world Indian e-commerce gifting dataset — covering 1,000 orders, 100 customers, 70 products, and 7 festive occasions across 12 months.

---

## 📌 Project Overview

This project simulates a **business intelligence report** for a mid-sized Indian gift e-commerce company. Starting from raw transactional CSV data, I performed full-cycle data analysis — cleaning, enrichment, KPI calculation, and executive dashboard design — entirely in Excel with Python (openpyxl + pandas).

The goal: give a **non-technical stakeholder** a single-file Excel workbook they can open and immediately understand sales performance, customer trends, and product insights.

---

## 📊 Key Business Insights

| Metric | Value |
|---|---|
| 💰 Total Revenue | ₹35.2 Lakh (₹3.52M) |
| 📦 Total Orders | 1,000 |
| 🛒 Average Order Value | ₹3,521 |
| 🚚 Average Delivery Time | 5.5 days |
| 👥 Unique Customers | 100 |
| 🏷️ Products | 70 SKUs across 8 categories |
| 📅 Peak Month | **August** (Raksha Bandhan season) |
| 🎉 Top Occasion | **Anniversary** (highest revenue) |
| 🏙️ Top City | **Ghaziabad** |
| ⏰ Peak Order Day | **Tuesday & Sunday** |

---

## 🗂️ Workbook Structure (7 Sheets)

| Sheet | Description |
|---|---|
| 📊 Dashboard | Executive KPI cards, occasion/category tables, city rankings |
| 📋 Orders Data | 1,000 cleaned orders with revenue column, data bars |
| 🛍️ Products | 70 products with price, category, occasion mapping |
| 📈 Monthly Analysis | Month-wise revenue + orders with bar & line charts |
| 🎉 Occasion Analysis | Occasion revenue share, pie chart, top 10 products |
| 👥 Customers | 100 customers with order history & revenue ranking |
| 📌 Summary Metrics | Quick-reference KPI table for reporting |

---

## 🛠️ Tech Stack

- **Python 3.12** — data wrangling & automation
- **pandas** — aggregations, groupby, merges
- **openpyxl** — Excel workbook generation, charts, conditional formatting
- **Excel Features Used** — Data bars, color scales, bar/line/pie charts, freeze panes, tab colors, conditional formatting, merged cells

---

## 📁 Dataset Description

| File | Records | Columns |
|---|---|---|
| `customers.csv` | 100 rows | ID, Name, City, Gender, Contact, Email, Address |
| `orders.csv` | 1,000 rows | Order ID, Customer, Product, Qty, Dates, Location, Occasion, Hour |
| `products.csv` | 70 rows | Product ID, Name, Category, Price, Occasion, Description |

**Occasions covered:** Valentine's Day · Holi · Anniversary · Birthday · Raksha Bandhan · Diwali · All Occasions

---

## 🚀 How to Run

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/giftmart-india-dashboard.git
cd giftmart-india-dashboard

# Install dependencies
pip install pandas openpyxl

# Generate the dashboard
python build_dashboard.py
# → Output: GiftMart_India_Sales_Dashboard.xlsx
```

---

## 📈 Analysis Highlights

### Revenue by Occasion
Anniversary orders drove the highest revenue, closely followed by Holi and Birthday. Diwali had the highest average order value per transaction.

### Seasonal Trend
August (Raksha Bandhan season) peaked in both order volume and revenue. February (Valentine's Day) was the second strongest month.

### Product Categories
**Colors** and **Sweets** together contribute over 50% of total revenue, making them the core product lines. Cakes have the highest average delivery days, suggesting supply chain opportunities.

### Customer Behavior
- Orders spike on **Tuesdays and Sundays**
- Peak ordering hours are **evenings (7–10 PM)**
- Average delivery time of **5.5 days** with room for optimization

---

## 🎯 Business Recommendations

1. **Double inventory for Colors & Sweets** ahead of August and February
2. **Target Ghaziabad, Haridwar, Imphal** with city-specific campaigns — these cities have the highest order density
3. **Run promotions on Mondays/Thursdays** to balance the Tuesday/Sunday order spikes
4. **Reduce delivery time for Cake category** — currently above average
5. **Upsell during Valentine's Day** — high traffic but lower average order value vs. Anniversary

---

## 📂 File Structure

```
giftmart-india-dashboard/
│
├── data/
│   ├── customers.csv
│   ├── orders.csv
│   └── products.csv
│
├── build_dashboard.py          # Main Python script
├── GiftMart_India_Sales_Dashboard.xlsx  # Output workbook
└── README.md
```

---

## 🙋 About

Built as a portfolio project to demonstrate **data analysis, business intelligence, and Excel automation** skills using Python.

**Skills demonstrated:** Data Cleaning · EDA · KPI Design · Dashboard Design · Excel Automation · Python (pandas, openpyxl) · Business Storytelling

---

*⭐ Star this repo if you found it useful!*
