<div align="center">

# 📊 Superstore Sales — Interactive Excel Dashboard

### *From raw CSV to a clean, formula-driven, fully interactive dashboard — zero formula errors.*

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Dataset](https://img.shields.io/badge/Dataset-Kaggle%20Superstore-20BEFF?style=for-the-badge&logo=kaggle&logoColor=white)
![Formulas](https://img.shields.io/badge/Formulas-85%2C000%2B-1F3A5F?style=for-the-badge)
![Errors](https://img.shields.io/badge/Formula%20Errors-0-2A9D8F?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge)

</div>

---

## 📑 Table of Contents

- [🎯 Project Overview](#-project-overview)
- [🖼️ Dashboard Preview](#️-dashboard-preview)
- [✨ Key Features](#-key-features)
- [📈 Key Results](#-key-results)
- [🗂️ Workbook Structure](#️-workbook-structure)
- [🧹 Data Cleaning](#-data-cleaning)
- [🧮 Excel Skills Demonstrated](#-excel-skills-demonstrated)
- [🚀 How to Use](#-how-to-use)
- [📁 Repository Structure](#-repository-structure)
- [⚠️ Notes & Assumptions](#️-notes--assumptions)
- [👤 Author](#-author)

---

## 🎯 Project Overview

This project is a **Final Practical Exam** submission that turns the public **Superstore Sales** dataset (Kaggle) into a professional, **interactive Excel dashboard**.

Starting from a raw `train.csv` (9,800 rows), the workbook walks through the complete analytics workflow:

> 📥 Import → 🧹 Clean → 🔄 Summarise → 🧮 Calculate → 📊 Visualise → 🎛️ Make Interactive → 📝 Document

Everything is **formula-driven** — no hard-coded results — so every KPI, chart and table recalculates instantly when a filter changes or new data is added.

| | |
|---|---|
| 📦 **Dataset** | Superstore Sales (Kaggle) — `train.csv` |
| 🗓️ **Period** | Jan 2015 – Dec 2018 |
| 📏 **Size** | 9,800 raw rows → **9,799** clean rows, 18 original columns |
| 🌍 **Scope** | United States — 4 regions, 49 states, 793 customers |
| 🛠️ **Tool** | Microsoft Excel (formulas, named ranges, tables, charts, conditional formatting, data validation) |

---

## 🖼️ Dashboard Preview

> 💡 *Add a screenshot of your dashboard to an `images/` folder and it will show up here.*

<div align="center">
(<img width="567" height="448" alt="Dashboard" src="https://github.com/user-attachments/assets/3550026a-3a40-4b5f-a638-09783bb4617b" />)

</div>

---

## ✨ Key Features

- 🎛️ **5 interactive dropdown filters** — Year, Quarter, Region, Segment, Category
- 💳 **6 live KPI cards** — Total Sales, Orders, Avg Order Value, Line Items, Avg Ship Days, YoY Growth
- 📉 **7 dynamic charts** — monthly trend, category donut, region, segment, ship mode, top sub-categories, top states
- 🏆 **Top 10 Customers** table with in-cell data bars
- 💬 **Auto-generated Key Insights** that rewrite themselves with every filter change
- 🔥 **Conditional formatting** — heat maps, data bars and arrow icon sets
- 🧮 **Advanced formulas** — `VLOOKUP`, `INDEX-MATCH`, nested `IF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `IFERROR`
- 🔁 **Dynamic named ranges** that grow with the data
- ✅ **Built-in reconciliation checks** proving all totals match
- 🚫 **Zero formula errors** across 85,000+ formulas

---

## 📈 Key Results

### 💰 Headline KPIs (all years, no filters)

| Metric | Value |
|---|---:|
| 💵 Total Sales | **$2,261,255** |
| 🧾 Total Orders | **4,922** |
| 🛒 Avg Order Value | **$459.42** |
| 📦 Line Items Sold | **9,799** |
| 🚚 Avg Ship Time | **4.0 days** |
| 📈 YoY Growth (2018 vs 2017) | **▲ 20.3%** |

### 🔍 Insights Uncovered

- 🥇 **West** is the top region with **31.4%** of sales.
- 💻 **Technology** is the top category with **36.6%** of sales.
- 📅 **November** is the peak month (**15.5%** of annual sales) — strong holiday seasonality.
- 📱 **Phones** is the best-selling sub-category (**$327,782**), followed closely by **Chairs**.
- 🌴 **California** and **New York** lead all states in sales.
- 👑 Top customer: **Sean Miller** (**$25,043**).
- 📊 Sales grew every year from 2016 onwards — **$459K (2016) → $722K (2018)**.

---

## 🗂️ Workbook Structure

| # | Sheet | Purpose |
|:-:|---|---|
| 1️⃣ | **Dashboard** | Interactive dashboard — filters, KPI cards, charts, top customers, insights |
| 2️⃣ | **Pivot_Tables** | 7 pivot-style summary tables with heat maps, data bars, icon sets & reconciliation check |
| 3️⃣ | **Formulas** | `VLOOKUP` vs `INDEX-MATCH`, logic tester, nested functions, dynamic ranges, editable thresholds |
| 4️⃣ | **Data_Clean** | Cleaned dataset as an Excel Table (`tblSales`) with 8 helper formula columns |
| 5️⃣ | **Cleaning_Log** | Every cleaning step with live before/after counts and validation checks |
| 6️⃣ | **Data_Raw** | Original `train.csv`, untouched |
| 7️⃣ | **Lookup_Tables** | Customer master (793 customers) powering lookups and tiers |
| 8️⃣ | **Calc_Helper** | Dropdown lists, filter criteria and chart source ranges *(do not edit)* |

---

## 🧹 Data Cleaning

| ✅ Step | Result |
|---|---|
| 🔁 Duplicate rows (identical except `Row ID`) | **1 removed** |
| 🕳️ Missing Postal Codes | **11 filled** (all Burlington, Vermont → `05401`) |
| 📅 Dates stored as text (`dd/mm/yyyy`) | **Converted to real Excel dates** |
| 🔤 Text hygiene | Leading/trailing spaces trimmed |
| 🔢 Postal codes | Stored as 5-digit text (leading zeros preserved) |
| 🔎 Logic checks | Ship date before order date = **0**, Sales ≤ 0 = **0** |
| ➕ Helper columns | Year, Month, Quarter, Ship Days, Ship Speed, Order Size, New Order flags |

All counts are live formulas on the **Cleaning_Log** sheet, comparing `Data_Raw` with `Data_Clean`.

---

## 🧮 Excel Skills Demonstrated

| Area | Techniques |
|---|---|
| 📥 **Import & Clean** | Duplicate removal, missing-value handling, data-type conversion |
| 🔄 **Analysis** | Cross-tab summaries with `SUMIFS` / `COUNTIFS` / `AVERAGEIFS`, % share, YoY growth |
| 🔎 **Lookups** | `VLOOKUP`, `INDEX` + `MATCH` (including left-lookup), `IFERROR` guards |
| 🧠 **Logic** | Nested `IF`, `AND`, `OR`, `RANK`, `LARGE` for dynamic Top-N lists |
| 🏷️ **Named ranges** | Fixed-capacity names (`dSales`, `dRegion`…), filter names (`fYear`…), thresholds (`tLarge`…), dynamic names (`dynSales`) |
| 📊 **Visualization** | Line, column, bar, pie and doughnut charts; heat maps; data bars; icon sets |
| 🎛️ **Interactivity** | Data-validation dropdowns driving every KPI, chart and insight |
| 📝 **Documentation** | Naming conventions, colour-coded inputs, assumptions, reconciliation checks |

---

## 🚀 How to Use

1. 📥 **Download** `Swayam_ExcelDashboard.xlsx` from this repository.
2. 💻 **Open it in desktop Microsoft Excel** and click *Enable Editing*.
3. 🎛️ Go to the **Dashboard** sheet and pick values in the yellow dropdowns (Year, Quarter, Region, Segment, Category).
4. 👀 Watch KPIs, charts, the Top-10 table and the insights update instantly. Choose **All** to clear a filter.
5. 🔧 *(Optional)* Change thresholds (e.g. "Large order ≥ $1,000") in **Formulas → Assumptions** — everything recalculates.
6. ➕ *(Optional)* Add native slicers: click inside `tblSales` → **Table Design → Insert Slicer**.

### 🔄 Adding new data
Paste new rows directly below the last row of `tblSales` on **Data_Clean**. The table, helper columns and all formulas extend automatically (capacity: **20,000 rows**).

---

## 📁 Repository Structure

```text
📦 superstore-excel-dashboard
 ┣ 📂 data
 ┃ ┗ 📄 train.csv                    # Original Kaggle dataset
 ┣ 📂 images
 ┃ ┗ 🖼️ dashboard.png                # Dashboard screenshot
 ┣ 📊 Swayam_ExcelDashboard.xlsx     # Final interactive dashboard workbook
 ┗ 📝 README.md                      # Project documentation
```

---

## ⚠️ Notes & Assumptions

- 🟢 **Use desktop Excel.** Dynamic named ranges are not supported by Google Sheets, so the "Dynamic named ranges" section on the `Formulas` sheet may show errors there.
- 💲 Sales are in **US dollars**. The dataset has no profit or quantity columns, so the analysis is sales-based.
- 🧾 **Orders** = distinct Order IDs. When a Category filter is active, orders are counted per order-category because one order can contain several categories.
- 📌 The summary tables are **formula-driven pivot-style tables** (not native PivotTable objects) — they always reconcile to total sales and never break on refresh.
- 🏠 The 11 blank postal codes (Burlington, VT) were filled with `05401` — a documented assumption.
- 📆 YoY growth compares the selected year with the previous year; with *Year = All* it compares 2018 vs 2017. FY2015 shows **n/a** (no 2014 data).

---

## 👤 Author

<div align="center">

**Swayam**

💼 *Excel • Data Analysis • Dashboards*

⭐ *If you found this project useful, consider giving it a star!* ⭐

</div>

---

<div align="center">

*Dataset: [Superstore Sales — Kaggle](https://www.kaggle.com/) • Built with ❤️ and Excel*

</div>
