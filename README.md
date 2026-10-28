<div align="center">

# 📊 Sales & Profit Interactive Dashboard

### An end-to-end Excel analytics project: from raw data to a one-click interactive dashboard

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Rows](https://img.shields.io/badge/Records-9,994-blue?style=for-the-badge)
![Years](https://img.shields.io/badge/Period-2014--2017-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Complete-success?style=for-the-badge)

**Built by Bhavika**

</div>

---

## 🎯 Project Overview

This project turns a **9,994-row retail sales dataset** into a fully interactive Excel dashboard. Pick a **Region, Category, Segment or Year** from a dropdown and every KPI and chart refreshes instantly. No macros, no VBA, only formulas.

It covers the complete analyst workflow: **data cleaning → pivot analysis → advanced formulas → visualization → interactive dashboard**.

---

## 📸 Dashboard Preview

> 💡 *Add a screenshot of your Dashboard sheet here:*
>
> `![Dashboard Preview](images/dashboard.png)`

---

## ✨ Key Features

| Feature | Description |
|---|---|
| 🎛️ **Dynamic Filters** | 4 dropdown slicers (Region, Category, Segment, Year) with an "All" option |
| 📈 **Live KPI Cards** | Total Sales, Total Profit, Profit Margin, Quantity Sold, Avg Discount |
| 📊 **5 Interactive Charts** | Category, Region, Segment, Sub-Category and Year trend |
| 🔍 **Lookup Tool** | Search any Row ID or Customer ID and get instant details |
| 🏷️ **Smart Columns** | Auto-labelled Profit/Loss and Discount Level for every order line |
| 🎨 **Conditional Formatting** | Colour cues that highlight strong and weak performers |

---

## 🗂️ Workbook Structure

```
📘 Bhavika_ExcelDashboard.xlsx
│
├── 📄 Raw_Data       → Original untouched dataset (9,994 rows × 21 columns)
├── 🧹 Clean_Data     → Cleaned Excel Table + 4 calculated columns
├── 🧮 Pivot_Tables   → 5 summary pivots (Category, Region, Segment, Sub-Category, Year)
├── ⚙️ Formulas       → VLOOKUP, INDEX-MATCH, nested IF, SUMIF, COUNTIF
├── 📊 Charts         → Chart source tables + 5 charts + conditional formatting
└── 🖥️ Dashboard      → Final interactive dashboard
```

---

## 🛠️ Skills & Excel Techniques Used

- **Data Cleaning**: structured Excel Tables, correct data types, helper columns
- **Calculated Columns**: `YEAR()`, `MONTH()`, `IF()`, nested `IF()`
- **Lookups**: `VLOOKUP` vs `INDEX-MATCH` comparison (including left-lookup, which VLOOKUP cannot do)
- **Conditional Aggregation**: `SUMIFS`, `COUNTIF`, `SUMIF`, `AVERAGEIFS`
- **Named Ranges**: `rng_Sales`, `rng_Region`, `rng_Category` and more, for readable formulas
- **Pivot Tables**: multi-dimension summaries of sales and profit
- **Data Validation**: dropdown lists with error alerts
- **Charts**: clustered bar, pie and line charts
- **Dashboard Design**: KPI cards, filter panel, wildcard-criteria logic (`"*"` for "All")

---

## 🧠 How the Dashboard Works

The filters use a **wildcard trick** that avoids macros entirely:

```excel
Region criteria   =IF($B$4="All","*",$B$4)

Total Sales       =SUMIFS(rng_Sales, rng_Region,   $S$4,
                                     rng_Category, $S$5,
                                     rng_Segment,  $S$6,
                                     rng_Year, ">="&$S$7,
                                     rng_Year, "<="&$S$8)
```

Choosing **"All"** turns the criterion into `*` (match everything). Choosing a specific value filters to just that value. Every KPI and chart reads from this logic, so the whole dashboard stays in sync.

---

## 📌 Key Insights from the Data

**Headline numbers**

| Total Sales | Total Profit | Profit Margin | Quantity Sold | Customers |
|:---:|:---:|:---:|:---:|:---:|
| **$2.30M** | **$286K** | **12.5%** | **37,873** | **793** |

**What the analysis reveals**

- 📈 **Steady growth**: sales rose from **$484K (2014)** to **$733K (2017)**, and profit nearly doubled from **$49.5K to $93.4K**.
- 💻 **Technology** leads in both sales ($836K) and profit ($145K) with a **17.4%** margin.
- 🪑 **Furniture** sells a lot ($742K) but earns only **2.5%** margin, the weakest category.
- 🌎 **West** is the top region for sales and profit (**14.9%** margin), while **Central** lags at **7.9%**.
- 🔻 **Tables, Bookcases and Supplies are loss-making**. Tables alone lose about **$17.7K**.
- ⭐ **Copiers, Paper, Labels and Envelopes** have the highest margins (roughly **37% to 44%**).
- 💸 **Heavy discounts hurt**: orders with no or low discount are profitable, while *Medium* and *High* discount lines produce net losses (about **-$36K** and **-$100K**).
- 👥 **Consumer** is the biggest segment by sales, while **Home Office** has the best margin (14.0%).

---

## 🚀 How to Use

1. **Download** `Bhavika_ExcelDashboard.xlsx`
2. Open it in **Microsoft Excel** (desktop version recommended)
3. Go to the **`Dashboard`** sheet
4. Use the **yellow dropdown cells** to filter by Region, Category, Segment or Year
5. Watch the KPIs and charts update instantly
6. Visit the **`Formulas`** sheet to try the lookup tool: change the yellow input cells (Row ID `1–9994` or a Customer ID such as `CG-12520`)

> ⚠️ **Note:** Open in Excel (not a web previewer) so formulas and charts calculate correctly. Please don't edit the "Chart Data & Filter Logic" block on the right side of the Dashboard.

---

## 📁 Dataset Description

| Column Group | Fields |
|---|---|
| **Order** | Row ID, Order ID, Order Date, Ship Date, Ship Mode |
| **Customer** | Customer ID, Customer Name, Segment |
| **Location** | Country, City, State, Postal Code, Region |
| **Product** | Product ID, Category, Sub-Category, Product Name |
| **Metrics** | Sales, Quantity, Discount, Profit |
| **Added in Clean_Data** | Year, Month, Profit Status, Discount Level |

---

## 🔮 Future Improvements

- [ ] Add a **Month / Quarter** filter and seasonality view
- [ ] Add a **State-level map** chart
- [ ] Add a **Top 10 Customers / Products** leaderboard
- [ ] Rebuild the same dashboard in **Power BI** or **Tableau** for comparison

---

## 👩‍💻 Author

**Bhavika**
Aspiring Data Analyst | Excel • Data Visualization • Business Insights

📫 *Add your LinkedIn / GitHub / Email here*

---

<div align="center">

⭐ **If you found this project useful, consider giving it a star!** ⭐

</div>
