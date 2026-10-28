# Sales_Profit_Dashboard__

<div align="center">

<img src="https://capsule-render.vercel.app/api?type=waving&color=0:1e1b4b,50:4f46e5,100:a78bfa&height=260&section=header&text=SALES%20%26%20PROFIT&fontSize=52&fontColor=ffffff&stroke=1e1b4b&strokeWidth=1&animation=twinkling&fontAlignY=36&desc=Excel%20Dashboard%20%7C%20Pivots%20%7C%20Lookups%20%7C%20Interactive%20Filters&descAlignY=58&descSize=19" width="100%"/>

<img src="https://readme-typing-svg.demolab.com/?lines=Sales+%26+Profit+Interactive+Dashboard;6+Sheets+%E2%80%A2+9%2C994+Orders+%E2%80%A2+0+Macros;Pivots+%7C+Lookups+%7C+Charts+%7C+Dynamic+Filters;Pick+a+filter+%E2%80%94+every+KPI+and+chart+refreshes;Sales+%E2%80%A2+Profit+%E2%80%A2+Regions+%E2%80%A2+Categories&font=Fira+Code&center=true&width=700&height=45&color=4F46E5&vCenter=true&size=22&pause=1200"/>

<br/>

![Excel](https://img.shields.io/badge/Microsoft-Excel-4f46e5?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Sheets](https://img.shields.io/badge/Worksheets-6-0ea5e9?style=for-the-badge)
![Orders](https://img.shields.io/badge/Order_Lines-9,994-8957e5?style=for-the-badge)
![Period](https://img.shields.io/badge/Period-2014--2017-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-8b5cf6?style=for-the-badge)

</div>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Project%20Overview&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Project Overview"/>
</div>

**Sales & Profit Dashboard** is a hands-on Excel project (`Bhavika_ExcelDashboard.xlsx`) that turns a **9,994-row retail sales dataset (2014 – 2017)** into a one-page interactive dashboard. It is backed by a cleaned data table, pivot-style summaries, an advanced-formulas lab and a full chart library. Each sheet answers real business questions such as *"Which region earns the most?"*, *"Which products lose money?"* and *"How do discounts affect profit?"*

Everything is **formula-driven with zero macros**. Pick a **Region, Category, Segment or Year** from a dropdown and every KPI and chart updates instantly.

> 📸 **Dashboard preview:** add your screenshot here
>
> `![Dashboard](screenshots/dashboard.png)`

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Workbook%20Structure&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Workbook Structure"/>
</div>

| # | Sheet | Records | What it covers |
|:-:|---|:-:|---|
| 1 | 📄 `Raw_Data` | 9,994 | Original untouched dataset (21 columns) |
| 2 | 🧹 `Clean_Data` | 9,994 | Excel Table + 4 calculated columns |
| 3 | 🧮 `Pivot_Tables` | 5 tables | Category, Region, Segment, Sub-Category, Year |
| 4 | ⚙️ `Formulas` | 3 tools | VLOOKUP vs INDEX-MATCH, nested IF, SUMIF / COUNTIF |
| 5 | 📊 `Charts` | 5 charts | Source tables, charts and conditional formatting |
| 6 | 🖥️ `Dashboard` | 5 KPIs | Filters, KPI cards and 5 interactive charts |

### 1️⃣ `Raw_Data`
The original dataset, kept untouched for reference: order, customer, location, product and sales metrics across **9,994 order lines, 793 customers, 5,009 orders and 49 states**.

### 2️⃣ `Clean_Data`
| Column | Description |
|---|---|
| Year | `=YEAR(Order Date)` |
| Month | `=MONTH(Order Date)` |
| Profit Status | `Profit` / `Loss` using `IF` |
| Discount Level | `No Discount` / `Low` / `Medium` / `High` using nested `IF` |

### 3️⃣ `Pivot_Tables`
Five summary tables: **Sales & Profit by Category**, **by Region**, **Sales by Segment**, **by Sub-Category** and **by Year**, each with grand totals.

### 4️⃣ `Formulas`
| Tool | Description |
|---|---|
| 🔎 Row ID lookup | `VLOOKUP` and `INDEX-MATCH` side by side with a "Both match?" check |
| ⬅️ Left lookup | `INDEX-MATCH` by Customer ID, which `VLOOKUP` cannot do, plus total sales via `SUMIF` |
| 🏷️ Margin verdict | Nested `IF` → Loss-making / Low / Medium / High margin |
| 📊 Counts & totals | `COUNTIF` / `SUMIF` by Profit Status and Discount Level |

### 5️⃣ `Charts`
`SUMIFS`-powered source tables for Category, Region, Segment, Year and Sub-Category with **Profit Margin**, five charts and **conditional formatting** to flag strong and weak performers.

### 6️⃣ `Dashboard`
| Section | Description |
|---|---|
| 🎛️ Filters | 4 dropdowns: Region · Category · Segment · Year (each with "All") |
| 📈 KPI Cards | Total Sales · Total Profit · Profit Margin · Quantity Sold · Avg Discount |
| 📊 Charts | Category bar, Region bar, Segment pie, Sub-Category bar, Yearly trend line |

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Features%20%2F%20Operations%20Covered&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Features / Operations Covered"/>
</div>

<table>
<tr>
<td width="50%" valign="top">

**🔹 Aggregation**
- `SUMIFS` for multi-condition totals
- `AVERAGEIFS` for average discount
- `SUMIF`, `COUNTIF`
- Named ranges (`rng_Sales`, `rng_Region`...)

</td>
<td width="50%" valign="top">

**🔹 Lookup & Logic**
- `VLOOKUP` and `INDEX` + `MATCH`
- `IF` and nested `IF`
- `IFERROR` for safe results
- Wildcard `"*"` trick for "All" filters

</td>
</tr>
<tr>
<td width="50%" valign="top">

**🔹 Data Preparation**
- Structured Excel Tables
- `YEAR`, `MONTH` helper columns
- Profit Status and Discount Level flags
- Data Validation dropdowns with error alerts

</td>
<td width="50%" valign="top">

**🔹 Visuals**
- Bar, Pie and Line charts
- Conditional formatting
- KPI cards
- Filter panel with live refresh

</td>
</tr>
</table>

### 🧠 How the Filters Work

```excel
Region criteria   =IF($B$4="All","*",$B$4)

Total Sales       =SUMIFS(rng_Sales, rng_Region,   $S$4,
                                     rng_Category, $S$5,
                                     rng_Segment,  $S$6,
                                     rng_Year, ">="&$S$7,
                                     rng_Year, "<="&$S$8)
```

Choosing **"All"** becomes `*` (match everything); choosing a value narrows the result. All KPIs and charts read from this same logic, so they always stay in sync.

### 🧮 Analysis Results

| Sheet | Task | Function | Result |
|---|---|---|:-:|
| Dashboard | Total sales | `SUMIFS` | **$2,297,201** |
| Dashboard | Total profit | `SUMIFS` | **$286,397** |
| Dashboard | Profit margin | `Profit ÷ Sales` | **12.5%** |
| Dashboard | Quantity sold | `SUMIFS` | **37,873** |
| Dashboard | Average discount | `AVERAGEIFS` | **15.6%** |
| Clean_Data | Loss-making order lines | `IF` | **1,871 (18.7%)** |

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Dataset%20Description&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Dataset Description"/>
</div>

| Column | Description |
|---|---|
| Row ID, Order ID | Unique line number and order code |
| Order Date, Ship Date, Ship Mode | When and how the order moved |
| Customer ID, Customer Name, Segment | Consumer · Corporate · Home Office |
| Country, City, State, Postal Code | Delivery location |
| Region | Central · East · South · West |
| Category, Sub-Category | Furniture · Office Supplies · Technology (17 sub-categories) |
| Product ID, Product Name | Product details |
| Sales, Quantity, Discount, Profit | Core business metrics |
| Year, Month, Profit Status, Discount Level | ✨ Added in `Clean_Data` |

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Tech%20Stack&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Tech Stack"/>
</div>

![Excel](https://img.shields.io/badge/Microsoft_Excel-4f46e5?style=flat-square&logo=microsoftexcel&logoColor=white)
![Formulas](https://img.shields.io/badge/Live_Formulas-a78bfa?style=flat-square)
![Charts](https://img.shields.io/badge/Charts-4f46e5?style=flat-square)
![Data Validation](https://img.shields.io/badge/Data_Validation-0ea5e9?style=flat-square)
![Conditional Formatting](https://img.shields.io/badge/Conditional_Formatting-8957e5?style=flat-square)

> ⚠️ **Note:** Open in desktop **Microsoft Excel** so formulas, filters and charts calculate correctly. Please do not edit the *"Chart Data & Filter Logic"* block on the right side of the `Dashboard`.

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=How%20to%20Run&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="How to Run"/>
</div>

1. Download or clone this repository.
2. Open `Bhavika_ExcelDashboard.xlsx` in Microsoft Excel (2016 or later).
3. Go to the `Dashboard` sheet.
4. Use the 🟨 **yellow dropdown cells** to filter by Region, Category, Segment or Year.
5. Watch the KPIs and charts update instantly.
6. Try the lookup tool in `Formulas`: change the 🟨 inputs (Row ID `1–9994` or Customer ID like `CG-12520`).

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Sample%20Insights&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Sample Insights"/>
</div>

- 📈 Sales grew from **$484K (2014)** to **$733K (2017)**; profit nearly doubled from **$49.5K to $93.4K**
- 💻 **Technology** leads in sales ($836K) and profit ($145K) with a **17.4%** margin
- 🪑 **Furniture** sells $742K but earns only a **2.5%** margin, the weakest category
- 🌎 **West** is the top region (**14.9%** margin); **Central** lags at **7.9%**
- 🔻 **Tables, Bookcases and Supplies** lose money; Tables alone lose about **$17.7K**
- ⭐ **Copiers, Paper, Labels and Envelopes** have the best margins (about **37% – 44%**)
- 💸 **Medium and High discount** lines are net losses (about **−$36K** and **−$100K**)
- 👥 **Consumer** is the biggest segment; **Home Office** has the best margin (**14.0%**)

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=File%20Structure&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="File Structure"/>
</div>

```
📦 Sales_Profit_Dashboard
 ┣ 📄 Bhavika_ExcelDashboard.xlsx   # Main workbook (6 sheets)
 ┣ 📁 screenshots                   # Dashboard images
 ┗ 📄 README.md                     # Project documentation
```

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Future%20Improvements&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Future Improvements"/>
</div>

- [ ] Add a Month / Quarter filter and seasonality view
- [ ] Add a state-level map chart
- [ ] Add Top 10 Customers / Products leaderboard
- [ ] Add a what-if discount simulator
- [ ] Build a Power BI / Tableau version for comparison

<br/>

<div align="center">

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=110&section=header&text=Author&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Author"/>
</div>

**Bhavika Thadani**

Made with ❤️ to turn raw sales data into live business insights.

⭐ **If you found this project useful, consider giving it a star!**

<img src="https://capsule-render.vercel.app/api?type=waving&color=0:4f46e5,100:a78bfa&height=120&section=footer&text=Thanks%20for%20visiting!&fontSize=24&fontColor=ffffff&animation=twinkling&fontAlignY=65" width="100%"/>

</div>
