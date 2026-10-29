<div align="center">

# 📊 Sales Performance Analysis

### State-Level Sales & Profitability Reporting using Advanced Excel

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](#)
[![Google Sheets](https://img.shields.io/badge/Google_Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)](#)
[![Status](https://img.shields.io/badge/Status-Completed-2ea44f?style=for-the-badge)](#)
[![Records](https://img.shields.io/badge/Records-2,538-0b5fff?style=for-the-badge)](#)
[![Period](https://img.shields.io/badge/Period-2014–2017-f59e0b?style=for-the-badge)](#)

**[📄 View the Full Report (PDF)](report/Sales_Analysis_Report.pdf)** &nbsp;|&nbsp; **[📊 Open the Live Workbook (Google Sheets)](https://docs.google.com/spreadsheets/d/1lVaDYh63HXofDTpDXYK8QjSj4Q_Bhc8ZrBwQWS5ctqM/edit?usp=sharing)**

</div>

---

## 📑 Table of Contents

1. [Executive Summary](#-executive-summary)
2. [Business Problem](#-business-problem)
3. [Dataset](#-dataset)
4. [Methodology](#-methodology)
5. [Formula Reference](#-formula-reference)
6. [Key Findings](#-key-findings)
7. [Recommendations](#-recommendations)
8. [Limitations & Future Work](#-limitations--future-work)
9. [Repository Structure](#-repository-structure)
10. [Getting Started](#-getting-started)
11. [Skills Demonstrated](#-skills-demonstrated)
12. [Author](#-author)

---

## 🧾 Executive Summary

This project converts **2,538 order-level sales records (2014–2017)** into a **20-state performance report**, built entirely with spreadsheet formulas. The report shows where revenue comes from, which customer segment drives it, and where profit is at risk.

| 💰 Total Sales | 📈 Net Profit | 🛒 Orders | 🗺 States |
|:---:|:---:|:---:|:---:|
| **$438K** | **$37.5K** | **2,538** | **20** |

**Headline insight:** sales volume does not guarantee profit. Five of the ten highest-selling states are loss-making, and profit is concentrated in just two states.

---

## 🎯 Business Problem

Leadership needs a quick, reliable view of regional performance. For every state, this report answers five questions:

| # | Question | Metric |
|:-:|---|---|
| 1 | How many orders did the state place? | Total orders |
| 2 | How much did it sell, and what share of the total is that? | Total sales, contribution % |
| 3 | How does it rank against other states? | Rank by sales |
| 4 | Which customer segment leads? | Leading segment |
| 5 | Is profit healthy? | Total profit, profit slab, red flags |

---

## 🗂 Dataset

| Attribute | Detail |
|---|---|
| **Granularity** | One row per order line |
| **Size** | 2,538 rows × 21 columns |
| **Time span** | 2014 – 2017 |
| **Geography** | 20 U.S. states |
| **Fields used** | Order Date, Segment, State, Region, Sales, Quantity, Discount, Profit |
| **Live workbook** | [Sales_Report_Analysis (Google Sheets, view-only)](https://docs.google.com/spreadsheets/d/1lVaDYh63HXofDTpDXYK8QjSj4Q_Bhc8ZrBwQWS5ctqM/edit?usp=sharing) |

> 📝 **Source:** *add the dataset source and link here before publishing.*

---

## ⚙️ Methodology

The workbook has three layers, following a standard raw → summary → insight flow.

| Layer | Sheet | Purpose |
|---|---|---|
| **1. Raw data** | `Sheet1` | 2,538 rows and 21 columns of order, customer, segment, state, region, product, sales, quantity, discount and profit |
| **2. Summary** | `Sheet2` | 20 states, with formulas in columns **B to U** and a totals row (row 22) |
| **3. Insights** | Output | Rank, segment mix, order recency, profit slab and red flags for low profit |

**Design decision:** the formulas use whole-column references (`Sheet1!I:I` for State, `P:P` for Sales, `S:S` for Profit). New rows added to `Sheet1` are picked up automatically, with no formula edits needed.

---

## 🧮 Formula Reference

### Function Inventory

| Column | Metric | Function(s) |
|:-:|---|---|
| B | Region | `VLOOKUP` |
| C | Region : State label | `CONCATENATE` |
| D | Total orders | `COUNTIF` |
| E, P | Total sales | `SUMIF` |
| F | Rank by sales | `RANK` |
| G | Contribution % | Division with absolute row (`$22`) |
| H–J | Quantity by segment | `SUMIFS` |
| K | Leading segment | `XLOOKUP` + `MAX` |
| L–M | First and last order date | `MINIFS`, `MAXIFS` |
| N | Time since last order | `DATEDIF` + `TODAY` |
| O | Last order month | `TEXT` |
| Q | Average sale | `AVERAGEIF` |
| R, U | Total profit | `SUMIF` (U adds red conditional formatting) |
| S | Profit after 18% tax | `IF` |
| T | Profit slab | `IF` + `AND` (nested) |

### Formulas in Action (Texas as the sample row)

**Lookup, counting and ranking**
```excel
=VLOOKUP(A2, Sheet1!I:J, 2, FALSE)           // Region              → Central
=CONCATENATE(B2, ":", A2)                    // Label               → Central:Texas
=COUNTIF(Sheet1!I:I, A2)                     // Total orders        → 232
=SUMIF(Sheet1!I:I, A2, Sheet1!P:P)           // Total sales         → 32,876
=RANK(E2, E$2:E$21, 0)                       // Rank by sales       → 3
=E2/E$22                                     // Contribution %      → 7.50%
```

**Segments and dates**
```excel
=SUMIFS(Sheet1!Q:Q, Sheet1!I:I, A2, Sheet1!G:G, $H$1)   // Quantity by segment → 259 / 406 / 160
=XLOOKUP(MAX(H2:J2), H2:J2, $H$1:$J$1)                  // Leading segment     → Consumer
=MINIFS(Sheet1!C:C, Sheet1!I:I, A2)                     // First order date    → 03-Jan-2014
=MAXIFS(Sheet1!C:C, Sheet1!I:I, A2)                     // Last order date     → 26-Dec-2017
=DATEDIF(M2, TODAY(), "Y") & " Years " & …              // Time since last order
=TEXT(M2, "MMM - YY")                                   // Last order month    → Dec - 17
```

**Profit and classification**
```excel
=AVERAGEIF(Sheet1!I:I, A2, Sheet1!P:P)       // Average sale        → 141.71
=SUMIF(Sheet1!I:I, A2, Sheet1!S:S)           // Total profit        → -3,323.43
=IF(R2>0, R2*(1-18%), "no tax")              // Profit after tax
=IF(R2<10000, "<10K", IF(AND(R2>=10000, R2<=20000), …)) // Profit slab
=SUM(D2:D21)                                 // Totals row (row 22)
```

**Conditional formatting (column U):** a red fill is applied when the cell value is below 10,000. This flags **18 of 20 states**.

---

## 📈 Key Findings

### 1. Revenue is concentrated in a few states

| Rank | State | Sales | Share of Total |
|:-:|---|--:|--:|
| 1 | California | $96,863 | 22.11% |
| 2 | New York | $75,834 | 17.31% |
| 3 | Texas | $32,876 | 7.50% |

- California and New York together generate **39.4%** of total sales.
- The top three states account for **46.9%** of total sales.

### 2. The Consumer segment dominates

- Consumer accounts for **1,317 of 2,538 orders (52%)**.
- It is the **leading segment in 19 of 20 states**.

### 3. Sales rank does not predict profit

Profit margin is calculated here as profit ÷ sales, using the figures in the report. Profit values are rounded to the nearest dollar.

| Sales Rank | State | Sales | Profit | Margin |
|:-:|---|--:|--:|--:|
| 1 | California | $96,863 | $15,134 | 15.6% |
| 2 | New York | $75,834 | $16,253 | 21.4% |
| 3 | Texas | $32,876 | **-$3,323** | -10.1% |
| 4 | Washington | $30,219 | $6,091 | 20.2% |
| 5 | Illinois | $28,992 | **-$6,758** | -23.3% |
| 6 | Pennsylvania | $25,085 | **-$3,587** | -14.3% |
| 7 | Michigan | $20,752 | $8,186 | 39.4% |
| 8 | Florida | $19,823 | $555 | 2.8% |
| 9 | Ohio | $19,111 | **-$1,822** | -9.5% |
| 10 | N. Carolina | $17,635 | **-$2,890** | -16.4% |

- **5 of the top 10 sales states are loss-making:** Texas, Illinois, Pennsylvania, Ohio and North Carolina.
- **Michigan** ranks 7th in sales but 3rd in profit, with the highest margin of the top 10.

### 4. Profit is concentrated and widely at risk

| Indicator | Result |
|---|:-:|
| States above $10K profit | **2 of 20** (New York, California) |
| States below $10K profit (red-flagged) | **18 of 20** |
| States with negative profit | **8 of 20** |
| Largest single loss | **Illinois, -$6.8K** |

Negative-profit states: Illinois, Pennsylvania, Texas, North Carolina, Arizona, Ohio, Tennessee and Colorado.

---

## 💡 Recommendations

1. **Investigate loss-making high-volume states first.** Illinois, Pennsylvania and Texas combine large sales with negative profit, so a small margin improvement there has an outsized effect on total profit.
2. **Test the role of discounting.** The dataset includes a `Discount` column. Comparing average discount against profit by state is the logical next step. *This is a hypothesis to test, not a conclusion.*
3. **Protect the profit engines.** New York and California carry most of the profit, so retention there is the highest-priority risk.
4. **Learn from Michigan.** Its margin is the strongest among the top 10 states, so it is worth examining what differs in its product mix or pricing.
5. **Check segment profitability.** Consumer leads on volume, so confirm whether the other segments earn better margins before shifting focus.

---

## ⚠️ Limitations & Future Work

**Known limitations**

| Area | Detail | Suggested Improvement |
|---|---|---|
| Tax rate | The 18% rate is hardcoded in the formula | Move it to a single named input cell |
| Volatile function | `TODAY()` in `DATEDIF` changes the result daily | Use a fixed reference date for reproducible results |
| Performance | Whole-column references can slow large workbooks | Convert the data to a Table with structured references |
| Scope | Profit is analysed at state level only | Extend to product, category and discount-band analysis |

**Roadmap**

- [ ] Add PivotTables and Pivot Charts by region and segment
- [ ] Build an interactive **Power BI** dashboard from the same data
- [ ] Reproduce the state aggregation in **SQL** to cross-validate the spreadsheet totals
- [ ] Analyse profit by product category and discount band
- [ ] Add a year-over-year trend analysis (2014 to 2017)

---

## 📁 Repository Structure

> Adjust the file names to match your repository.

```text
Sales-Performance-Analysis/
│
├── data/
│   └── Sales_Analysis.xlsx            # Optional: downloaded copy of the Google Sheet
│
├── report/
│   └── Sales_Analysis_Report.pdf      # Full presentation of findings
│
├── images/
│   ├── sales_by_state.png             # Chart: top 10 states by sales
│   └── profit_by_state.png            # Chart: profit by state
│
└── README.md
```

---

## 🚀 Getting Started

### Option 1: Explore online (no download)

👉 **[Open the workbook in Google Sheets](https://docs.google.com/spreadsheets/d/1lVaDYh63HXofDTpDXYK8QjSj4Q_Bhc8ZrBwQWS5ctqM/edit?usp=sharing)** (view-only)

- **Sheet1** holds the 2,538 raw order records.
- **Sheet2** holds the state-level summary. Click any cell in columns **B to U** to inspect its formula.

### Option 2: Work with your own copy

1. Open the Google Sheet and sign in to your Google account.
2. Go to **File → Make a copy** to edit and experiment freely.
3. To use Excel instead, go to **File → Download → Microsoft Excel (.xlsx)**.

> **Compatibility:** the formulas work in Google Sheets and in Excel 365 / 2021 or later. `XLOOKUP`, `MINIFS` and `MAXIFS` are not available in older Excel versions.

New rows added to **Sheet1** are picked up by the summary automatically.

---

## 🛠 Skills Demonstrated

| Category | Skills |
|---|---|
| **Spreadsheets** | Excel and Google Sheets: lookup functions, conditional aggregation (`SUMIFS`, `COUNTIF`, `MINIFS`), nested logic, date functions, conditional formatting |
| **Analysis** | Ranking, contribution analysis, segmentation, profitability and margin review |
| **Data Quality** | Formula auditing and documenting limitations |
| **Communication** | Stakeholder-ready report with clear KPIs, findings and recommendations |

---

## 👩‍💻 Author

**Akanksha**
*Aspiring Data Analyst · Excel · Power BI · SQL · Python · Statistics*

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](www.linkedin.com/in/akanksha-badoni-9b4ab83a9)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-181717?style=flat-square&logo=github&logoColor=white)]([https://github.com/your-username](https://github.com/akankshabadoni))

---

<div align="center">

⭐ **If you found this project useful, please consider giving it a star.** ⭐

</div>
