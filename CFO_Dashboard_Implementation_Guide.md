# Corporate Executive CFO Dashboard — Power BI Implementation Guide

A complete, buildable specification for a Revenue → Net Income CFO command center, with a matching sample dataset (`CFO_Dashboard_Sample_Data.xlsx`) you can load straight into Power BI Desktop and follow along step by step.

---

## 0. What You're Building

One Power BI model (star schema) driving three report pages:

1. **Executive Overview** — KPI cards, Revenue→Net Income waterfall, expense decomposition tree, trend charts
2. **Department Detail** — drill-through page filtered to a single department
3. **Department Tooltip** — a small hover page shown on the trend/matrix visuals

Everything below assumes you load the attached sample workbook. If you're connecting to a real GL export instead, the column names and logic transfer directly — just adjust the Power Query source step.

---

## 1. Sample Dataset

`CFO_Dashboard_Sample_Data.xlsx` contains 32 months of synthetic GL activity (Jan 2024–Aug 2026) across 4 operating regions, 7 departments, and 16 chart-of-accounts line items — 2,944 transaction rows in total. It's built so the model reconciles cleanly (2025 sanity check: ~58% gross margin, ~21% operating margin, ~15% net margin) and so every measure in Section 6 returns real numbers the moment you load it.

### Sheet: `Fact_GL`

| Column | Type | Notes |
|---|---|---|
| TransactionID | Whole number | Surrogate key, not used in the model |
| Date | Date | First of month; one aggregated line per account/region/department/month |
| AccountKey | Whole number | FK → `Dim_Account` |
| RegionKey | Whole number | FK → `Dim_Region` |
| DepartmentKey | Whole number | FK → `Dim_Department` |
| Amount | Decimal | **Always a positive magnitude**, with one exception — see below |

> **Sign convention:** every line in `Fact_GL` is stored as a positive number representing the size of that line item. Whether it *adds to* or *subtracts from* profit is determined entirely by `Dim_Account[AccountType]`, not by the sign of `Amount`. This is the single most important modeling decision in this guide — see Section 14 for why. The one exception is `Net Change in Working Capital`, which can be genuinely positive (source of cash) or negative (use of cash), because that's what it represents economically.

### Sheet: `Dim_Account`

| AccountKey | AccountName | AccountType | ExpenseCategory | SortOrder |
|---|---|---|---|---|
| 1 | Product Revenue | Revenue | — | 10 |
| 2 | Service Revenue | Revenue | — | 11 |
| 3 | Materials & Supplies | COGS | COGS | 20 |
| 4 | Direct Labor | COGS | COGS | 21 |
| 5 | Digital Advertising | Operating Expense | Marketing | 30 |
| 6 | Marketing Events & Sponsorships | Operating Expense | Marketing | 31 |
| 7 | Salaries | Operating Expense | Salaries & Wages | 32 |
| 8 | Employee Benefits | Operating Expense | Salaries & Wages | 33 |
| 9 | Rent & Facilities | Operating Expense | Other Operating | 34 |
| 10 | IT & Software Subscriptions | Operating Expense | Other Operating | 35 |
| 11 | Travel & Entertainment | Operating Expense | Other Operating | 36 |
| 12 | Professional & Legal Fees | Operating Expense | Other Operating | 37 |
| 13 | Interest Expense | Interest Expense | — | 40 |
| 14 | Income Tax Expense | Tax Expense | — | 41 |
| 15 | Depreciation & Amortization | Non-Cash Adjustment | — | 50 |
| 16 | Net Change in Working Capital | Working Capital Change | — | 51 |

`AccountType` drives every P&L placement measure in Section 6. `ExpenseCategory` is the third level of the Region → Department → Expense Category decomposition tree, and is intentionally blank for Revenue, Interest, Tax, and the two cash-flow-only accounts, so they never appear as expense branches. `SortOrder` drives the waterfall chart's left-to-right sequence.

### Sheet: `Dim_Region`

| RegionKey | RegionName |
|---|---|
| 1 | East |
| 2 | West |
| 3 | North |
| 4 | South |
| 5 | Corporate |

`Corporate` is a bookkeeping region used only for company-level items (Interest, Tax, D&A, Working Capital Change) that don't belong to any single operating region.

### Sheet: `Dim_Department`

| DepartmentKey | DepartmentName |
|---|---|
| 1 | Sales |
| 2 | Marketing |
| 3 | Engineering |
| 4 | Customer Support |
| 5 | G&A |
| 6 | Operations |
| 7 | Corporate Finance |

`Corporate Finance` pairs exclusively with `Region = Corporate`.

**No date table is in the workbook on purpose** — Section 4 builds it natively in DAX, which is the correct approach for a production model (see Section 12, mistake #1, for why).

---

## 2. Data Model — Star Schema

```
                     DimDate
                        |
DimAccount ---- FactFinancials ---- DimRegion
                        |
                  DimDepartment
```

Rename the four loaded queries in Power Query (Section 3) to match Power BI naming conventions: `Fact_GL` → **FactFinancials**, `Dim_Account` → **DimAccount**, `Dim_Region` → **DimRegion**, `Dim_Department` → **DimDepartment**. Then build **DimDate** natively (Section 4).

| From (dimension) | To (fact) | Cardinality | Cross-filter direction |
|---|---|---|---|
| DimDate[Date] | FactFinancials[Date] | One-to-many | Single (Dim → Fact) |
| DimAccount[AccountKey] | FactFinancials[AccountKey] | One-to-many | Single (Dim → Fact) |
| DimRegion[RegionKey] | FactFinancials[RegionKey] | One-to-many | Single (Dim → Fact) |
| DimDepartment[DepartmentKey] | FactFinancials[DepartmentKey] | One-to-many | Single (Dim → Fact) |

**Why a star schema, and why single-direction:** FactFinancials is the only "many" side table; every dimension filters into it and never the reverse. This keeps every filter path unambiguous — Power BI never has to guess which route a filter should travel — and it's what makes CALCULATE()-based measures (Section 6) behave predictably. Bidirectional relationships are avoided entirely here; they're rarely needed in a single-fact star schema and are a common source of circular-dependency errors (Section 12).

**If you later add budget data:** bring in a `FactBudget` table with the *same* AccountKey / RegionKey / DepartmentKey / Date grain and relate it to the *same* four dimension tables. This is a legitimate extension of a star schema (sometimes called a "fact constellation") — both fact tables share conformed dimensions, so a single set of slicers filters both simultaneously. See the Budget vs Actual measures in Section 6.

---

## 3. Power Query Transformation Steps

1. **Get Data → Excel Workbook** → select `CFO_Dashboard_Sample_Data.xlsx`.
2. Load `Fact_GL`, `Dim_Account`, `Dim_Region`, `Dim_Department` (skip `ReadMe`, or load and immediately disable "Enable Load" on it).
3. In each query, confirm the auto-detected data types: `Date` → Date, `Amount` → Decimal Number, all Keys → Whole Number, text columns → Text. Power Query usually gets this right from Excel, but verify — a Key silently typed as text will break relationship cardinality.
4. On `Fact_GL`: **Transform → Replace Errors → 0** on the Amount column as a safety net (guards against a blank cell in a future refresh from a real GL export).
5. On the two smaller dimension queries (`Dim_Region`, `Dim_Department`): **Transform → Trim** and **Transform → Clean** on the name columns, in case a real-world export has trailing whitespace.
6. Rename each query: right-click → Rename → `FactFinancials`, `DimAccount`, `DimRegion`, `DimDepartment`.
7. **Home → Close & Apply.**
8. In Model view, Power BI will usually auto-detect the four relationships from matching key names; if any are missing, drag `AccountKey`/`RegionKey`/`DepartmentKey` from each dimension onto the matching column in `FactFinancials`. Confirm each is **One-to-many, single direction** per the table in Section 2.

**If you're connecting to a real GL export instead of the sample file:** the most common extra step is un-pivoting. Many GL exports come with one column per month (`Jan-24`, `Feb-24`, …) rather than one row per month. Select all the month columns → **Transform → Unpivot Columns** → rename the resulting columns to `Date` (after parsing the header text into an actual date) and `Amount`. You'll also typically need a **Merge Queries** step against your chart-of-accounts mapping to attach `AccountType` and `ExpenseCategory` — that mapping *is* `DimAccount` in this model.

---

## 4. Date Table — Complete DAX

Build this as a new table (Modeling → New Table), not an imported sheet — that's what makes every time-intelligence function in Section 6 work correctly against a guaranteed-contiguous calendar.

```dax
DimDate =
VAR FiscalYearStartMonth = 1   -- 1 = fiscal year matches calendar year; set to e.g. 7 for a July start
VAR MinDate = DATE(YEAR(MIN(FactFinancials[Date])), 1, 1)
VAR MaxDate = DATE(YEAR(MAX(FactFinancials[Date])), 12, 31)
RETURN
ADDCOLUMNS(
    CALENDAR(MinDate, MaxDate),
    "Year", YEAR([Date]),
    "MonthNumber", MONTH([Date]),
    "MonthName", FORMAT([Date], "MMMM"),
    "MonthShort", FORMAT([Date], "MMM"),
    "MonthYear", FORMAT([Date], "MMM YYYY"),
    "YearMonthSort", YEAR([Date]) * 100 + MONTH([Date]),
    "Quarter", "Q" & QUARTER([Date]),
    "YearQuarter", YEAR([Date]) & " Q" & QUARTER([Date]),
    "FiscalYear",
        YEAR([Date]) + IF(FiscalYearStartMonth = 1, 0, IF(MONTH([Date]) >= FiscalYearStartMonth, 1, 0)),
    "FiscalMonthNumber", MOD(MONTH([Date]) - FiscalYearStartMonth, 12) + 1,
    "FiscalQuarter", "FQ" & ROUNDUP((MOD(MONTH([Date]) - FiscalYearStartMonth, 12) + 1) / 3, 0),
    "IsCurrentMonth", AND(YEAR([Date]) = YEAR(TODAY()), MONTH([Date]) = MONTH(TODAY()))
)
```

**Mark it as the official Date table:** select `DimDate` in the Fields pane → Table tools (ribbon) → **Mark as Date Table** → choose the `Date` column. Do this before writing any time-intelligence measure — several DAX time functions fail silently or throw ambiguous errors without it.

**Relationship:** `DimDate[Date]` (1) → `FactFinancials[Date]` (many), single direction, as already listed in Section 2.

**On fiscal years:** the sample data models fiscal year = calendar year (`FiscalYearStartMonth = 1`). If your real company's fiscal year starts on a different month, change that one variable — the `FiscalYear` and `FiscalQuarter` columns recompute correctly. Note that `SAMEPERIODLASTYEAR` (used throughout Section 6) works off the calendar `Date` column, which is correct as long as your fiscal periods are still full calendar months (true for the vast majority of companies, even non-calendar-fiscal ones). If your fiscal calendar uses 4-4-5 weeks instead of calendar months, that's a materially different date table pattern — flag it and it can be adapted separately.

---

## 5. DAX Measures Library

Create a **Measures** table (Modeling → New Table → `Measures = {}` then delete the placeholder column, or use the "Home table" trick) purely to hold these measures so they don't clutter your fact/dimension tables in the Fields pane. Organize into display folders as noted.

### 5.1 Core Financial Measures (folder: `01 Core`)

```dax
Total Revenue = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Revenue")

Total COGS = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "COGS")

Gross Profit = [Total Revenue] - [Total COGS]

Gross Profit Margin % = DIVIDE([Gross Profit], [Total Revenue])

Total Operating Expenses = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Operating Expense")

Operating Income = [Gross Profit] - [Total Operating Expenses]

Operating Margin % = DIVIDE([Operating Income], [Total Revenue])

Interest Expense = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Interest Expense")

Tax Expense = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Tax Expense")

Net Income = [Operating Income] - [Interest Expense] - [Tax Expense]

Net Profit Margin % = DIVIDE([Net Income], [Total Revenue])

Depreciation and Amortization = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Non-Cash Adjustment")

Net Working Capital Change = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[AccountType] = "Working Capital Change")

Operating Cash Flow = [Net Income] + [Depreciation and Amortization] + [Net Working Capital Change]

Total Expenses = [Total COGS] + [Total Operating Expenses]
```

`Total Expenses` deliberately excludes Interest and Tax — those are "below the line" financing/tax items, not operating spend, and shouldn't appear in the expense decomposition tree (Section 6.3) or the "largest expense category" analysis (Section 6.6). `Operating Cash Flow` uses the standard indirect-method shape (Net Income + non-cash add-backs +/− working capital movement) — a simplified version that's directionally correct and audit-explainable without needing a full cash-flow-statement fact table.

### 5.2 Expense Category Breakouts (folder: `02 Expense Detail`)

```dax
Total Marketing Expense = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[ExpenseCategory] = "Marketing")

Total Salaries Expense = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[ExpenseCategory] = "Salaries & Wages")

Total Other Operating Expense = CALCULATE(SUM(FactFinancials[Amount]), DimAccount[ExpenseCategory] = "Other Operating")
```

### 5.3 Time Intelligence (folder: `03 Time Intelligence`)

```dax
Revenue LY = CALCULATE([Total Revenue], SAMEPERIODLASTYEAR(DimDate[Date]))

Revenue YoY = [Total Revenue] - [Revenue LY]

Revenue YoY % = DIVIDE([Revenue YoY], [Revenue LY])

Net Income LY = CALCULATE([Net Income], SAMEPERIODLASTYEAR(DimDate[Date]))

Net Income YoY = [Net Income] - [Net Income LY]

Net Income YoY % = DIVIDE([Net Income YoY], [Net Income LY])

Gross Profit LY = CALCULATE([Gross Profit], SAMEPERIODLASTYEAR(DimDate[Date]))

Gross Profit YoY % = DIVIDE([Gross Profit] - [Gross Profit LY], [Gross Profit LY])

Operating Cash Flow LY = CALCULATE([Operating Cash Flow], SAMEPERIODLASTYEAR(DimDate[Date]))

Operating Cash Flow YoY % = DIVIDE([Operating Cash Flow] - [Operating Cash Flow LY], [Operating Cash Flow LY])
```

`SAMEPERIODLASTYEAR` requires `DimDate` to be marked as the official date table (Section 4) and needs a contiguous, gap-free date range — both are true here. Every YoY % measure uses `DIVIDE()`, not `/`, specifically so a department or period with zero prior-year activity returns a blank instead of an error (Section 12, mistake #9).

### 5.4 % of Revenue (folder: `04 Ratios`)

```dax
Expense % of Revenue = DIVIDE([Total Expenses], [Total Revenue])

COGS % of Revenue = DIVIDE([Total COGS], [Total Revenue])

Salary % of Revenue = DIVIDE([Total Salaries Expense], [Total Revenue])

Marketing % of Revenue = DIVIDE([Total Marketing Expense], [Total Revenue])
```

### 5.5 Budget vs Actual (folder: `05 Budget`) — only if a `FactBudget` table exists

```dax
Budget Revenue = CALCULATE(SUM(FactBudget[Amount]), DimAccount[AccountType] = "Revenue")

Revenue vs Budget = [Total Revenue] - [Budget Revenue]

Revenue vs Budget % = DIVIDE([Revenue vs Budget], [Budget Revenue])
```

Repeat the pattern for any other line item you need to budget-check (COGS, Opex, Net Income) by swapping the `AccountType`/`ExpenseCategory` filter.

### 5.6 Waterfall Support Measure (folder: `06 Waterfall`)

This needs a small disconnected table first — see Section 6.2 for why and how it's built — then:

```dax
PL Bridge Value =
VAR CurrentStep = SELECTEDVALUE(PLBridge[StepName])
RETURN
SWITCH(
    TRUE(),
    CurrentStep = "Revenue", [Total Revenue],
    CurrentStep = "COGS", -[Total COGS],
    CurrentStep = "Gross Profit", [Gross Profit],
    CurrentStep = "Marketing Expenses", -[Total Marketing Expense],
    CurrentStep = "Salaries & Wages", -[Total Salaries Expense],
    CurrentStep = "Other Operating Expenses", -[Total Other Operating Expense],
    CurrentStep = "Operating Income", [Operating Income],
    CurrentStep = "Interest Expense", -[Interest Expense],
    CurrentStep = "Taxes", -[Tax Expense],
    CurrentStep = "Net Income", [Net Income],
    BLANK()
)
```

Every cost line is returned as a **negative** number; every subtotal (Gross Profit, Operating Income, Net Income) is returned as its true positive value. That combination is what makes the waterfall visual's running total land exactly on the real subtotal at each checkpoint — see Section 6.2.

---

## 6. Visual-by-Visual Build Instructions

### 6.1 KPI Cards (Row 1)

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Card (new)** ×8 | — | Data value: one of `[Total Revenue]`, `[Gross Profit]`, `[Gross Profit Margin %]`, `[Net Income]`, `[Net Profit Margin %]`, `[Operating Cash Flow]`, `[Revenue YoY %]`, `[Net Income YoY %]`. Reference value: the matching `LY` measure (e.g., `[Revenue LY]`) with "Show as % difference" turned on | Page-level Fiscal Year / Region / Department slicers | Display units = Auto (Millions for $ cards), 1 decimal max; Callout value font 28–32pt Segoe UI Semibold; conditional font color on the reference-value delta (green ≥0, red <0) via the card's built-in comparison coloring | Set to **no cross-filtering** of other visuals (Format → Edit interactions → None) — KPI cards should summarize, not act as slicers |

Use the modern **Card (new)** visual specifically — its "Reference label" section is built for exactly this pattern: primary value, a comparison value, an automatic % delta, and an automatic up/down indicator color, with no extra measures beyond the LY measure needed.

### 6.2 Revenue → Net Income Waterfall (Row 2, left)

**Setup (one-time, in addition to the measure in 5.6):**

1. **Modeling → New Table**, paste:
```dax
PLBridge = DATATABLE(
    "StepOrder", INTEGER, "StepName", STRING,
    {
        {1, "Revenue"}, {2, "COGS"}, {3, "Gross Profit"},
        {4, "Marketing Expenses"}, {5, "Salaries & Wages"}, {6, "Other Operating Expenses"},
        {7, "Operating Income"}, {8, "Interest Expense"}, {9, "Taxes"}, {10, "Net Income"}
    }
)
```
2. This table has **no relationship** to anything else in the model — it exists purely to drive `SELECTEDVALUE()` in the `PL Bridge Value` measure.
3. Select `PLBridge[StepName]` → Modeling ribbon → **Sort by Column** → `StepOrder`, so the waterfall always renders left-to-right in the correct financial sequence regardless of alphabetical order.

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Waterfall chart** | Category = `PLBridge[StepName]` | Y-axis = `[PL Bridge Value]` | Page-level Fiscal Year / Region / Department slicers | Sentiment colors: Increase = green, Decrease = coral/red, Total = navy; data labels on, currency format in millions | Highlights on click; typically excluded from being filterable by clicking other visuals, since it's the master summary chart |

**Critical step:** right-click the **Gross Profit**, **Operating Income**, and **Net Income** bars in the rendered visual and choose **Set as total**. Without this, the waterfall treats those three as additional deltas and compounds them on top of the running total — a doubled, wrong chart. With it, they render as flat subtotal bars exactly where the eye expects them, and the sign convention in the measure (costs negative, subtotals positive) makes the running math land correctly at each one.

### 6.3 Expense Decomposition Tree (Row 3, left)

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Decomposition tree** | Explain by (in this order): `DimRegion[RegionName]`, `DimDepartment[DepartmentName]`, `DimAccount[ExpenseCategory]` | Analyze: `[Total Expenses]` | Visual-level filter: `DimAccount[ExpenseCategory]` is not blank (removes Revenue/Interest/Tax/D&A/NWC accounts, which have no expense category, from ever appearing as branches) | Root node styled with the navy header color; increase/decrease coloring on branch contribution | Cross-filters the trend charts and matrix on Row 2/4 when a branch is selected, so drilling into "East → Sales → Salaries & Wages" simultaneously updates the trend line to that slice |

Limiting the **Explain by** list to exactly Region, Department, and Expense Category keeps users inside the intended hierarchy while still letting them branch in any order (e.g., Department first, then Region) or use the built-in **AI Splits** button, which auto-suggests the highest-impact next split — genuinely useful for a CFO asking "why did expenses spike this quarter." If you want Account-level granularity as a fourth level, add `DimAccount[AccountName]` to the same list.

### 6.4 Revenue & Profit Trend (Row 2, right)

**Field Parameter setup (gives the CFO a Monthly / Quarterly / Yearly toggle):**

Modeling → New Parameter → **Fields**, add `DimDate[MonthYear]`, `DimDate[YearQuarter]`, `DimDate[Year]` as the three options, name it `Date Granularity`. Power BI creates a small parameter table and a matching slicer automatically.

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Line chart** | X-axis: `Date Granularity` field parameter | Y-values: `[Total Revenue]`, `[Gross Profit]`, `[Operating Income]`, `[Net Income]` | Page-level slicers | Legend on top, distinct line colors per metric (navy/teal/gold/green), data labels off by default (rely on tooltips) | Sync with the `Date Granularity` slicer placed just above it; responds to cross-filtering from the decomposition tree |

### 6.5 Margin Trend (Row 4, left)

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Line chart** | X-axis: same `Date Granularity` field parameter | Y-values: `[Gross Profit Margin %]`, `[Operating Margin %]`, `[Net Profit Margin %]` | Page-level slicers | Y-axis formatted as %, 1 decimal; reference line at 0% if margins can go negative | Same as 6.4 |

### 6.6 Expense Category Analysis (Row 3, right)

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Horizontal bar chart** (not pie/donut) | Axis: `DimAccount[ExpenseCategory]` | Values: `[Total Expenses]`, sorted descending | Region / Department / Expense Category / Fiscal Year / Month slicers all apply | Data labels on, single accent color with the largest bar highlighted in navy and the rest in slate gray | Clicking a bar cross-filters the trend charts and matrix |

### 6.7 Regional / Department Performance (Row 4, right)

| Visual Type | Fields | Measures | Filters | Formatting | Interaction |
|---|---|---|---|---|---|
| **Matrix** | Rows: `DimRegion[RegionName]` expandable to `DimDepartment[DepartmentName]` | Values: `[Total Revenue]`, `[Operating Income]`, `[Operating Margin %]` | Page-level slicers | Conditional formatting: data bars on Revenue, color scale (red→green) on Operating Margin %; row subtotals on | Right-click a Department row → **Drill through → Department Detail** (Section 7) |

---

## 7. Drill-Through Setup: Department Detail Page

1. Add a new report page named **Department Detail**.
2. In the Filters pane for that page, drag `DimDepartment[DepartmentName]` into the **Drillthrough** well at the top (this well only appears once you've added a field to it — Power BI auto-creates the section).
3. Power BI automatically adds a **Back** button in the top-left corner; leave it — it returns the user to whichever page and filter context they drilled from.
4. Build these visuals on the page, all of which will inherit the selected department automatically:
   - Card row: `[Total Revenue]`, `[Total Expenses]`, `[Operating Income]`, `[Operating Margin %]`
   - Expense breakdown bar chart: `DimAccount[AccountName]` × `[Total Expenses]`
   - Monthly trend line chart: `DimDate[MonthYear]` × `[Total Revenue]`, `[Operating Income]`
   - Top expense categories table: `DimAccount[ExpenseCategory]` × `[Total Expenses]`, sorted descending
5. On the source page, right-click any department (in the matrix from 6.7, or a branch of the decomposition tree) → **Drillthrough → Department Detail**.

---

## 8. Slicers & Reset Filters

| Slicer | Field | Style |
|---|---|---|
| Fiscal Year | `DimDate[FiscalYear]` | Dropdown |
| Month | `DimDate[MonthYear]` | Dropdown or between-style date slicer on `DimDate[Date]` |
| Region | `DimRegion[RegionName]` | Vertical list, multi-select |
| Department | `DimDepartment[DepartmentName]` | Vertical list, multi-select |
| Expense Category | `DimAccount[ExpenseCategory]` | Vertical list, multi-select |

Place all five on the Executive Overview page only; use **View → Sync Slicers** to also apply them to any additional pages you add later, but leave the **Department Detail** drillthrough page out of the sync group — it should stay driven purely by drillthrough context, not the page-level slicers.

**Reset Filters button:**
1. Set every slicer to its default ("all selected") state.
2. **View → Bookmarks → Add**, name it `Default View`. In the bookmark's options, confirm "Data" is checked so it captures slicer state.
3. **Insert → Buttons → Blank**, place it near the slicers, label it "Reset Filters."
4. Select the button → Format → Action → Type = **Bookmark** → Bookmark = `Default View`.

---

## 9. Dashboard Page Layout — Executive Overview

```
┌──────────────────────────────────────────────────────────────────────┐
│  CORPORATE CFO EXECUTIVE DASHBOARD                                   │
│  Financial Performance | Profitability | Cash Flow | Expense Intel.  │
├──────────────────────────────────────────────────────────────────────┤
│ [Revenue] [Gross Profit Margin] [Net Profit Margin] [Op. Cash Flow]  │  ← Row 1
├───────────────────────────────────┬──────────────────────────────────┤
│  Revenue → Net Income Waterfall   │   Revenue & Profit Trend         │  ← Row 2
├───────────────────────────────────┼──────────────────────────────────┤
│  Total Expense Decomposition Tree │   Expense Category Analysis      │  ← Row 3
├───────────────────────────────────┼──────────────────────────────────┤
│  Margin Trend                     │   Regional / Department Matrix   │  ← Row 4
└───────────────────────────────────┴──────────────────────────────────┘
   Slicers (Fiscal Year | Month | Region | Department | Expense Category | Reset)
```

Slicers can sit in a thin persistent band along the top or left rail, whichever your page canvas size favors — top works better on a standard 16:9 canvas.

---

## 10. Formatting & Design System

**Palette** (apply via a saved custom Theme JSON — View → Themes → Browse for themes — so it's consistent across every page):

| Role | Color |
|---|---|
| Background | `#F7F8FA` (very light gray) or white |
| Primary / headers | `#1F3864` (deep navy) |
| Accent | `#2CA6A4` (teal) |
| Positive | `#2E7D32` (green) |
| Negative | `#C62828` (red) |
| Neutral bars | `#5B6470` (slate gray) |

**Typography:** Segoe UI throughout. Page title 20pt bold navy; subtitle 11pt regular gray; KPI callout values 28–32pt Segoe UI Semibold; body/axis labels 9–10pt.

**Number formatting:** millions with one decimal (`$12.4M`, not `$12,384,910`). Use Format pane → Data labels → Display units = Millions, Value decimal places = 1 on every currency visual; percentages to one decimal (`21.2%`); avoid unrounded values anywhere on the page.

**Conditional formatting:**
- KPI card reference-value color: green/red automatically via the Card (new) visual's built-in comparison formatting.
- Matrix Operating Margin % column: color scale (red low → green high).
- Matrix Revenue column: data bars.
- Any variance table: icon set (▲/▼) bound to the sign of the underlying YoY measure.

**Avoid:** pie/donut charts anywhere (per the spec — use bars), more than 4–5 colors on any one visual, unnecessary decimal precision, and dense borders/gridlines — let whitespace and the navy/teal palette carry the hierarchy.

---

## 11. Interactions & Tooltip Design

**Edit interactions** (Format ribbon → Edit Interactions, then click each visual in turn):
- KPI cards → **None** on every other visual (they're read-only summaries, not filters).
- Decomposition tree, matrix, and expense bar chart → **Filter** on the trend charts and each other (this is the default; just don't disable it).
- Slicers → filter everything on the page by default; leave as-is.

**Custom tooltip page:**
1. New page, canvas size **Tooltip** (Format → Canvas settings → Type = Tooltip, ~320×240px), name it `Department Tooltip`.
2. Add compact visuals: department name (text), `[Total Revenue]`, `[Total Expenses]`, `[Operating Margin %]` as small cards.
3. On the trend chart and matrix (6.4, 6.5, 6.7): Format → Tooltips → toggle **Report page** → select `Department Tooltip`.

This lets a CFO hover over any department bar or matrix row and see a clean profitability snapshot without clicking through to the full drill-through page.

---

## 12. Common Mistakes to Avoid

1. **Not marking DimDate as the official Date table.** Time-intelligence functions like `SAMEPERIODLASTYEAR` fail or behave unpredictably without it.
2. **Using calculated columns instead of measures** for anything that should respond to filter context (Revenue, margins, etc.) — bloats the model and produces static, non-interactive numbers.
3. **Bidirectional relationships** "just to make a slicer work" — almost always creates an ambiguous filter path instead. This model doesn't need any.
4. **Inconsistent sign conventions** — mixing "expenses stored negative" in one account and "expenses stored positive" in another silently breaks the waterfall and every margin calc. This guide stores everything as a positive magnitude and lets `AccountType`/`ExpenseCategory` do the classification — keep it that way if you extend the model.
5. **Forgetting to mark the waterfall's subtotal bars as "Set as total."** Without it, Gross Profit/Operating Income/Net Income compound on top of the running total instead of landing on the real number.
6. **Hardcoding "current year"** as a literal (e.g., `= 2026`) instead of deriving it dynamically (`MAX(DimDate[Year])`) — quietly breaks on January 1st of next year.
7. **Using `/` instead of `DIVIDE()`** for any margin or ratio — throws an error the moment a filter context produces a zero-revenue period or department.
8. **Overcrowding a single page.** Eight KPI cards plus six large visuals is already a lot — resist adding more without a strong reason; it hurts both readability and refresh performance.
9. **Skipping the decomposition tree's "Explain by" filter list.** Leaving every field in the model available lets users drill into meaningless combinations (e.g., Region by Account Name directly, skipping Department).
10. **Not testing edge cases before presenting:** the very first month in the data (no LY comparison exists — confirm it shows blank, not an error), a department with zero expenses in a given slice, and the "all filters cleared" state.

---

## 13. Final Validation Checklist

- [ ] `DimDate` marked as Date Table, marked column = `Date`
- [ ] All four relationships are one-to-many, single-direction, dimension → fact
- [ ] Model view shows no bidirectional or ambiguous relationships
- [ ] `Total Revenue − Total Expenses − Interest Expense − Tax Expense = Net Income` reconciles exactly in every filter context tested
- [ ] Every YoY/margin measure uses `DIVIDE()`, confirmed by testing a zero-denominator filter context
- [ ] Waterfall: Gross Profit, Operating Income, and Net Income bars are marked "Set as total" and land on the correct running values
- [ ] Decomposition tree's Explain-by list is limited to Region, Department, Expense Category (+ optional Account)
- [ ] Drillthrough page filters correctly on Department and the Back button returns to the source page/filter state
- [ ] Slicers affect every intended visual; sync settings correct across pages
- [ ] Reset Filters button returns every slicer to its default state
- [ ] Custom theme (colors/fonts) applied consistently across all pages
- [ ] Negative YoY and negative margins are visually obvious (red / down arrow) without hovering
- [ ] No pie/donut charts anywhere on the report
- [ ] Performance Analyzer shows no visual taking more than ~1–2 seconds to render on a full data refresh

---

## Final CFO Dashboard Blueprint

**Page 1 — Executive Overview**
- Header: title + subtitle
- Row 1: 8 KPI cards (Revenue, Gross Profit, Gross Profit Margin %, Net Income, Net Profit Margin %, Operating Cash Flow, Revenue YoY %, Net Income YoY %) — each with LY comparison and color-coded delta
- Row 2: Revenue→Net Income waterfall (left) | Revenue & Profit trend with Monthly/Quarterly/Yearly toggle (right)
- Row 3: Total Expense decomposition tree, Region→Department→Expense Category (left) | Expense category bar chart (right)
- Row 4: Margin trend — GP%/OM%/NM% (left) | Region/Department matrix with drillthrough (right)
- Persistent slicer band: Fiscal Year, Month, Region, Department, Expense Category, Reset Filters button

**Page 2 — Department Detail** *(drillthrough, entered from the matrix or decomposition tree)*
- KPI card row: Revenue, Expenses, Operating Income, Operating Margin % for the selected department
- Expense breakdown bar chart by account
- Monthly trend line (Revenue, Operating Income)
- Top expense categories table
- Back button (auto-generated)

**Page 3 — Department Tooltip** *(hidden hover page, ~320×240px)*
- Department name, Revenue, Expenses, Operating Margin % as compact cards — attached as the custom tooltip on the trend charts and matrix

---

The attached sample workbook already matches every table and column name used throughout this guide — load it, build the four relationships in Section 2, add the DAX in Sections 4–5, and every visual in Section 6 will render immediately.
