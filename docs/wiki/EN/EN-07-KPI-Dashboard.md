# 07 – KPI Dashboard

## Opening the KPI Dashboard

Click **≋ KPI** in the main toolbar. The window opens maximised.

---

## Window Structure

The top section contains the time filter controls:

### Period Filter (Rolling Window)

The **1M**, **3M**, **12M**, **3Y**, **5Y**, **10Y**, **All** buttons filter data using a rolling window calculated back from today:

| Button | Range |
|--------|-------|
| 1M | Last 30 days |
| 3M | Last 90 days |
| 12M | Last 365 days |
| 3Y | Last 3 years |
| 5Y | Last 5 years |
| 10Y | Last 10 years |
| All | All data, no time limit |

### Year Filter (Calendar Year)

The **Year** dropdown lets you select a specific calendar year. When the Year filter is active, you get **exactly 12 fixed months** from January to December, regardless of today's date.

> The Period and Year filters are **mutually exclusive**: selecting one automatically deselects the other.

### How the Filter Applies to Savings and Cost Avoidance KPIs

> **Important:** the time filter for Savings and Cost Avoidance KPIs acts on the **economic competence period** (i.e. the month in which the impact is realised), not on the date the event was created.

This means that:
- An event from January with a 12-month distribution has impacts through December.
- Filtering on "current year" includes impacts from events created in the previous year, provided those impacts fall within the current year.
- This accurately reflects the procurement accounting reality.

---

## RFQ Tab

Shows KPIs related to RFQ issuance activity.

### Available KPI Cards

| KPI | Meaning |
|-----|---------|
| Active RFQs | Number of RFQs with "active" status in the period |
| Archived RFQs | Number of RFQs with "archived" status |
| Total RFQs | Active + Archived |
| Non-Expired RFQs | RFQs with a future expiry date |
| Expired RFQs | RFQs with a past expiry date |
| Work Order | RFQs of work-order type |
| Full Supply | RFQs of full supply type |

### Chart

A bar chart showing the number of RFQs issued per month. Each bar corresponds to one month in the selected period. Months with no activity show a zero bar.

### Details Table

Below the chart, a `Period | RFQs Issued` table displays the numerical data sorted most-recent first.

---

## Savings Tab

### Available KPI Cards

| KPI | Meaning |
|-----|---------|
| **Theoretical Savings** | Sum of the theoretical monthly value of all Savings events in the filtered period |
| **Actual Savings** | Sum of the actual value (Theoretical × % Realisation / 100) |
| **Average Savings %** | Weighted average of the savings percentage across all events |
| **Best Savings %** | Highest percentage recorded among all events in the period |
| **Worst Savings %** | Lowest percentage recorded |
| **Median Savings %** | Median value of savings percentages |
| **Recurring Impact (€)** | Savings from events with Recurring OPEX enabled |
| **Non-Recurring Impact (€)** | Savings from one-time events |

### How Average Savings % Is Calculated

This is not a simple arithmetic average — it is a **weighted average**:

$$\text{Average Savings\%} = \frac{\sum \text{Savings}_{event}}{\sum \text{Base}_{event}} \times 100$$

The "Base" is `Budget Amount × Annual Quantity` for a Price driver, or `Annual Spend` for a Payment Terms driver. This prevents a small event with a high percentage from skewing the overall result.

### Carry-Over (Year Filter Only)

When the **Year filter** is active, an additional KPI appears: **Carry-over to next year (€)**.

This value represents the total economic impact from events already created in the selected year (or earlier) that will materialise in the **following year**. It is useful for budget projections and for demonstrating to management the value already "in the pipeline" for next year.

For example, a savings event from a multi-year contract signed in November will have impacts in the eleven months that follow — the portion falling in the next year is the carry-over.

### Chart

A dual bar chart: blue bars for Theoretical Savings and orange bars for Actual Savings, side by side for each month in the period.

---

## Cost Avoidance Tab

Identical structure to the Savings tab. The corresponding KPIs are:

| KPI | Meaning |
|-----|---------|
| **Theoretical Cost Avoidance** | Sum of the theoretical value of Cost Avoidance events in the period |
| **Actual Cost Avoidance** | Actual value after applying % Realisation |
| **Average CA %** | Weighted average of avoidance percentages |
| **Best / Worst / Median CA %** | Per-event statistics |
| **Recurring / Non-Recurring Impact** | Breakdown by event type |
| **Carry-over to next year** | Only available when the Year filter is active |

---

## Derisking Tab

### Available KPI Cards

| KPI | Meaning |
|-----|---------|
| **Total Potential Suppliers** | Number of suppliers registered in the period |
| **Unique Categories** | Number of distinct product categories |
| **New** / **Under evaluation** / **Qualified** / **Rejected** | Count per status |

### Chart

A bar chart showing the number of new suppliers registered per month.

### Details Table

A summary by category showing the number of suppliers in each.

---

## Exporting KPIs to Excel

Click **📥 Export Excel** in the top-right corner of the KPI window.

**Step 1 – Choose the scope:**
- **📋 Current section** – exports only the active tab's data
- **📊 All sections** – exports all four tabs in a single Excel file

**Step 2 – Choose the language:** Italian or English.

**Step 3 – Choose where to save the file.**

The generated Excel file contains:
- A **Summary** sheet with metadata (export date, active filter, scope)
- One sheet per exported section with numerical values
- Number formats: monetary `€ 1,234.56`; percentage `12.34%`

---

## Refreshing Data

The data in the KPI window reflects the state of the database at the time the window was opened. To update values after adding new events, close and reopen the KPI window, or change the time filter to force a recalculation.
