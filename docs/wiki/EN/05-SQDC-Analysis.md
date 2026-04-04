# 05 – SQDC Analysis

## What is the SQDC Analysis

SQDC is a weighted multi-criteria evaluation framework for supplier selection. The four dimensions assessed are:

| Criterion | Meaning in DataFlow |
|-----------|---------------------|
| **S – Safety** | Compliance with safety standards, regulatory risk, financial and geopolitical stability |
| **Q – Quality** | Ability to meet the agreed technical specifications |
| **D – Delivery** | On-time delivery, flexibility, and responsiveness |
| **C – Cost** | Price competitiveness, payment terms, and ancillary costs |

The result is a **weighted score** for each supplier. The supplier with the highest score is the one recommended by the analysis.

---

## Opening the SQDC Analysis

1. Open the control panel of an RFQ.
2. Click **📊 SQDC** in the top toolbar.
3. The **"SQDC Analysis – RFQ N° [number]"** window opens.

If the RFQ has no suppliers, the SQDC button is unavailable.

---

## Tab 1 – Weights (%)

The first tab lets you set the relative importance of the four criteria:

1. Enter a value between 1 and 100 in each of the four fields (Safety, Quality, Delivery, Cost).
2. The four values **must sum to exactly 100**. If they do not, switching to the next tab is blocked.
3. The default values are **25% for each criterion** (uniform distribution).

**Example of custom weights:**
- Safety: 20%
- Quality: 30%
- Delivery: 20%
- Cost: 30%

---

## Tab 2 – Scores (1–10)

The second tab shows a table with one row per supplier and four score columns:

| Column | Type |
|--------|------|
| **Supplier** | Read-only (determined by the RFQ) |
| **Safety** | Enter a score from 1 to 10 |
| **Quality** | Enter a score from 1 to 10 |
| **Delivery** | Enter a score from 1 to 10 |
| **Cost** | Enter a score from 1 to 10 |
| **TOTAL** | Calculated automatically |

Scores must be **integers from 1 to 10**. Invalid values (letters, decimals, out-of-range numbers) are flagged and rejected.

The **TOTAL** column is calculated using the formula:

$$\text{TOTAL} = \frac{(S_{safety} \times W_{safety}) + (S_{quality} \times W_{quality}) + (S_{delivery} \times W_{delivery}) + (S_{cost} \times W_{cost})}{100}$$

Where $W$ are the weights set in the previous tab and $S$ are the scores entered.

Once all scores are complete, **the row of the highest-scoring supplier is highlighted in green**. In the event of a tie (difference less than 0.01), all tied suppliers are highlighted.

---

## Automatic Cost Score Calculation

The **🔄 Calculate Cost Automatically** button retrieves prices from the RFQ grid and calculates the Cost score proportionally:

- The supplier with the lowest total cost receives a **score of 10**.
- Other suppliers receive proportionally lower scores.
- Suppliers with missing prices receive a **score of 0**, and a red warning banner is displayed.

This function is useful as a starting point for the cost criterion; the buyer can still manually adjust the automatic score.

---

## Saving the Analysis

1. Complete the weights and all scores.
2. Click **💾 Save SQDC**.
3. The analysis is saved as an **Internal Document** attached to the RFQ.

After saving, the button in the RFQ toolbar changes to **📊 SQDC ✓** to indicate that a saved analysis exists.

---

## Exporting to Excel

Click **📊 Export Excel** in the SQDC window.

The generated Excel file contains:
- The weights / scores / totals matrix
- The winning supplier highlighted in green
- A missing-price warning, if applicable

---

## Read-Only Behaviour

If the RFQ belongs to another user, the SQDC window opens in read-only mode: you can review the saved weights and scores and export to Excel, but cannot make any changes.
