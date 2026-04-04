# 06 – Value Stream Mapping

## Overview

The **Value Stream Mapping** module lets you record and track economic improvement activities in procurement. Activities are divided into three types:

| Tab | Type | Description |
|-----|------|-------------|
| **Savings** | Cost reduction | Reduction of costs against budget or historical spend |
| **Cost Avoidance** | Preventing cost increases | Blocking a cost increase requested by a supplier or driven by market conditions |
| **Derisking** | Supply chain risk reduction | Qualifying new suppliers to reduce dependence on a single source |

---

## Creating a New Savings Event

1. Click the **Savings** tab.
2. Click **➕ New Event**.
3. The **"New Value Stream Mapping Event"** window opens with the type preset to "Savings".

### General Information Section

| Field | Notes |
|-------|-------|
| **Event Date** | The date on which the negotiation was concluded or formalised (required) |
| **Event Type** | Preset to "Savings"; not editable |
| **Action** | Select: `Negotiation` / `Derisking` / `Other` |
| **User** | Filled automatically with your name; not editable |

### Description Section

Free text field. Briefly describe the subject of the negotiation, the supplier, and the context.

### Reference Section

Text field to link the event to an RFQ, a PO, a contract, or a specific supplier.

### Economic Data Section – Price Driver

Fill this in when the saving comes from a unit price reduction:

| Field | Notes |
|-------|-------|
| **Budget Amount (€)** | Unit price or total budgeted amount (use a comma as the decimal separator) |
| **Negotiated Amount (€)** | Unit price or amount actually negotiated |
| **Annual Quantity** | Number of units per year. Default: 1 |
| **% Realisation** | Expected percentage of the saving that will actually be achieved. Default: 100 |

The theoretical value is calculated automatically on save:

$$V_{theoretical} = Q_{annual} \times (\text{Budget Amount} - \text{Negotiated Amount})$$

### Economic Data Section – Payment Terms Driver

Fill this in when the saving comes from an improvement in payment terms (e.g. from 30 to 90 days):

| Field | Notes |
|-------|-------|
| **Annual Spend (€)** | Annual spend with the supplier to which the payment saving applies |
| **Current Payment Terms (days)** | E.g. `30` |
| **Negotiated Payment Terms (days)** | E.g. `90` |
| **Financial Impact (% per 30 days)** | Financial coefficient per 30 days. Default: `0.50%` (configured in Settings) |

$$V_{theoretical} = \text{Annual Spend} \times \frac{(\text{Neg. Days} - \text{Curr. Days})}{30} \times \text{Coefficient}$$

---

## Creating a New Cost Avoidance Event

The workflow is identical to Savings. The difference is in the Economic Data – Price Driver section:

| Field | Notes |
|-------|-------|
| **Initial Requested Amount (€)** | Price or amount originally requested by the supplier |
| **Negotiated Amount (€)** | Amount actually agreed after negotiation |
| **Annual Quantity** | Default: 1 |
| **% Realisation** | Default: 100 |

$$V_{theoretical} = Q_{annual} \times (\text{Initial Requested Amount} - \text{Negotiated Amount})$$

> The Payment Terms driver is **not available** for Cost Avoidance events.

---

## Distributing Value Over Time (Recurring OPEX)

By default, the economic value of an event is recorded **entirely in the month of the event date** (one-time impact).

If the negotiation generates a benefit that will recur every month (typically OPEX: service contracts, subscription-based supply agreements, multi-year framework agreements), enable the option:

- **☑ Recurring OPEX (multi-month distribution)**

With this option active, DataFlow distributes the theoretical value over **up to 24 months** starting from the event month, applying a **first-month pro-rata**:

$$\text{First-month coefficient} = \frac{30 - \text{event day} + 1}{30}$$

**Practical example:**  
Savings from a maintenance service negotiation: €12,000 per year = €1,000/month.  
Event date: 15 March → first-month coefficient = (30 − 15 + 1) / 30 = 0.533  
- March: 1,000 × 0.533 = €533
- April through February (23 months): €1,000/month
- Last month: adjusted to ensure the total sum equals exactly €24,000

The actual value for each month = theoretical monthly value × (% Realisation / 100).

---

## The Derisking Tab – Potential Supplier Registry

The Derisking tab does not record economic events; instead, it builds a **registry of potential suppliers** for evaluation and qualification.

### Adding a New Potential Supplier

1. Click the **Derisking** tab.
2. Click **➕ New Event**.
3. The **"New Supplier"** window opens.

### Supplier Record Fields

| Section | Field | Notes |
|---------|-------|-------|
| General Information | **Supplier** | Registered company name (required) |
| | **Category** | Select from the existing category catalogue |
| | **New category** | Enter here if the category does not yet exist (created automatically) |
| | **Status** | `New` / `Under evaluation` / `Qualified` / `Rejected` |
| Contacts | **Contact** | Name of the commercial contact |
| | **Email** | Clickable to open the mail client |
| | **Phone** | |
| | **Website** | URL (clickable to open in browser) |
| Notes | | Free text |

Click **💾 Save** to save the record.

### Updating a Supplier's Status

1. Double-click the supplier row in the Derisking tab.
2. Change the **Status** field.
3. Click **💾 Save**.

Status typically progresses from `New` → `Under evaluation` → `Qualified` (or `Rejected`).

---

## Managing Supplier Categories

Categories allow you to group suppliers by product family.

### Accessing Category Management

In the supplier window, click **Manage Categories**.

### Renaming a Category

1. Select the category from the list.
2. Enter the new name in the **New name** field.
3. Click **Rename**.

The rename is **pending** until you click **💾 Save**.

### Merging Two Categories

1. Select the source category from the list.
2. In the **Merge with** field, choose the target category.
3. Click **Merge**.

All suppliers in the source category are moved to the target category. The source category is removed.

### Deleting a Category

A category can only be deleted if it **has no associated suppliers**. The counter **"Associated suppliers: N"** shows how many suppliers belong to the selected category. If N > 0, the Delete button is blocked.

All changes remain pending until **💾 Save** is clicked. Click **❌ Cancel** to discard all pending changes.

---

## Editing or Deleting an Event

1. Select the event row in the corresponding tab.
2. Click **⚡ Actions**:
   - Choose **Edit** to open the edit window.
   - Choose **Delete** to remove the event.

> Deletion also removes all monthly impacts associated with the event. This action cannot be undone.

When an event is edited, DataFlow **automatically recalculates and recreates** all monthly impacts. The previous calculation is discarded.

---

## Viewing Another User's Event

Events belonging to other users are visible in the list but open in **read-only mode**. The window displays all data, but all fields are disabled and the only available button is **✖ Close**.
