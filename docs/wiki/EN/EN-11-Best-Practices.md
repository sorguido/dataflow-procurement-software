# 11 – Best Practices

## RFQ Management

### Use the Reference field consistently

The **Reference** field is the primary free-text search field. Always enter a value that will make the RFQ easy to find in the future:
- The project or job order name
- A customer code or contract number
- A short but unique description

Avoid generic entries like "Various", "Test", or "April quote".

### Archive closed RFQs

Once an RFQ is complete — order placed, tender closed, or no longer relevant — **archive it** rather than leaving it in the Active RFQs tab. This reduces clutter in the main view and improves search performance.

### Always enter purchase order numbers

As soon as a purchase order is issued, enter the PO number in the RFQ using the **📋 Insert PO** button. This lets you search by PO number and trace the pipeline from RFQ to order.

### Attach received quotes

Save supplier quotes as attachments in the RFQ. This centralises all tender documentation and makes it accessible to the entire team, even months later.

---

## Recording in the Value Stream Mapping Module

### Record all savings, including small ones

The Value Stream Mapping module is most useful when it is complete. Record even modest savings, as they contribute to aggregate statistics and to management KPI reports.

### Choose the event date carefully

The event date determines how monthly impacts are distributed. Use the actual date on which the negotiation was formally concluded — the contract signing date, the approval date for a new price list, or the date of the supplier's confirmation email.

### Use Recurring OPEX for multi-year contracts

Whenever a negotiation produces a benefit that will recur every month (services, subscription-based supplies, multi-year agreements), enable **☑ Recurring OPEX**. This distributes the value across the actual months of economic competence and makes KPI charts significantly more accurate.

### Set % Realisation realistically

% Realisation is not a cosmetic field: it directly affects the **Effective Savings** figure shown in the KPIs. If a negotiation is expected to materialise only partially — for example, because the order has not yet been placed, or because the counterpart has not yet confirmed — set a value below 100 and update it once the outcome is definitive.

### Use the Reference field to link Value Stream Mapping events to RFQs

In the **Reference** field of a Value Stream Mapping event, enter the corresponding RFQ number (e.g. `RFQ-124` or `RFQ 124`). This makes it easy to find the event from the RFQ and vice versa.

---

## Value Stream Mapping – Derisking

### Keep potential supplier statuses up to date

The value of the Derisking register depends on the quality of the updates. When a supplier advances in the evaluation process, update the **Status** field in the supplier record. An up-to-date status lets the Derisking KPIs give an accurate picture of the supplier portfolio.

### Use category names consistently

Before adding new suppliers, check whether the required category already exists in the catalogue. Using consistent naming avoids duplicates (`CNC Machining` vs `Machining CNC`) and keeps the per-category KPI charts accurate.

---

## Filters and Search

### Use global search for quick lookups

To quickly find an RFQ when you remember the supplier name or a material code, type directly into the global search field. This is the fastest approach.

### Use advanced filters for periodic reports

For monthly or quarterly reviews, use the **advanced filters** with an issue date range. At the end of the year, for example, filter by "Issue Date: from 01/01 to 31/12" to see all RFQs for the year with the correct totals.

### Always clear filters after a specific search

After using particular filters for a focused search, click **🔎 Clear Filters** to return to the full view. Leaving filters active can create the false impression that data is missing.

---

## Backup

### Run a manual backup before any critical operation

Before:
- Changing the database location
- Migrating to a new server
- Updating DataFlow to a new version

...always run a manual backup from **⚙️ Settings → 💾 Manual Backup**.

### Configure automatic backup

Enable the daily automatic backup and point it to a folder different from the database folder — ideally on a separate drive or server. This ensures a recent copy is always available in case of hardware failure.

---

## Periodic Cleanup

Every three to six months, it is good practice to:

1. **Archive any obsolete RFQs** still sitting in the Active tabs.
2. **Review and update potential suppliers** in the Derisking tab — remove definitively rejected suppliers if they are no longer needed for historical reference, or keep them for audit purposes.
3. **Revisit Value Stream Mapping events with a persistently low % Realisation** — either update the figure if the negotiation has since been confirmed, or document why the realisation remained partial.
