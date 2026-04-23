# 02 – Main Screen

## General Layout

The main screen is organised vertically into five areas:

| Area | Content |
|------|---------|
| Toolbar | Logo + operational buttons |
| Global search bar | Single search field + advanced filters toggle |
| Advanced filters | Collapsible panel with dedicated filter fields |
| Notebook | 5 tabs: Active RFQs, Archived RFQs, Savings, Cost Avoidance, Derisking |
| Footer | Version, author, licence |

---

## Toolbar

| Button | Action |
|--------|--------|
| **➕ New Event** | Creates a new RFQ (on RFQ tabs) or a new Value Stream Mapping event (on Savings / Cost Avoidance / Derisking tabs) |
| **⚡ Actions** | Context menu with actions for the selected row (active only when a row is selected) |
| **📥 Export Excel** | Exports the full RFQ list to an Excel file |
| **≋ KPI** | Opens the KPI analysis window |
| **⚙️ Settings** | Opens the application settings |
| **≡ License** | Displays the software licence |
| **❓ Help** | Opens the built-in user guide |

### The ⚡ Actions Menu

The Actions menu is **always disabled** until a row is selected. It becomes active when you click on an item:

- **On RFQ tabs**: allows you to delete, duplicate, archive, or reactivate the selected RFQ.
- **On Value Stream Mapping tabs (Savings / Cost Avoidance / Derisking)**: allows you to edit or delete the selected event.

---

## Global Search Bar

The wide search field at the centre of the bar is the fastest way to find anything. On RFQ tabs, it searches simultaneously across:

- RFQ number
- Project reference
- Supplier name
- Material code
- Drawing / attachment
- Material description
- Purchase order number
- Raw code
- Raw attachment
- Material for processing

**How to use it:**
1. Type a keyword (e.g. a supplier name, a part code, a project reference).
2. Press **Enter**.
3. Results are shown in the active tab. All other tabs update in parallel.
4. To clear the search, empty the field and press **Enter**.

On the Savings, Cost Avoidance, and Derisking tabs, the same search bar checks the main visible text fields of the active list.

The search is **case-insensitive** and uses **OR** logic: a result is shown if it contains the search text in at least one searched field.

---

## Advanced Filters

For more precise searches, click the **⌄ Advanced Filters** label (to the right of the search bar). The panel expands to reveal individual filter fields for each criterion.

> The Advanced Filters toggle is disabled when the **Derisking** tab is active, because supplier searching happens directly within the list.

### RFQ Filters

| Field | Description |
|-------|-------------|
| RFQ Number | Search by identifier |
| RFQ Type | Dropdown: All / Full Supply / Work Order |
| Reference | Free text on the project reference |
| Supplier | Supplier name (partial match supported) |
| Material Code | Item code |
| Material Description | Item description |
| PO Number | Purchase order number |
| Raw Code | Work Order RFQs only |
| Raw Attachment | Work Order RFQs only |
| Material for Processing | Work Order RFQs only |
| User | Filter by buyer (dropdown listing all users) |
| Issue Date From / To | Issue date range |
| Expiry Date From / To | Expiry date range |

### Value Stream Mapping Event Filters (Savings / Cost Avoidance tabs)

| Field | Description |
|-------|-------------|
| User | Buyer who owns the event |
| From / To | Event date range |
| Action | Negotiation / Derisking / Other |
| Recurring | Yes / No (filters recurring OPEX events) |
| Theoretical Value From / To | Theoretical amount range |
| Actual Value From / To | Actual amount range |

After setting the filters, click **🔍 Search**. To return to the full view, click **🔎 Clear Filters**.

Filters use **AND** logic: each active field adds an additional constraint to the results.

---

## The Five Main Tabs

### Active RFQs and Archived RFQs

Display the list of requests for quotation in a table with the following columns:

- **N°** – Sequential number assigned automatically
- **Type** – Full Supply / Work Order
- **Issue Date**
- **Expiry Date** – Expired RFQs are highlighted in red
- **Reference** – Project name or brief description
- **User** – Buyer who owns the RFQ

**Double-click** a row to open the RFQ control panel. RFQs belonging to other users open in **read-only mode** (red banner at the bottom: *"You are viewing an RFQ belonging to another user"*).

Columns are **sortable**: click a column header to sort ascending or descending.

### Savings and Cost Avoidance Tabs

Display the list of economic Value Stream Mapping events, with the following columns:

- Date, Type, Action, Description, Reference, Driver, Theoretical Value, Actual Value, % Realisation, Recurring, User

**Double-click** to open the event details. See the [Value Stream Mapping](06-Value-Stream-Mapping.md) section for full instructions.

### Derisking Tab

Displays the potential supplier registry, with the following columns:

- Supplier, Category, Status, Contact, Email, Phone, User

**Double-click** to open the supplier record.

---

## Sorting Columns

Click a column header to sort the list in ascending order. Click again to reverse the order. Sorting is visual only and does not affect the underlying data.

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Start a search | Type in the search field + **Enter** |
| Open an item | **Double-click** the row |
| Open the help guide | **❓ Help** button |

---
[← Previous page](EN-01-Getting-Started) | [Next page →](EN-03-Create-a-New-RFQ)
