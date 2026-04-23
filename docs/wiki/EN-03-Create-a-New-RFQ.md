# 03 – Create a New RFQ

## RFQ Types

DataFlow handles two types of Request for Quotation:

| Type | When to use |
|------|-------------|
| **Full Supply** | Supply of components or finished products. The supplier provides everything (material + processing). |
| **Work Order** | The raw material is supplied by the company; the supplier performs the processing only. Requires entering the raw material code, attachment reference, and material description. |

---

## Starting the Creation Process

1. Click **➕ New Event** in the toolbar while on the **Active RFQs** or **Archived RFQs** tab.
2. In the type selection window, choose:
   - **📦 Full Supply**
   - **🔧 Work Order**
   - **❌ Cancel** to go back
3. The **control panel** for the new RFQ opens automatically.

---

## RFQ Control Panel

The window title reads: `Control Panel - User: [name] - Request N° [number] - [type]`.

### Header Fields

| Field | Notes |
|-------|-------|
| **Issue Date** | Date picker. Saved automatically on selection or when the field loses focus. |
| **Expiry Date** | Date picker. Same auto-save behaviour. |
| **Reference** | Clickable label. Click it to open the edit window and enter a reference (e.g. project name, work order, customer). |

### Top Toolbar Buttons

| Button | Action |
|--------|--------|
| **📄 Manage Supplier Quotes** | Opens the attachment manager for documents received from suppliers |
| **📁 Manage Internal Documents** | Opens the attachment manager for internal documents (specs, drawings, etc.) |
| **Suppliers (N)** or **➕ Add Suppliers** | Opens the window to enter or modify the supplier list |
| **📝 Note** or **📝 Add Note** | Opens the rich-text note editor (bold, italic, underline) |
| **📊 Export** | Opens the RFQ export menu with **Excel** and **PDF** options |
| **📊 SQDC** or **📊 SQDC ✓** | Opens the SQDC analysis (✓ indicates a saved analysis already exists) |

---

## Adding Suppliers

It is recommended to set up suppliers before entering prices in the grid:

1. Click **➕ Add Suppliers** (or **Suppliers (N)** if some already exist).
2. Enter supplier names in the text field, **separated by commas** (e.g. `Supplier A, Supplier B, Supplier C`).
3. Click **💾 Save**.

> DataFlow does not accept duplicate supplier names (case-insensitive). If the same name is entered twice, saving is blocked with a warning.
>
> While typing, DataFlow can suggest supplier names already used in RFQs or Derisking. Similar names may also trigger a non-blocking warning before saving.

Each supplier adds **one price column** to the grid.

---

## Entering Items Manually

The **price grid** occupies the central area of the control panel. To add items:

1. Click **➕ Add Item** (button at the bottom left).
2. A new empty row is added.
3. Click any cell and type the value.

### Columns for Full Supply RFQs

| Column | Content |
|--------|---------|
| **Pos.** | Position number (automatic) |
| **Drawing** | Reference to the technical drawing or document |
| **Description** | Item description |
| **Qty** | Required quantity (use a comma as the decimal separator) |
| **[Supplier 1]** | Unit price quoted by supplier 1 |
| **[Supplier 2…]** | One column per supplier entered |

### Additional Columns for Work Order RFQs

Three extra columns appear after the base columns:

| Column | Content |
|--------|---------|
| **Raw Code** | Code of the raw material supplied by the company |
| **Raw Attachment** | Attachment or drawing reference of the raw material |
| **Material for Processing** | Description of the material to be processed |

### Entering Prices

Prices are entered directly into the cells of the corresponding supplier column. Rules:

- Use a **comma** as the decimal separator (e.g. `12,50`)
- Do not use a period as the decimal separator
- Do not include the currency symbol (€)
- An empty cell means "no quote received"

---

## Importing Items from Excel

If an Excel file with the item list already exists, rows can be imported without entering them manually:

1. Click **📊 Import from Excel** (button at the bottom, next to Add Item).
2. Select the Excel file in the file dialog.
3. The system reads the file and inserts the rows into the grid.

> The Excel file must match the expected format (the same format produced by DataFlow's **Export → Excel** option). If the format is invalid, an error message will describe the problem.

---

## Removing an Item

1. Click the row to select it.
2. Click **🗑 Remove Selected Item**.

> Removal is immediate and does not ask for confirmation. Proceed with care.

---

## Saving the RFQ

The RFQ is **saved automatically** whenever a value is changed (date, reference, price in the grid). There is no explicit Save button. You can close the panel at any time.

---

## Adding Notes to the RFQ

For negotiation context, received communication summaries, or technical observations:

1. Click **📝 Add Note**.
2. Type your text in the editor. Use the formatting buttons: **𝐁** bold, **𝑰** italic, **U̲** underline.
3. Click **💾 Save Note**.

Notes support formatted text with styles. The note window disables the underlying screen while editing; it is re-enabled upon closing.

---

## Entering a Purchase Order Number

Once a purchase order is issued as a result of the RFQ:

1. Click **📋 Insert PO** in the panel header.
2. Enter the **PO Number** in the text field.
3. Select the **Supplier** from the dropdown (only suppliers already on the RFQ are available).
4. Click **➕ Add**.
5. Click **Close** to save.

Multiple orders can be linked to the same RFQ (one per supplier or per delivery tranche).

---
[← Previous page](EN-02-Main-Screen) | [Next page →](EN-04-Manage-an-Existing-RFQ)
