# 04 – Manage an Existing RFQ

## Opening an RFQ

1. In the **Active RFQs** tab (or **Archived RFQs**), locate the desired RFQ.
2. **Double-click** the row to open the control panel.

If the RFQ belongs to another user, it opens in **read-only mode**: all edit fields and buttons are disabled, and a red banner appears at the bottom of the window. You can review all data but cannot make changes.

---

## Viewing and Editing the Reference

The **Reference** field in the header acts as a clickable label:

1. Click the reference label.
2. A small window opens with a pre-filled text field.
3. Edit the text and click **💾 Save**, or click **❌ Cancel** to discard the change.

---

## Editing the Price Grid

The price grid is fully editable:

- **Click a cell** to select it and edit its content.
- Use **Tab** or **Enter** to move to the next cell.
- Prices use a comma as the decimal separator (`12,50`).
- Empty price cells indicate that no quote was received from that supplier.

For Subcontracting RFQs, the **Raw Code**, **Raw Drawing**, and **Sub. Material** columns are editable in the same way as the others.

---

## Adding and Removing Suppliers

The suppliers already on the RFQ determine the price columns. To modify them:

1. Click **Suppliers (N)** at the top.
2. Edit the supplier list (names separated by commas).
3. Click **💾 Save**.

> Warning: **removing a supplier from the list also deletes all prices entered for that supplier**. This action cannot be undone.

---

## Managing Attachments

DataFlow distinguishes two types of attachments:

| Type | When to use |
|------|-------------|
| **Supplier Quote** | Quotes received in PDF, Excel, or other formats |
| **Internal Document** | Technical specifications, drawings, authorisations, SQDC analyses |

### Adding an Attachment

1. Click **📄 Manage Supplier Quotes** or **📁 Manage Internal Documents**.
2. For supplier quotes: first select the supplier from the dropdown.
3. Click **➕ Add...**
4. Select the file in the file dialog.
5. The file is copied to the `Attachments/{RFQ number}/` folder and the relative path is saved in the database.

> The original file is **not moved or deleted**. DataFlow keeps its own copy.

### Opening an Attachment

1. Click the attachment row to select it.
2. Click **📂 Open Selected**.
3. The file opens with the operating system's default application.

### Downloading an Attachment

1. Select the attachment row.
2. Click **⬇️ Download...**
3. Choose the destination folder.

### Deleting an Attachment

1. Select the attachment row.
2. Click **❌ Delete Selected**.

> For RFQs belonging to other users, the Add and Delete buttons are disabled, but opening and downloading attachments always remain available.

---

## Formatted Notes

Notes allow you to record negotiation context, received communications, or technical considerations:

1. Click **📝 Note** (or **📝 Add Note** if none exists yet).
2. Type your text in the editor. You can apply:
   - **𝐁 Bold**
   - **𝑰 Italic**
   - **U̲ Underline**
3. Click **💾 Save Note**.

Notes have no practical length limit (technical cap: 1 MB of content). Notes with more than 10,000 internal formatting elements cannot be saved.

---

## Purchase Order Numbers

To record orders issued as a result of the negotiation:

1. Click **📋 Insert PO**.
2. Enter the order number and select the supplier.
3. Click **➕ Add**.
4. The table immediately updates with the new order.
5. Click **Close** — data is saved automatically on closing.

To edit an existing order, double-click its cell in the table. To delete it, select the row and click **❌ Delete**.

---

## Archiving an RFQ

Completed or closed RFQs can be archived:

1. Go to the **Active RFQs** tab.
2. Select the RFQ (single click).
3. Click **⚡ Actions** → **📦 Archive**.

The RFQ will no longer appear in the Active RFQs tab, but can still be found in **Archived RFQs**. To reactivate it, select it in the archived tab and use **⚡ Actions** → **♻️ Reactivate**.

---

## Duplicating an RFQ

To create a new RFQ based on an existing one (same items, same suppliers):

1. Select the RFQ to copy.
2. Click **⚡ Actions** → **📋 Duplicate**.

A copy is created with a new issue date and sequential number. Prices entered in the grid are copied. Notes and attachments are **not** copied.

---

## Deleting an RFQ

1. Select the RFQ.
2. Click **⚡ Actions** → **🗑 Delete**.
3. Confirm in the dialog window.

> Deletion is **irreversible** and removes all items, prices, notes, and PO metadata. Physical attachment files remain in the `Attachments/` folder and must be removed manually if no longer needed.

---

## Exporting the Price Grid to Excel

To share the price comparison with colleagues or management:

1. In the RFQ panel, click **📊 Export Excel**.
2. Choose the file language (Italian / English).
3. Select the destination folder and file name.
4. The Excel file is generated with bold headers, grey background, formatted prices, and one column per supplier.

The **📊 Export Excel** button in the main toolbar exports **all** RFQs in the database into a single file.
