"""Derisking supplier sheet helpers."""

from __future__ import annotations


def build_supplier_rows_and_metadata(*, suppliers, current_username, translate_status):
    """Build sheet rows and ownership metadata for PotentialSupplier records."""
    data_rows = []
    metadata = []

    for supplier in suppliers:
        data_rows.append([
            supplier.supplier_name or "",
            supplier.category or "",
            translate_status(supplier.supplier_status) if supplier.supplier_status else "",
            supplier.contact_name or "",
            supplier.email or "",
            supplier.phone or "",
            supplier.website or "",
            supplier.notes or "",
            supplier.username or "",
        ])
        is_mine = (supplier.username or "").lower() == (current_username or "").lower()
        metadata.append({
            "supplier_id": supplier.id,
            "username": supplier.username or "",
            "is_mine": is_mine,
        })

    return data_rows, metadata


def auto_size_supplier_sheet(*, sheet, data_rows, notes_header_text):
    """Apply dynamic column widths to supplier sheet (Notes fixed after first calc)."""
    headers = getattr(sheet, "_vsm_headers", [])
    if not headers:
        return

    note_idx = next((i for i, h in enumerate(headers) if h == notes_header_text), None)

    try:
        import tkinter.font as tkfont

        cell_font = tkfont.Font(family="Calibri", size=11, weight="normal")
        header_font = tkfont.Font(family="Calibri", size=11, weight="bold")
    except Exception:
        if hasattr(sheet, "_vsm_col_widths"):
            sheet.set_column_widths(sheet._vsm_col_widths)
        return

    padding = 20

    if note_idx is not None and not hasattr(sheet, "_note_col_width"):
        base_widths = getattr(sheet, "_vsm_col_widths", None)
        if base_widths and note_idx < len(base_widths):
            sheet._note_col_width = int(base_widths[note_idx] * 1.5)
        else:
            sheet._note_col_width = 450

    widths = []
    for col_idx, header_text in enumerate(headers):
        if col_idx == note_idx:
            widths.append(sheet._note_col_width)
            continue

        width = header_font.measure(header_text) + padding
        for row in data_rows:
            if col_idx < len(row):
                cell_width = cell_font.measure(str(row[col_idx])) + padding
                if cell_width > width:
                    width = cell_width

        widths.append(max(80, width))

    sheet.set_column_widths(widths)


def populate_supplier_sheet(*, sheet, data_rows, metadata, resize_columns, notes_header_text):
    """Populate Derisking sheet preserving metadata semantics."""
    sheet.set_sheet_data(data_rows, reset_col_positions=False)
    if resize_columns:
        auto_size_supplier_sheet(sheet=sheet, data_rows=data_rows, notes_header_text=notes_header_text)
    if hasattr(sheet, "_vsm_align_cols"):
        sheet.align_columns(columns=sheet._vsm_align_cols, align="center")
    sheet.redraw()

    sheet._supplier_metadata = metadata
    sheet._event_metadata = []
