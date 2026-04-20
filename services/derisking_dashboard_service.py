"""Derisking supplier sheet helpers."""

from __future__ import annotations

from database_manager import DatabaseManager
from services.app_paths import get_db_path
from services.supplier_persistence import get_all_suppliers


def get_derisking_dataset(*, derisking_username_filter, current_username):
    """Load Derisking supplier dataset according to active username scope."""
    with DatabaseManager(get_db_path()) as db_manager:
        if derisking_username_filter is None:
            raw = db_manager.get_all_potential_suppliers_aggregated(get_db_path())
            suppliers = [supplier for supplier, _src in raw]
            extra_meta = [
                {
                    "is_mine": (
                        src == "local"
                        and (supplier.username or "").lower() == (current_username or "").lower()
                    ),
                    "source_file": src,
                }
                for supplier, src in raw
            ]
        elif derisking_username_filter == (current_username or "").lower():
            suppliers = get_all_suppliers(db_manager, username=current_username)
            extra_meta = None
        else:
            raw = db_manager.get_all_potential_suppliers_aggregated(
                get_db_path(),
                username=derisking_username_filter,
            )
            suppliers = [supplier for supplier, _src in raw]
            extra_meta = [
                {
                    "is_mine": (
                        src == "local"
                        and (supplier.username or "").lower() == (current_username or "").lower()
                    ),
                    "source_file": src,
                }
                for supplier, src in raw
            ]
    return suppliers, extra_meta


def build_supplier_rows_and_metadata(*, suppliers, current_username, translate_status, extra_metadata=None):
    """Build sheet rows and ownership metadata for PotentialSupplier records."""
    data_rows = []
    metadata = []

    for i, supplier in enumerate(suppliers):
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
        source_file = "local"
        if extra_metadata is not None and i < len(extra_metadata):
            _meta = extra_metadata[i] or {}
            is_mine = _meta.get("is_mine", is_mine)
            source_file = _meta.get("source_file", "local")
        metadata.append({
            "supplier_id": supplier.id,
            "username": supplier.username or "",
            "is_mine": is_mine,
            "source_file": source_file,
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
    # Baseline per riallineamento metadata dopo sort visuale (tksheet sort nativo).
    sheet._supplier_metadata_source = [dict(meta) for meta in metadata]
    sheet._supplier_rows_data_source = [tuple(row) for row in data_rows]
    sheet._metadata_needs_resync = False
    sheet._event_metadata = []
