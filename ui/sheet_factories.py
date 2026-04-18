"""Factories for dashboard tksheet widgets.

Pure UI construction helpers with explicit callbacks.
"""

from __future__ import annotations

from tkinter import ttk

from tksheet import Sheet, natural_sort_key

from utils.i18n_utils import tr


def create_cell_select_handler(update_button_visibility_cb):
    """Create a cell-select callback that refreshes action state."""

    def handler(_event_data):
        update_button_visibility_cb()

    return handler


def create_row_select_handler(update_button_visibility_cb):
    """Create a row-select callback that refreshes action state."""

    def handler(_event_data):
        update_button_visibility_cb()

    return handler


def create_request_sheet(*, parent, on_cell_select, on_row_select, on_double_click_cb):
    """Create RFQ request sheet with canonical dashboard behavior."""
    tree_frame = ttk.Frame(parent)
    tree_frame.pack(fill="both", expand=True)

    sheet = Sheet(
        tree_frame,
        theme="light blue",
        header_font=("Calibri", 11, "bold"),
        font=("Calibri", 11, "normal"),
        headers=[tr("RfQ No."), tr("RfQ Type"), tr("Issue Date"), tr("Expiry Date"), tr("Reference"), tr("User")],
        show_header=True,
        show_row_index=False,
    )

    sheet.set_column_widths([80, 120, 120, 120, 300, 140])
    sheet.align_columns(columns=[0, 1, 2, 3, 5], align="center")
    sheet.enable_bindings()

    for col_idx in range(6):
        sheet.readonly_columns(columns=[col_idx], readonly=True)

    sheet.extra_bindings("cell_select", on_cell_select)
    sheet.extra_bindings("row_select", on_row_select)
    # Se l'utente ordina la vista, i metadata esterni vanno riallineati al primo uso.
    def mark_metadata_dirty(_event_data):
        sheet._metadata_needs_resync = True
    sheet.extra_bindings("begin_sort_rows", mark_metadata_dirty)
    sheet._last_click_time = 0
    sheet._last_click_row = None
    sheet.bind("<Double-Button-1>", lambda event: on_double_click_cb(sheet, event))

    sheet.pack(fill="both", expand=True)
    sheet._sheet_data = []
    sheet._metadata_needs_resync = False
    return sheet


def create_vsm_event_sheet(*, parent, event_type, on_cell_select, on_row_select, on_double_click_cb):
    """Create VSM event sheet for Saving/Cost Avoidance tabs."""
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)

    if event_type == "Saving":
        headers = [
            tr("Date"), tr("Type"), tr("Action"), tr("Description"),
            tr("Theoretical Savings"), tr("Actual Savings"),
            tr("Realization %"), tr("Variance %"), tr("Repetitive"), tr("User"),
        ]
        align_cols = [0, 1, 2, 6, 7, 8, 9]
        amount_cols = [4, 5]
        n_cols = 10
    elif event_type == "Cost Avoidance":
        headers = [
            tr("Date"), tr("Type"), tr("Action"), tr("Description"),
            tr("CA Theoretical"), tr("CA Actual"),
            tr("Realization %"), tr("Variance %"), tr("Repetitive"), tr("User"),
        ]
        align_cols = [0, 1, 2, 6, 7, 8, 9]
        amount_cols = [4, 5]
        n_cols = 10
    else:
        headers = [
            tr("Date"), tr("New Supplier"), tr("Description"), tr("Repetitive"), tr("User"),
        ]
        align_cols = [0, 3, 4]
        amount_cols = []
        n_cols = 5

    header_padding = 30
    desc_col_idx = 2 if event_type is None else 3
    desc_col_width = 400
    date_col_idx = 0
    action_col_idx = 2 if event_type is not None else None
    action_min_width = 150
    type_col_idx = 1 if event_type is not None else None
    new_supplier_col_idx = 1 if event_type is None else None

    try:
        import tkinter.font as tkfont

        hfont = tkfont.Font(family="Calibri", size=11, weight="bold")
        cfont = tkfont.Font(family="Calibri", size=11)
        date_min = cfont.measure("dd/mm/YYYY") + header_padding
        type_min = hfont.measure(tr("Realization %")) + header_padding
        new_supplier_min = hfont.measure(tr("Theoretical Value")) + header_padding
        col_widths = [
            desc_col_width if i == desc_col_idx
            else max(date_min, hfont.measure(h) + header_padding) if i == date_col_idx
            else max(action_min_width, hfont.measure(h) + header_padding) if i == action_col_idx
            else max(type_min, hfont.measure(h) + header_padding) if i == type_col_idx
            else max(new_supplier_min, hfont.measure(h) + header_padding) if i == new_supplier_col_idx
            else max(60, hfont.measure(h) + header_padding)
            for i, h in enumerate(headers)
        ]
    except Exception:
        col_widths = [
            400 if i == desc_col_idx else 150 if i in (action_col_idx, type_col_idx) else 120
            for i in range(len(headers))
        ]

    sheet = Sheet(
        frame,
        theme="light blue",
        header_font=("Calibri", 11, "bold"),
        font=("Calibri", 11, "normal"),
        headers=headers,
        show_header=True,
        show_row_index=False,
    )

    sheet._vsm_event_type = event_type
    sheet._vsm_headers = headers
    sheet._vsm_col_widths = col_widths
    sheet._vsm_align_cols = align_cols
    sheet._vsm_amount_cols = amount_cols

    def currency_numeric_sort_key(value):
        if isinstance(value, (int, float)):
            return natural_sort_key(float(value))
        if not isinstance(value, str):
            return natural_sort_key(value)
        text = value.replace("\xa0", " ").strip()
        if not text:
            return natural_sort_key(value)
        normalized = text
        for prefix in ("CHF ", "$", "£"):
            if normalized.startswith(prefix):
                normalized = normalized[len(prefix):].strip()
                break
        if normalized.endswith("€"):
            normalized = normalized[:-1].strip()
        if "," in normalized and "." in normalized:
            if normalized.rfind(",") > normalized.rfind("."):
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
        elif "," in normalized:
            parts = normalized.split(",")
            if len(parts[-1]) in (1, 2):
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
        try:
            return natural_sort_key(float(normalized))
        except (TypeError, ValueError):
            return natural_sort_key(value)

    def configure_vsm_sort_key(_event_data):
        sheet._metadata_needs_resync = True
        selected = sheet.get_currently_selected()
        column = selected.column if selected is not None else None
        if column in amount_cols:
            sheet.set_options(redraw=False, sort_key=currency_numeric_sort_key)
        else:
            sheet.set_options(redraw=False, sort_key=natural_sort_key)

    sheet.extra_bindings("begin_sort_rows", configure_vsm_sort_key)
    sheet.set_column_widths(col_widths)
    sheet.align_columns(columns=align_cols, align="center")
    if amount_cols:
        sheet.align_columns(columns=amount_cols, align="right")

    sheet.enable_bindings()
    sheet.extra_bindings("cell_select", on_cell_select)
    sheet.extra_bindings("row_select", on_row_select)
    sheet.bind("<Double-Button-1>", lambda event: on_double_click_cb(sheet, event))

    for col_idx in range(n_cols):
        sheet.readonly_columns(columns=[col_idx], readonly=True)

    sheet.pack(fill="both", expand=True)
    sheet._event_metadata = []
    sheet._metadata_needs_resync = False
    return sheet


def create_supplier_sheet(*, parent, on_cell_select, on_row_select, on_double_click_cb):
    """Create Derisking supplier sheet."""
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)

    headers = [
        tr("Supplier"), tr("Category"), tr("Status"), tr("Contact"), tr("E-mail"),
        tr("Phone"), tr("Web"), tr("Notes"), tr("User"),
    ]
    align_cols = [2, 5, 8]
    n_cols = len(headers)

    try:
        import tkinter.font as tkfont

        hfont = tkfont.Font(family="Calibri", size=11, weight="bold")
        header_padding = 30
        note_width = 300
        note_idx = 7
        col_widths = [
            note_width if i == note_idx else max(80, hfont.measure(h) + header_padding)
            for i, h in enumerate(headers)
        ]
    except Exception:
        col_widths = [300 if i == 8 else 140 for i in range(n_cols)]

    sheet = Sheet(
        frame,
        theme="light blue",
        header_font=("Calibri", 11, "bold"),
        font=("Calibri", 11, "normal"),
        headers=headers,
        show_header=True,
        show_row_index=False,
    )

    sheet._vsm_headers = headers
    sheet._vsm_col_widths = col_widths
    sheet._vsm_align_cols = align_cols

    sheet.set_column_widths(col_widths)
    sheet.align_columns(columns=align_cols, align="center")

    sheet.enable_bindings()
    sheet.extra_bindings("cell_select", on_cell_select)
    sheet.extra_bindings("row_select", on_row_select)
    # Se l'utente ordina la vista, i metadata esterni vanno riallineati al primo uso.
    def mark_metadata_dirty(_event_data):
        sheet._metadata_needs_resync = True
    sheet.extra_bindings("begin_sort_rows", mark_metadata_dirty)
    sheet.bind("<Double-Button-1>", lambda event: on_double_click_cb(sheet, event))

    for col_idx in range(n_cols):
        sheet.readonly_columns(columns=[col_idx], readonly=True)

    sheet.pack(fill="both", expand=True)

    sheet._event_metadata = []
    sheet._supplier_metadata = []
    sheet._metadata_needs_resync = False
    return sheet
