"""Selection and ownership policy helpers for dashboard sheets."""

from __future__ import annotations

import logging


logger = logging.getLogger(__name__)


def get_selected_row_indices(sheet):
    """Return selected row indices from tksheet, supporting cell or row selection."""
    row_indices = []

    currently_selected = sheet.get_currently_selected()
    if currently_selected:
        if hasattr(currently_selected, "row") and currently_selected.row is not None:
            row_indices.append(currently_selected.row)
        elif isinstance(currently_selected, tuple) and len(currently_selected) >= 1:
            row_indices.append(currently_selected[0])

    if not row_indices:
        selected_rows = sheet.get_selected_rows()
        if selected_rows:
            if isinstance(selected_rows, (list, set, tuple)):
                row_indices.extend(selected_rows)
            else:
                row_indices.append(selected_rows)

    return row_indices


def check_all_selected_are_mine(*, sheet, selected_indices, metadata_attr, owner_key="is_mine", entity_label="record"):
    """Check ownership for selected rows using sheet metadata.

    Returns False when metadata is missing or no rows are selected (fail-safe).
    """
    if not selected_indices:
        return False

    if not hasattr(sheet, metadata_attr):
        logger.warning("Metadati %s non disponibili - blocco operazioni per sicurezza", metadata_attr)
        return False

    metadata_rows = getattr(sheet, metadata_attr)

    for idx in selected_indices:
        if idx >= len(metadata_rows):
            logger.warning(
                "Indice %s %d fuori range metadati (len=%d)",
                entity_label,
                idx,
                len(metadata_rows),
            )
            continue

        if not metadata_rows[idx].get(owner_key, False):
            return False

    return True
