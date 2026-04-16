"""RFQ dashboard data helpers extracted from MainWindow."""

from __future__ import annotations

import logging

from database_manager import DatabaseManager
from services.app_paths import get_db_path


logger = logging.getLogger(__name__)


def load_requests_by_status(*, status, username_filter, tipo_canonico=None, pre_fetched_rows=None):
    """Load RFQ rows from aggregated datasets and apply canonical filters."""
    if pre_fetched_rows is not None:
        all_rows = pre_fetched_rows
        logger.info("[MULTI-DB] Caricamento da dati pre-caricati (filtro utente: %s)...", username_filter)
    else:
        logger.info("[MULTI-DB] Caricamento da tutti i database (filtro utente: %s)...", username_filter)
        with DatabaseManager(get_db_path()) as db_manager:
            all_rows = db_manager.get_all_richieste_aggregated(get_db_path())

    filtered_rows = [row for row in all_rows if row[6] == status]

    if username_filter is not None:
        filtered_rows = [row for row in filtered_rows if row[5] and row[5].lower() == username_filter.lower()]
        logger.info(
            "[MULTI-DB] Trovate %d RdO in stato '%s' per utente '%s'",
            len(filtered_rows),
            status,
            username_filter,
        )
    else:
        logger.info("[MULTI-DB] Trovate %d RdO in stato '%s' da tutti gli utenti", len(filtered_rows), status)

    if tipo_canonico:
        filtered_rows = [row for row in filtered_rows if row[1] == tipo_canonico]
        logger.info("[MULTI-DB] Filtro tipo RdO '%s' applicato: %d risultati", tipo_canonico, len(filtered_rows))

    return filtered_rows


def build_rfq_sheet_payload(*, requests, translate_rfq_type, format_date_for_display):
    """Build RFQ tksheet rows and metadata preserving current semantics."""
    data_rows = []
    metadata_rows = []
    max_ref_length = 0

    for i, req in enumerate(requests):
        tipo_rdo_tradotto = translate_rfq_type(req[1])
        riferimento = req[4] if req[4] else ""
        username_value = ""
        if len(req) > 5 and req[5]:
            username_value = str(req[5]).strip()

        if len(req) > 8:
            is_mine = req[7]
            source_file = req[8]
            logger.debug("Riga %d (ID %s): is_mine=%s, source=%s", i, req[0], is_mine, source_file)
        else:
            is_mine = True
            source_file = "local"
            if len(req) < 6:
                logger.warning(
                    "Riga %d: tuple troppo corta (%d elementi), dati incompleti. Usando default is_mine=True",
                    i,
                    len(req),
                )

        metadata_rows.append({
            "is_mine": is_mine,
            "source_file": source_file,
        })

        if riferimento:
            max_ref_length = max(max_ref_length, len(riferimento))

        data_rows.append([
            str(req[0]),
            tipo_rdo_tradotto,
            format_date_for_display(req[2]),
            format_date_for_display(req[3]),
            riferimento,
            username_value,
        ])

    return data_rows, metadata_rows, max_ref_length
