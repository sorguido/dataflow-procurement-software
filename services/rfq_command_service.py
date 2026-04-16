"""RFQ command operations extracted from MainWindow handlers."""

from __future__ import annotations

import logging
import os

from database_manager import DatabaseManager
from services.app_paths import get_db_path


logger = logging.getLogger(__name__)


def update_request_status(*, request_ids, new_status):
    """Bulk update RFQ status for selected ids."""
    params = [(new_status, req_id) for req_id in request_ids]
    with DatabaseManager(get_db_path()) as db_manager:
        db_manager.update_stato_richieste(params)


def delete_requests_with_attachments(*, request_ids, archive_path):
    """Delete RFQs and attempt to delete external attachment files first."""
    with DatabaseManager(get_db_path()) as db_manager:
        for req_id in request_ids:
            try:
                rows = db_manager.conn.execute(
                    "SELECT percorso_esterno FROM allegati_richiesta WHERE id_richiesta = ? AND percorso_esterno IS NOT NULL",
                    (req_id,),
                ).fetchall()
            except Exception:
                rows = []

            for row in rows:
                percorso = row[0]
                if not percorso:
                    continue
                if archive_path and not os.path.isabs(percorso):
                    file_to_delete = os.path.join(archive_path, percorso)
                else:
                    file_to_delete = percorso

                try:
                    if os.path.exists(file_to_delete):
                        os.remove(file_to_delete)
                        logger.info("Allegato eliminato dal disco durante cancellazione RdO: %s", file_to_delete)
                    else:
                        logger.info("File allegato non trovato durante cancellazione RdO: %s", file_to_delete)
                except Exception as disk_error:
                    logger.warning("Impossibile eliminare il file allegato %s: %s", file_to_delete, disk_error)

        return db_manager.delete_richieste_batch(request_ids)


def duplicate_request_full(*, original_id):
    """Duplicate a RFQ fully and return the new request id."""

    def get_columns(table_name, exclude):
        with DatabaseManager(get_db_path()) as db_mgr:
            cols_info = db_mgr.get_table_columns(table_name)
        excluded = set(exclude)
        return [row[1] for row in cols_info if row[1] not in excluded]

    with DatabaseManager(get_db_path()) as db_manager:
        return db_manager.duplicate_richiesta_full(original_id, get_columns)


def create_request_shell(*, tipo_rdo, status, issue_date, username):
    """Create a minimal RFQ head row and return its generated id."""
    with DatabaseManager(get_db_path()) as db_manager:
        return db_manager.insert_richiesta_offerta(
            tipo_rdo,
            status,
            issue_date,
            username=username,
        )
