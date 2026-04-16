"""Operational VSM command helpers extracted from MainWindow handlers."""

from __future__ import annotations

from database_manager import DatabaseManager
from services.app_paths import get_db_path


def status_to_event_type(status):
    """Map VSM tab status to event_type."""
    return {
        "vsm_saving": "Saving",
        "vsm_cost_avoidance": "Cost Avoidance",
    }.get(status)


def delete_vsm_events_by_ids(event_ids):
    """Delete VSM events (and impacts) by ids."""
    from services.vsm_persistence import delete_event_and_impacts

    with DatabaseManager(get_db_path()) as db_manager:
        for event_id in event_ids:
            delete_event_and_impacts(db_manager, event_id)


def delete_suppliers_by_ids(supplier_ids):
    """Delete Derisking suppliers by ids."""
    from services.supplier_persistence import delete_supplier

    with DatabaseManager(get_db_path()) as db_manager:
        for supplier_id in supplier_ids:
            delete_supplier(db_manager, supplier_id)


def duplicate_vsm_event_by_id(event_id):
    """Duplicate a VSM event 1:1 and return new event id."""
    from models.vsm_event import VSMEvent
    from services.vsm_persistence import get_event_with_impacts, save_event_with_impacts

    with DatabaseManager(get_db_path()) as db_manager:
        original_event, _impacts = get_event_with_impacts(db_manager, event_id)

        duplicate_event = VSMEvent(
            id=None,
            event_date=original_event.event_date,
            username=original_event.username,
            buyer=original_event.buyer,
            event_type=original_event.event_type,
            action=original_event.action,
            description=original_event.description,
            reference=original_event.reference,
            importo_bdg=original_event.importo_bdg,
            importo_negoziato=original_event.importo_negoziato,
            importo_richiesto_iniziale=original_event.importo_richiesto_iniziale,
            quantita_annua=original_event.quantita_annua,
            percent_realizzo=original_event.percent_realizzo,
            driver=original_event.driver,
            giorni_pagamento_attuali=original_event.giorni_pagamento_attuali,
            giorni_pagamento_negoziati=original_event.giorni_pagamento_negoziati,
            spending_annuo=original_event.spending_annuo,
            opex_ripetitivo=original_event.opex_ripetitivo,
            note=original_event.note,
        )

        return save_event_with_impacts(db_manager, duplicate_event)
