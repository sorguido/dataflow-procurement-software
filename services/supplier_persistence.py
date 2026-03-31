"""
Potential Supplier Persistence Layer

CRUD e query di aggregazione per l'anagrafica fornitori potenziali.
Entità separata dal modulo VSM: nessuna dipendenza da vsm_persistence, VSMEvent
o vsm_events.

Pattern di utilizzo:
    from services.supplier_persistence import (
        create_supplier,
        update_supplier,
        get_supplier_by_id,
        get_all_suppliers,
        delete_supplier,
        get_distinct_macrocategories,
        get_supplier_kpi,
    )
"""

import logging
from datetime import date as _date
from typing import Optional

from database_manager import DatabaseManager, DatabaseError
from models.potential_supplier import PotentialSupplier

logger = logging.getLogger('DataFlow.SupplierPersistence')


class SupplierError(Exception):
    """Eccezione per errori di business logic del modulo fornitori potenziali."""
    pass


# ---------------------------------------------------------------------------
# CRUD
# ---------------------------------------------------------------------------

def create_supplier(db_manager: DatabaseManager, supplier: PotentialSupplier) -> int:
    """
    Inserisce un nuovo fornitore potenziale nel database.

    Args:
        db_manager: Istanza attiva di DatabaseManager
        supplier:   PotentialSupplier con id=None e supplier_name valorizzato

    Returns:
        int: supplier_id assegnato dal database

    Raises:
        SupplierError: se supplier_name vuoto o supplier ha già un id
        DatabaseError: se errore DB
    """
    if supplier.id is not None:
        raise SupplierError(
            f"create_supplier richiede supplier.id=None, ricevuto id={supplier.id}"
        )
    if not supplier.supplier_name or not supplier.supplier_name.strip():
        raise SupplierError("supplier_name obbligatorio e non può essere vuoto.")

    logger.info(
        "Creazione fornitore potenziale: name=%s, username=%s",
        supplier.supplier_name,
        supplier.username,
    )
    try:
        supplier_id = db_manager.insert_potential_supplier(supplier)
        logger.info("Fornitore potenziale creato con ID %s", supplier_id)
        return supplier_id
    except DatabaseError:
        raise
    except Exception as e:
        raise SupplierError(f"Errore durante la creazione del fornitore: {e}") from e


def update_supplier(db_manager: DatabaseManager, supplier: PotentialSupplier) -> None:
    """
    Aggiorna un fornitore potenziale esistente.

    Args:
        db_manager: Istanza attiva di DatabaseManager
        supplier:   PotentialSupplier con id valido

    Raises:
        SupplierError: se supplier.id è None/0 o supplier_name vuoto
        DatabaseError: se errore DB
    """
    if not supplier.id:
        raise SupplierError(
            "update_supplier richiede supplier.id valido (intero > 0)."
        )
    if not supplier.supplier_name or not supplier.supplier_name.strip():
        raise SupplierError("supplier_name obbligatorio e non può essere vuoto.")

    logger.info("Aggiornamento fornitore potenziale ID %s", supplier.id)
    try:
        db_manager.update_potential_supplier(supplier)
        logger.info("Fornitore potenziale ID %s aggiornato", supplier.id)
    except DatabaseError:
        raise
    except Exception as e:
        raise SupplierError(f"Errore durante l'aggiornamento del fornitore: {e}") from e


def get_supplier_by_id(
    db_manager: DatabaseManager, supplier_id: int
) -> Optional[PotentialSupplier]:
    """
    Recupera un fornitore potenziale per ID.

    Returns:
        PotentialSupplier oppure None se non trovato
    """
    row = db_manager.get_potential_supplier_by_id(supplier_id)
    if row is None:
        return None
    return PotentialSupplier.from_row(row)


def get_all_suppliers(
    db_manager: DatabaseManager, username: str = None
) -> list:
    """
    Restituisce tutti i fornitori potenziali come lista di PotentialSupplier.

    Args:
        username: se fornito, filtra per utente; None = tutti

    Returns:
        list[PotentialSupplier] ordinata per supplier_name
    """
    rows = db_manager.get_all_potential_suppliers(username=username)
    return [PotentialSupplier.from_row(r) for r in rows]


def delete_supplier(db_manager: DatabaseManager, supplier_id: int) -> None:
    """
    Elimina un fornitore potenziale per ID.

    Raises:
        SupplierError: se supplier_id non valido
        DatabaseError: se errore DB
    """
    if not supplier_id or supplier_id <= 0:
        raise SupplierError("supplier_id deve essere un intero > 0.")

    logger.info("Eliminazione fornitore potenziale ID %s", supplier_id)
    try:
        db_manager.delete_potential_supplier(supplier_id)
        logger.info("Fornitore potenziale ID %s eliminato", supplier_id)
    except DatabaseError:
        raise
    except Exception as e:
        raise SupplierError(f"Errore durante l'eliminazione del fornitore: {e}") from e


def get_distinct_macrocategories(db_manager: DatabaseManager) -> list:
    """
    Restituisce la lista ordinata di macrocategorie distinte (non vuote).

    Returns:
        list[str]
    """
    return db_manager.get_distinct_macrocategories()


# ---------------------------------------------------------------------------
# QUERY KPI — backend pronto per futuri calcoli nella KPI window
# ---------------------------------------------------------------------------

def get_supplier_kpi(
    db_manager: DatabaseManager,
    username: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
) -> dict:
    """
    Calcola i KPI aggregati per i fornitori potenziali.

    Parametri filtro (tutti opzionali):
        username:   filtra per utente specifico (None = tutti)
        date_from:  filtro su created_at (formato ISO YYYY-MM-DD)
        date_to:    filtro su created_at (formato ISO YYYY-MM-DD)

    Returns::

        {
            "total_suppliers":           int,   # totale record
            "distinct_macrocategories":  int,   # macrocategorie distinte (non vuote)
            "unclassified_suppliers":    int,   # senza macrocategoria
            "by_status":                 dict,  # {status_str: count}
            "by_macrocategory":          dict,  # {macrocategory_str: count}
            "new_in_period":             int,   # inseriti nel periodo date_from/date_to
        }

    Non lancia eccezioni: in caso di errore restituisce valori a zero.
    """
    result = {
        "total_suppliers":          0,
        "distinct_macrocategories": 0,
        "unclassified_suppliers":   0,
        "by_status":                {},
        "by_macrocategory":         {},
        "new_in_period":            0,
    }

    try:
        cursor = db_manager.cursor

        # --- Clausole base ---
        base_clauses = []
        base_params = []
        if username:
            base_clauses.append("username = ?")
            base_params.append(username)

        def _where(clauses):
            if not clauses:
                return ""
            return "WHERE " + " AND ".join(clauses)

        w_base = _where(base_clauses)

        # 1. Totale fornitori
        cursor.execute(
            f"SELECT COUNT(*) FROM potential_suppliers {w_base}",
            tuple(base_params),
        )
        row = cursor.fetchone()
        result["total_suppliers"] = row[0] if row and row[0] is not None else 0

        # 2. Macrocategorie distinte (non vuote)
        cursor.execute(
            f"""
            SELECT COUNT(DISTINCT macrocategory)
            FROM potential_suppliers
            {_where(base_clauses + ["macrocategory IS NOT NULL", "TRIM(macrocategory) != ''"])}
            """,
            tuple(base_params),
        )
        row = cursor.fetchone()
        result["distinct_macrocategories"] = row[0] if row and row[0] is not None else 0

        # 3. Fornitori senza classificazione (macrocategory vuota o NULL)
        cursor.execute(
            f"""
            SELECT COUNT(*)
            FROM potential_suppliers
            {_where(base_clauses + ["(macrocategory IS NULL OR TRIM(macrocategory) = '')"])}
            """,
            tuple(base_params),
        )
        row = cursor.fetchone()
        result["unclassified_suppliers"] = row[0] if row and row[0] is not None else 0

        # 4. Conteggio per stato fornitore
        cursor.execute(
            f"""
            SELECT supplier_status, COUNT(*) AS cnt
            FROM potential_suppliers
            {w_base}
            GROUP BY supplier_status
            ORDER BY cnt DESC
            """,
            tuple(base_params),
        )
        result["by_status"] = {
            (row[0] or ""): row[1] for row in cursor.fetchall()
        }

        # 5. Conteggio per macrocategoria (inclusa "" per non classificati)
        cursor.execute(
            f"""
            SELECT COALESCE(macrocategory, '') AS mc, COUNT(*) AS cnt
            FROM potential_suppliers
            {w_base}
            GROUP BY mc
            ORDER BY cnt DESC
            """,
            tuple(base_params),
        )
        result["by_macrocategory"] = {
            row[0]: row[1] for row in cursor.fetchall()
        }

        # 6. Nuovi fornitori nel periodo (filtro su created_at)
        period_clauses = list(base_clauses)
        period_params = list(base_params)
        if date_from:
            period_clauses.append("created_at >= ?")
            period_params.append(date_from)
        if date_to:
            period_clauses.append("created_at <= ?")
            period_params.append(date_to + "T23:59:59")  # inclusive end-of-day
        cursor.execute(
            f"SELECT COUNT(*) FROM potential_suppliers {_where(period_clauses)}",
            tuple(period_params),
        )
        row = cursor.fetchone()
        result["new_in_period"] = row[0] if row and row[0] is not None else 0

    except Exception as e:
        logger.error("get_supplier_kpi: %s", e, exc_info=True)

    return result
