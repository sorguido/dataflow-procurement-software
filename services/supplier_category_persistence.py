"""
Supplier Category Persistence Layer

Gestione centralizzata delle categorie per i fornitori potenziali.
Le operazioni agiscono sulla tabella `supplier_categories` (catalogo ufficiale)
e sincronizzano potential_suppliers.category dove necessario.

Entità separata da vsm_persistence e supplier_persistence.
"""

import logging
from database_manager import DatabaseManager, DatabaseError

logger = logging.getLogger('DataFlow.SupplierCategoryPersistence')


class CategoryError(Exception):
    """Eccezione per errori di business logic del modulo categorie fornitori."""
    pass


def get_all_supplier_categories(db_manager: DatabaseManager) -> list:
    """
    Restituisce la lista ordinata di categorie dal catalogo ufficiale.

    Returns:
        list[str] — ordinata alfabeticamente, vuota se nessuna categoria
    """
    try:
        return db_manager.get_all_supplier_categories()
    except DatabaseError:
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante il recupero delle categorie: {e}") from e


def ensure_supplier_category_exists(db_manager: DatabaseManager, name: str) -> None:
    """
    Crea la categoria se non esiste già.

    - Applica trim automatico
    - Se name è vuoto dopo trim: no-op silenzioso
    - Idempotente: nessun errore se la categoria esiste già

    Args:
        db_manager: Istanza attiva di DatabaseManager
        name:       Nome della categoria (verrà trimmato)
    """
    name = name.strip() if name else ""
    if not name:
        return
    try:
        db_manager.ensure_supplier_category_exists(name)
        logger.debug("Categoria garantita in catalogo: '%s'", name)
    except DatabaseError:
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante la creazione della categoria: {e}") from e


def rename_supplier_category(
    db_manager: DatabaseManager, old_name: str, new_name: str
) -> None:
    """
    Rinomina una categoria in modo transazionale.

    Regole:
    - old_name e new_name vengono trimmati
    - old_name deve esistere nel catalogo
    - new_name non può essere vuoto
    - se old_name == new_name: no-op
    - se new_name esiste già: blocca con CategoryError (suggerisce di usare Unisci)
    - UPDATE massivo potential_suppliers: old_name → new_name
    - DELETE old_name dal catalogo, garantisce new_name nel catalogo

    Raises:
        CategoryError: regole business violate
        DatabaseError: errore DB
    """
    old_name = old_name.strip() if old_name else ""
    new_name = new_name.strip() if new_name else ""

    if not old_name:
        raise CategoryError("Il nome della categoria sorgente non può essere vuoto.")
    if not new_name:
        raise CategoryError("Il nuovo nome della categoria non può essere vuoto.")
    if old_name == new_name:
        return  # no-op

    logger.info("Rinomina categoria: '%s' → '%s'", old_name, new_name)
    try:
        db_manager.rename_supplier_category(old_name, new_name)
        logger.info("Categoria rinominata correttamente: '%s' → '%s'", old_name, new_name)
    except DatabaseError as e:
        msg = str(e)
        # Traduce eccezioni DB con message di business logic chiaro
        if "esiste già" in msg or "Usa la funzione Unisci" in msg:
            raise CategoryError(msg) from e
        if "non trovata" in msg:
            raise CategoryError(
                f"La categoria '{old_name}' non esiste nel catalogo."
            ) from e
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante la rinomina: {e}") from e


def merge_supplier_categories(
    db_manager: DatabaseManager, source: str, target: str
) -> None:
    """
    Unisce la categoria source verso target in modo transazionale.

    Tutti i supplier con category == source vengono aggiornati a target.
    La categoria source viene rimossa dal catalogo; target rimane.

    Regole:
    - trim su source e target
    - source deve esistere nel catalogo
    - source != target
    - target viene garantito nel catalogo (robustezza)

    Raises:
        CategoryError: regole business violate
        DatabaseError: errore DB
    """
    source = source.strip() if source else ""
    target = target.strip() if target else ""

    if not source:
        raise CategoryError("La categoria sorgente non può essere vuota.")
    if not target:
        raise CategoryError("La categoria destinazione non può essere vuota.")
    if source == target:
        raise CategoryError("Sorgente e destinazione devono essere diverse.")

    logger.info("Unione categoria: '%s' → '%s'", source, target)
    try:
        db_manager.merge_supplier_categories(source, target)
        logger.info("Unione completata: '%s' → '%s'", source, target)
    except DatabaseError as e:
        msg = str(e)
        if "non trovata" in msg:
            raise CategoryError(
                f"La categoria '{source}' non esiste nel catalogo."
            ) from e
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante l'unione: {e}") from e


def delete_supplier_category_if_unused(
    db_manager: DatabaseManager, name: str
) -> int:
    """
    Elimina la categoria dal catalogo solo se non è usata da alcun supplier.

    Args:
        db_manager: Istanza attiva di DatabaseManager
        name:       Nome categoria (verrà trimmato)

    Returns:
        int: 0 se eliminata; >0 se bloccata (numero di supplier ancora associati)

    Raises:
        CategoryError: se name vuoto
        DatabaseError: errore DB
    """
    name = name.strip() if name else ""
    if not name:
        raise CategoryError("Il nome della categoria non può essere vuoto.")

    logger.info("Richiesta eliminazione categoria '%s'", name)
    try:
        count = db_manager.delete_supplier_category_if_unused(name)
        if count == 0:
            logger.info("Categoria '%s' eliminata.", name)
        else:
            logger.info(
                "Categoria '%s' non eliminata: ancora usata da %d supplier.", name, count
            )
        return count
    except DatabaseError:
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante l'eliminazione: {e}") from e


def count_suppliers_by_category(db_manager: DatabaseManager, name: str) -> int:
    """
    Restituisce il numero di supplier associati a questa categoria.

    Args:
        db_manager: Istanza attiva di DatabaseManager
        name:       Nome categoria (verrà trimmato)

    Returns:
        int: numero di supplier
    """
    name = name.strip() if name else ""
    try:
        return db_manager.count_suppliers_by_category(name)
    except DatabaseError:
        raise
    except Exception as e:
        raise CategoryError(f"Errore durante il conteggio: {e}") from e
