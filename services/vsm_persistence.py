"""
VSM Persistence Layer
Modulo per la persistenza degli eventi VSM e dei relativi impatti mensili.

Pattern obbligatorio: DELETE-REGENERATE-SAVE
Gli impatti NON devono mai essere aggiornati in-place.
Ogni modifica richiede: DELETE vecchi impatti → REGENERATE con engine → SAVE nuovi impatti

Questo garantisce:
- Nessun duplicato
- Aggiornamenti idempotenti
- Consistenza dei dati
- Debug semplificato
"""

import logging
from typing import List

from models.vsm_event import VSMEvent
from models.vsm_impact import VSMImpact
from services.vsm_engine import generate_impacts_for_event
from database_manager import DatabaseError

logger = logging.getLogger('DataFlow.VSMPersistence')


class VSMError(Exception):
    """Eccezione personalizzata per errori di business logic VSM"""
    pass


def save_event_with_impacts(db_manager, event: VSMEvent) -> int:
    """
    Salva un nuovo evento VSM e genera automaticamente gli impatti mensili.
    
    ATOMICITÀ: Tutte le operazioni sono eseguite in UNA SINGOLA TRANSAZIONE.
    Se qualsiasi step fallisce, viene eseguito ROLLBACK completo.
    
    Pattern di esecuzione:
    1. BEGIN TRANSACTION
    2. INSERT vsm_event (senza commit)
    3. Ottieni event_id da lastrowid
    4. Genera impatti con VSM Engine
    5. INSERT impatti in batch (senza commit)
    6. COMMIT unico
    
    Args:
        db_manager: Istanza di DatabaseManager
        event: VSMEvent da salvare (event_id deve essere None o 0)
    
    Returns:
        int: event_id assegnato dal database
    
    Raises:
        VSMError: Se event ha già un event_id (deve essere nuovo)
        DatabaseError: Se errore durante operazioni DB
    """
    # Validazione: evento deve essere nuovo
    if event.id is not None and event.id != 0:
        raise VSMError(
            f"save_event_with_impacts richiede evento nuovo (event_id=None o 0), "
            f"ricevuto event_id={event.id}"
        )
    
    logger.info(
        f"Salvataggio nuovo evento VSM: username={event.username}, "
        f"data={event.event_date}, tipo={event.event_type}"
    )
    
    try:
        # ========================================
        # TRANSAZIONE ATOMICA
        # ========================================
        db_manager.cursor.execute("BEGIN TRANSACTION")
        
        # Step 1: INSERT evento (SENZA COMMIT)
        event_id = db_manager._insert_vsm_event_no_commit(event)
        logger.debug(f"Evento inserito con ID {event_id}")
        
        # Step 2: Genera impatti con VSM Engine
        event.id = event_id  # Aggiorna evento con ID assegnato
        impacts = generate_impacts_for_event(event)
        logger.debug(f"Generati {len(impacts)} impatti per evento {event_id}")
        
        # Step 3: INSERT impatti in batch (SENZA COMMIT)
        if impacts:
            db_manager._insert_vsm_impacts_no_commit(impacts)
            logger.info(
                f"Salvati {len(impacts)} impatti per evento {event_id} "
                f"(periodo: {impacts[0].year}/{impacts[0].month} - "
                f"{impacts[-1].year}/{impacts[-1].month})"
            )
        else:
            logger.warning(f"Nessun impatto generato per evento {event_id}")
        
        # COMMIT UNICO: tutte le operazioni hanno successo
        db_manager.conn.commit()
        
        return event_id
        
    except Exception as e:
        # ROLLBACK: annulla TUTTE le operazioni
        db_manager.conn.rollback()
        logger.error(f"Errore durante salvataggio atomico, ROLLBACK eseguito: {e}")
        raise VSMError(f"Errore durante salvataggio evento VSM: {e}") from e


def update_event_with_impacts(db_manager, event: VSMEvent) -> None:
    """
    Aggiorna un evento VSM esistente e rigenera tutti gli impatti mensili.
    
    ATOMICITÀ: Tutte le operazioni sono eseguite in UNA SINGOLA TRANSAZIONE.
    Se qualsiasi step fallisce, viene eseguito ROLLBACK completo.
    
    Pattern obbligatorio DELETE-REGENERATE-SAVE (protetto da transazione):
    1. BEGIN TRANSACTION
    2. UPDATE vsm_event (senza commit)
    3. DELETE vecchi impatti (senza commit)
    4. REGENERATE impatti con VSM Engine
    5. INSERT nuovi impatti in batch (senza commit)
    6. COMMIT unico
    
    Questo pattern garantisce:
    - Nessun duplicato (vecchi impatti sempre eliminati)
    - Idempotenza (multiple chiamate producono stesso risultato)
    - Consistenza (impatti sempre allineati con evento)
    - Atomicità (tutto-o-niente, no stati intermedi)
    
    Args:
        db_manager: Istanza di DatabaseManager
        event: VSMEvent aggiornato (event_id deve essere valido)
    
    Raises:
        VSMError: Se event_id non valido o evento non esiste
        DatabaseError: Se errore durante operazioni DB
    """
    # Validazione: evento deve esistere
    if event.id is None or event.id == 0:
        raise VSMError(
            f"update_event_with_impacts richiede evento esistente con event_id valido, "
            f"ricevuto event_id={event.id}"
        )
    
    logger.info(
        f"Aggiornamento evento VSM {event.id}: username={event.username}, "
        f"data={event.event_date}"
    )
    
    try:
        # ========================================
        # TRANSAZIONE ATOMICA
        # ========================================
        db_manager.cursor.execute("BEGIN TRANSACTION")
        
        # Step 1: UPDATE evento (SENZA COMMIT)
        db_manager._update_vsm_event_no_commit(event)
        logger.debug(f"Evento {event.id} aggiornato")
        
        # Step 2: DELETE vecchi impatti (SENZA COMMIT)
        db_manager._delete_vsm_impacts_no_commit(event.id)
        logger.debug(f"Vecchi impatti eliminati per evento {event.id}")
        
        # Step 3: REGENERATE impatti
        impacts = generate_impacts_for_event(event)
        logger.debug(f"Rigenerati {len(impacts)} impatti per evento {event.id}")
        
        # Step 4: INSERT nuovi impatti (SENZA COMMIT)
        if impacts:
            db_manager._insert_vsm_impacts_no_commit(impacts)
            logger.info(
                f"Rigenerati e salvati {len(impacts)} impatti per evento {event.id} "
                f"(periodo: {impacts[0].year}/{impacts[0].month} - "
                f"{impacts[-1].year}/{impacts[-1].month})"
            )
        else:
            logger.warning(f"Nessun impatto generato per evento {event.id}")
        
        # COMMIT UNICO: tutte le operazioni hanno successo
        db_manager.conn.commit()
            
    except Exception as e:
        # ROLLBACK: ripristina stato precedente completo
        db_manager.conn.rollback()
        logger.error(f"Errore durante aggiornamento atomico evento {event.id}, ROLLBACK eseguito: {e}")
        raise VSMError(f"Errore durante aggiornamento evento VSM: {e}") from e


def delete_event_and_impacts(db_manager, event_id: int) -> None:
    """
    Elimina un evento VSM e tutti i relativi impatti mensili.
    
    Pattern di esecuzione (senza CASCADE, gestione esplicita):
    1. DELETE impatti per event_id (figli prima)
    2. DELETE evento (padre dopo)
    
    Nota: Eliminazione è definitiva (no soft delete)
    
    Args:
        db_manager: Istanza di DatabaseManager
        event_id: ID dell'evento da eliminare
    
    Raises:
        VSMError: Se event_id non valido
        DatabaseError: Se errore durante operazioni DB
    """
    if event_id is None or event_id <= 0:
        raise VSMError(f"delete_event_and_impacts richiede event_id valido, ricevuto {event_id}")
    
    logger.info(f"Eliminazione evento VSM {event_id} e relativi impatti")
    
    try:
        # Step 1: DELETE impatti (figli prima)
        db_manager.delete_vsm_impacts_by_event_id(event_id)
        logger.debug(f"Impatti eliminati per evento {event_id}")
        
        # Step 2: DELETE evento (padre dopo)
        db_manager.delete_vsm_event(event_id)
        logger.info(f"Evento {event_id} eliminato con successo")
        
    except DatabaseError as e:
        logger.error(f"Errore database durante eliminazione evento {event_id}: {e}")
        raise
    except Exception as e:
        logger.error(f"Errore imprevisto durante eliminazione evento {event_id}: {e}")
        raise VSMError(f"Errore durante eliminazione evento VSM: {e}")


def get_event_with_impacts(db_manager, event_id: int) -> tuple[VSMEvent, List[VSMImpact]]:
    """
    Recupera un evento VSM con tutti i relativi impatti mensili.
    
    Args:
        db_manager: Istanza di DatabaseManager
        event_id: ID dell'evento da recuperare
    
    Returns:
        tuple: (VSMEvent, List[VSMImpact])
    
    Raises:
        VSMError: Se evento non trovato
        DatabaseError: Se errore durante operazioni DB
    """
    if event_id is None or event_id <= 0:
        raise VSMError(f"get_event_with_impacts richiede event_id valido, ricevuto {event_id}")
    
    try:
        event = db_manager.get_vsm_event_by_id(event_id)
        if event is None:
            raise VSMError(f"Evento {event_id} non trovato")
        
        impacts = db_manager.get_vsm_impacts_by_event_id(event_id)
        logger.debug(f"Recuperati evento {event_id} con {len(impacts)} impatti")
        
        return event, impacts
        
    except VSMError:
        raise
    except DatabaseError as e:
        logger.error(f"Errore database durante recupero evento {event_id}: {e}")
        raise
    except Exception as e:
        logger.error(f"Errore imprevisto durante recupero evento {event_id}: {e}")
        raise VSMError(f"Errore durante recupero evento VSM: {e}")
