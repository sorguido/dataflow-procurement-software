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
    
    Pattern di esecuzione:
    1. INSERT vsm_event (senza event_id)
    2. Ottieni event_id da lastrowid
    3. Genera impatti con VSM Engine
    4. Batch INSERT impatti con transazione
    
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
        # Step 1: INSERT evento
        event_id = db_manager.insert_vsm_event(event)
        logger.debug(f"Evento inserito con ID {event_id}")
        
        # Step 2: Genera impatti con VSM Engine
        event.id = event_id  # Aggiorna evento con ID assegnato
        impacts = generate_impacts_for_event(event)
        logger.debug(f"Generati {len(impacts)} impatti per evento {event_id}")
        
        # Step 3: Batch INSERT impatti
        if impacts:
            db_manager.insert_vsm_impacts_batch(impacts)
            logger.info(
                f"Salvati {len(impacts)} impatti per evento {event_id} "
                f"(periodo: {impacts[0].year}/{impacts[0].month} - "
                f"{impacts[-1].year}/{impacts[-1].month})"
            )
        else:
            logger.warning(f"Nessun impatto generato per evento {event_id}")
        
        return event_id
        
    except DatabaseError as e:
        logger.error(f"Errore database durante salvataggio evento: {e}")
        raise
    except Exception as e:
        logger.error(f"Errore imprevisto durante salvataggio evento: {e}")
        raise VSMError(f"Errore durante salvataggio evento VSM: {e}")


def update_event_with_impacts(db_manager, event: VSMEvent) -> None:
    """
    Aggiorna un evento VSM esistente e rigenera tutti gli impatti mensili.
    
    Pattern obbligatorio DELETE-REGENERATE-SAVE:
    1. UPDATE vsm_event (event deve avere event_id valido)
    2. DELETE vecchi impatti per event_id
    3. REGENERATE impatti con VSM Engine
    4. Batch INSERT nuovi impatti con transazione
    
    Questo pattern garantisce:
    - Nessun duplicato (vecchi impatti sempre eliminati)
    - Idempotenza (multiple chiamate producono stesso risultato)
    - Consistenza (impatti sempre allineati con evento)
    
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
        # Step 1: UPDATE evento
        db_manager.update_vsm_event(event)
        logger.debug(f"Evento {event.id} aggiornato")
        
        # Step 2: DELETE vecchi impatti
        db_manager.delete_vsm_impacts_by_event_id(event.id)
        logger.debug(f"Vecchi impatti eliminati per evento {event.id}")
        
        # Step 3: REGENERATE impatti
        impacts = generate_impacts_for_event(event)
        logger.debug(f"Rigenerati {len(impacts)} impatti per evento {event.id}")
        
        # Step 4: SAVE nuovi impatti
        if impacts:
            db_manager.insert_vsm_impacts_batch(impacts)
            logger.info(
                f"Rigenerati e salvati {len(impacts)} impatti per evento {event.id} "
                f"(periodo: {impacts[0].year}/{impacts[0].month} - "
                f"{impacts[-1].year}/{impacts[-1].month})"
            )
        else:
            logger.warning(f"Nessun impatto generato per evento {event.id}")
            
    except DatabaseError as e:
        logger.error(f"Errore database durante aggiornamento evento {event.id}: {e}")
        raise
    except Exception as e:
        logger.error(f"Errore imprevisto durante aggiornamento evento {event.id}: {e}")
        raise VSMError(f"Errore durante aggiornamento evento VSM: {e}")


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
