"""
Gestione percorsi e directory dell'applicazione DataFlow.
"""
import os
import sys
import configparser
import logging
from utils.user_utils import get_config_file, load_user_identity

logger = logging.getLogger('DataFlow')

# Cache per verifica struttura
_DATAFLOW_STRUCTURE_VERIFIED = False

# Cache per percorso database
_PERCORSO_DB_CACHE = None


def get_user_documents_dataflow_dir():
    """Restituisce la directory dati utente principale per DataFlow.
    Rispetta dataflow_base_dir se presente nel config.ini.
    """
    # Determina lo username dal config.ini
    config_file = get_config_file()
    config = configparser.ConfigParser(interpolation=None)
    username = None
    dataflow_base_dir = None
    
    if os.path.exists(config_file):
        config.read(config_file)
        if 'User' in config and config.get('User', 'username', fallback=None):
            username = config.get('User', 'username').strip().lower()
        
        # ✅ LEGGI dataflow_base_dir se presente
        if 'Settings' in config:
            dataflow_base_dir = config.get('Settings', 'dataflow_base_dir', fallback=None)

    # Se non c'è username, NON creare nessuna cartella e restituisci None
    if not username:
        logger.warning("Username non presente: la cartella utente non viene creata.")
        return None

    base_folder = f"DataFlow_{username}"
    
    # ✅ USA dataflow_base_dir se presente, altrimenti default a Documents
    if dataflow_base_dir and os.path.exists(dataflow_base_dir):
        chosen_dir = os.path.join(dataflow_base_dir, base_folder)
        logger.info(f"Usando directory DataFlow personalizzata: {chosen_dir}")
    else:
        # Windows: usa ~/Documents/DataFlow_username (comportamento standard Windows)
        # Linux/macOS: usa ~/DataFlow_username (direttamente nella home)
        if sys.platform == 'win32':
            documents_dir = os.path.join(os.path.expanduser('~'), 'Documents')
            chosen_dir = os.path.join(documents_dir, base_folder)
        else:
            chosen_dir = os.path.join(os.path.expanduser('~'), base_folder)
        logger.info(f"Usando directory DataFlow standard: {chosen_dir}")
    
    try:
        os.makedirs(chosen_dir, exist_ok=True)
    except OSError as e:
        logger.error(f"Impossibile creare la cartella DataFlow utente '{chosen_dir}': {e}")
        return None

    global _DATAFLOW_STRUCTURE_VERIFIED
    if not _DATAFLOW_STRUCTURE_VERIFIED:
        # Leggi config per capire se è stato già impostato un DB personalizzato.
        custom_db = None
        try:
            if os.path.exists(config_file):
                config.read(config_file)
                custom_db = config.get('Settings', 'custom_db_path', fallback=None)
        except Exception:
            custom_db = None

        required = [
            os.path.join(chosen_dir, 'Database'),
            os.path.join(chosen_dir, 'Attachments')
        ]
        if any(not os.path.exists(path) for path in required):
            try:
                from services.startup_service import initialize_dataflow_directory_structure
                initialize_dataflow_directory_structure(chosen_dir)
            except Exception as e:
                logger.error(f"Errore nel ripristino automatico della struttura DataFlow: {e}")
        _DATAFLOW_STRUCTURE_VERIFIED = True

    return chosen_dir


def get_fixed_db_dir():
    """Restituisce la cartella fissa per il database."""
    db_dir = os.path.join(get_user_documents_dataflow_dir(), 'Database')
    os.makedirs(db_dir, exist_ok=True)
    return db_dir


def get_fixed_attachments_dir():
    """Restituisce la cartella fissa per gli allegati (Attachments)."""
    base_dir = get_user_documents_dataflow_dir()
    if not base_dir:
        return None
    
    new_dir = os.path.join(base_dir, 'Attachments')
    old_dir = os.path.join(base_dir, 'Allegati')
    
    if os.path.exists(old_dir) and not os.path.exists(new_dir):
        try:
            import shutil
            shutil.move(old_dir, new_dir)
            logger.info(f"Cartella allegati rinominata da '{old_dir}' a '{new_dir}'")
        except Exception as e:
            logger.error(f"Impossibile rinominare cartella Allegati: {e}")
    
    os.makedirs(new_dir, exist_ok=True)
    return new_dir


def reset_db_cache():
    """
    Invalida la cache del percorso DB per forzare il ricaricamento.
    Chiamare questa funzione quando si modifica il percorso del database
    o si vuole forzare il ricalcolo.
    """
    global _PERCORSO_DB_CACHE
    _PERCORSO_DB_CACHE = None
    logger.info("Cache percorso DB invalidata")


def get_db_path():
    """
    Determina il percorso del database da usare per la sessione corrente.
    Alla prima chiamata, legge il file config.ini per decidere se usare il DB
    personalizzato o quello standard. Alle chiamate successive, restituisce 
    il percorso già memorizzato (cache) per garantire coerenza durante tutta la sessione.
    
    Priorità:
    1. Directory DataFlow personalizzata (dataflow_base_dir) - permanente
    2. Database personalizzato (custom_db_path)
    3. Database standard (Documents/DataFlow/Database)
    
    Returns:
        str: Percorso assoluto al file database da usare (estensione .db)
    """
    global _PERCORSO_DB_CACHE
    if _PERCORSO_DB_CACHE is not None:
        return _PERCORSO_DB_CACHE

    config = configparser.ConfigParser(interpolation=None)
    config_file = get_config_file()
    legacy_custom_path = None
    dataflow_override = None
    
    if os.path.exists(config_file):
        try:
            config.read(config_file)
            dataflow_override = config.get('Settings', 'dataflow_base_dir', fallback=None)
            legacy_custom_path = config.get('Settings', 'custom_db_path', fallback=None)
        except Exception as e:
            logger.error(f"Errore lettura config per percorsi DB: {e}")
    
    if legacy_custom_path:
        percorso_da_usare = legacy_custom_path
        logger.info(f"Usando database personalizzato: {legacy_custom_path}")
        if not os.path.exists(percorso_da_usare):
            try:
                os.makedirs(os.path.dirname(percorso_da_usare), exist_ok=True)
                logger.info(f"Creata directory per database legacy: {os.path.dirname(percorso_da_usare)}")
            except OSError as e:
                logger.error(f"Impossibile creare cartella per database legacy: {e}")
                percorso_da_usare = None
    else:
        # Prova a ricavare lo username e costruire il percorso
        identity = load_user_identity()
        username = identity.get('username')
        if username:
            base_dir = get_user_documents_dataflow_dir()
            percorso_da_usare = os.path.join(base_dir, 'Database', f'dataflow_db_{username}.db')
            logger.info(f"Usando database utente: {percorso_da_usare}")
        else:
            logger.error("Nessun username trovato: impossibile determinare percorso DB.")
            percorso_da_usare = None
    _PERCORSO_DB_CACHE = percorso_da_usare
    logger.info(f"Percorso database finale: {percorso_da_usare}")
    return percorso_da_usare
