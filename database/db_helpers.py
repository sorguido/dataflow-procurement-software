"""
Helper per inizializzazione e gestione database.
"""
import os
import configparser
import logging
from database_manager import DatabaseManager, DatabaseError
from utils.user_utils import get_config_file
from services.app_paths import get_db_path

logger = logging.getLogger('DataFlow')


def crea_database_v4():
    """Inizializza il database creando le tabelle necessarie."""
    logger.info("Inizializzazione database")
    # Determina il percorso DB preferito (rispetta custom_db_path se presente nel config)
    db_file = get_db_path()
    try:
        config_file = get_config_file()
        if os.path.exists(config_file):
            cfg = configparser.ConfigParser(interpolation=None)
            cfg.read(config_file)
            custom = cfg.get('Settings', 'custom_db_path', fallback=None)
            if custom:
                # Forza l'uso del DB personalizzato per evitare di creare il DB standard
                db_file = custom
                logger.info(f"Usando custom_db_path per inizializzazione DB: {db_file}")
    except Exception as e:
        # BUG #36 FIX: Log eccezioni invece di silenziarle completamente
        logger.debug(f"Nessun custom_db_path configurato o errore lettura config: {e}")
    
    is_new_db = not os.path.exists(db_file)
    
    # BUG #32 FIX: Usa try-finally per garantire chiusura DB anche in caso di eccezione
    db_manager = None
    try:
        # Usa il DatabaseManager per creare le tabelle
        db_manager = DatabaseManager(db_file)
        db_manager.create_tables()
        
        if is_new_db: 
            print("Nuovo database creato. Imposto il contatore RdO a 0.")
            logger.info("Nuovo database creato")
        
        logger.info("Database inizializzato con successo")
        
    except DatabaseError as e:
        logger.error(f"Errore critico inizializzazione database: {e}", exc_info=True)
        print(f"ERRORE CRITICO: Impossibile inizializzare il database.\n{e}")
        raise
    finally:
        # BUG #32 FIX: Garantisce chiusura connessione anche in caso di eccezione
        if db_manager is not None:
            try:
                db_manager.close()
            except Exception as close_error:
                logger.warning(f"Errore chiusura database in finally: {close_error}")
