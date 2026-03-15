"""
Servizi di inizializzazione e startup dell'applicazione.
"""
import os
import sys
import tempfile
import time
import glob
import shutil
import logging
from logging.handlers import RotatingFileHandler


def cleanup_temp_on_startup():
    """Pulisce le directory temporanee di PyInstaller rimaste da precedenti esecuzioni."""
    try:
        temp_dir = tempfile.gettempdir()
        
        # Cerca cartelle _MEI* create da PyInstaller
        for item in os.listdir(temp_dir):
            if item.startswith('_MEI'):
                temp_path = os.path.join(temp_dir, item)
                try:
                    if os.path.isdir(temp_path):
                        shutil.rmtree(temp_path, ignore_errors=True)
                except Exception:
                    pass  # Ignora errori di pulizia
        
        # Pulisci anche file temporanei di DataFlow vecchi (>24 ore)
        # Pattern per file temporanei creati da AttachmentWindow
        pattern = os.path.join(temp_dir, 'tmp*')
        current_time = time.time()
        for temp_file in glob.glob(pattern):
            try:
                # Elimina file più vecchi di 24 ore (86400 secondi)
                if os.path.isfile(temp_file) and (current_time - os.path.getmtime(temp_file)) > 86400:
                    os.remove(temp_file)
            except Exception:
                pass  # Ignora errori di pulizia
                
    except Exception:
        pass  # Ignora completamente errori di pulizia


def setup_logging():
    """Configura il sistema di logging con file rotanti."""
    # Usa la directory locale dell'utente invece della directory corrente
    if getattr(sys, 'frozen', False):
        # Se eseguito come EXE/MSIX/Flatpak frozen
        if sys.platform == 'win32':
            log_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DataFlow')
        else:
            # Linux/macOS frozen (Flatpak)
            if 'XDG_DATA_HOME' in os.environ:
                log_dir = os.path.join(os.environ['XDG_DATA_HOME'], 'DataFlow')
            else:
                log_dir = os.path.join(os.path.expanduser('~'), '.local', 'share', 'DataFlow')
    else:
        # Se eseguito come script Python (NON più dalla directory dello script)
        if 'XDG_DATA_HOME' in os.environ:
            log_dir = os.path.join(os.environ['XDG_DATA_HOME'], 'DataFlow')
        else:
            log_dir = os.path.join(os.path.expanduser('~'), '.local', 'share', 'DataFlow')
    
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'dataflow.log')
    
    logger = logging.getLogger('DataFlow')
    logger.setLevel(logging.INFO)
    
    # Rimuovi handler esistenti per evitare duplicati in caso di riavvio
    if logger.handlers:
        logger.handlers.clear()
    
    handler = RotatingFileHandler(
        log_file, 
        maxBytes=5*1024*1024,  # 5MB
        backupCount=3,
        encoding='utf-8'
    )
    
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(funcName)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    
    return logger


def initialize_dataflow_directory_structure(base_dir=None):
    """
    Crea la struttura standard DataFlow (Database, Allegati, ecc.) e
    inizializza un database SQLite vuoto con le tabelle richieste.
    """
    logger = logging.getLogger('DataFlow')
    
    try:
        if base_dir:
            base_dir = os.path.normpath(os.path.abspath(base_dir))
        else:
            from services.app_paths import get_user_documents_dataflow_dir
            base_dir = get_user_documents_dataflow_dir()
    except Exception as e:
        logger.error(f"Impossibile determinare la cartella DataFlow: {e}")
        raise
    
    # Gestione migrazione vecchia cartella "Allegati" -> "Attachments"
    old_attachments_dir = os.path.join(base_dir, 'Allegati')
    new_attachments_dir = os.path.join(base_dir, 'Attachments')
    if os.path.exists(old_attachments_dir) and not os.path.exists(new_attachments_dir):
        try:
            shutil.move(old_attachments_dir, new_attachments_dir)
            logger.info(f"Cartella Allegati migrata in Attachments: {new_attachments_dir}")
        except Exception as e:
            logger.error(f"Impossibile migrare cartella Allegati: {e}")
    
    subfolders = ['Database', 'Attachments']
    try:
        os.makedirs(base_dir, exist_ok=True)
        for sub in subfolders:
            os.makedirs(os.path.join(base_dir, sub), exist_ok=True)
        logger.info(f"Struttura DataFlow creata/in ripristino in: {base_dir}")
    except OSError as e:
        logger.error(f"Errore nella creazione delle cartelle DataFlow: {e}")
        raise
    
    # NON creare nessun database qui! Solo la struttura cartelle.
    # Il database verrà creato solo dopo l'inserimento dell'identità utente.
    return None
