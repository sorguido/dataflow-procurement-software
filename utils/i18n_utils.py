"""
Modulo per la gestione del sistema di internazionalizzazione (i18n).
Gestisce l'inizializzazione di gettext e le funzioni helper per traduzioni condizionali.
"""

import os
import sys
import gettext
import configparser
import logging

# Logger locale per questo modulo
logger = logging.getLogger(__name__)


def init_i18n(language_code='en'):
    """
    Inizializza il sistema di internazionalizzazione (gettext).
    Legge la lingua preferita dal config.ini o usa 'en' come default.
    """
    # Import locale per evitare dipendenze circolari
    from utils.user_utils import get_config_file
    from utils.resource_utils import resource_path
    
    # Leggi la lingua dal config.ini (solo se esiste e ha la chiave)
    try:
        config_file = get_config_file()
        if os.path.exists(config_file):
            config = configparser.ConfigParser(interpolation=None)
            config.read(config_file)
            if 'Settings' in config and config.has_option('Settings', 'language'):
                language_code = config.get('Settings', 'language', fallback='en')
    except Exception as e:
        logger.warning(f"Errore nel leggere config.ini per la lingua: {e}, uso default 'en'")
        language_code = 'en'
    
    # Validazione: accetta solo 'en' o 'it', default sempre 'en'
    if language_code not in ['en', 'it']:
        language_code = 'en'
    
    # Determina il percorso dei file di traduzione
    try:
        # In PyInstaller, usa resource_path per trovare i file nella directory _MEIPASS
        if getattr(sys, 'frozen', False):
            locale_dir = resource_path('locale')
        else:
            # In sviluppo, usa la directory corrente del progetto (parent di utils/)
            locale_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'locale')
        
        # Inizializza gettext
        try:
            logger.info(f"Tentativo di caricare traduzioni per '{language_code}' da: {locale_dir}")
            mo_path = os.path.join(locale_dir, language_code, 'LC_MESSAGES', 'dataflow.mo')
            
            if os.path.exists(mo_path):
                trans = gettext.translation('dataflow', localedir=locale_dir, languages=[language_code], fallback=False)
                trans.install()  # Installa _ in builtins
                logger.info(f"✓ File traduzioni caricato con successo: {mo_path}")
            else:
                logger.warning(f"File .mo non trovato: {mo_path}, uso fallback")
                trans = gettext.NullTranslations()
                trans.install()
        except Exception as e:
            # Se il file .mo non esiste o c'è errore, usa gettext.NullTranslations (fallback silenzioso)
            trans = gettext.NullTranslations()
            trans.install()
            logger.error(f"ERRORE nel caricare traduzioni per '{language_code}': {e}", exc_info=True)
    except Exception as e:
        # In caso di errore, usa NullTranslations come fallback
        trans = gettext.NullTranslations()
        trans.install()
        logger.error(f"Errore nel caricamento delle traduzioni: {e}")
    
    return language_code


def get_current_language():
    """Restituisce il codice lingua corrente ('it' o 'en').
    Gestisce correttamente il caso in cui il config non esista ancora o non sia inizializzato.
    """
    # Import locale per evitare dipendenze circolari
    from utils.user_utils import get_config_file
    
    try:
        config_file = get_config_file()
        if config_file and os.path.exists(config_file):
            config = configparser.ConfigParser(interpolation=None)
            config.read(config_file, encoding='utf-8')
            if 'Settings' in config and config.has_option('Settings', 'language'):
                lang = config.get('Settings', 'language', fallback='en')
                # Validazione: accetta solo 'en' o 'it'
                if lang in ['en', 'it']:
                    return lang
    except (configparser.Error, OSError, IOError, AttributeError) as e:
        # Log solo se non è un errore di file non esistente (normale all'avvio)
        try:
            if config_file and os.path.exists(config_file):
                logger.debug(f"Errore lettura config per lingua: {e}")
        except (NameError, UnboundLocalError):
            # config_file potrebbe non essere definito in caso di errore precoce
            pass
    except Exception as e:
        # Log altri errori inattesi
        logger.debug(f"Errore inatteso in get_current_language: {e}")
    # Fallback sempre a 'en' se qualcosa va storto
    return 'en'


def get_pos_column_text():
    """Restituisce il testo per la colonna Posizione: 'Item' in inglese, 'Pos.' in italiano."""
    return "Item" if get_current_language() == 'en' else "Pos."


def get_qty_column_text():
    """Restituisce il testo per la colonna Quantità: 'Q.ty' in inglese, 'Q.tà' in italiano."""
    return "Q.ty" if get_current_language() == 'en' else "Q.tà"


def normalize_rfq_type(rfq_type):
    """
    Normalizza un tipo di RFQ da qualsiasi lingua al valore canonico italiano.
    Gestisce sia i valori vecchi (tradotti) che quelli nuovi (canonici).
    
    BUG #7 FIX: Validazione robusta con gestione completa dei casi edge.
    """
    # Gestione valori None, vuoti o non-stringa
    if not rfq_type:
        logger.warning("normalize_rfq_type: valore None/vuoto ricevuto, uso default 'Fornitura piena'")
        return "Fornitura piena"
    
    # Converti a stringa e pulisci whitespace
    try:
        rfq_type = str(rfq_type).strip()
    except Exception as e:
        logger.error(f"normalize_rfq_type: impossibile convertire a stringa '{rfq_type}': {e}")
        return "Fornitura piena"
    
    # Se dopo strip è vuoto, usa default
    if not rfq_type:
        logger.warning("normalize_rfq_type: stringa vuota dopo strip, uso default")
        return "Fornitura piena"
    
    # Mappa tutte le possibili varianti ai valori canonici italiani
    type_map = {
        # Valori canonici italiani (già corretti)
        "Fornitura piena": "Fornitura piena",
        "Conto lavoro": "Conto lavoro",
        # Traduzioni inglesi
        "Full Supply": "Fornitura piena",
        "Work Order": "Conto lavoro",
        # Varianti possibili (case-insensitive)
        "fornitura piena": "Fornitura piena",
        "conto lavoro": "Conto lavoro",
        "full supply": "Fornitura piena",
        "work order": "Conto lavoro",
    }
    
    # Cerca corrispondenza esatta (case-sensitive prima)
    if rfq_type in type_map:
        return type_map[rfq_type]
    
    # Cerca corrispondenza case-insensitive
    rfq_type_lower = rfq_type.lower()
    for key, value in type_map.items():
        if key.lower() == rfq_type_lower:
            return value
    
    # Se non trovato, logga warning dettagliato e ritorna default
    logger.warning(f"normalize_rfq_type: tipo RFQ non riconosciuto: '{rfq_type}' (len={len(rfq_type)}), uso default 'Fornitura piena'")
    return "Fornitura piena"


def translate_rfq_type(rfq_type):
    """
    Traduce un tipo di RFQ dal valore canonico italiano alla lingua corrente.
    """
    # Prima normalizza per gestire anche valori vecchi
    canonical = normalize_rfq_type(rfq_type)
    
    # Traduci il valore canonico
    if canonical == "Fornitura piena":
        return _("Fornitura piena")
    elif canonical == "Conto lavoro":
        return _("Conto lavoro")
    else:
        # Fallback: ritorna il valore normalizzato tradotto
        return _(canonical)
