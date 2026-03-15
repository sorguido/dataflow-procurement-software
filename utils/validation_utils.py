"""
Utility per validazione e sanitizzazione input utente.
Helper functions per DataFlow.
"""

import re
from datetime import datetime
import logging

# Logger locale per questo modulo
logger = logging.getLogger(__name__)


def sanitize_filename(name):
    """Rimuove caratteri non validi da nome file.
    
    Args:
        name: Nome file originale
        
    Returns:
        str: Nome file sanitizzato (rimuove caratteri Windows/Unix non validi)
    """
    if not name:
        return ""
    return re.sub(r'[\\/*?:"<>|]', "", str(name))


def format_date_for_db(display_date):
    """Converte data in formato visualizzazione (dd/mm/yyyy) in formato DB (YYYY-MM-DD).
    
    Args:
        display_date: Data in formato dd/mm/yyyy
        
    Returns:
        str | None: Data in formato YYYY-MM-DD o None se invalida
    """
    if not display_date:
        return None
    try:
        return datetime.strptime(display_date, '%d/%m/%Y').strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        return None


def format_price_display(num):
    """Formatta il prezzo per la visualizzazione con 4 decimali e virgola.
    
    BUG #6 FIX: Gestione completa degli errori di conversione.
    
    Args:
        num: Numero o stringa rappresentante prezzo
        
    Returns:
        str: Prezzo formattato (es: "123,4500") o stringa vuota se invalido
    """
    # Import locale per evitare dipendenze circolari
    from utils.format_utils import parse_float_from_comma_string
    
    if num is None or num == '':
        return ''
    
    # Converti in float gestendo errori
    try:
        if isinstance(num, str):
            num_float = parse_float_from_comma_string(num)
        else:
            num_float = float(num)
        return f"{num_float:.4f}".replace('.', ',')
    except (ValueError, TypeError) as e:
        # In caso di errore, restituisci stringa vuota o valore originale
        logger.warning(f"Impossibile formattare prezzo '{num}': {e}")
        return str(num) if num else ''
