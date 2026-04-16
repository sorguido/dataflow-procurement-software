"""
Utility per formattazione numeri secondo convenzioni italiane.
Gestione conversione virgola/punto decimale.
"""

import configparser
import os

from utils.user_utils import get_config_file


_VALID_CURRENCY_CODES = {"NONE", "EUR", "USD", "GBP", "CHF"}


def parse_float_from_comma_string(s):
    """Converte una stringa con virgola decimale in float, con validazione robusta.
    
    BUG #5 FIX: Validazione completa per gestire None, stringhe vuote e malformate.
    """
    # Gestione None e tipi numerici
    if s is None:
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    
    # Converti a stringa e pulisci
    s = str(s).strip()
    
    # Gestione stringa vuota
    if not s or s == '':
        return 0.0
    
    # Validazione: accetta solo numeri, virgola e segno
    if not all(c.isdigit() or c in ',-' for c in s):
        raise ValueError(f"Formato numero non valido: '{s}'. Usare solo cifre e virgola come separatore decimale.")
    
    # Validazione: no punto decimale
    if '.' in s:
        raise ValueError("Usare la virgola, non il punto, come separatore decimale.")
    
    # Validazione: massimo una virgola
    if s.count(',') > 1:
        raise ValueError(f"Formato numero non valido: '{s}'. Troppi separatori decimali.")
    
    # Conversione sicura
    try:
        return float(s.replace(',', '.'))
    except ValueError as e:
        raise ValueError(f"Impossibile convertire '{s}' in numero: {e}")


def format_quantity_display(val):
    """Formatta la quantità per la visualizzazione con gestione errori robusta.
    
    BUG #6 FIX: Gestione completa degli errori di conversione.
    """
    if val is None or val == '':
        return ''
    
    # Se è già numero, formatta direttamente
    if isinstance(val, (int, float)):
        if val == int(val):
            return str(int(val))
        else:
            return str(val).replace('.', ',')
    
    # Se è stringa, prova a convertire
    try:
        val_float = parse_float_from_comma_string(val)
        if val_float == int(val_float):
            return str(int(val_float))
        else:
            return str(val_float).replace('.', ',')
    except (ValueError, TypeError):
        # Se la conversione fallisce, restituisci la stringa originale
        return str(val)


def format_amount_display(val):
    """Formatta un valore numerico con esattamente 2 decimali e virgola come separatore.

    Usato per mostrare importi e percentuali nei campi di input al caricamento,
    indipendentemente dalla precisione con cui sono stati salvati.

    Esempi:
        20000.3   → "20000,30"
        70.6434   → "70,64"
        100.0     → "100,00"
        1         → "1,00"
    """
    if val is None or val == '':
        return ''
    try:
        return f"{float(val):.2f}".replace('.', ',')
    except (ValueError, TypeError):
        return str(val)


def get_currency_code(default="NONE"):
    """Legge il codice valuta globale da config.ini (Settings.currency_code)."""
    fallback = default if default in _VALID_CURRENCY_CODES else "NONE"
    try:
        config_file = get_config_file()
        config = configparser.ConfigParser(interpolation=None)
        if os.path.exists(config_file):
            config.read(config_file, encoding="utf-8")
        value = config.get("Settings", "currency_code", fallback=fallback).strip().upper()
        return value if value in _VALID_CURRENCY_CODES else fallback
    except Exception:
        return fallback


def _format_number_it(value):
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _format_number_en(value):
    return f"{value:,.2f}"


def format_currency_display(val, currency_code=None):
    """Formatta un importo in base alla preferenza valuta globale (solo display)."""
    code = (currency_code or get_currency_code()).upper()
    if code not in _VALID_CURRENCY_CODES:
        code = "NONE"
    try:
        value = float(val or 0.0)
    except (ValueError, TypeError):
        return str(val)

    if code in {"NONE", "EUR"}:
        numeric = _format_number_it(value)
    else:
        numeric = _format_number_en(value)

    if code == "NONE":
        return numeric
    if code == "EUR":
        return f"{numeric} €"
    if code == "USD":
        return f"${numeric}"
    if code == "GBP":
        return f"£{numeric}"
    # CHF
    return f"CHF {numeric}"


def get_currency_excel_number_format(currency_code=None):
    """Restituisce il number_format Excel per importi numerici secondo valuta scelta."""
    code = (currency_code or get_currency_code()).upper()
    if code == "EUR":
        return '#,##0.00 "€"'
    if code == "USD":
        return "$#,##0.00"
    if code == "GBP":
        return "£#,##0.00"
    if code == "CHF":
        return '#,##0.00 "CHF"'
    return "#,##0.00"
