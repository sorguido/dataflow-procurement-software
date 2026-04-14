"""
Utility per formattazione numeri secondo convenzioni italiane.
Gestione conversione virgola/punto decimale.
"""


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


def format_currency_display(val, show_symbol=True):
    """Formatta un valore monetario con separatore migliaia e virgola decimale.

    Convenzione italiana: punto come separatore migliaia, virgola come decimale.

    Esempi:
        2000.0       → "€ 2.000,00" (show_symbol=True)
        38700.5      → "38.700,50"  (show_symbol=False)
        -2.16        → "-2,16"      (show_symbol=False)
        0.32         → "0,32"       (show_symbol=False)
    """
    if val is None:
        return '€ 0,00' if show_symbol else '0,00'
    try:
        # :,.2f usa virgola come migliaia e punto come decimali (en_US)
        # swap: virgola↔punto per convenzione italiana
        formatted = f"{float(val):,.2f}"
        numeric = formatted.replace(',', 'X').replace('.', ',').replace('X', '.')
        return f"€ {numeric}" if show_symbol else numeric
    except (ValueError, TypeError):
        return str(val)
