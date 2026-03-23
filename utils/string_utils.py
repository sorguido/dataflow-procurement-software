"""
Utility per manipolazione stringhe e generazione username.
Helper functions per DataFlow.
"""

import unicodedata


def _strip_accents(value):
    """Rimuove gli accenti da una stringa mantenendo solo caratteri ASCII."""
    if not value:
        return ""
    normalized = unicodedata.normalize('NFKD', value)
    return ''.join(ch for ch in normalized if not unicodedata.combining(ch))


def generate_username(first_name, last_name):
    """
    Genera lo username secondo le regole: prima lettera del nome + cognome,
    senza spazi, senza accenti e tutto in minuscolo.
    """
    if not first_name or not last_name:
        raise ValueError("Nome e cognome sono obbligatori per generare lo username.")
    
    first_clean = ''.join(ch for ch in _strip_accents(first_name).strip() if ch.isalpha())
    last_clean = ''.join(ch for ch in _strip_accents(last_name) if ch.isalnum())
    
    if not first_clean or not last_clean:
        raise ValueError("Nome e cognome devono contenere caratteri alfabetici validi.")
    
    username = (first_clean[0] + last_clean).lower()
    return username
