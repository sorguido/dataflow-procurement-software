"""Servizi i18n centralizzati per DataFlow.

Questo modulo espone il punto unico ufficiale per la localizzazione UI:
- `tr(text)` per tradurre stringhe a runtime
- `get_translation_service()` per accesso al service centralizzato
- compatibilita' retro tramite `_` (alias di `tr`) per i moduli legacy
"""

import os
import sys
import gettext
import configparser
import logging
import builtins
from typing import Callable

# Logger locale per questo modulo
logger = logging.getLogger(__name__)

_ALLOWED_LANGUAGES = {"en", "it"}


class TranslationService:
    """Service centralizzato per gettext.

    Mantiene un traduttore attivo a runtime e pubblica un'interfaccia stabile
    (`translate`) da usare in tutti i moduli UI.
    """

    def __init__(self):
        self._language_code = "en"
        self._translator: Callable[[str], str] = lambda text: text

    def __call__(self, text):
        return self.translate(text)

    def translate(self, text):
        """Traduce `text` usando il catalogo attivo."""
        try:
            return self._translator(text)
        except Exception:
            return text

    @staticmethod
    def _normalize_language(language_code):
        if language_code in _ALLOWED_LANGUAGES:
            return language_code
        return "en"

    @staticmethod
    def _read_language_from_config(default="en"):
        # Import locale per evitare dipendenze circolari
        from utils.user_utils import get_config_file

        language_code = default
        try:
            config_file = get_config_file()
            if os.path.exists(config_file):
                config = configparser.ConfigParser(interpolation=None)
                config.read(config_file, encoding="utf-8")
                if "Settings" in config and config.has_option("Settings", "language"):
                    language_code = config.get("Settings", "language", fallback=default)
        except Exception as e:
            logger.warning("Errore lettura lingua da config.ini: %s", e)
            language_code = default
        return language_code

    @staticmethod
    def _resolve_locale_dir():
        # Import locale per evitare dipendenze circolari
        from utils.resource_utils import resource_path

        if getattr(sys, "frozen", False):
            return resource_path("locale")
        return os.path.join(os.path.dirname(os.path.dirname(__file__)), "locale")

    def initialize(self, language_code="en"):
        """Inizializza gettext e installa il traduttore runtime."""
        configured = self._read_language_from_config(default=language_code)
        self._language_code = self._normalize_language(configured)
        locale_dir = self._resolve_locale_dir()
        mo_path = os.path.join(locale_dir, self._language_code, "LC_MESSAGES", "dataflow.mo")

        try:
            logger.info(
                "Tentativo caricamento traduzioni '%s' da %s",
                self._language_code,
                locale_dir,
            )
            if os.path.exists(mo_path):
                trans = gettext.translation(
                    "dataflow",
                    localedir=locale_dir,
                    languages=[self._language_code],
                    fallback=False,
                )
                self._translator = trans.gettext
                logger.info("File traduzioni caricato con successo: %s", mo_path)
            else:
                logger.warning("File .mo non trovato: %s, uso fallback", mo_path)
                self._translator = gettext.NullTranslations().gettext
        except Exception as e:
            logger.error(
                "Errore nel caricamento traduzioni per '%s': %s",
                self._language_code,
                e,
                exc_info=True,
            )
            self._translator = gettext.NullTranslations().gettext

        # Retrocompatibilita': alcuni moduli legacy usano builtins._.
        builtins._ = self.translate
        return self._language_code

    def get_current_language(self):
        """Restituisce la lingua configurata corrente (`it` o `en`)."""
        # Legge da config per allinearsi allo stato persistito applicativo.
        configured = self._read_language_from_config(default=self._language_code)
        self._language_code = self._normalize_language(configured)
        return self._language_code


_translation_service = TranslationService()


def get_translation_service():
    """Ritorna il singleton TranslationService."""
    return _translation_service


def tr(text):
    """API ufficiale per tradurre testo UI."""
    return _translation_service.translate(text)


def _(text):
    """Alias retrocompatibile di `tr`."""
    return tr(text)


def init_i18n(language_code="en"):
    """Inizializza il service i18n centralizzato."""
    return _translation_service.initialize(language_code=language_code)


def get_current_language():
    """Restituisce il codice lingua corrente (`it` o `en`)."""
    return _translation_service.get_current_language()


def get_pos_column_text():
    """Restituisce il testo per la colonna Posizione: 'Item' in inglese, 'Pos.' in italiano."""
    return "Item" if get_current_language() == 'en' else "Pos."


def get_qty_column_text():
    """Restituisce il testo per la colonna Quantità: 'Q.ty' in inglese, 'Q.tà' in italiano."""
    return "Q.ty" if get_current_language() == 'en' else "Q.tà"


_DERISKING_STATUS_IT_TO_EN = {
    "Nuovo": "New",
    "In valutazione": "Under Evaluation",
    "Qualificato": "Qualified",
    "Scartato": "Rejected",
}
_DERISKING_STATUS_EN_TO_IT = {v: k for k, v in _DERISKING_STATUS_IT_TO_EN.items()}


def normalize_derisking_status(value):
    """Normalizza uno status Derisking al valore canonico italiano."""
    if value is None:
        return value

    try:
        raw = str(value).strip()
    except Exception:
        return value

    if not raw:
        return raw

    if raw in _DERISKING_STATUS_IT_TO_EN:
        return raw
    if raw in _DERISKING_STATUS_EN_TO_IT:
        return _DERISKING_STATUS_EN_TO_IT[raw]

    raw_lower = raw.lower()
    for canonical_it in _DERISKING_STATUS_IT_TO_EN:
        if canonical_it.lower() == raw_lower:
            return canonical_it
    for label_en, canonical_it in _DERISKING_STATUS_EN_TO_IT.items():
        if label_en.lower() == raw_lower:
            return canonical_it

    return raw


def _derisking_status_msgid_en(value):
    """Ritorna il msgid EN stabile per uno status Derisking."""
    canonical_it = normalize_derisking_status(value)
    return _DERISKING_STATUS_IT_TO_EN.get(canonical_it, canonical_it)


def translate_derisking_status(value, *, language_code=None):
    """Traduce uno status Derisking con fallback coerente al dominio."""
    if value is None:
        return value

    msgid_en = _derisking_status_msgid_en(value)

    if language_code is not None:
        code = str(language_code).strip().lower()
        if code in {"en", "eng"}:
            return msgid_en
        if code in {"it", "ita"}:
            return normalize_derisking_status(value)

    return tr(msgid_en)


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
        return tr("Full Supply")
    elif canonical == "Conto lavoro":
        return tr("Work Order")
    else:
        # Fallback: ritorna il valore normalizzato tradotto
        return tr(canonical)
