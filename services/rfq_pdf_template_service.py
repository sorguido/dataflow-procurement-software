"""Servizi template testuale esterno per export RFQ PDF."""

import os
import subprocess
import sys
from typing import Dict, Optional

from services.app_paths import get_user_documents_dataflow_dir
from utils.user_utils import get_app_data_dir

TABLE_PLACEHOLDER = "{{TABLE}}"
_LANG_ITA = "ita"
_LANG_ENG = "eng"

_DEFAULT_TEMPLATE_ENG = """Dear Supplier,
I kindly ask you to provide your best quotation for the following items:

{{TABLE}}

I look forward to your kind reply.
Best regards."""

_DEFAULT_TEMPLATE_ITA = """Gentile Fornitore,
con la presente sono a richiedere la Vs. migliore quotazione per il seguente materiale:

{{TABLE}}

In attesa di un Vs. gentile riscontro, porgo cordiali saluti."""


def _normalize_language(language_code: Optional[str]) -> str:
    value = (language_code or "").strip().lower()
    if value in ("it", "ita"):
        return _LANG_ITA
    return _LANG_ENG


def get_template_storage_dir() -> str:
    """Ritorna la cartella template RFQ_PDF, creandola se assente."""
    base_dir = get_user_documents_dataflow_dir()
    if not base_dir:
        base_dir = get_app_data_dir()
    template_dir = os.path.join(base_dir, "Assets", "RFQ_PDF")
    os.makedirs(template_dir, exist_ok=True)
    return template_dir


def get_template_path(language_code: Optional[str] = None) -> str:
    """Ritorna il path template in base alla lingua corrente."""
    normalized = _normalize_language(language_code)
    filename = "pdf_template_ita.txt" if normalized == _LANG_ITA else "pdf_template_eng.txt"
    return os.path.join(get_template_storage_dir(), filename)


def get_default_template_content(language_code: Optional[str] = None) -> str:
    """Ritorna contenuto di default del template per lingua."""
    normalized = _normalize_language(language_code)
    return _DEFAULT_TEMPLATE_ITA if normalized == _LANG_ITA else _DEFAULT_TEMPLATE_ENG


def ensure_template_file(language_code: Optional[str] = None) -> str:
    """Crea il template con contenuto default se non esiste."""
    template_path = get_template_path(language_code=language_code)
    if not os.path.exists(template_path):
        default_content = get_default_template_content(language_code=language_code)
        with open(template_path, "w", encoding="utf-8") as file_handle:
            file_handle.write(default_content)
    return template_path


def validate_template_content(template_text: Optional[str]) -> Dict[str, object]:
    """Valida testo template. Valido solo se {{TABLE}} appare esattamente una volta."""
    if template_text is None:
        return {"valid": False, "reason": "invalid_file"}

    normalized = str(template_text).replace("\r\n", "\n").replace("\r", "\n")
    if not normalized.strip():
        return {"valid": False, "reason": "empty_template"}

    placeholder_count = normalized.count(TABLE_PLACEHOLDER)
    if placeholder_count != 1:
        if placeholder_count == 0:
            return {"valid": False, "reason": "missing_placeholder"}
        return {"valid": False, "reason": "invalid_placeholder_count"}

    before_text, after_text = normalized.split(TABLE_PLACEHOLDER)
    return {
        "valid": True,
        "reason": None,
        "before_text": before_text,
        "after_text": after_text,
    }


def load_template_parts(language_code: Optional[str] = None) -> Dict[str, object]:
    """Carica template esterno; in caso di errore usa fallback interno."""
    template_path = ensure_template_file(language_code=language_code)
    default_content = get_default_template_content(language_code=language_code)

    try:
        with open(template_path, "r", encoding="utf-8") as file_handle:
            external_text = file_handle.read()
    except Exception:
        fallback_parts = validate_template_content(default_content)
        fallback_parts.update(
            {
                "template_path": template_path,
                "used_external_template": False,
                "fallback_reason": "invalid_file",
            }
        )
        return fallback_parts

    validation = validate_template_content(external_text)
    if validation.get("valid"):
        validation.update(
            {
                "template_path": template_path,
                "used_external_template": True,
                "fallback_reason": None,
            }
        )
        return validation

    fallback_parts = validate_template_content(default_content)
    fallback_parts.update(
        {
            "template_path": template_path,
            "used_external_template": False,
            "fallback_reason": validation.get("reason"),
        }
    )
    return fallback_parts


def open_template_with_system_app(language_code: Optional[str] = None) -> str:
    """Apre template con app di sistema (Windows/Linux/macOS)."""
    template_path = ensure_template_file(language_code=language_code)

    try:
        if sys.platform.startswith("win"):
            os.startfile(template_path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", template_path])
        else:
            subprocess.Popen(["xdg-open", template_path])
    except Exception as exc:
        raise RuntimeError(f"Impossibile aprire il file template: {exc}") from exc

    return template_path
