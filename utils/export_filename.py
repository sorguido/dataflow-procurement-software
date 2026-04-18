"""Helpers for user-facing Excel export filenames.

This module centralizes filename construction while keeping existing
export flows unchanged.
"""

from __future__ import annotations

import re
from datetime import datetime


_INVALID_CHARS_RE = re.compile(r"[^A-Za-z0-9_]+")
_MULTI_UNDERSCORE_RE = re.compile(r"_+")


def _sanitize_token(value) -> str:
    """Return a cross-platform safe token for filenames."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = _INVALID_CHARS_RE.sub("_", text)
    text = _MULTI_UNDERSCORE_RE.sub("_", text).strip("_")
    return text


def normalize_export_lang(lang_code) -> str:
    """Normalize language code/label to IT|EN."""
    if lang_code is None:
        return "EN"
    code = str(lang_code).strip().lower()
    if code in {"it", "ita", "italian", "italiano"}:
        return "IT"
    if code in {"en", "eng", "english"}:
        return "EN"
    return "EN"


def build_excel_export_filename(*parts, when: datetime | None = None) -> str:
    """Build a safe Excel filename with second-level timestamp.

    Format: <parts...>_YYYY-MM-DD_HH-MM-SS.xlsx
    """
    ts = when or datetime.now()
    safe_parts = []
    for part in parts:
        token = _sanitize_token(part)
        if token:
            safe_parts.append(token)
    safe_parts.append(ts.strftime("%Y-%m-%d"))
    safe_parts.append(ts.strftime("%H-%M-%S"))
    return "_".join(safe_parts) + ".xlsx"


def build_rfq_context(rfq_number) -> str:
    """Build the RfQ context token using business-visible request number."""
    token = _sanitize_token(rfq_number)
    if not token:
        return "RfQUnknown"
    return f"RfQ{token}"
