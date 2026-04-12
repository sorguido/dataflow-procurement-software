"""
Utility conservative per matching soft dei nomi fornitore.

NOTA:
- usate solo per confronto/suggerimento (mai per riscrivere dati storici)
- approccio volutamente semplice per ridurre falsi positivi
"""

import re


_MULTISPACE_RE = re.compile(r"\s+")
_PUNCT_RE = re.compile(r"[^\w\s]", re.UNICODE)
_SUFFIXES = {"srl", "srls", "spa", "snc", "sas", "sapa"}
_DOTTED_SRL_RE = re.compile(r"\bs\.\s*r\.\s*l\.?\b", re.IGNORECASE)
_DOTTED_SPA_RE = re.compile(r"\bs\.\s*p\.\s*a\.?\b", re.IGNORECASE)


def normalize_supplier_name_for_match(value: str) -> str:
    """
    Normalizza in modo soft una ragione sociale per matching.

    Regole:
    - trim + lower
    - rimozione punteggiatura
    - spazi multipli -> spazio singolo
    - equivalenza varianti come s.r.l -> srl, s.p.a -> spa (deriva dalla punteggiatura)
    """
    if not value:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""

    # Micro-fix conservativo: converte varianti puntate comuni prima della
    # pulizia punteggiatura, così "s.r.l." e "s.p.a." convergono a token canonici.
    text = _DOTTED_SRL_RE.sub("srl", text)
    text = _DOTTED_SPA_RE.sub("spa", text)

    text = _PUNCT_RE.sub(" ", text)
    text = _MULTISPACE_RE.sub(" ", text).strip()
    return text


def build_supplier_match_keys(value: str) -> tuple[str, str, bool]:
    """
    Ritorna (strict_key, base_key, has_legal_suffix).

    - strict_key: chiave normalizzata completa
    - base_key: strict_key senza suffisso legale finale (se presente)
    - has_legal_suffix: True se il nome termina con suffisso societario noto
    """
    strict_key = normalize_supplier_name_for_match(value)
    if not strict_key:
        return "", "", False

    tokens = strict_key.split(" ")
    if not tokens:
        return strict_key, strict_key, False

    last_token = tokens[-1]
    has_suffix = last_token in _SUFFIXES
    if has_suffix and len(tokens) > 1:
        base_key = " ".join(tokens[:-1]).strip()
    else:
        base_key = strict_key
    return strict_key, base_key, has_suffix
