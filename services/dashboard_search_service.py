"""Search/filter helpers for dashboard orchestration."""

from __future__ import annotations


_GLOBAL_SEARCH_EN_TO_IT_EXACT = {
    # Closed vocabulary: VSM action
    "negotiation": "negoziazione",
    "other": "altro",
    "derisking": "derisking",
    # Closed vocabulary: Derisking supplier status
    "new": "nuovo",
    "under evaluation": "in valutazione",
    "qualified": "qualificato",
    "rejected": "scartato",
    # Backward-compatible alias for historical EN wording.
    "approved": "qualificato",
}


def _normalize_global_search_query_for_closed_domains(query):
    """Normalize known EN labels to canonical IT values for closed i18n vocabularies."""
    if query is None:
        return ""

    normalized = str(query).strip().lower()
    if not normalized:
        return normalized

    # Keep normalization explicit and predictable: exact-label replacement only.
    return _GLOBAL_SEARCH_EN_TO_IT_EXACT.get(normalized, normalized)


def has_active_search_filters(*, search_values, search_tipo_value, all_label, date_values):
    """Return True when any RFQ search filter is active."""
    for value in search_values.values():
        if (value or "").strip():
            return True

    if search_tipo_value != all_label:
        return True

    for value in date_values.values():
        if (value or "").strip():
            return True

    return False


def filter_derisking_suppliers_by_query(*, suppliers, query, fields=None):
    """Filter suppliers using case-insensitive substring match across configured fields."""
    normalized_query = _normalize_global_search_query_for_closed_domains(query)
    if not normalized_query:
        return list(suppliers)

    target_fields = fields or (
        "supplier_name",
        "category",
        "supplier_status",
        "contact_name",
        "email",
        "phone",
        "website",
        "notes",
        "username",
    )
    return [
        supplier
        for supplier in suppliers
        if any(normalized_query in (getattr(supplier, field_name) or "").lower() for field_name in target_fields)
    ]


def split_vsm_events_by_type(*, events, metadata, event_type):
    """Split events (and optional metadata) by event type preserving index alignment."""
    if metadata is not None:
        pairs = [(ev, meta) for ev, meta in zip(events, metadata) if ev.event_type == event_type]
        return [pair[0] for pair in pairs], [pair[1] for pair in pairs]
    return [event for event in events if event.event_type == event_type], None


def filter_vsm_events_by_query(*, events, metadata, query, fields):
    """Apply text search on VSM events while preserving metadata alignment."""
    normalized_query = _normalize_global_search_query_for_closed_domains(query)
    if metadata is not None:
        pairs = [
            (ev, meta)
            for ev, meta in zip(events, metadata)
            if any(normalized_query in (getattr(ev, field_name) or "").lower() for field_name in fields)
        ]
        return [pair[0] for pair in pairs], [pair[1] for pair in pairs]
    return [
        ev
        for ev in events
        if any(normalized_query in (getattr(ev, field_name) or "").lower() for field_name in fields)
    ], None
