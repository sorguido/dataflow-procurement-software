"""VSM dashboard data pipeline helpers.

Conservative extraction from dataflow.MainWindow keeping semantics unchanged.
"""

from __future__ import annotations

from datetime import datetime

from database_manager import DatabaseManager
from services.app_paths import get_db_path
from utils.i18n_utils import tr, translate_vsm_action


def get_vsm_dataset(*, vsm_username_filter, current_username):
    """Load raw VSM dataset according to active username scope."""
    with DatabaseManager(get_db_path()) as db_manager:
        if vsm_username_filter is None:
            raw = db_manager.get_all_vsm_events_aggregated(get_db_path())
            all_events = [ev for ev, _im, _src in raw]
            extra_meta = [{"is_mine": im, "source_file": src} for _, im, src in raw]
        elif vsm_username_filter == (current_username or "").lower():
            all_events = db_manager.get_all_vsm_events(username=current_username)
            extra_meta = None
        else:
            raw = db_manager.get_all_vsm_events_aggregated(get_db_path(), username=vsm_username_filter)
            all_events = [ev for ev, _im, _src in raw]
            extra_meta = [{"is_mine": im, "source_file": src} for _, im, src in raw]
    return all_events, extra_meta


def apply_vsm_filters(*, events, event_type, extra_meta=None, filters=None):
    """Apply VSM advanced filters preserving current semantics."""
    filters = filters or {}

    date_from_str = (filters.get("date_from") or "").strip()
    date_to_str = (filters.get("date_to") or "").strip()
    action_filter = (filters.get("action") or "").strip()
    repetitive_filter = (filters.get("repetitive") or "").strip()
    theoretical_from_str = (filters.get("theoretical_from") or "").strip()
    theoretical_to_str = (filters.get("theoretical_to") or "").strip()
    actual_from_str = (filters.get("actual_from") or "").strip()
    actual_to_str = (filters.get("actual_to") or "").strip()

    if not any([
        date_from_str,
        date_to_str,
        action_filter,
        repetitive_filter,
        theoretical_from_str,
        theoretical_to_str,
        actual_from_str,
        actual_to_str,
    ]):
        return events, extra_meta

    date_from = date_to = None
    fmt = "%d/%m/%Y"
    try:
        if date_from_str:
            date_from = datetime.strptime(date_from_str, fmt).date()
    except ValueError:
        pass
    try:
        if date_to_str:
            date_to = datetime.strptime(date_to_str, fmt).date()
    except ValueError:
        pass

    def parse_amount(value):
        if not value:
            return None
        value = value.strip()
        if "," in value:
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
        try:
            return float(value)
        except ValueError:
            return None

    theoretical_from = parse_amount(theoretical_from_str)
    theoretical_to = parse_amount(theoretical_to_str)
    actual_from = parse_amount(actual_from_str)
    actual_to = parse_amount(actual_to_str)

    use_dual_value = event_type in ("Saving", "Cost Avoidance")
    meta_iter = extra_meta if extra_meta is not None else [None] * len(events)
    filtered_pairs = []

    for event, meta in zip(events, meta_iter):
        if event.event_date:
            ev_date = event.event_date.date() if hasattr(event.event_date, "date") else event.event_date
            if date_from and ev_date < date_from:
                continue
            if date_to and ev_date > date_to:
                continue
        elif date_from or date_to:
            continue

        if action_filter and use_dual_value:
            if translate_vsm_action(event.action) != action_filter:
                continue

        if repetitive_filter:
            want = repetitive_filter == tr("Yes")
            if event.opex_ripetitivo != want:
                continue

        if theoretical_from is not None or theoretical_to is not None:
            tval = event.calculate_theoretical_value()
            if theoretical_from is not None and tval < theoretical_from:
                continue
            if theoretical_to is not None and tval > theoretical_to:
                continue

        if use_dual_value and (actual_from is not None or actual_to is not None):
            aval = event.calculate_effective_value()
            if actual_from is not None and aval < actual_from:
                continue
            if actual_to is not None and aval > actual_to:
                continue

        filtered_pairs.append((event, meta))

    filtered_events = [pair[0] for pair in filtered_pairs]
    filtered_meta = [pair[1] for pair in filtered_pairs] if extra_meta is not None else None
    return filtered_events, filtered_meta
