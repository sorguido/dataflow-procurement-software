"""Policy helpers for Actions button state and menu composition."""

from __future__ import annotations


def compute_actions_capabilities(*, status, selected_count, all_mine):
    """Compute capability flags used by the Actions button/menu."""
    has_selection = selected_count > 0

    if status.startswith("vsm_"):
        can_delete = has_selection and all_mine
        can_duplicate = (selected_count == 1) and all_mine
        can_change_status = False
        can_act = has_selection and all_mine
    else:
        can_delete = has_selection and all_mine
        can_duplicate = (selected_count == 1) and all_mine
        can_change_status = has_selection and all_mine
        can_act = has_selection and all_mine

    return {
        "has_selection": has_selection,
        "can_delete": can_delete,
        "can_duplicate": can_duplicate,
        "can_change_status": can_change_status,
        "can_act": can_act,
    }


def build_actions_menu_spec(*, status, can_delete=False, can_duplicate=False, can_change_status=False):
    """Return a declarative menu spec consumed by UI code.

    Item formats:
    - ("command", key, enabled)
    - ("separator",)
    """
    if status.startswith("vsm_"):
        spec = [("command", "delete", can_delete)]
        if status != "vsm_derisking":
            spec.append(("command", "duplicate", can_duplicate))
        return spec

    spec = [
        ("command", "delete", can_delete),
        ("command", "duplicate", can_duplicate),
        ("separator",),
    ]
    if status == "attiva":
        spec.append(("command", "archive", can_change_status))
    else:
        spec.append(("command", "reactivate", can_change_status))
    return spec
