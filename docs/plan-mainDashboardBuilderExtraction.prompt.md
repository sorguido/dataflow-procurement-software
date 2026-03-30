# Plan: Conservative UI Extraction — `MainWindow.__init__` → `ui/main_dashboard_builder.py`

Extract the UI-construction block from `MainWindow.__init__` into a single `build_main_dashboard(app)` function in a new module, replacing it with one call. All widget assignments (`app.xxx = ...`) and method calls remain identical — only the location of the code changes.

---

## Steps

1. **Create `ui/main_dashboard_builder.py`** — new module with required imports (all reused from `dataflow.py`, no new packages) and the `build_main_dashboard(app)` function. The function body is the verbatim UI construction block from `__init__`, with `self.` → `app.` substitution.

2. **Modify `dataflow.py`** — two minimal changes:
   - Add `from ui.main_dashboard_builder import build_main_dashboard` near the other UI component imports (~line 88)
   - In `MainWindow.__init__`, delete the UI construction block and replace it with `build_main_dashboard(self)`

3. **Verify** with error checking on both files.

---

## Exact demarcation in `__init__`

| What | Decision |
|---|---|
| `self.root = root` … `self.active_db_path = ...` (state init) | **STAYS** in `__init__` |
| `frame_top = ttk.Frame(self.root)` … last `root.bind(...)` | **MOVES** to `build_main_dashboard` |
| VSM data loading loop + `populate_vsm_username_filter()` | **STAYS** in `__init__` — runtime data init, not UI construction |
| `self.refresh_data(); update_button_visibility(); check_for_autobackup()` | **STAYS** in `__init__` |

---

## Relevant files

- [dataflow.py](dataflow.py) — `MainWindow.__init__` lines ~3558–3851; imports block ~line 88
- [ui/main_dashboard_builder.py](ui/main_dashboard_builder.py) — new file to create
- [ui/components/main_dashboard_toolbar.py](ui/components/main_dashboard_toolbar.py) — referenced, unchanged
- [ui/components/collapsible_filters.py](ui/components/collapsible_filters.py) — referenced, unchanged

---

## Imports for the new module (all already used in `dataflow.py` — zero new dependencies)

```python
import os
import tkinter as tk
from tkinter import ttk
import webbrowser
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from utils.resource_utils import resource_path
from utils.i18n_utils import _, get_current_language
from ui.components.main_dashboard_toolbar import MainDashboardToolbar
from ui.components.collapsible_filters import CollapsibleFilters
```

---

## `__init__` skeleton after refactoring

```python
def __init__(self, root):
    # --- State initialization ---
    self.root = root
    set_window_icon(self.root)
    self.root.title(...)
    # window maximize (Linux/Windows compat)
    self.all_users_placeholder = ...
    self.username_filter_var = None
    self.user_filter_combo = None
    self.vsm_username_filter_var = None
    self.vsm_user_filter_combos = []
    self._load_identity_from_config()
    self.last_backup_date = None; self.db_path_standard = ...
    self._autobackup_timer_id = None
    self._sql_warning_after_id = None
    self._opening_request = False
    self.db_manager = DatabaseManager(get_db_path())
    self.active_db_path = get_db_path()

    # --- UI construction ---
    build_main_dashboard(self)

    # --- Initial data load (runtime init, not UI) ---
    for event_type, sheet in [
        ("Saving", self.sheet_saving),
        ("Cost Avoidance", self.sheet_cost_avoidance),
        ("Derisking", self.sheet_derisking),
    ]:
        self._load_vsm_events(event_type, sheet)
    self.populate_vsm_username_filter()

    # --- Post-build finalization ---
    self.refresh_data()
    self.update_button_visibility()
    self.check_for_autobackup()
```

---

## `build_main_dashboard(app)` internal structure

```
build_main_dashboard(app):
  frame_top = ttk.Frame(app.root)
  [logo loading → app.logo_photo]
  [buttons: app.btn_new_rdo, app.btn_actions, app.actions_menu,
            app.btn_mega_export, app.btn_kpi,
            app.btn_guida, app.btn_license, app.btn_settings]

  app.root.grid_rowconfigure / grid_columnconfigure
  frame_top.grid(row=0, ...)
  app.main_dashboard_toolbar = MainDashboardToolbar(app.root, app)   → grid(row=1)
  app.collapsible_filters = CollapsibleFilters(app.root, ...)        → grid(row=2)

  search_frame = app.collapsible_filters.filters_frame

  # RFQ filter subframe
  app.rfq_filter_subframe = ttk.Frame(search_frame)
  app.search_vars = {10 StringVar}; app.search_tipo = StringVar
  [all label/entry/combobox widgets]
  app.username_filter_var = StringVar(value=...)
  app.user_filter_combo = ttk.Combobox(...) → bound to app.refresh_data()
  app.date_entries = {4 DateEntry}

  # VSM filter subframe
  app.vsm_filter_subframe = ttk.Frame(search_frame)  → grid_remove()
  app.vsm_username_filter_var; app.vsm_action_var; app.vsm_repetitive_var
  app.vsm_theoretical_from_var; app.vsm_theoretical_to_var
  app.vsm_actual_from_var; app.vsm_actual_to_var
  app.vsm_date_from_entry; app.vsm_date_to_entry
  app.vsm_user_filter_combos = [_vsm_cb]
  [_vsm_sc_spec frame: action/repetitive/theoretical/actual filters]
  [_vsm_dr_spec frame: derisking-only repetitive filter] → grid_remove()
  app._vsm_spec_frames = {'vsm_saving': ..., 'vsm_cost_avoidance': ..., 'vsm_derisking': ...}

  # Shared search buttons
  btn_search_frame → Cerca / Pulisci Filtri

  # Notebook + tabs
  app.notebook = ttk.Notebook(app.root)  → grid(row=3)
  app.tab_attive, app.tab_archiviate, app.tab_saving,
  app.tab_cost_avoidance, app.tab_derisking
  app.sheet_saving = app._create_vsm_event_sheet(app.tab_saving, "Saving")
  app.sheet_cost_avoidance = app._create_vsm_event_sheet(app.tab_cost_avoidance, "Cost Avoidance")
  app.sheet_derisking = app._create_vsm_event_sheet(app.tab_derisking)

  # Footer
  footer_frame → version label + hyperlink → grid(row=4)

  # RFQ treeviews
  app.tree_attive = app.create_request_treeview(app.tab_attive)
  app.tree_archiviate = app.create_request_treeview(app.tab_archiviate)

  # Bindings
  app.notebook.bind("<<NotebookTabChanged>>", app.on_tab_changed)
  app.root.bind("<Button-1>", app._on_root_click, add="+")
```

---

## Decisions

- `_load_vsm_events()` loop and `populate_vsm_username_filter()` **stay in `__init__`** — they are runtime data initialization, not UI construction. `build_main_dashboard` is a pure widget builder: creates widgets, assigns `app.xxx`, sets layout, registers bindings — nothing more.
- `refresh_data()`, `update_button_visibility()`, `check_for_autobackup()` stay in `__init__` — they are clearly post-build finalization.
- `set_window_icon()`, `title()`, window maximize stay in `__init__` — they are root window configuration, not dashboard widget construction.
- All method bodies (handlers, `_load_vsm_events`, etc.) stay **untouched** in `MainWindow`.

---

## What is NOT touched

- `AttachmentWindow`, `PurchaseOrderWindow`, `SettingsWindow` — completely untouched
- All backup / restart / path-switching logic — untouched
- `search_requests()`, `refresh_data()`, `clear_filters()` — untouched
- `_load_vsm_events()` — untouched (only called from builder, not moved)
- All VSM CRUD / export logic — untouched
- All method signatures — unchanged

---

## Why this is low-risk

1. **Zero semantic changes** — the extracted code is a verbatim copy with `self.` → `app.` substitution; no logic is rewritten.
2. **Zero new dependencies** — all imports in the new module already exist in `dataflow.py`.
3. **Fully reversible** — to revert, inline `build_main_dashboard` back into `__init__` and replace `app.` with `self.`.
4. **All attribute assignments preserved** — every `app.xxx = widget` in `build_main_dashboard` corresponds 1:1 to the original `self.xxx = widget`, so all downstream methods work without modification.
5. **Import order preserved** — `init_i18n()` is still called at module level in `dataflow.py` before any UI is imported, so i18n strings in `build_main_dashboard` are already initialized when the function runs.
