# Plan: Dashboard Controller Extraction — `services/dashboard_controller.py`

Extract 5 dashboard-orchestration methods from `MainWindow` into a new `DashboardController` class. The recommended approach uses **delegation stubs** — the most conservative variant — so that zero external call sites need updating.

---

## Recommended approach: delegation-stub pattern

Move the real implementations to `DashboardController`; replace each method body in `MainWindow` with a single-line stub that forwards to the controller. This means:
- `main_dashboard_toolbar.py` — zero changes (its `hasattr` guards + `self.main_window.search_requests()` continue to work)
- `main_dashboard_builder.py` — zero changes (`app.search_requests`, `app.clear_filters`, `app.refresh_data` still exist on the app object)
- All 8 `self.refresh_data()` call sites inside other MainWindow methods — zero changes
- `on_tab_changed` call to `self._update_filter_panel_for_current_tab()` — zero changes

---

## Phase 1 — Create `services/dashboard_controller.py`

### 5 methods to receive (verbatim copy, `self.` → `self.app.`)

| Method | Line range in `dataflow.py` | Internal cross-calls within the 5 targets |
|---|---|---|
| `populate_username_filter` | 3716–3758 | none |
| `_update_filter_panel_for_current_tab` | 5211–5239 | none |
| `refresh_data` | 5408–5430 | → `self.search_requests()` (conditional), → `self.populate_username_filter()` |
| `search_requests` | 5685–6004 | none of the 5 |
| `clear_filters` | 6206–6238 | → `self.refresh_data()` |

Cross-calls **between the 5 extracted methods** (`refresh_data→search_requests`, `refresh_data→populate_username_filter`, `clear_filters→refresh_data`) stay as `self.xxx()` — they are all in the same `DashboardController` class, no `self.app.` needed.

The **one cross-module call** requiring `self.app.`: `clear_filters` calls `self._load_vsm_events(...)` → becomes `self.app._load_vsm_events(...)`.

All `self.xxx` **attribute accesses** (`self.search_vars`, `self.tree_attive`, `self.root`, etc.) become `self.app.xxx`.

### Imports for the new module (zero new dependencies — all already in `dataflow.py`)

```python
import re
import logging
from tkinter import messagebox
from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.i18n_utils import _, normalize_rfq_type
from utils.validation_utils import format_date_for_db

logger = logging.getLogger(__name__)
```

---

## Phase 2 — Modify `dataflow.py` (minimal)

### a) Add import (near other service imports)

```python
from services.dashboard_controller import DashboardController
```

### b) Add in `__init__` after `build_main_dashboard(self)`

```python
self.dashboard_controller = DashboardController(self)
```

Must be after `build_main_dashboard` — the controller's methods access UI widgets created by the builder.
`DashboardController.__init__` itself does nothing besides `self.app = app`, so ordering is safe.

### c) Replace 5 method bodies with delegation stubs (bodies only, signatures unchanged)

```python
def populate_username_filter(self):
    self.dashboard_controller.populate_username_filter()

def _update_filter_panel_for_current_tab(self):
    self.dashboard_controller._update_filter_panel_for_current_tab()

def refresh_data(self):
    self.dashboard_controller.refresh_data()

def search_requests(self):
    self.dashboard_controller.search_requests()

def clear_filters(self):
    self.dashboard_controller.clear_filters()
```

No other changes in `dataflow.py`. No changes to `main_dashboard_builder.py` or `main_dashboard_toolbar.py`.

---

## Relevant files

- [dataflow.py](dataflow.py) — `MainWindow` methods at lines 3716, 5211, 5408, 5685, 6206; `__init__` at line 3604 (builder call); imports block ~line 100
- [services/dashboard_controller.py](services/dashboard_controller.py) — new file to create
- [ui/main_dashboard_builder.py](ui/main_dashboard_builder.py) — unchanged
- [ui/components/main_dashboard_toolbar.py](ui/components/main_dashboard_toolbar.py) — unchanged

---

## Verification

1. `get_errors` on `services/dashboard_controller.py` and `dataflow.py` — zero new errors
2. Delegation stubs are syntactically correct (no method body left empty)
3. `hasattr` guards in `main_dashboard_toolbar.py` lines 204 and 219 still resolve correctly against `MainWindow` (they will, since stubs exist)
4. `clear_filters` in the controller calls `self.app._load_vsm_events(...)` not `self._load_vsm_events(...)`

---

## Decisions

- **Delegation stubs chosen over full call-site migration**: reduces change surface from ~15 edits across 3 files to 5 stub replacements in 1 file. Fully reversible by inlining stubs back.
- `_load_vsm_events` stays in `MainWindow` — not in scope; referenced via `self.app._load_vsm_events(...)` inside the controller only from `clear_filters`.
- `logger` in the new module uses `logging.getLogger(__name__)` — standard, functionally identical to the module-level `logger` in `dataflow.py`.
- Controller instantiated after `build_main_dashboard(self)` to preserve initialization order (UI must exist before controller is used at runtime).

---

## What is NOT touched

- `_load_vsm_events`, all VSM CRUD, export logic
- `main_dashboard_builder.py` (just refactored — untouched)
- `main_dashboard_toolbar.py` — its `hasattr` + `self.main_window.xxx()` calls continue to resolve through the delegation stubs
- All 8 `self.refresh_data()` call sites in other `MainWindow` methods
- `on_tab_changed` and all other `MainWindow` methods
- All database logic, backup logic, settings, attachment/PO windows
