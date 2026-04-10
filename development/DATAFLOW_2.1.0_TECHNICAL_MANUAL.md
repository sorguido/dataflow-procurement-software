# DataFlow 2.1.0 — Technical Manual

> **Audience:** Software developers who need to understand, maintain, or extend the DataFlow codebase.  
> **Scope:** Full architecture description, file-by-file analysis, data flows, and design rationale.  
> **Version:** 2.1.0 (current codebase as of April 2026)

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Repository Structure](#2-repository-structure)
3. [Architecture Overview](#3-architecture-overview)
4. [Directory-by-Directory Deep Dive](#4-directory-by-directory-deep-dive)
5. [Core Entry Points](#5-core-entry-points)
6. [UI Architecture](#6-ui-architecture)
7. [Business Logic](#7-business-logic)
8. [Data Layer](#8-data-layer)
9. [Internationalization](#9-internationalization)
10. [Build & Distribution](#10-build--distribution)
11. [Testing Strategy](#11-testing-strategy)
12. [Assets & Resources](#12-assets--resources)
13. [Utilities Layer](#13-utilities-layer)
14. [Execution Flow](#14-execution-flow)
15. [Design Principles](#15-design-principles)

---

## 1. Project Overview

### Purpose

DataFlow is a desktop procurement management application targeting purchasing departments (uffici acquisti). It centralizes the entire Request for Quotation (RFQ) workflow — from creation to supplier comparison — and extends into strategic purchasing activities via Value Stream Mapping (VSM) and KPI dashboards.

### Typology

- **Desktop application** — Python 3, Tkinter UI toolkit
- **Local-first** — SQLite database, stored on the user's local file system or a shared network drive
- **Multi-user capable** — WAL-mode SQLite allows concurrent reads from multiple users; each user's records are tagged with a `username`
- **Cross-platform in development** — the codebase runs on Linux and macOS in source form; the primary distribution target is Windows (MSIX/EXE via PyInstaller)

### Primary Features

| Feature | Description |
|---|---|
| **RFQ Management** | Create, edit, archive, and duplicate Requests for Quotation (full supply or work order type) |
| **Supplier Comparison** | Compare unit prices from multiple suppliers per RFQ line item in a spreadsheet-like grid |
| **Attachments** | Attach documents to RFQs either as external file links or as BLOBs embedded in the database |
| **Notes with Formatting** | Rich text notes (bold, italic, underline) per RFQ, serialized as JSON |
| **Purchase Orders** | Associate purchase order numbers to closed RFQs |
| **SQDC Analysis** | Safety/Quality/Delivery/Cost weighted scoring analysis per RFQ, with Excel export |
| **VSM — Saving** | Track negotiated savings against budget; OPEX events distribute value over up to 24 months with first-month pro-rata |
| **VSM — Cost Avoidance** | Track avoided cost increases |
| **VSM — Derisking** | Manage a registry of potential/new suppliers being evaluated for supply chain risk reduction |
| **KPI Dashboard** | Aggregated KPI cards and bar charts for RFQ throughput, Saving, Cost Avoidance, and Derisking, with Excel export |
| **Multi-user Aggregation** | The dashboard can aggregate data from multiple users' databases on a shared drive |
| **Settings** | Language selection, backup path, auto-backup scheduling, DataFlow folder location |
| **i18n** | Full UI translation: English and Italian, switchable at runtime |

---

## 2. Repository Structure

```
dataflow-dev/
├── app.manifest.xml               # Windows DPI-awareness manifest (root copy, used by PyInstaller)
├── CHANGELOG.md                   # User-facing changelog
├── constants.py                   # UI layout constants (pixel dimensions, percentages)
├── database_manager.py            # SQLite database access layer (single class)
├── dataflow.py                    # Application entry point + MainWindow + SettingsWindow
├── dataflow.spec                  # PyInstaller build spec for Windows EXE
├── LICENSE
├── README.md
├── requirements.txt               # Runtime Python dependencies
│
├── add_data/                      # Bundled static assets
│   ├── DataFlow.ico               # Application icon
│   ├── Logo*.png                  # Logo variants (multiple sizes)
│   ├── logo_dataflow.png          # Toolbar logo
│   ├── template_rdo.xlsx          # Excel template: full-supply RFQ (Italian)
│   ├── template_rdo_cl.xlsx       # Excel template: work-order RFQ (Italian)
│   ├── template_rdo_eng.xlsx      # Excel template: full-supply RFQ (English)
│   ├── template_rdo_eng_cl.xlsx   # Excel template: work-order RFQ (English)
│   ├── template_sqdc.xlsx         # Excel template: SQDC analysis (Italian)
│   └── template_sqdc_eng.xlsx     # Excel template: SQDC analysis (English)
│
├── database/
│   ├── __init__.py
│   └── db_helpers.py              # crea_database_v4(): initializes DB tables on first run
│
├── dev_tools/
│   ├── compile_translations.py    # Compiles .po → .mo using polib
│   └── tools_build_WIN/
│       ├── exe/
│       │   ├── app.manifest.xml   # Windows manifest for EXE build
│       │   └── DataFlow.spec      # PyInstaller spec for EXE
│       └── msix/
│           └── AppxManifest.xml   # Windows Store MSIX package manifest
│
├── docs/
│   ├── CHANGELOG_2.1.0_ITA.md
│   ├── dataflow_structure.txt     # Directory tree snapshot
│   ├── README_2.1.0_ITA.md
│   ├── screenshot/
│   ├── video/
│   └── wiki/                      # GitHub Wiki source (EN + IT)
│       ├── EN/
│       └── IT/
│
├── locale/
│   ├── en/LC_MESSAGES/
│   │   ├── dataflow.po            # English translation source
│   │   └── dataflow.mo            # Compiled binary catalog
│   └── it/LC_MESSAGES/
│       ├── dataflow.po            # Italian translation source
│       └── dataflow.mo            # Compiled binary catalog
│
├── models/
│   ├── __init__.py
│   ├── potential_supplier.py      # PotentialSupplier dataclass
│   ├── vsm_event.py               # VSMEvent dataclass + calculation methods
│   └── vsm_impact.py              # VSMImpact dataclass
│
├── services/
│   ├── __init__.py
│   ├── app_paths.py               # Path resolution: data dirs, DB path, config
│   ├── dashboard_controller.py    # Dashboard orchestration: search, refresh, filters
│   ├── kpi_chart_data.py          # Time-series data for KPI charts
│   ├── kpi_engine.py              # KPI calculation engine (read-only, no UI)
│   ├── kpi_excel_export.py        # KPI Excel workbook builder (openpyxl)
│   ├── startup_service.py         # Logging setup, temp cleanup, directory init
│   ├── supplier_category_persistence.py  # CRUD for supplier_categories table
│   ├── supplier_persistence.py    # CRUD for potential_suppliers table
│   ├── vsm_engine.py              # VSM impact generation engine
│   └── vsm_persistence.py        # VSM transactional save/update/delete
│
├── tests/
│   ├── __init__.py
│   ├── test_supplier_category_persistence.py
│   ├── test_vsm_engine.py
│   ├── test_vsm_event_model.py
│   ├── test_vsm_persistence.py
│   └── db_test_tool/
│       ├── generate_test_db_eng.py
│       └── generate_test_db_it.py
│
├── ui/
│   ├── kpi_chart.py               # Pure-canvas bar chart rendering (no external deps)
│   ├── kpi_window.py              # KPI Analysis window (tabs: RFQ/Saving/CA/Derisking)
│   ├── main_dashboard_builder.py  # Builds all MainWindow widgets (pure UI, no logic)
│   ├── window_launchers.py        # Opens Help, License, KPI windows
│   ├── components/
│   │   ├── __init__.py
│   │   ├── collapsible_filters.py # Collapsible "Advanced Filters" container
│   │   └── main_dashboard_toolbar.py  # Global search bar component
│   ├── dialogs/
│   │   ├── __init__.py
│   │   ├── common_dialogs.py      # Reusable dialogs: message, yes/no, splash, user identity
│   │   ├── manage_supplier_categories_dialog.py  # Category CRUD dialog
│   │   ├── potential_supplier_dialog.py           # Create/edit supplier dialog
│   │   └── vsm_event_dialog.py                    # Create/edit VSM event dialog
│   └── windows/
│       ├── __init__.py
│       ├── attachment_window.py   # Attachment management (BLOB + external link)
│       ├── edit_reference_window.py  # Edit RFQ reference field
│       ├── edit_suppliers_window.py  # Edit supplier list for an RFQ
│       ├── notes_window.py        # Rich-text note editor
│       ├── purchase_order_window.py  # Purchase order number management
│       ├── sqdc_analysis_window.py   # SQDC weighted scoring + export
│       └── view_request_window.py    # Main RFQ detail/edit window
│
└── utils/
    ├── __init__.py
    ├── format_utils.py            # Number/currency formatting (Italian locale conventions)
    ├── i18n_utils.py              # gettext init, language detection, helper functions
    ├── resource_utils.py          # resource_path() for PyInstaller, window icon setter
    ├── string_utils.py            # Username generation, accent stripping
    ├── user_utils.py              # config.ini read/write, user identity persistence
    ├── validation_utils.py        # Input sanitization: filenames, dates, prices, email, URL
    ├── vsm_config.py              # VSM configurables: pagamenti_coefficient
    └── window_utils.py            # Window sizing, centering, DPI-aware geometry
```

---

## 3. Architecture Overview

### Layers

```
┌──────────────────────────────────────────────────────────┐
│   UI Layer                                               │
│   dataflow.py (MainWindow, SettingsWindow)               │
│   ui/kpi_window.py                                       │
│   ui/main_dashboard_builder.py                           │
│   ui/components/    ui/dialogs/    ui/windows/           │
└──────────────────────┬───────────────────────────────────┘
                       │ calls
┌──────────────────────▼───────────────────────────────────┐
│   Services Layer                                         │
│   dashboard_controller.py   kpi_engine.py               │
│   vsm_engine.py             vsm_persistence.py          │
│   supplier_persistence.py   supplier_category_...py     │
│   kpi_chart_data.py         kpi_excel_export.py         │
│   app_paths.py              startup_service.py          │
└──────────────────────┬───────────────────────────────────┘
                       │ uses
┌──────────────────────▼───────────────────────────────────┐
│   Models Layer                                           │
│   vsm_event.py   vsm_impact.py   potential_supplier.py  │
└──────────────────────┬───────────────────────────────────┘
                       │ persisted via
┌──────────────────────▼───────────────────────────────────┐
│   Database Layer                                         │
│   database_manager.py   database/db_helpers.py          │
│   SQLite WAL file on disk                                │
└──────────────────────────────────────────────────────────┘
                  cross-cutting
┌──────────────────────────────────────────────────────────┐
│   Utils Layer                                            │
│   constants.py   utils/*.py                             │
└──────────────────────────────────────────────────────────┘
```

### Layer Responsibilities

| Layer | Responsibility | What it must NOT do |
|---|---|---|
| **UI** | Build widgets, bind events, display data, collect user input | Contain business rules, access DB directly (exception: `dataflow.py` legacy code) |
| **Services** | Business logic, orchestration, DB queries, file generation | Build Tkinter widgets |
| **Models** | Data structures, field-level calculations, validation on init | Access DB, build UI |
| **Database** | Connection management, SQL execution, schema migrations | Business logic |
| **Utils** | Stateless helper functions, formatting, path resolution | Maintain state, access DB |

### General Data Flow

1. User interacts with a widget in the UI layer.
2. The UI calls a method on `MainWindow` (or a sub-window), which delegates to `DashboardController` or directly instantiates a service function.
3. The service opens a `DatabaseManager` context manager (`with DatabaseManager(get_db_path()) as db:`), executes queries, and returns plain Python data (lists of tuples, dataclasses, dicts).
4. The UI receives the data and updates widgets (Treeview rows, Sheet cells, Label text).
5. Mutations (INSERT/UPDATE/DELETE) go through service functions that build model objects and call `DatabaseManager` methods; transactional guarantees are enforced at the service layer.

---

## 4. Directory-by-Directory Deep Dive

### 4.1 Root-level files

#### `dataflow.py`

See [Section 5.1](#51-dataflowpy) for full analysis.

#### `database_manager.py`

See [Section 5.2](#52-database_managerpy) for full analysis.

#### `constants.py`

- **Responsibility:** Single source of truth for layout constants used across the UI.
- **Input:** None (pure declarations).
- **Output:** Imported by `window_utils.py`, `dataflow.py`, and any window that needs geometry calculations.
- **Constants defined:**

| Constant | Value | Purpose |
|---|---|---|
| `TASKBAR_BUFFER` | 100 px | Anti-overlap buffer below taskbar |
| `BASE_ARTICLE_WIDTH` | 470 px | Width of article columns in ViewRequestWindow |
| `CONTO_LAVORO_WIDTH` | 350 px | Extra width for work-order columns |
| `SUPPLIER_COLUMN_WIDTH` | 120 px | Width per supplier column |
| `PADDING` | 140 px | Total lateral padding (margins + scrollbar + DPI safety) |
| `BUTTONS_MIN_WIDTH` | 1150 px | Minimum width to fit all toolbar buttons |
| `MIN_WINDOW_WIDTH` | 850 px | Absolute minimum for any window |
| `SCREEN_WIDTH_PERCENTAGE` | 0.95 | ViewRequestWindow target width as fraction of screen |
| `SCREEN_HEIGHT_PERCENTAGE` | 0.80 | ViewRequestWindow target height as fraction of screen |

#### `requirements.txt`

Runtime dependencies (no dev/test extras listed):

| Package | Version | Purpose |
|---|---|---|
| `openpyxl` | 3.1.5 | Read/write `.xlsx` templates and export files |
| `Pillow` | 12.1.1 | Load PNG/ICO images for window icons and logo |
| `polib` | 1.2.0 | Compile `.po` → `.mo` (dev tool only, not imported at runtime) |
| `tkcalendar` | 1.6.1 | `DateEntry` widget for date fields |
| `tksheet` | 7.6.0 | Spreadsheet-like grid for article/price data |

> **Note:** `sqlite3` is part of the Python standard library and is not listed.

#### `app.manifest.xml`

Windows DPI-awareness manifest embedded in the EXE by PyInstaller. Declares `PerMonitorV2` DPI awareness so that Tkinter scales correctly on high-DPI monitors without blurring. A copy lives at `dev_tools/tools_build_WIN/exe/app.manifest.xml` for the Windows-specific build pipeline.

---

### 4.2 `add_data/`

Static assets bundled into the PyInstaller distribution via the spec file's `datas` list. Accessed at runtime via `resource_path()`.

| File | Purpose |
|---|---|
| `DataFlow.ico` | Window icon (all Toplevel windows via `set_window_icon()`) |
| `Logo.png`, `Logo_*.png` | Logo variants (150×150, 44×44, 50×50 — for different contexts) |
| `logo_dataflow.png` | Horizontal logo displayed in the main dashboard toolbar |
| `template_rdo.xlsx` | RFQ export template — Italian, full supply type |
| `template_rdo_cl.xlsx` | RFQ export template — Italian, work-order (Conto Lavoro) type |
| `template_rdo_eng.xlsx` | RFQ export template — English, full supply type |
| `template_rdo_eng_cl.xlsx` | RFQ export template — English, work-order type |
| `template_sqdc.xlsx` | SQDC analysis export template — Italian |
| `template_sqdc_eng.xlsx` | SQDC analysis export template — English |

The export logic (`ViewRequestWindow`, `SQDCAnalysisWindow`) opens the appropriate template, populates it with data from the database, and saves/opens the result. Template selection is driven by `get_current_language()` and the RFQ `tipo_rdo` field.

---

### 4.3 `database/`

#### `__init__.py`

Empty package marker.

#### `db_helpers.py`

- **Responsibility:** `crea_database_v4()` — the single function called at application startup to ensure the database schema exists.  
- **Input:** None directly; reads `config.ini` via `get_config_file()` and `get_db_path()` to locate the database file.  
- **Output:** Side effect — creates or migrates the SQLite database file and all tables.  
- **Dependencies:** `DatabaseManager`, `get_db_path()`, `get_config_file()`.  
- **Behavior:**
  1. Resolves the DB file path, preferring `custom_db_path` from config if set.
  2. Instantiates `DatabaseManager` and calls `create_tables()`.
  3. Logs success or raises on critical failure.
  4. Does **not** create the DB file if the user identity is not yet set (the caller controls this gate).

---

### 4.4 `dev_tools/`

Development-only utilities, not imported by the application at runtime.

#### `compile_translations.py`

- **Responsibility:** Iterates over `locale/en/` and `locale/it/`, loads each `dataflow.po` with `polib`, and saves the compiled `dataflow.mo` binary catalog.
- **Usage:** Run manually by the developer after editing `.po` files: `python dev_tools/compile_translations.py`
- **Dependencies:** `polib` (listed in `requirements.txt` for this reason).

#### `tools_build_WIN/exe/DataFlow.spec`

Windows-specific PyInstaller spec. Differs from root `dataflow.spec` in Windows-targeting details (Tcl/Tk path detection, manifest embedding). See [Section 10](#10-build--distribution).

#### `tools_build_WIN/msix/AppxManifest.xml`

Windows Store MSIX package manifest. Declares the package identity, display name, capabilities, and application entry point for distribution through the Microsoft Store or enterprise MSIX deployment.

---

### 4.5 `locale/`

See [Section 9](#9-internationalization) for full analysis.

---

### 4.6 `models/`

Pure data containers — `@dataclass` classes with no DB access and no Tkinter imports.

#### `vsm_event.py` — `VSMEvent`

- **Responsibility:** Represents a single VSM action (Saving, Cost Avoidance, or Derisking) with all its economic attributes.  
- **Input:** Keyword arguments; `__post_init__` converts ISO-format string dates to `datetime` objects.  
- **Key fields:**

| Field | Type | Description |
|---|---|---|
| `id` | `Optional[int]` | DB primary key; `None` for new, unsaved events |
| `username` | `str` | Owner username (required for multi-user isolation) |
| `event_type` | `str` | `"Saving"`, `"Cost Avoidance"`, or `"Derisking"` |
| `driver` | `str` | `"Prezzo"` (price-based) or `"Pagamenti"` (payment-terms-based) |
| `opex_ripetitivo` | `bool` | If True, impacts repeat monthly for up to 24 months |
| `importo_bdg` | `float` | Budget amount |
| `importo_negoziato` | `float` | Negotiated amount |
| `percent_realizzo` | `float` | Realization rate (0–100); scales effective vs theoretical value |
| `giorni_pagamento_*` | `Optional[int]` | Current/negotiated payment days (driver=Pagamenti only) |
| `payments_rate` | `Optional[float]` | Override for the monthly cost-of-capital rate |

- **Methods:**
  - `calculate_theoretical_value()` — computes the annual economic value before realization adjustment, based on driver type.
  - `calculate_effective_value()` — applies `percent_realizzo / 100` to the theoretical value.

#### `vsm_impact.py` — `VSMImpact`

- **Responsibility:** Represents one monthly economic impact generated by a `VSMEvent`.  
- **Generated by:** `vsm_engine.generate_impacts_for_event()`, never created by hand in application code.  
- **Key fields:** `event_id`, `username`, `year`, `month`, `value_type`, `valore_teorico`, `valore_effettivo`.  
- **Validation:** `__post_init__` raises `ValueError` if `month` is outside 1–12.  
- **Helpers:** `to_dict()`, `from_dict()`, `period_key` (→ `"YYYY-MM"`), `get_realizzo_percentage()`.

#### `potential_supplier.py` — `PotentialSupplier`

- **Responsibility:** Represents an entry in the derisking supplier registry.  
- **Key fields:** `supplier_name`, `category`, `supplier_status` (enum: Nuovo / In valutazione / Qualificato / Scartato), contact details, `username`.  
- **Status constants:** `SUPPLIER_STATUS_CHOICES` list is used by UI comboboxes and KPI queries.  
- **Dependencies:** None (pure dataclass).

---

### 4.7 `services/`

The business and orchestration layer. Each module has a narrow, declared scope.

#### `app_paths.py`

- **Responsibility:** Resolve all file system paths used by the application.
- **Input:** `config.ini` (via `get_config_file()`).
- **Output:** String paths.
- **Key functions:**

| Function | Returns | Notes |
|---|---|---|
| `get_user_documents_dataflow_dir()` | `str \| None` | `~/Documents/DataFlow_<username>` on Windows; `~/DataFlow_<username>` on Linux/macOS. Respects `dataflow_base_dir` override in config. Returns `None` if no username set yet. |
| `get_fixed_db_dir()` | `str` | `<dataflow_dir>/Database/` |
| `get_fixed_attachments_dir()` | `str \| None` | `<dataflow_dir>/Attachments/` (migrates old `Allegati/` name automatically) |
| `get_db_path()` | `str` | Session-cached DB path. Priority: (1) `dataflow_base_dir`, (2) `custom_db_path` from config, (3) standard `Database/dataflow_db.db`. Cached after first call for session consistency. |
| `reset_db_cache()` | `None` | Invalidates `_PERCORSO_DB_CACHE`; called after user changes DB location. |

- **Caching strategy:** `_PERCORSO_DB_CACHE` (module-level singleton) ensures that `get_db_path()` returns the same path throughout a session, even if config changes mid-session.

#### `startup_service.py`

- **Responsibility:** Application initialization tasks that run once at startup.
- **Functions:**

| Function | Description |
|---|---|
| `cleanup_temp_on_startup()` | Deletes leftover `_MEI*` PyInstaller temp folders and temporary DataFlow files older than 24 hours. |
| `setup_logging()` | Creates a `RotatingFileHandler` (5 MB × 3 backups) writing to `~/.local/share/DataFlow/dataflow.log` (Linux/macOS) or `~/AppData/Local/DataFlow/dataflow.log` (Windows). Returns the configured `logging.Logger`. |
| `initialize_dataflow_directory_structure(base_dir)` | Creates `Database/` and `Attachments/` subdirectories under the DataFlow base folder. Migrates `Allegati/` → `Attachments/` if needed. Does **not** create the database file. |

#### `dashboard_controller.py` — `DashboardController`

- **Responsibility:** Orchestrates the main dashboard's data refresh, search, and filter logic. Extracted from `MainWindow` in v2.1.0 to reduce `dataflow.py` size.
- **Input:** `app` — a reference to the live `MainWindow` instance.
- **Key methods:**

| Method | Description |
|---|---|
| `refresh_data()` | Reloads RFQ and VSM data preserving active filters. Uses `get_all_richieste_aggregated()` for a single DB round-trip. |
| `search_requests()` | Dispatches to `_search_vsm_events()`, `_search_derisking_suppliers()`, or the RFQ search handler based on the active tab. |
| `populate_username_filter(all_requests)` | Builds the username dropdown from all records, including all known users. Falls back to local DB on aggregation failure. |
| `_update_filter_panel_for_current_tab()` | Shows/hides the RFQ or VSM sub-frame inside the collapsible filter panel depending on which Notebook tab is active. |

#### `kpi_engine.py`

- **Responsibility:** Pure computation — reads DB, calculates KPI values, returns dictionaries. No UI, no writes.
- **Input:** `db_path`, `date_from`, `date_to`, `year` (all optional).
- **Output:** Flat `dict` with typed KPI values.
- **Public API:**

| Function | Returns | Contains |
|---|---|---|
| `get_rfq_kpi(...)` | `dict` | Total RFQs, active/archived split, supplier count, RFQ type breakdown, average response time, on-time %, supplier coverage |
| `get_saving_kpi(...)` | `dict` | Event count, theoretical/effective saving (annual + total), weighted avg realization %, OPEX vs CAPEX split |
| `get_cost_avoidance_kpi(...)` | `dict` | Event count, total avoided cost theoretical/effective |
| `get_derisking_kpi(...)` | `dict` | Total suppliers, status breakdown (Nuovo/In valutazione/Qualificato/Scartato), category count |

- **Internal helpers:** `_build_date_filter()`, `_build_impact_period_filter()`, `_where()`, `_scalar()`, `_safe_pct()`, `_pct_stats()` — all private, not part of the public API.
- **Error handling:** All functions catch all exceptions and return zero-value dicts; they never raise to the caller.

#### `kpi_chart_data.py`

- **Responsibility:** Prepares time-series data for the KPI window's bar charts using a deterministic bucket-first approach.
- **Design principle:** The time domain (list of monthly buckets) is built from the filter parameters, **not** from the data. Every bucket in the range is guaranteed to appear; missing months have value 0. This prevents charts from silently hiding gaps.
- **Public functions:**

| Function | Returns | Description |
|---|---|---|
| `get_rfq_chart_data(date_from, date_to, year, db_path)` | `list[{'label', 'count'}]` | RFQ creation count per month bucket |
| `get_saving_chart_data(...)` | `list[{'label', 'theoretical', 'actual'}]` | Monthly Saving impact (€) |
| `get_cost_avoidance_chart_data(...)` | `list[{'label', 'theoretical', 'actual'}]` | Monthly Cost Avoidance impact (€) |
| `get_derisking_chart_data(...)` | `list[{'label', 'count'}]` | New supplier registrations per month |

#### `kpi_excel_export.py`

- **Responsibility:** Builds an `openpyxl.Workbook` from pre-computed KPI data. No DB access, no UI.
- **Input:** Four dicts from `kpi_engine` (`rfq_data`, `saving_data`, `ca_data`, `derisking_data`) plus locale and filter parameters.
- **Output:** `openpyxl.Workbook` object (caller saves it to disk).
- **Public function:** `build_kpi_workbook(rfq_data, saving_data, ca_data, derisking_data, is_ita, date_from, date_to, year)`.
- **Internal formatters:** `_hdr()`, `_dat()`, `_t()`, `_i()`, `_f()`, `_FMT_MONEY`, `_FMT_PCT` — private styling helpers.

#### `vsm_engine.py`

See [Section 7.2](#72-vsm-engine) for full analysis.

#### `vsm_persistence.py`

See [Section 7.3](#73-vsm-persistence) for full analysis.

#### `supplier_persistence.py`

- **Responsibility:** CRUD operations for the `potential_suppliers` table.
- **Input/Output:** `DatabaseManager` instance + `PotentialSupplier` model objects.
- **Public functions:** `create_supplier()`, `update_supplier()`, `get_supplier_by_id()`, `get_all_suppliers()`, `delete_supplier()`, `get_distinct_macrocategories()`, `get_supplier_kpi()`.
- **Error hierarchy:** Raises `SupplierError` for business rule violations (empty name, missing ID for update); propagates `DatabaseError` as-is for DB failures.
- **`get_supplier_kpi()`:** Returns aggregate stats: total count, count by status, count by category — used by `kpi_engine.get_derisking_kpi()`.

#### `supplier_category_persistence.py`

- **Responsibility:** Manage the `supplier_categories` table (official category catalog) and keep `potential_suppliers.category` synchronized.
- **Key operations:**

| Function | Description |
|---|---|
| `get_all_supplier_categories(db)` | Returns sorted list of category names |
| `ensure_supplier_category_exists(db, name)` | Idempotent upsert; trims whitespace; no-op on empty input |
| `rename_supplier_category(db, old, new)` | Transactional: updates all supplier records, renames in catalog. Blocks if `new` already exists (suggests merge). |
| `merge_supplier_categories(db, source, target)` | Moves all suppliers from `source` → `target`, deletes `source` from catalog |
| `delete_supplier_category_if_unused(db, name)` | Deletes only if no suppliers reference it; raises `CategoryError` otherwise |
| `count_suppliers_by_category(db, name)` | Returns count of suppliers in a category |

---

### 4.8 `tests/`

See [Section 11](#11-testing-strategy) for full analysis.

---

### 4.9 `ui/`

See [Section 6](#6-ui-architecture) for full analysis.

---

### 4.10 `utils/`

See [Section 13](#13-utilities-layer) for full analysis.

---

## 5. Core Entry Points

### 5.1 `dataflow.py`

**Role:** Application entry point, root `tk.Tk` instance owner, and host of the two largest classes: `MainWindow` and `SettingsWindow`.

**Responsibilities:**

1. **DPI Awareness (lines 1–27):** Before any Tkinter import, on Windows it calls `SetProcessDpiAwarenessContext(PER_MONITOR_V2)` via `ctypes.windll`. This prevents Tkinter from rendering blurry on high-DPI displays. The call is isolated in a try/except so failure is non-fatal.

2. **Global imports:** All third-party and internal modules are imported. The i18n system (`init_i18n()`) is initialized before any UI module import to ensure `_()` is available at module-level string evaluation time.

3. **`SettingsWindow` class:** A `tk.Toplevel` dialog that allows the user to:
   - Change the DataFlow folder location (persisted in `config.ini` as `dataflow_base_dir`)
   - Configure and trigger manual backup
   - Configure hourly auto-backup scheduling with a 3-copy rotation
   - Switch UI language (requires application restart)

4. **`MainWindow` class** (the application core):
   - Inherits from nothing; wraps a `tk.Tk` instance passed as `root`.
   - Owns all dashboard state: active Treeview trees (`tree_attive`, `tree_archiviate` for RFQs; VSM trees for Saving, Cost Avoidance, Derisking), filter variables, identity variables, button references.
   - Delegates widget construction to `build_main_dashboard(app)`.
   - Delegates search/refresh to `DashboardController(self)`.
   - Contains large in-line methods for: RFQ creation, duplication, archiving, Excel mega-export, multi-user DB aggregation, auto-backup scheduler, VSM event operations, derisking operations.
   - The `restart_program()` method re-launches the Python process via `subprocess` and calls `sys.exit()`.

5. **Entry block** (`if __name__ == '__main__':`):
   - Calls `cleanup_temp_on_startup()`.
   - Configures logging.
   - Creates `tk.Tk()` root window.
   - Checks for user identity; shows `UserIdentityDialog` on first run.
   - Calls `crea_database_v4()` to ensure schema exists.
   - Creates and runs `MainWindow(root)`.
   - Starts `root.mainloop()`.

**Input:** OS process launch (no CLI arguments used by the application logic).  
**Output:** Running Tkinter event loop; all state is side-effect-driven via DB and config file.  
**Key dependencies:** All of `utils/`, `services/`, `database/`, `ui/`, `database_manager.py`, `constants.py`.

---

### 5.2 `database_manager.py`

**Role:** The single database access class for the entire application. All SQL is executed through this class.

**Class: `DatabaseManager`**

**Constructor:** `__init__(db_name, read_only=False)`
- Opens a SQLite connection.
- In read-write mode: enables WAL journal mode, `synchronous=NORMAL`, 64 MB cache, temp store in memory, `busy_timeout=10000 ms`.
- In read-only mode: opens via URI (`file:...?mode=ro`) for concurrent read access from other users' processes.
- Sets `row_factory = sqlite3.Row` for dict-like row access.
- Raises `DatabaseError` (the application-specific exception) on connection failure.

**Context manager:** Implements `__enter__` / `__exit__`; the canonical usage pattern throughout the codebase is:
```python
with DatabaseManager(get_db_path()) as db:
    rows = db.some_query(...)
```

**Schema (`create_tables()`):** Idempotent method using `CREATE TABLE IF NOT EXISTS` and `ALTER TABLE ... ADD COLUMN` in try/except blocks for migrations. Tables created:

| Table | Primary Key | Description |
|---|---|---|
| `fornitori` | `id_fornitore AUTOINCREMENT` | Supplier registry (name only) |
| `richieste_offerta` | `id_richiesta INTEGER` (year-driven, not AUTOINCREMENT) | RFQ header |
| `dettagli_richiesta` | `id_dettaglio AUTOINCREMENT` | RFQ line items (articles) |
| `richiesta_fornitori` | `(id_richiesta, nome_fornitore)` | M:N link: which suppliers per RFQ |
| `offerte_ricevute` | `(id_dettaglio, nome_fornitore)` | Unit price per article per supplier |
| `allegati_richiesta` | `id_allegato AUTOINCREMENT` | Attachments (BLOB or external link) |
| `vsm_events` | `event_id AUTOINCREMENT` | VSM event headers |
| `vsm_impacts` | `impact_id AUTOINCREMENT` | Monthly economic impacts |
| `potential_suppliers` | `supplier_id AUTOINCREMENT` | Derisking supplier registry |
| `supplier_categories` | `id AUTOINCREMENT` | Official category catalog |

**RFQ ID generation strategy (`insert_richiesta_offerta()`):** IDs are not pure AUTOINCREMENT. The formula is:
```
year_base = YY * 100_000   (e.g., 2026 → 2600000)
next_id = max(year_base, max_existing_id + 1)
```
This ensures IDs carry year information (e.g., `2600001`) and reset logically at year boundaries.

**Methods overview (CRUD by category):**

- `insert_*` methods: return the new primary key (`lastrowid`); commit immediately.
- `update_*` methods: no return value; commit immediately.
- `get_*` methods: return tuples or lists of `sqlite3.Row` objects.
- `delete_*` methods: commit immediately.
- VSM-specific `_insert_vsm_event_no_commit()` / `_insert_vsm_impacts_no_commit()`: used exclusively by `vsm_persistence.py` within an explicit transaction.
- `get_all_richieste_aggregated(db_path)`: scans the parent directory of `db_path` for other user databases and performs a `UNION ALL` across all of them, returning a combined view of all RFQs for the multi-user aggregated dashboard.

**`DatabaseError` exception:** A custom exception class wrapping all SQLite errors, isolating callers from the underlying `sqlite3` module.

---

## 6. UI Architecture

### 6.1 Tkinter Structure

The application uses a single `tk.Tk` root window managed by `MainWindow`. All secondary interfaces are `tk.Toplevel` subclasses, ensuring they are children of the root and are destroyed when the root exits.

The main dashboard uses `ttk.Notebook` (tabbed interface) to separate:
- **RFQ tabs:** Active RFQs, Archived RFQs
- **VSM tabs:** Saving, Cost Avoidance, Derisking

Each tab contains a `ttk.Treeview` for the list, sized to fill available space.

Layout within `MainWindow` uses `.grid()` (not `.pack()`) for the main container frames, which allows dynamic show/hide via `grid_remove()` / `grid()` — specifically used by `CollapsibleFilters`.

### 6.2 `ui/main_dashboard_builder.py`

**Design contract:** Pure widget builder. `build_main_dashboard(app)` creates every widget in the main dashboard, assigns them to `app.<attribute>`, configures grid layout, and binds events. It does **not** load data, trigger refreshes, or contain conditional logic beyond layout decisions.

**Widgets built:**
- Top frame: logo, New Event button, Actions dropdown (disabled by default), Export Excel button, KPI button, Settings/License/Help buttons (right-side).
- Grid row 1: `MainDashboardToolbar` (global search bar).
- Grid row 2: `CollapsibleFilters` (advanced filter panel, starts collapsed).
- Grid row 3: `ttk.Notebook` with all tabs.
- Filter sub-frames for RFQ and VSM are inserted into the `CollapsibleFilters.filters_frame`.
- VSM-specific filter frames are created per tab type and stored in `app._vsm_spec_frames`.

### 6.3 `ui/components/`

#### `main_dashboard_toolbar.py` — `MainDashboardToolbar`

A `ttk.Frame` subclass implementing the global search bar.
- Contains a single `tk.Entry` with a placeholder text mechanism (grey italic text that clears on focus, restores on blur when empty).
- Pressing Enter triggers `main_window.dashboard_controller.search_requests()`.
- The search operates across 6 fields (RFQ number, reference, supplier name, article code, description, purchase order) with OR logic per field and AND with any active advanced filters.
- An "Advanced Filters" toggle button on the right calls `CollapsibleFilters.toggle()`.

#### `collapsible_filters.py` — `CollapsibleFilters`

A `ttk.Frame` subclass that wraps the advanced filter panel in a collapsible container.
- Internally creates a `ttk.LabelFrame` (accessible as `.filters_frame`) where the caller inserts filter widgets.
- `expand()` / `collapse()` / `toggle()` control visibility via `grid()` / `grid_remove()`.
- The wrapper itself is always in the grid hierarchy when expanded; `grid_remove()` leaves zero-height gap.
- `set_grid_config(**kwargs)` saves the grid parameters so `expand()` can restore them correctly.

### 6.4 `ui/dialogs/`

All dialogs are `tk.Toplevel` subclasses following a consistent pattern:
1. `self.withdraw()` on creation (invisible)
2. `set_window_icon(self)` — applies the application icon
3. `self.transient(parent)` — modality relationship
4. Build UI content
5. `center_window(self)` — compute geometry and call `self.deiconify()`
6. `self.wait_window()` — block until dialog closes (modal behavior)

#### `common_dialogs.py`

| Class | Description |
|---|---|
| `SimpleMessageDialog` | Modal OK dialog with a text message and info/warning/error indication |
| `SimpleYesNoDialog` | Modal Yes/No dialog; result stored in `self.result` (bool) |
| `LanguagePrompt` | First-run language selection dialog (English / Italiano) |
| `NewRdOTypeDialog` | Dialog to choose RFQ type: "Fornitura piena" or "Conto lavoro" |
| `UserIdentityDialog` | First-run user identity form (first name, last name → auto-generates username) |
| `CopyProgressWindow` | Progress bar dialog for long-running DB copy operations |
| `SplashScreen` | Startup splash screen shown during initialization |
| `LicenseAcceptanceDialog` | License agreement accept/decline dialog |

#### `vsm_event_dialog.py` — `VSMEventDialog`

Form dialog for creating/editing VSM events. Key behaviors:
- `event_type` ComboBox drives dynamic field visibility: fields irrelevant to the selected type (e.g., payment fields for Saving, `importo_richiesto_iniziale` for Cost Avoidance only) are shown/hidden via `grid()` / `grid_remove()`.
- All monetary inputs use comma as decimal separator (Italian convention), validated by `parse_float_from_comma_string()`.
- In read-only mode (other user's event), all inputs are disabled; Save button is hidden.
- On save: constructs a `VSMEvent` dataclass, calls `save_event_with_impacts()` or `update_event_with_impacts()` from `vsm_persistence`.

#### `potential_supplier_dialog.py` — `PotentialSupplierDialog`

Form dialog for the derisking supplier registry. Validates `email` via `is_valid_email()` and `website` via `is_valid_website()`. Status values are displayed as translated labels but stored as canonical English strings. Calls `create_supplier()` or `update_supplier()` on save.

#### `manage_supplier_categories_dialog.py` — `ManageSupplierCategoriesDialog`

In-memory category management dialog: accumulates rename/merge/delete operations in `_pending_ops` and only writes to DB on the "Save" button. Annulling or closing the window discards all pending operations.

### 6.5 `ui/windows/`

Sub-windows opened from `ViewRequestWindow` or directly from `MainWindow`. All are `tk.Toplevel`.

| Window | Trigger | Key behavior |
|---|---|---|
| `ViewRequestWindow` | Double-click RFQ in dashboard | Main RFQ editing interface with tksheet grid for article/price data; opens child windows for attachments, notes, suppliers, POs, SQDC |
| `AttachmentWindow` | "Allegati" / "Offerte" button in ViewRequestWindow | Lists attachments with BLOB/link type; supports open (temp file), add (link or embed), delete; threaded file open for responsiveness |
| `NotesWindow` | "Note" button in ViewRequestWindow | Rich-text editor (tk.Text with bold/italic/underline tags) serializing formatted content as JSON; loads/saves via `DatabaseManager` |
| `PurchaseOrderWindow` | "N° Ordine" button | tksheet grid for managing multiple PO numbers per RFQ; stored as JSON in `numeri_ordine` column |
| `EditSuppliersWindow` | "Modifica Fornitori" button | Single Entry widget for comma-separated supplier list; updates `richiesta_fornitori` table |
| `EditReferenceWindow` | "Modifica Riferimento" button | Single Entry widget for the RFQ reference field |
| `SQDCAnalysisWindow` | "SQDC" button | Tabbed notebook with weight inputs and score entry per supplier; computes weighted SQDC score; exports to pre-formatted Excel template |

### 6.6 `ui/kpi_window.py` — `KpiWindow`

Top-level window with 4 tabs (RFQ, Saving, Cost Avoidance, Derisking).

**Structure per tab:**
1. **KPI Cards row** — Label widgets in a frame, populated from `kpi_engine.get_*_kpi()`.
2. **Chart canvas** — `tk.Canvas` rendered by `ui/kpi_chart.py`.
3. **Details area** (placeholder in current version).

**Filter bar:** Period presets (Last 3/6/12 months), specific year dropdown (`get_available_years()`), and free date range pickers. Changing any filter triggers `_refresh_all()` which reloads all four tabs.

**Export:** `KpiExportScopeDialog` asks whether to export the current tab or all tabs; calls `build_kpi_workbook()` and saves via `filedialog.asksaveasfilename()`.

### 6.7 `ui/kpi_chart.py`

- **Responsibility:** Pure canvas rendering. No data fetching, no Tkinter widget creation outside the `Canvas` passed in.
- **Functions:** `draw_bar_chart(canvas, data, y_fmt, title, y_label, x_label)` and `draw_dual_bar_chart(canvas, data, label1, label2)`.
- **Empty state:** Both functions show a "No data available" message when `data` is empty.
- **Sizing:** All dimensions are derived from the canvas's current pixel size at draw time — the chart is fully responsive.
- **No external graphing library** (matplotlib, etc.) — intentional to keep the dependency footprint minimal.

---

## 7. Business Logic

### 7.1 KPI Engine

**File:** `services/kpi_engine.py`

The engine executes read-only SQL queries directly against the SQLite database. Each of the four public functions (`get_rfq_kpi`, `get_saving_kpi`, `get_cost_avoidance_kpi`, `get_derisking_kpi`) opens a `DatabaseManager` context, runs a set of aggregate queries, and assembles a result dict.

**Date filtering logic (`_build_date_filter`, `_build_impact_period_filter`):**
- If `year` is provided → `strftime('%Y', date_col) = ?` (event date) or `vi.anno = ?` (impact period).
- If `date_from`/`date_to` are provided → range comparison.
- If neither → no date filter (all-time).
- For impact-based KPIs (Saving, Cost Avoidance), filtering is applied on the **impact period** (`vsm_impacts.anno`/`mese`) not on the event date — this correctly attributes economic value to when it materializes, not when the negotiation was recorded.

**Saving KPI calculation (illustrative):**
```sql
SELECT SUM(vi.valore_teorico), SUM(vi.valore_effettivo)
FROM vsm_impacts vi
JOIN vsm_events ve ON vi.event_id = ve.event_id
WHERE ve.event_type = 'Saving'
  AND vi.anno = ?  -- year filter
```

### 7.2 VSM Engine

**File:** `services/vsm_engine.py`

Stateless calculation module. Given a `VSMEvent`, produces a list of `VSMImpact` objects representing how the economic value is distributed over time.

**Core logic (`generate_impacts_for_event`):**

1. **Validate** the event (non-null date, username, valid event_type).
2. **Derisking events** → return empty list (no monetary impact calculated).
3. **Calculate `first_month_coefficient`:** Pro-rata for the first month using commercial 30-day convention: `(30 - day + 1) / 30`. Day 1 → 1.0; Day 16 → 0.5; Day 30 → 1/30.
4. **Calculate distribution months:**
   - `opex_ripetitivo=True` → up to 24 months starting from the event month.
   - `opex_ripetitivo=False` → single month (CAPEX one-shot).
5. **Distribute value:** The annual theoretical value (from `VSMEvent.calculate_theoretical_value()`) is divided by 12 to get a monthly base. The first month is multiplied by `first_month_coefficient`. Effective value applies `percent_realizzo / 100`.
6. **Assign `username` and `event_id`** to each impact (for multi-user isolation).

**Driver = Pagamenti (payment terms):** The theoretical saving is `spending_annuo * (giorni_negoziati - giorni_attuali) / 30 * payments_rate`. `payments_rate` defaults to `get_pagamenti_coefficient()` (0.5%/month = 0.005) if not overridden on the event.

### 7.3 VSM Persistence

**File:** `services/vsm_persistence.py`

**Mandatory pattern: DELETE-REGENERATE-SAVE**

The persistence layer never updates impacts in-place. Any change to a VSM event must:
1. Delete all existing impacts for that `event_id`.
2. Re-generate impacts from the updated event using `vsm_engine`.
3. Save the new impacts in a single transaction.

This guarantees idempotency, prevents ghost impacts, and simplifies debugging.

**Transactions:** `save_event_with_impacts()` and `update_event_with_impacts()` issue an explicit `BEGIN TRANSACTION`, use the `_no_commit` variants of the DB insert methods, and issue a single `COMMIT`. On any exception, `ROLLBACK` is called.

**Public functions:**

| Function | Description |
|---|---|
| `save_event_with_impacts(db, event)` | New event only (`event.id` must be None); saves event, generates impacts, commits atomically |
| `update_event_with_impacts(db, event)` | Existing event; updates event record, deletes old impacts, regenerates, commits |
| `delete_event_and_impacts(db, event_id)` | Deletes event and all its impacts atomically |
| `get_event_with_impacts(db, event_id)` | Returns `(VSMEvent, List[VSMImpact])` for a given event |

---

## 8. Data Layer

### 8.1 Database Engine

**SQLite 3** with **WAL (Write-Ahead Logging)** mode.

WAL is chosen over the default journal mode because it allows:
- Multiple concurrent readers while one writer is active.
- A writer does not block readers.
- Essential for the multi-user scenario where User A opens their DB in read-write mode and User B reads it in read-only mode via `get_all_richieste_aggregated()`.

### 8.2 Database File Location

```
~/Documents/DataFlow_<username>/Database/dataflow_db.db   (Windows default)
~/DataFlow_<username>/Database/dataflow_db.db              (Linux/macOS default)
```

Three side-car files exist alongside the `.db`:
- `dataflow_db.db-wal` — WAL write-ahead log
- `dataflow_db.db-shm` — shared memory file (WAL index)

Manual backup must copy all three files. The built-in backup function in `SettingsWindow` handles this.

### 8.3 Schema Summary

```
fornitori (id_fornitore PK, nome_fornitore UNIQUE)

richieste_offerta (
    id_richiesta PK,          -- year-driven integer
    data_emissione VARCHAR,
    data_scadenza VARCHAR,
    riferimento VARCHAR,
    note_generali VARCHAR,
    stato VARCHAR,            -- 'attiva' | 'archiviata'
    numeri_ordine VARCHAR,    -- JSON array of {supplier, po} objects
    tipo_rdo VARCHAR,         -- 'Fornitura piena' | 'Conto lavoro'
    note_formattate VARCHAR,  -- JSON rich-text structure
    username VARCHAR
)

dettagli_richiesta (
    id_dettaglio PK,
    id_richiesta FK,
    codice_materiale VARCHAR,
    descrizione_materiale VARCHAR,
    quantita VARCHAR,
    disegno VARCHAR,
    data_consegna_richiesta VARCHAR,
    codice_grezzo VARCHAR,
    disegno_grezzo VARCHAR,
    materiale_conto_lavoro VARCHAR
)

richiesta_fornitori (id_richiesta FK, nome_fornitore VARCHAR, PK composite)

offerte_ricevute (
    id_dettaglio FK,
    nome_fornitore VARCHAR,
    prezzo_unitario VARCHAR,   -- TEXT (not REAL) to preserve exact user input
    PK composite
)

allegati_richiesta (
    id_allegato PK,
    id_richiesta FK,
    nome_file VARCHAR,
    dati_file BLOB,            -- NULL for external links
    tipo_allegato VARCHAR,     -- 'Offerta Fornitore' | 'Documento Interno'
    nome_fornitore VARCHAR,
    percorso_esterno VARCHAR,  -- non-NULL for external links
    data_inserimento VARCHAR
)

vsm_events (
    event_id PK,
    username TEXT,
    event_date TEXT,
    buyer TEXT,
    event_type TEXT,           -- 'Saving' | 'Cost Avoidance' | 'Derisking'
    action TEXT,
    description TEXT,
    reference TEXT,
    importo_bdg REAL,
    importo_negoziato REAL,
    importo_richiesto_iniziale REAL,
    quantita_annua REAL,
    percent_realizzo REAL,
    driver TEXT,
    giorni_pagamento_attuali INTEGER,
    giorni_pagamento_negoziati INTEGER,
    spending_annuo REAL,
    opex_ripetitivo INTEGER,   -- 0 | 1 (SQLite boolean)
    note TEXT,
    payments_rate REAL,
    new_supplier TEXT,
    created_at TEXT,
    updated_at TEXT
)

vsm_impacts (
    impact_id PK,
    event_id FK → vsm_events,
    username TEXT,
    anno INTEGER,
    mese INTEGER,              -- 1–12
    tipo_valore TEXT,          -- 'Saving' | 'Cost Avoidance'
    valore_teorico REAL,
    valore_effettivo REAL
)

potential_suppliers (
    supplier_id PK,
    supplier_name TEXT,
    macrocategory TEXT,        -- legacy; superceded by category
    merchandise_class TEXT,
    supplier_status TEXT,      -- 'Nuovo' | 'In valutazione' | 'Qualificato' | 'Scartato'
    contact_name TEXT,
    email TEXT,
    phone TEXT,
    website TEXT,
    notes TEXT,
    username TEXT,
    category TEXT,             -- current canonical category field
    created_at TEXT,
    updated_at TEXT
)

supplier_categories (id PK, name TEXT UNIQUE)
```

### 8.4 Multi-user Aggregation

`DatabaseManager.get_all_richieste_aggregated(db_path)` discovers peer databases by:
1. Taking the directory containing `db_path` (i.e., `Database/`).
2. Going up one level to the `DataFlow_<username>` folder.
3. Going up one more level to the parent (e.g., `Documents/` or a shared drive root).
4. Scanning all `DataFlow_*/Database/*.db` siblings.
5. Building a `UNION ALL` query that reads from all discovered DB files (using `ATTACH DATABASE`).

Read-only connections are used for non-owned databases to respect WAL semantics.

---

## 9. Internationalization

### Structure

```
locale/
├── en/LC_MESSAGES/
│   ├── dataflow.po    # Source: msgid (English) = msgstr (English — passthrough)
│   └── dataflow.mo    # Compiled binary; loaded at runtime by gettext
└── it/LC_MESSAGES/
    ├── dataflow.po    # Source: msgid (English) = msgstr (Italian)
    └── dataflow.mo    # Compiled binary
```

### Runtime Mechanism (`utils/i18n_utils.py`)

1. **`init_i18n(language_code='en')`** is called once at the very start of `dataflow.py`, before any UI module is imported.
2. It reads the user's language preference from `config.ini` (`Settings.language` = `'en'` or `'it'`).
3. Locates the `.mo` file via `resource_path('locale')` (handles both development and PyInstaller frozen modes).
4. Calls `gettext.translation(...).install()`, which installs the translation function in `builtins._`.
5. The `_()` function defined in `i18n_utils.py` is a **dynamic forwarder**: it calls `builtins._()` at call time, not at import time. This solves the stale-binding problem that would occur with `from gettext import _`.

### Locale-specific UI helpers

| Function | Description |
|---|---|
| `get_current_language()` | Returns `'en'` or `'it'` from config |
| `get_pos_column_text()` | Returns "Pos." or "Pos." — language-aware |
| `get_qty_column_text()` | Returns "Q.tà" or "Qty" |
| `normalize_rfq_type(t)` | Normalizes a possibly-translated RFQ type to the canonical Italian form stored in DB |
| `translate_rfq_type(t)` | Translates the canonical Italian form to the current UI language |

### Development Workflow

1. Add/modify `_("string")` calls in source code. English strings serve as `msgid`.
2. Extract strings to `dataflow.po` using `xgettext` or equivalent.
3. Translate in `locale/it/LC_MESSAGES/dataflow.po`.
4. Run `python dev_tools/compile_translations.py` to compile `.po` → `.mo`.
5. Test by switching language in Settings and restarting.

---

## 10. Build & Distribution

### `dataflow.spec` (root — Linux/development)

PyInstaller spec file for building a one-folder distribution. Key configuration:
- **Entry point:** `dataflow.py`
- **`datas`:** Bundles `add_data/` (assets) and `locale/` (translations).
- **`hiddenimports`:** Explicit list of modules that PyInstaller's static analysis might miss (e.g., `tkinter.ttk`, `tkcalendar`, `tksheet`).
- **Mode:** One-folder (not one-file) for better Tcl/Tk compatibility (`init.tcl` path resolution).
- **Platform note:** Linux Tkinter has no special Tcl/Tk detection issues; the spec is simpler than the Windows version.

### `dev_tools/tools_build_WIN/exe/DataFlow.spec`

Windows-specific spec. Additional steps:
- **Automatic Tcl/Tk detection:** Uses `tkinter.__file__` to find the Tcl/Tk library directory and adds `tcl8.6/` and `tk8.6/` to `datas`. This prevents the `Can't find a usable init.tcl` error on Windows.
- **Manifest embedding:** References `app.manifest.xml` (DPI-awareness) via `EXE(manifest=...)`.
- **`console=False`:** Produces a Windows GUI executable with no console window.

### `dev_tools/tools_build_WIN/msix/AppxManifest.xml`

MSIX package declaration for Windows Store / enterprise distribution. Declares:
- Package identity (name, version, publisher)
- Application entry point
- Visual assets paths
- Required capabilities

### Build commands (assumed standard — not verified from source)

```bash
# Install dependencies
pip install -r requirements.txt pyinstaller

# Compile translations
python dev_tools/compile_translations.py

# Build (Linux/dev)
pyinstaller dataflow.spec

# Build (Windows)
pyinstaller dev_tools/tools_build_WIN/exe/DataFlow.spec
```

---

## 11. Testing Strategy

### Test Location

All tests are in `tests/`. No test runner configuration file (e.g., `pytest.ini`) is present in the repository; tests are run with `unittest` or `pytest` in discovery mode.

### Test Modules

#### `test_vsm_engine.py`

- **Class:** `TestVSMEngineHelpers`, `TestVSMEngineIntegration`
- **Scope:** Pure unit tests on `vsm_engine.py`; no DB access.
- **Cases covered:**
  - `_calculate_first_month_coefficient()` for day 1, 16, 30 (boundary and mid-month).
  - `_calculate_distribution_months()` for repetitive (24 months) vs one-shot (1 month) events.
  - `generate_impacts_for_event()`: correct month count, pro-rata correctness, mathematical conservation of total value.
  - **Derisking events** → empty impact list.
  - Propagation of `username` and `event_id` to all generated impacts.
  - Chronological ordering of impacts.
  - `VSMError` raised on missing `event_date`, missing `username`, unsupported `event_type`.

#### `test_vsm_persistence.py`

- **Class:** `TestVSMPersistence`
- **Scope:** Integration tests against a temporary SQLite file (not in-memory, to allow multi-connection scenarios).
- **Setup:** Creates a temp `.db`, calls `create_tables()`, inserts a `test_user` into the `utenti` table (FK requirement).
- **Cases covered:**
  - `save_event_with_impacts()`: new event persisted with correct impacts count.
  - `update_event_with_impacts()`: DELETE-REGENERATE-SAVE pattern verified (no duplicate impacts).
  - `delete_event_and_impacts()`: both event and impacts removed.
  - `get_event_with_impacts()`: round-trip fidelity.
  - `VSMError` raised if `save` receives an event with an existing `id`.

#### `test_vsm_event_model.py`

- **Class:** `TestVSMEventCalculations`
- **Scope:** Unit tests on `VSMEvent.calculate_theoretical_value()` and `calculate_effective_value()`.
- **Cases covered:** Saving + Prezzo at qty=1 (baseline), qty=20000, partial realization (80%), driver=Pagamenti, Cost Avoidance.

#### `test_supplier_category_persistence.py`

- **Class:** `TestSupplierCategoryPersistence`
- **Scope:** Integration tests against a temporary SQLite file.
- **Cases covered (per spec):**
  1. Migration: categories from `potential_suppliers` imported into `supplier_categories`.
  2. `ensure_supplier_category_exists`: idempotency, trim behavior.
  3. Rename: all suppliers updated, old category removed, new present.
  4. Rename towards existing category → `CategoryError`.
  5. Merge: suppliers moved, source deleted.
  6. Delete unused category → allowed.
  7. Delete used category → `CategoryError`.
  8. Trim: `" Tornerie "` → stored as `"Tornerie"`, no duplicate.

### `tests/db_test_tool/`

Development utility scripts (not automated tests):
- `generate_test_db_eng.py` — creates a populated test database with English content.
- `generate_test_db_it.py` — creates a populated test database with Italian content.
- Used for manual UI testing and demo purposes.

### Test Coverage Gaps (Assumption)

Based on the visible test files, the following areas have no automated tests:
- `kpi_engine.py` (complex SQL aggregations)
- `kpi_chart_data.py` (bucket generation logic)
- `services/app_paths.py` (path resolution requiring file system mocking)
- All UI code (Tkinter UIs are not unit-tested)
- `database_manager.py` CRUD methods directly

---

## 12. Assets & Resources

### `add_data/` — Icon and logos

| File | Size hint | Usage |
|---|---|---|
| `DataFlow.ico` | Multi-res ICO | `set_window_icon()` — every `Toplevel` window |
| `logo_dataflow.png` | Landscape | Main dashboard toolbar (left of button row) |
| `Logo.png` | Standard | General-purpose |
| `Logo_150x150.png` | 150×150 | Future/contextual use |
| `Logo_44x44.png` | 44×44 | MSIX tile assets |
| `Logo_50x50.png` | 50×50 | MSIX tile assets |

`resource_path(relative)` (from `utils/resource_utils.py`) resolves paths correctly both in source mode (`base = project_root/`) and in PyInstaller frozen mode (`base = sys._MEIPASS`).

### `add_data/` — Excel Templates

Templates are pre-formatted `.xlsx` files that the export logic opens, populates, and saves. The application selects the correct template at export time based on language and RFQ type:

| Template | Language | RFQ Type |
|---|---|---|
| `template_rdo.xlsx` | Italian | Fornitura piena |
| `template_rdo_cl.xlsx` | Italian | Conto lavoro |
| `template_rdo_eng.xlsx` | English | Full supply |
| `template_rdo_eng_cl.xlsx` | English | Work order |
| `template_sqdc.xlsx` | Italian | SQDC analysis |
| `template_sqdc_eng.xlsx` | English | SQDC analysis |

Conto Lavoro (work-order) templates include extra columns: `Cod.Grezzo` (raw part code), `Dis.Grezzo` (raw part drawing), `Mat.C/L` (subcontract material).

---

## 13. Utilities Layer

All utility modules are stateless (no module-level mutable state except `vsm_config.py` which reads from a file). They expose pure functions and raise standard Python exceptions or domain-specific ones where noted.

### `utils/format_utils.py`

**Purpose:** Number and currency formatting following Italian locale conventions (comma as decimal separator, period as thousands separator).

| Function | Input | Output | Notes |
|---|---|---|---|
| `parse_float_from_comma_string(s)` | String with comma decimal | `float` | Raises `ValueError` on point decimal or multiple commas; handles None, int, float |
| `format_quantity_display(val)` | float/str | `str` | Removes trailing decimals for whole numbers; uses comma decimal |
| `format_amount_display(val)` | float/str | `str` | Always 2 decimal places, comma; used for pre-filling amount fields |
| `format_currency_display(val)` | float/str | `str` | `"€ 2.000,00"` format |

### `utils/i18n_utils.py`

See [Section 9](#9-internationalization).

### `utils/resource_utils.py`

| Function | Description |
|---|---|
| `resource_path(relative_path)` | Returns absolute path; uses `sys._MEIPASS` in frozen mode, `project_root/` otherwise. The double `dirname()` accounts for this module living in `utils/`. |
| `set_window_icon(window)` | On Windows: `window.iconbitmap(path)`; on Linux/macOS: opens with Pillow, sets via `window.iconphoto(True, photo)`; keeps a reference on `window._icon_photo` to prevent GC. |

### `utils/string_utils.py`

| Function | Description |
|---|---|
| `generate_username(first_name, last_name)` | `first[0].lower() + last.lower()` after stripping accents via NFKD normalization and removing non-alphanumeric characters. Example: `"Guido", "Sorrentino"` → `"gsorrentino"`. |

### `utils/user_utils.py`

| Function | Description |
|---|---|
| `get_app_data_dir()` | Platform-aware application data directory. Frozen: `AppData/Local/DataFlow` (Win) or `~/.local/share/DataFlow` (Linux). Dev: same `~/.local/share/DataFlow` path. |
| `get_config_file()` | `get_app_data_dir() / config.ini`; creates the directory if it doesn't exist. |
| `load_user_identity()` | Reads `[User]` section from `config.ini`; returns dict with `first_name`, `last_name`, `username`, `full_name`. |
| `save_user_identity(first, last, username)` | Writes or overwrites `[User]` section in `config.ini` with UTF-8 encoding. |

### `utils/validation_utils.py`

| Function | Description |
|---|---|
| `sanitize_filename(name)` | Removes Windows/Unix forbidden characters (`\/*?:"<>|`) via regex. |
| `format_date_for_db(display_date)` | `dd/mm/yyyy` → `YYYY-MM-DD`; returns `None` on failure. |
| `format_price_display(num)` | `float`/`str` → `"123,4500"` (4 decimal places, comma). |
| `is_valid_email(value)` | Regex validation; returns `True` for empty (optional field). |
| `is_valid_website(value)` | Accepts `http://`, `https://`, `www.`, plain domain patterns. Returns `True` for empty. |

### `utils/window_utils.py`

| Function | Description |
|---|---|
| `calculate_center_position(win)` | Returns `"{w}x{h}+{x}+{y}"` geometry string; accounts for taskbar buffer. |
| `calculate_optimal_window_size(win, num_suppliers, is_conto_lavoro)` | Computes `ViewRequestWindow` width from article columns + supplier columns + padding, capped at 95% screen width, minimum `BUTTONS_MIN_WIDTH`. |
| `center_window(win)` | Calls `calculate_center_position`, applies geometry, calls `deiconify()`. |

### `utils/vsm_config.py`

| Function | Description |
|---|---|
| `get_pagamenti_coefficient()` | Reads `Settings.vsm_pagamenti_coefficient` from `config.ini`; initializes with default `0.005` if missing. Default = 0.5%/month cost of capital. |
| `set_pagamenti_coefficient(value)` | Writes the coefficient to `config.ini`. Returns `bool` success. |

The coefficient represents the monthly opportunity cost of capital for the "Pagamenti" VSM driver: `saving = spending * delta_days / 30 * coefficient`.

---

## 14. Execution Flow

### Step 1 — OS Process Launch

The OS executes `python dataflow.py` (or the frozen EXE).

**Line 1–27 (DPI):** Before any other import, on Windows, `ctypes.windll.shcore.SetProcessDpiAwarenessContext(-4)` (`PER_MONITOR_V2`) is called. This must happen before Tkinter initializes.

### Step 2 — Module Initialization

All `import` statements execute at the module level:
- Standard library modules load.
- Third-party packages load (`tkinter`, `tksheet`, `tkcalendar`, `openpyxl`, `PIL`).
- `constants.py` loads layout constants.
- `utils/` modules load (stateless, no side effects except logger creation).
- `init_i18n()` is called — reads `config.ini`, loads the appropriate `.mo` file, installs `_()` in `builtins`.
- UI modules load (they call `_()` at import time for default string values; `_()` must already be available).
- `cleanup_temp_on_startup()` runs — deletes stale PyInstaller temp folders.
- `setup_logging()` runs — creates the rotating file logger.

### Step 3 — Application Startup

`if __name__ == '__main__':` block:

1. `root = tk.Tk()` — creates the Tcl/Tk interpreter and root window.
2. Root window is hidden (`root.withdraw()`).
3. `root.title("DataFlow")` and icon are set.
4. **First-run check:** `load_user_identity()` is called. If no username exists, `UserIdentityDialog` is shown. The user enters name/surname; `generate_username()` produces the username; `save_user_identity()` writes to `config.ini`.
5. **Language check:** If no language is set, `LanguagePrompt` is shown. Chosen language is persisted in `config.ini`.
6. **License check:** If license has not been accepted, `LicenseAcceptanceDialog` is shown. Declining exits the app.
7. **Directory structure:** `initialize_dataflow_directory_structure()` ensures `Database/` and `Attachments/` subdirectories exist under the user's DataFlow folder.
8. **Database initialization:** `crea_database_v4()` calls `DatabaseManager.create_tables()` — creates all tables and runs column migrations if upgrading from an older version.
9. **Splash screen:** `SplashScreen` is shown briefly during remaining initialization.

### Step 4 — UI Construction

`MainWindow(root)` is instantiated:
1. `build_main_dashboard(app)` runs — creates all Tkinter widgets.
2. `DashboardController(self)` is created.
3. Initial data load: `dashboard_controller.refresh_data()` is called, which queries the database and populates all Treeview widgets.
4. `populate_username_filter()` fills the user dropdown.
5. Root window is shown (`root.deiconify()`).

### Step 5 — Event Loop

`root.mainloop()` starts the Tkinter event loop. The application is now event-driven.

### Step 6 — User Interactions

**Opening an RFQ:**
1. User double-clicks a row in `tree_attive` or `tree_archiviate`.
2. `MainWindow._on_row_double_click()` extracts the `request_id` from the row data.
3. `ViewRequestWindow(root, request_id)` is instantiated as a `Toplevel`.
4. `ViewRequestWindow.__init__` determines the correct DB path (local or remote for other-user RFQs).
5. `DatabaseManager` reads the RFQ header, line items, supplier list, and prices.
6. The `tksheet.Sheet` widget is populated with article/price data.

**Creating a VSM Event:**
1. User clicks "➕ Nuovo Evento" while on a VSM tab.
2. `MainWindow.open_new_event()` determines the event type from the active tab.
3. `VSMEventDialog(root, current_username, event_type)` opens.
4. User fills the form; on save, a `VSMEvent` dataclass is constructed.
5. `save_event_with_impacts(db, event)` (from `vsm_persistence`) runs a transaction: INSERT event, generate impacts, INSERT impacts.
6. Dialog closes; `dashboard_controller.refresh_data()` reloads the VSM tree.

### Step 7 — Data Persistence

All writes go through `DatabaseManager` methods. The canonical pattern:
```python
with DatabaseManager(get_db_path()) as db:
    db.some_insert_or_update(...)
# connection is committed and closed by __exit__
```

For VSM operations requiring atomicity, the `_no_commit` variants are used within an explicit `BEGIN TRANSACTION` / `COMMIT` block managed by `vsm_persistence.py`.

### Step 8 — Export

**RFQ Excel Export:**
1. User triggers export in `ViewRequestWindow`.
2. Template selected based on `tipo_rdo` and `get_current_language()`.
3. `openpyxl.load_workbook(resource_path(template))` opens the template.
4. Data from the database is written into the workbook cells.
5. `openpyxl.Workbook.save(temp_path)` writes to a temp file; `os.startfile()` (Windows) or `subprocess.Popen(['xdg-open', ...])` (Linux) opens the file.

**KPI Excel Export:**
1. User clicks Export in `KpiWindow`.
2. `KpiExportScopeDialog` asks scope (current tab or all).
3. `kpi_engine.get_*_kpi(...)` is called for the relevant tabs.
4. `build_kpi_workbook(...)` builds an `openpyxl.Workbook` from scratch.
5. `filedialog.asksaveasfilename()` prompts for save location.
6. `workbook.save(path)` writes the file.

---

## 15. Design Principles

### Modularity

The codebase has been progressively refactored (release 2.1.0 explicitly marks multiple extractions) to move logic out of the monolithic `dataflow.py` into focused modules. The refactoring comments (`# REFACTORING: ...`) in `dataflow.py` document what was moved and where.

Remaining in `dataflow.py` (as of 2.1.0): `MainWindow`, `SettingsWindow`, and some legacy inline methods that have not yet been extracted. New features are to be implemented in dedicated service/ui modules, not added to `dataflow.py`.

### Separation of Responsibilities

The architecture enforces strict boundaries:
- **Services never import `tkinter`** — they are testable in isolation.
- **UI modules never import `sqlite3` directly** — all DB access goes through `DatabaseManager`.
- **Models contain no I/O** — they are pure data + field-level calculations.
- **`kpi_engine`** is read-only: it never writes to the database.
- **`vsm_engine`** is stateless: it receives a `VSMEvent` and returns a list of `VSMImpact`; no DB access.

### Context Manager Pattern

`DatabaseManager` implements `__enter__`/`__exit__`, and all non-VSM DB access uses the `with` statement to guarantee connection closure even on exceptions. This prevents connection leaks, which are particularly harmful in a WAL-mode SQLite setup with peer readers.

### First-class Multi-user Support

The `username` field is propagated to every business entity (`richieste_offerta.username`, `vsm_events.username`, `vsm_impacts.username`, `potential_suppliers.username`). The dashboard can aggregate data across multiple users' databases on a shared drive with minimal configuration.

### Read-only Mode

Any window that opens another user's RFQ (via `source_db_path`) automatically sets `read_only=True`, which disables all mutation operations in the UI and opens the database in URI read-only mode — preventing accidental writes to another user's data.

### Defensive Error Handling

- Business functions raise typed exceptions (`VSMError`, `SupplierError`, `CategoryError`) that carry human-readable messages.
- KPI calculations never raise to the caller — they return zero-value dicts on any DB error, ensuring the KPI window always renders.
- UI event handlers wrap DB calls in try/except and display `SimpleMessageDialog` on failure rather than crashing.
- All `ALTER TABLE` migration calls are wrapped in individual try/except (not a single rollback block) so that a missing column migration does not abort table creation for other columns.

### DPI Awareness

Three-layer approach for Windows high-DPI:
1. `app.manifest.xml` (static, embedded in EXE) declares `PerMonitorV2`.
2. `ctypes` call at `dataflow.py` line 11 sets DPI awareness programmatically as early as possible.
3. `window_utils.py` geometry calculations incorporate `PADDING = 140 px` as a DPI safety margin.

### Internationalization Architecture

`_(text)` is a dynamic forwarder to `builtins._`, not a static binding. This design allows modules imported before `init_i18n()` is called to still have `_()` available (they import the forwarder), and ensures all translations are resolved against the installed catalog at call time rather than at import time.

---

*Last updated: April 2026 — DataFlow v2.1.0*
