# DataFlow 2.3.0 — Technical Manual

> **Audience:** Software developers who need to understand, maintain, or extend the DataFlow codebase.  
> **Scope:** Full architecture description, file-by-file analysis, data flows, and design rationale.  
> **Version:** 2.3.0 (current codebase as of May 2026)

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

DataFlow is a desktop procurement management application for purchasing teams. It centralizes the operational RFQ workflow and extends it with strategic purchasing modules for Value Stream Mapping (VSM), Derisking, KPI analysis, Excel export, and RFQ PDF generation.

### Typology

- **Desktop application** — Python 3 with Tkinter/ttk UI
- **Local-first** — SQLite database per user, stored in a dedicated DataFlow folder
- **Multi-user capable** — each user works on a local database, while dashboard aggregation reads sibling databases in read-only mode
- **Cross-platform in source form** — Linux/macOS/Windows can run the code; packaging is primarily oriented to Windows via PyInstaller

### Primary Features

| Feature | Description |
|---|---|
| **RFQ Management** | Create, edit, archive, reactivate, duplicate, and delete RFQs |
| **Supplier Comparison** | Spreadsheet-like RFQ detail grid with supplier price comparison |
| **Attachments** | Supplier offers and internal documents stored as BLOBs or external file paths |
| **Rich Notes** | Formatted notes per RFQ, persisted in the database |
| **Purchase Orders** | Purchase order list stored on the RFQ header |
| **SQDC Analysis** | Safety/Quality/Delivery/Cost scoring with Excel export and attachment integration |
| **VSM — Saving** | Negotiation savings tracking with monthly impact generation |
| **VSM — Cost Avoidance** | Avoided-cost tracking with impact distribution by competence period |
| **VSM — Derisking** | Potential supplier registry with category/status/contact management |
| **KPI Dashboard** | RFQ, Saving, Cost Avoidance, and Derisking KPI cards, charts, and Excel export |
| **Global Search & Advanced Filters** | Context-aware search/filtering for RFQ, VSM, and Derisking tabs |
| **Excel Export** | RFQ, VSM, Derisking, and KPI export flows |
| **PDF Export** | RFQ PDF generation with optional persistent company logo and editable text templates |
| **Settings & Maintenance** | Language, currency, manual backup, daily auto-backup, DataFlow folder migration |
| **Runtime i18n** | English/Italian catalogs with widespread `tr(...)` usage across UI and services |

---

## 2. Repository Structure

```text
dataflow-procurement-software/
├── .github/
│   └── workflows/
│       └── build-windows.yml                  # Windows CI packaging workflow
├── app.manifest.xml                           # Windows DPI-awareness manifest
├── constants.py                               # Shared UI geometry/layout constants
├── database_manager.py                        # Central SQLite access layer
├── dataflow.py                                # Application entry point, MainWindow, SettingsWindow
├── dataflow.spec                              # Main PyInstaller spec (Windows one-folder build)
├── dataflow_appimage.spec                     # Root PyInstaller spec for Linux/AppImage-oriented build
├── LICENSE
├── README.md
├── requirements.txt                           # Runtime and packaging dependencies
│
├── add_data/
│   ├── DataFlow.ico
│   ├── Logo.png
│   ├── Logo_150x150.png
│   ├── Logo_44x44.png
│   ├── Logo_50x50.png
│   ├── logo_dataflow.png
│   ├── template_rdo.xlsx
│   ├── template_rdo_cl.xlsx
│   ├── template_rdo_eng.xlsx
│   ├── template_rdo_eng_cl.xlsx
│   ├── template_sqdc.xlsx
│   └── template_sqdc_eng.xlsx
│
├── database/
│   ├── __init__.py
│   └── db_helpers.py                          # First-run schema initialization wrapper
│
├── development/
│   └── dev_tools/
│       ├── compile_translations.py            # .po -> .mo compiler
│       └── tools_build_WIN/
│           ├── AppImage/
│           │   └── dataflow_appimage.spec     # Alternate AppImage-oriented spec
│           ├── exe/
│           │   ├── DataFlow.spec              # Windows EXE-oriented spec variant
│           │   └── app.manifest.xml
│           └── msix/
│               └── AppxManifest.xml           # MSIX package manifest
│
├── locale/
│   ├── en/LC_MESSAGES/
│   │   ├── dataflow.mo
│   │   └── dataflow.po
│   └── it/LC_MESSAGES/
│       ├── dataflow.mo
│       └── dataflow.po
│
├── models/
│   ├── __init__.py
│   ├── potential_supplier.py
│   ├── vsm_event.py
│   └── vsm_impact.py
│
├── services/
│   ├── __init__.py
│   ├── app_paths.py
│   ├── dashboard_actions_policy.py
│   ├── dashboard_controller.py
│   ├── dashboard_search_service.py
│   ├── dashboard_selection_policy.py
│   ├── dataflow_location_service.py
│   ├── derisking_dashboard_service.py
│   ├── excel_export_service.py
│   ├── kpi_chart_data.py
│   ├── kpi_engine.py
│   ├── kpi_excel_export.py
│   ├── restart_lifecycle_service.py
│   ├── rfq_command_service.py
│   ├── rfq_dashboard_service.py
│   ├── rfq_pdf_export_service.py
│   ├── rfq_pdf_logo_service.py
│   ├── rfq_pdf_template_service.py
│   ├── settings_maintenance_service.py
│   ├── settings_preferences_service.py
│   ├── startup_service.py
│   ├── supplier_category_persistence.py
│   ├── supplier_name_suggestion_service.py
│   ├── supplier_persistence.py
│   ├── vsm_command_service.py
│   ├── vsm_dashboard_service.py
│   ├── vsm_engine.py
│   └── vsm_persistence.py
│
├── tests/
│   ├── __init__.py
│   └── db_test_tool/
│       ├── generate_test_db_eng.py
│       ├── generate_test_db_it.py
│       ├── test_dataflow_full.db
│       └── test_dataflow_full_it.db
│
├── ui/
│   ├── kpi_chart.py
│   ├── kpi_window.py
│   ├── main_dashboard_builder.py
│   ├── sheet_factories.py
│   ├── window_launchers.py
│   ├── components/
│   │   ├── __init__.py
│   │   ├── collapsible_filters.py
│   │   ├── main_dashboard_toolbar.py
│   │   └── supplier_name_suggest.py
│   ├── dialogs/
│   │   ├── __init__.py
│   │   ├── common_dialogs.py
│   │   ├── manage_supplier_categories_dialog.py
│   │   ├── potential_supplier_dialog.py
│   │   ├── rfq_pdf_export_dialog.py
│   │   └── vsm_event_dialog.py
│   └── windows/
│       ├── __init__.py
│       ├── attachment_window.py
│       ├── edit_reference_window.py
│       ├── edit_suppliers_window.py
│       ├── notes_window.py
│       ├── purchase_order_window.py
│       ├── sqdc_analysis_window.py
│       └── view_request_window.py
│
└── utils/
    ├── __init__.py
    ├── export_filename.py
    ├── format_utils.py
    ├── i18n_utils.py
    ├── resource_utils.py
    ├── string_utils.py
    ├── supplier_name_normalization.py
    ├── user_utils.py
    ├── validation_utils.py
    ├── vsm_config.py
    └── window_utils.py
```

**Tree filtering note:** the tree above intentionally excludes `__pycache__/`, compiled cache artifacts, temporary files, raw documentation folders, local build outputs, and other non-application repository content. It represents the logical application structure plus relevant build/test support.

---

## 3. Architecture Overview

### Layers

```text
┌────────────────────────────────────────────────────────────┐
│ UI Layer                                                   │
│ dataflow.py                                                │
│ ui/main_dashboard_builder.py                               │
│ ui/kpi_window.py                                           │
│ ui/components/  ui/dialogs/  ui/windows/  ui/sheet_factories.py │
└───────────────────────────┬────────────────────────────────┘
                            │ calls / orchestrates
┌───────────────────────────▼────────────────────────────────┐
│ Services Layer                                             │
│ dashboard_*  rfq_*  vsm_*  kpi_*  settings_*               │
│ supplier_*  startup_service  app_paths  restart_lifecycle  │
└───────────────────────────┬────────────────────────────────┘
                            │ uses
┌───────────────────────────▼────────────────────────────────┐
│ Models Layer                                               │
│ VSMEvent  VSMImpact  PotentialSupplier                     │
└───────────────────────────┬────────────────────────────────┘
                            │ persisted through
┌───────────────────────────▼────────────────────────────────┐
│ Data Layer                                                 │
│ database_manager.py  database/db_helpers.py  SQLite files  │
└────────────────────────────────────────────────────────────┘
               cross-cutting: utils/, locale/, add_data/
```

### Layer Responsibilities

| Layer | Responsibility | Must not do |
|---|---|---|
| **UI** | Build widgets, collect input, render data, open dialogs/windows | Encode core business rules or schema migrations |
| **Services** | Encapsulate orchestration, exports, search/filtering, KPI logic, path/settings/backup flows | Build Tkinter widgets directly |
| **Models** | Represent domain entities and local calculations | Execute SQL or depend on UI state |
| **Data** | Manage connections, SQL, schema evolution, aggregation | Own UI logic |
| **Utils** | Stateless formatting, validation, resource/path helpers | Persist domain state directly |

### General Data Flow

1. `dataflow.py` initializes i18n, startup services, configuration, and the root window.
2. `MainWindow` delegates dashboard widget construction to `ui/main_dashboard_builder.py`.
3. User actions are routed to `MainWindow`, which either calls an extracted service or opens a specialized UI dialog/window.
4. Services access the SQLite database through `DatabaseManager`.
5. UI sheets and dialogs are refreshed using plain Python structures: rows, dicts, dataclasses, and metadata lists.
6. Exports and backups are produced by dedicated services rather than by inlined UI code.

### Architectural Direction in 2.3.0

The codebase still uses `dataflow.py` as the main orchestration host, but several operational areas have been extracted into services:

- Dashboard search, selection, and actions policy
- RFQ/VSM/Derisking dashboard data pipelines
- Excel export
- RFQ PDF export, logo persistence, and template management
- Settings persistence, backup maintenance, folder migration, and restart lifecycle

The result is a hybrid architecture: legacy orchestration remains in `MainWindow`, while new or recently isolated logic lives in smaller service modules.

---

## 4. Directory-by-Directory Deep Dive

### 4.1 Root-level files

#### `dataflow.py`

See [Section 5.1](#51-dataflowpy) for full analysis.

#### `database_manager.py`

See [Section 5.2](#52-database_managerpy) for full analysis.

#### `constants.py`

Shared UI geometry constants used by window sizing logic and RFQ detail layout. The file remains a pure declaration module with no runtime side effects.

#### `requirements.txt`

| Package | Version | Purpose |
|---|---|---|
| `openpyxl` | `3.1.5` | Excel import/export |
| `Pillow` | `12.1.1` | Images, icons, logo handling |
| `polib` | `1.2.0` | Translation catalog compilation |
| `reportlab` | `4.2.2` | RFQ PDF export |
| `tkcalendar` | `1.6.1` | `DateEntry` widgets |
| `tksheet` | `7.6.0` | Spreadsheet-like grids |
| `tkinterdnd2` | `0.4.2` | Optional drag-and-drop support |

#### `app.manifest.xml`

Windows manifest embedded by the main EXE build. It enables PerMonitorV2 DPI awareness for correct scaling on high-DPI displays.

#### `dataflow.spec` / `dataflow_appimage.spec`

Primary packaging specs. The root Windows spec collects `babel`, `reportlab`, `add_data/`, and `locale/`. The root AppImage-oriented spec also collects `reportlab` and adds `PIL._tkinter_finder` to hidden imports.

---

### 4.2 `add_data/`

Bundled static assets used at runtime through `resource_path()`:

- Application icon and logo variants
- RFQ Excel templates in Italian and English
- SQDC Excel templates in Italian and English

These assets are repository-managed. By contrast, user-specific RFQ PDF assets are created outside the repository under the DataFlow user folder.

---

### 4.3 `database/`

#### `db_helpers.py`

This module exposes `crea_database_v4()`, the startup entry used to ensure schema availability. It delegates real work to `DatabaseManager.create_tables()` and does not implement business logic on its own.

---

### 4.4 `development/dev_tools/`

Development-only support utilities relevant to packaging and localization:

| File | Role |
|---|---|
| `compile_translations.py` | Compiles `locale/*/LC_MESSAGES/dataflow.po` into `.mo` catalogs |
| `tools_build_WIN/exe/DataFlow.spec` | Alternate Windows EXE-oriented PyInstaller spec |
| `tools_build_WIN/exe/app.manifest.xml` | Windows manifest copy for the EXE path |
| `tools_build_WIN/AppImage/dataflow_appimage.spec` | Alternate AppImage-oriented PyInstaller spec |
| `tools_build_WIN/msix/AppxManifest.xml` | MSIX package manifest |

The build variants are part of the repository, but the GitHub Actions workflow currently invokes the root `dataflow.spec`.

---

### 4.5 `locale/`

See [Section 9](#9-internationalization) for full analysis.

---

### 4.6 `models/`

The `models/` package contains dataclasses used by the service and persistence layers.

#### `vsm_event.py`

`VSMEvent` models Saving, Cost Avoidance, and Derisking events. Key characteristics:

- supports both price-driven and payment-terms-driven logic
- stores optional `payments_rate`
- stores `new_supplier` for Derisking events
- computes theoretical and effective values directly on the model

#### `vsm_impact.py`

`VSMImpact` represents monthly economic impacts generated from a `VSMEvent`. It enforces month validity and provides lightweight conversion helpers.

#### `potential_supplier.py`

`PotentialSupplier` is the Derisking registry model. The canonical statuses persisted in the current codebase are Italian domain values:

- `Nuovo`
- `In valutazione`
- `Qualificato`
- `Scartato`

---

### 4.7 `services/`

The `services/` package is the main extraction area for operational logic. Responsibility is now split into focused modules rather than concentrated inside `dataflow.py`.

#### Path, startup, settings, and lifecycle

| File | Responsibility |
|---|---|
| `app_paths.py` | Resolve DataFlow base folder, `Database/`, `Attachments/`, and session-cached DB path |
| `startup_service.py` | Temp cleanup, rotating log setup, initial directory structure creation |
| `settings_preferences_service.py` | Persist language, currency, and auto-backup preferences |
| `settings_maintenance_service.py` | Manual backup copy and daily auto-backup retention logic |
| `dataflow_location_service.py` | Validate destination path, writability, and username conflicts during folder migration |
| `restart_lifecycle_service.py` | Resolve script/executable path and launch detached restart |

#### Dashboard orchestration

| File | Responsibility |
|---|---|
| `dashboard_controller.py` | Main dashboard refresh/search/filter orchestration |
| `dashboard_search_service.py` | Closed-domain normalization and in-memory filtering helpers |
| `dashboard_selection_policy.py` | Row selection extraction and ownership checks |
| `dashboard_actions_policy.py` | Declarative enablement rules for the Actions menu |
| `rfq_dashboard_service.py` | RFQ dataset loading and sheet payload construction |
| `vsm_dashboard_service.py` | VSM dataset loading and advanced filter application |
| `derisking_dashboard_service.py` | Derisking dataset loading, sheet row metadata, and auto-sizing |

#### RFQ, export, and PDF services

| File | Responsibility |
|---|---|
| `rfq_command_service.py` | RFQ archive/reactivate/delete/duplicate/create shell operations |
| `excel_export_service.py` | RFQ/VSM/Derisking Excel export flows |
| `rfq_pdf_export_service.py` | ReportLab-based RFQ PDF generation |
| `rfq_pdf_logo_service.py` | Persistent logo validation/storage for PDF export |
| `rfq_pdf_template_service.py` | Editable language-specific text templates for PDF export |

#### VSM, KPI, and supplier services

| File | Responsibility |
|---|---|
| `vsm_engine.py` | Monthly impact generation from `VSMEvent` |
| `vsm_persistence.py` | Transactional save/update/delete for VSM events and impacts |
| `vsm_command_service.py` | Operational delete/duplicate helpers for VSM and Derisking |
| `kpi_engine.py` | Read-only KPI calculation engine |
| `kpi_chart_data.py` | Deterministic chart bucket generation |
| `kpi_excel_export.py` | Workbook builder for KPI export |
| `supplier_persistence.py` | CRUD for `potential_suppliers` |
| `supplier_category_persistence.py` | CRUD-like operations for the category catalog |
| `supplier_name_suggestion_service.py` | In-memory suggestion index derived from RFQ and Derisking supplier names |

---

### 4.8 `tests/`

The current repository no longer includes application-level automated test modules. The remaining `tests/` content is a database-generation toolset used to produce rich English/Italian sample databases for manual validation and non-production checks.

See [Section 11](#11-testing-strategy) for full analysis.

---

### 4.9 `ui/`

The UI package is organized into builders, reusable components, dialogs, and operational windows.

| Area | Role |
|---|---|
| `main_dashboard_builder.py` | Creates the dashboard shell and shared widgets |
| `sheet_factories.py` | Builds configured `tksheet` instances for RFQ, VSM, and Derisking |
| `components/` | Toolbar, collapsible filters, supplier suggestion controller |
| `dialogs/` | Modal dialogs and standardized prompts |
| `windows/` | RFQ operational windows and editors |
| `kpi_window.py` / `kpi_chart.py` | KPI UI and custom chart rendering |
| `window_launchers.py` | External help/license URLs and KPI window launch |

See [Section 6](#6-ui-architecture) for full analysis.

---

### 4.10 `utils/`

The `utils/` package contains low-level stateless helpers for formatting, translation, validation, resource resolution, username generation, export filenames, and window sizing.

See [Section 13](#13-utilities-layer) for full analysis.

---

## 5. Core Entry Points

### 5.1 `dataflow.py`

**Role:** main application entry point and host of the two largest orchestration classes: `SettingsWindow` and `MainWindow`.

**Startup responsibilities:**

1. Enable DPI-aware behavior on supported Windows environments.
2. Initialize i18n before importing UI modules that call `tr(...)` at import time.
3. Run startup maintenance helpers:
   - `cleanup_temp_on_startup()`
   - `setup_logging()`
4. Create the root `tk.Tk()` instance.
5. Show the splash screen and first-run dialogs as needed.
6. Ensure user identity exists.
7. Ensure schema creation through `crea_database_v4()`.
8. Instantiate `MainWindow(root)`.
9. Enter `mainloop()`.
10. If a restart has been scheduled, launch the detached post-mainloop restart process.

**`SettingsWindow` responsibilities:**

- DataFlow folder relocation
- language selection
- currency selection
- manual backup
- daily auto-backup configuration
- controlled restart prompt after settings requiring restart

**`MainWindow` responsibilities:**

- own the root dashboard state and Tk variables
- build the main dashboard through `build_main_dashboard(app)`
- delegate search/refresh to `DashboardController`
- load RFQ, Saving, Cost Avoidance, and Derisking sheets
- manage actions menus, selection state, dialogs, exports, and subwindows
- schedule daily auto-backup checks
- coordinate restart requests via `restart_lifecycle_service`

**Notable orchestration characteristics:**

- Excel export is dispatched to `services/excel_export_service.py`
- dashboard action rules are computed by policy helpers
- tab-specific search/filter behavior is centralized in the controller/services
- Derisking is handled as a supplier-based module, separate from monetary VSM events

---

### 5.2 `database_manager.py`

**Role:** single access class for nearly all SQL in the application.

#### Connection model

- SQLite 3 backend
- read-write mode enables WAL and tuning pragmas
- read-only mode uses SQLite URI access (`mode=ro`) for aggregation scenarios
- `row_factory = sqlite3.Row`
- context manager support is used throughout the repository

#### Schema management

`create_tables()` performs conservative schema creation and migration through `CREATE TABLE IF NOT EXISTS` and guarded `ALTER TABLE` steps.

Current core tables:

| Table | Purpose |
|---|---|
| `fornitori` | RFQ supplier registry |
| `richieste_offerta` | RFQ header |
| `dettagli_richiesta` | RFQ detail lines |
| `richiesta_fornitori` | RFQ-to-supplier link |
| `offerte_ricevute` | Supplier price per detail line |
| `allegati_richiesta` | Attachments and external references |
| `vsm_events` | Saving / Cost Avoidance / Derisking event headers |
| `vsm_impacts` | Monthly monetary impacts |
| `potential_suppliers` | Derisking supplier registry |
| `supplier_categories` | Category catalog for Derisking |

#### Current schema-relevant characteristics

- RFQ IDs follow a year-based generated range, not a raw AUTOINCREMENT sequence.
- `offerte_ricevute.prezzo_unitario` remains `VARCHAR`, preserving UI-entered decimal formatting.
- `vsm_events` now includes `payments_rate` and `new_supplier`.
- `potential_suppliers` includes `category` and `created_at`; legacy rows may still have `created_at = NULL`.
- `supplier_categories` is maintained without a formal foreign key from `potential_suppliers.category`.

#### Aggregation model

Multi-user aggregation is implemented directly in `DatabaseManager`:

- `get_all_richieste_aggregated(...)`
- `get_all_vsm_events_aggregated(...)`
- `get_all_potential_suppliers_aggregated(...)`

The code scans sibling databases only in the standard DataFlow layout:

```text
<shared root>/DataFlow_*/Database/dataflow_db_*.db
```

It does not recursively scan the entire shared root. RFQ aggregation uses `ATTACH DATABASE` for SQL-side union logic; VSM and Derisking aggregation use direct read-only connections.

#### Important accuracy note

Some comments in `database_manager.py` still mention DuckDB, but the current implementation is fully SQLite-based. The runtime behavior is unambiguously SQLite.

---

## 6. UI Architecture

### 6.1 Tkinter Structure

The application is rooted in a single `tk.Tk` instance. Secondary interfaces are implemented as `tk.Toplevel` windows and dialogs.

The main dashboard is a `ttk.Notebook` with five tabs:

- `Active RfQs`
- `Archived RfQs`
- `Saving`
- `Cost Avoidance`
- `Derisking`

The UI relies on `ttk` for structure and on `tksheet` for data-heavy grids.

### 6.2 `ui/main_dashboard_builder.py`

This file remains a pure widget builder. It creates:

- the top action bar
- the global search toolbar
- the collapsible advanced filters block
- the main notebook tabs
- the footer

The top toolbar exposes:

- New Event
- Actions
- Export Excel
- KPI
- Help
- License
- Settings

The footer currently renders the application version string `v.2.3.0`.

### 6.3 `ui/components/`

#### `main_dashboard_toolbar.py`

Implements the shared search bar:

- global search entry with placeholder management
- advanced filters toggle
- Enter key routing to `search_requests()`
- behavior adapted to RFQ versus VSM/Derisking contexts

#### `collapsible_filters.py`

Wraps the advanced filters panel in a show/hide container using `grid_remove()` semantics. This allows the dashboard to collapse unused filter space without rebuilding widgets.

#### `supplier_name_suggest.py`

Provides the UI controller for supplier-name suggestions in entry fields. It is used in supplier-editing flows and is backed by `SupplierNameSuggestionService`.

### 6.4 `ui/sheet_factories.py`

This module centralizes `tksheet` creation for:

- RFQ summary sheets
- VSM event sheets
- Derisking supplier sheets

It also sets sheet headers, widths, sort behavior, alignment, and metadata placeholders used later by selection and ownership policies.

### 6.5 `ui/dialogs/`

#### `common_dialogs.py`

This module now acts as the standardized dialog layer. It provides reusable dialog classes and wrapper helpers such as:

- `SimpleMessageDialog`
- `SimpleYesNoDialog`
- `SimpleOkCancelDialog`
- `LanguagePrompt`
- `NewRdOTypeDialog`
- `UserIdentityDialog`
- `CopyProgressWindow`
- `SplashScreen`
- `LicenseAcceptanceDialog`
- `show_info`, `show_success`, `show_error`, `show_warning`, `show_confirm`, `show_ok_cancel`

#### `vsm_event_dialog.py`

Dynamic form for VSM events:

- adapts fields to `Saving`, `Cost Avoidance`, and `Derisking`
- supports read-only mode for events opened from external aggregated databases
- distinguishes price-driven versus payment-driven logic
- surfaces the `new_supplier` field for Derisking

#### `potential_supplier_dialog.py`

Potential supplier CRUD dialog:

- category selection
- translated status display over canonical persisted values
- contact validation
- supplier-name suggestions
- read-only compatibility

#### `manage_supplier_categories_dialog.py`

Category management dialog with pending in-memory operations followed by atomic apply.

#### `rfq_pdf_export_dialog.py`

Dialog dedicated to RFQ PDF export:

- logo configuration/removal
- external template access
- PDF generation trigger

### 6.6 `ui/windows/`

#### `view_request_window.py`

Main RFQ operational window. It:

- loads RFQ header/detail/supplier data
- adapts geometry using `calculate_optimal_window_size(...)`
- supports read-only mode when an RFQ belongs to another aggregated user
- opens child windows for attachments, notes, suppliers, PO numbers, SQDC, and PDF export
- exports RFQ data to Excel and PDF

#### `attachment_window.py`

Attachment management window:

- handles BLOB-stored and path-based attachments
- resolves external paths relative to the originating database folder when necessary
- optionally enables drag-and-drop through runtime `tkdnd` or `tkinterdnd2`
- disables mutating actions in read-only mode

#### `notes_window.py`

Rich-text note editor:

- serializes formatting through `Text.dump(...)`
- saves data via `repr(content_dump)`
- reloads with `json.loads(...)` first and `ast.literal_eval(...)` fallback for legacy content

#### Other RFQ support windows

| Window | Purpose |
|---|---|
| `edit_suppliers_window.py` | Supplier list editing with suggestions and soft duplicate warnings |
| `purchase_order_window.py` | Purchase order management persisted in `numeri_ordine` |
| `edit_reference_window.py` | RFQ reference editor |
| `sqdc_analysis_window.py` | SQDC scoring, export, and internal attachment save flow |

### 6.7 `ui/kpi_window.py` and `ui/kpi_chart.py`

`KpiWindow` is a full analytical top-level window with:

- rolling period presets: `1M`, `3M`, `12M`, `3Y`, `5Y`, `10Y`, `ALL`
- mutually exclusive year filter
- notebook tabs for RFQ, Saving, Cost Avoidance, and Derisking
- KPI cards, charts, details tables, and export flow

`ui/kpi_chart.py` renders charts directly on `tk.Canvas`, avoiding external plotting dependencies.

### 6.8 `ui/window_launchers.py`

Centralized launcher helpers:

- opens language-sensitive GitHub wiki URLs for Help
- opens the LICENSE URL
- creates `KpiWindow`

---

## 7. Business Logic

### 7.1 Dashboard, RFQ, KPI, and Derisking orchestration

#### Dashboard controller

`services/dashboard_controller.py` is the central search/refresh coordinator for the main dashboard.

Key current behaviors:

- refresh reuses a single aggregated RFQ load for active tab, archived tab, and username filter population
- filter panel content changes by active tab
- Derisking hides VSM-only date widgets
- search dispatch is module-aware:
  - RFQ search
  - VSM search
  - Derisking supplier search

#### RFQ search behavior

The RFQ search path combines:

- structural filters: status, type, username
- a global multi-field OR block
- historical per-field filters combined with AND logic

The code additionally:

- sanitizes forbidden characters from search inputs
- limits search field length to 100 characters
- normalizes RFQ type before DB comparison

#### Dashboard policy helpers

- `dashboard_selection_policy.py` extracts selected row indices and blocks actions when metadata is missing
- `dashboard_actions_policy.py` computes whether delete, duplicate, archive, or reactivate actions are allowed

Ownership-sensitive actions are intentionally fail-safe: if a selected row is not confirmed as belonging to the current user, the action is disabled.

#### RFQ dashboard service

`rfq_dashboard_service.py` loads RFQs by status and prepares sheet payloads. Metadata tracks whether a row is owned locally and, in aggregated mode, from which source file it was loaded.

#### VSM and Derisking dashboard services

- `vsm_dashboard_service.py` loads either local or aggregated VSM datasets and applies advanced filters for dates, action, repetitive flag, and theoretical/effective ranges
- `derisking_dashboard_service.py` loads `PotentialSupplier` datasets and builds supplier-sheet rows plus row metadata

#### KPI engine

`kpi_engine.py` is a read-only calculation layer with four public entry points:

- `get_rfq_kpi(...)`
- `get_saving_kpi(...)`
- `get_cost_avoidance_kpi(...)`
- `get_derisking_kpi(...)`

Important current semantics:

- Saving and Cost Avoidance are calculated from `vsm_impacts`, not only from event dates
- available year lists are differentiated between general KPI data and Derisking supplier creation data
- Derisking KPI counts rely on `potential_suppliers.created_at`, which means legacy rows with `NULL` creation date are excluded from time-based KPI calculations

#### KPI chart data

`kpi_chart_data.py` creates deterministic month buckets based on the selected filter range, ensuring empty months remain visible in charts instead of disappearing from the x-axis.

### 7.2 VSM Engine

`services/vsm_engine.py` converts a `VSMEvent` into monthly `VSMImpact` rows.

Current rules:

1. Derisking events do not generate monetary impacts.
2. Price-driven events compute theoretical value from negotiated versus baseline price deltas.
3. Payment-driven events compute theoretical value from annual spend, payment-day delta, and the active coefficient.
4. Repetitive OPEX events distribute value for up to 24 months.
5. The first month is prorated using a 30-day convention.
6. Effective value applies realization only where the event/model logic requires it.

### 7.3 VSM Persistence

`services/vsm_persistence.py` uses an explicit transactional pattern:

1. save or update the VSM event
2. delete existing impacts when applicable
3. regenerate impacts from the current event state
4. insert regenerated impacts
5. commit atomically

This module deliberately avoids incremental impact patching.

### 7.4 RFQ, export, and supplier-specific business flows

#### RFQ command service

`rfq_command_service.py` implements:

- bulk status changes
- deletion with best-effort cleanup of external attachment files
- full RFQ duplication
- creation of a minimal RFQ shell before the detail window is opened

#### Excel export service

`excel_export_service.py` now centralizes the dashboard export flows:

- RFQ export
- VSM export
- Derisking export

It uses:

- `LanguagePrompt`
- translation helpers
- safe filename builders
- currency-aware Excel formatting

#### RFQ PDF export services

The PDF export feature is split across three services:

| File | Role |
|---|---|
| `rfq_pdf_export_service.py` | Build the ReportLab document and data table |
| `rfq_pdf_logo_service.py` | Persist and validate a user-configured company logo |
| `rfq_pdf_template_service.py` | Persist and validate editable text templates with a single `{{TABLE}}` placeholder |

#### Supplier persistence and suggestions

- `supplier_persistence.py` owns CRUD for `potential_suppliers`
- `supplier_category_persistence.py` owns rename/merge/delete-if-unused operations for categories
- `supplier_name_suggestion_service.py` builds a read-only suggestion index from both RFQ suppliers and potential suppliers, enabling autocomplete and soft duplicate warnings

---

## 8. Data Layer

### 8.1 Database Engine

The current application uses **SQLite 3** with **WAL** in read-write mode.

Why this matters in the current architecture:

- local writes stay simple and embedded
- read-only aggregation from sibling databases remains possible
- WAL sidecar files are part of backup considerations

### 8.2 Database File Location

Current path behavior is determined by `services/app_paths.py`.

Default DataFlow folder:

```text
Windows:   ~/Documents/DataFlow_<username>/
Linux/mac: ~/DataFlow_<username>/
```

Current default database path:

```text
<DataFlow folder>/Database/dataflow_db_<username>.db
```

Important path rules:

- `Settings.dataflow_base_dir` can relocate the entire DataFlow folder
- legacy `custom_db_path` is still honored
- `Attachments/` is the canonical attachment folder name
- legacy `Allegati/` is migrated conservatively when detected

### 8.3 Schema Summary

```text
fornitori
  - id_fornitore
  - nome_fornitore

richieste_offerta
  - id_richiesta
  - data_emissione
  - data_scadenza
  - riferimento
  - note_generali
  - stato
  - numeri_ordine
  - tipo_rdo
  - note_formattate
  - username

dettagli_richiesta
  - id_dettaglio
  - id_richiesta
  - codice_materiale
  - descrizione_materiale
  - quantita
  - disegno
  - data_consegna_richiesta
  - codice_grezzo
  - disegno_grezzo
  - materiale_conto_lavoro

richiesta_fornitori
  - id_richiesta
  - nome_fornitore

offerte_ricevute
  - id_dettaglio
  - nome_fornitore
  - prezzo_unitario

allegati_richiesta
  - id_allegato
  - id_richiesta
  - nome_file
  - dati_file
  - tipo_allegato
  - nome_fornitore
  - percorso_esterno
  - data_inserimento

vsm_events
  - event_id
  - username
  - event_date
  - buyer
  - event_type
  - action
  - description
  - reference
  - importo_bdg
  - importo_negoziato
  - importo_richiesto_iniziale
  - quantita_annua
  - percent_realizzo
  - driver
  - giorni_pagamento_attuali
  - giorni_pagamento_negoziati
  - spending_annuo
  - opex_ripetitivo
  - note
  - created_at
  - updated_at
  - payments_rate
  - new_supplier

vsm_impacts
  - impact_id
  - event_id
  - username
  - year/month
  - valore_teorico
  - valore_effettivo
  - value_type

potential_suppliers
  - supplier_id
  - supplier_name
  - category
  - supplier_status
  - contact_name
  - email
  - phone
  - website
  - notes
  - username
  - created_at
  - updated_at

supplier_categories
  - id
  - name
  - created_at
```

### 8.4 Data Access Principles

- No ORM is used.
- SQL is centralized in `DatabaseManager`.
- Services own transaction boundaries for multi-step business operations.
- Read-only aggregation paths are separate from local write paths.
- `PotentialSupplier` data is not stored in `vsm_events`; Derisking has its own dedicated registry table.

---

## 9. Internationalization

### 9.1 Core Mechanism

`utils/i18n_utils.py` is the official localization entry point.

Public runtime API:

- `tr(text)`
- `_` as backward-compatible alias
- `init_i18n(...)`
- `get_current_language()`

The central `TranslationService` loads the active `.mo` file and publishes the translator through both the module API and `builtins._`.

### 9.2 Initialization Order

`dataflow.py` explicitly initializes i18n before importing UI modules. This is necessary because several UI modules evaluate translated strings during import/build time.

### 9.3 Domain Normalization Helpers

The current codebase does more than generic gettext translation. It also provides closed-domain normalization helpers for:

- RFQ types
- VSM actions
- Derisking statuses

This matters because the persisted canonical values are not always the same as the currently displayed UI labels.

### 9.4 Catalogs

The repository contains:

- `locale/en/LC_MESSAGES/dataflow.po`
- `locale/en/LC_MESSAGES/dataflow.mo`
- `locale/it/LC_MESSAGES/dataflow.po`
- `locale/it/LC_MESSAGES/dataflow.mo`

Compilation is supported by `development/dev_tools/compile_translations.py`.

### 9.5 Current i18n Usage Pattern

The current 2.3.0 codebase uses `tr(...)` extensively across:

- `dataflow.py`
- dashboard builders and dialogs
- KPI window and charts
- export services
- PDF export flow
- search normalization/display logic

This is a broader and more systematic runtime translation usage than a purely UI-label-only approach.

On first launch, the license acceptance dialog remains English-only and includes an interface-language selector. When the user accepts the license, `dataflow.py` stores both `license_accepted=True` and the selected language in `config.ini`, then reinitializes i18n before the dashboard is created. Later language changes remain handled through Settings.

---

## 10. Build & Distribution

### 10.1 Runtime Dependencies

Packaging now requires not only Excel/UI dependencies but also:

- `reportlab` for RFQ PDF export
- `tkinterdnd2` for optional attachment drag-and-drop

### 10.2 Main Packaging Files

| File | Role |
|---|---|
| `dataflow.spec` | Main one-folder Windows-oriented build |
| `dataflow_appimage.spec` | Root Linux/AppImage-oriented build spec |
| `.github/workflows/build-windows.yml` | GitHub Actions packaging workflow |
| `app.manifest.xml` | Windows DPI manifest |

### 10.3 GitHub Actions Workflow

The repository workflow:

- runs on Windows
- uses Python 3.10
- installs `requirements.txt` and `pyinstaller`
- builds with `pyinstaller dataflow.spec`
- verifies `dist/dataflow/dataflow.exe`
- zips the produced folder
- uploads it as artifact
- publishes a release asset on version tags

### 10.4 Build Variants Under `development/dev_tools`

The repository also includes alternate build specs and packaging metadata under `development/dev_tools/tools_build_WIN/`.

Two points are worth noting because they are directly verifiable in the current files:

1. The root specs and the dev-tool specs are not perfectly identical.
2. The root `dataflow.spec` collects `reportlab`, while `development/dev_tools/tools_build_WIN/exe/DataFlow.spec` does not currently collect `reportlab`.

This manual records the current file state; it does not normalize or reinterpret it.

### 10.5 MSIX Metadata

`development/dev_tools/tools_build_WIN/msix/AppxManifest.xml` is present and currently declares an MSIX package identity version of `2.3.0.0`.

---

## 11. Testing Strategy

### 11.1 Current State

The current repository does **not** contain application-level automated test modules for the main services/UI/business logic paths.

What remains under `tests/` is:

- `tests/db_test_tool/generate_test_db_eng.py`
- `tests/db_test_tool/generate_test_db_it.py`
- two generated SQLite fixture databases

### 11.2 Role of the DB Test Tool

The generator scripts create dense sample databases with:

- hundreds of RFQs
- multiple suppliers per RFQ
- Saving / Cost Avoidance / Derisking sample data
- date distribution over multiple years

This is useful for:

- manual dashboard validation
- KPI verification
- export smoke testing
- visual regression checks on large datasets

### 11.3 What Is Verifiable from the Repository

The repository supports **manual and fixture-driven validation**, not a repository-contained automated regression suite for the current 2.3.0 application logic.

No current `pytest`-style module set for the active business logic is present in the analyzed codebase.

---

## 12. Assets & Resources

### 12.1 Repository-managed resources

- icons and logos in `add_data/`
- Excel templates in `add_data/`
- translation catalogs in `locale/`

### 12.2 Runtime-generated user resources

The PDF export feature persists user-specific resources outside the repository tree under the DataFlow user folder:

```text
<DataFlow folder>/Assets/RFQ_PDF/
```

These resources include:

- persistent company logo
- `pdf_template_eng.txt`
- `pdf_template_ita.txt`

### 12.3 Resource Resolution

`utils/resource_utils.py` resolves resource paths correctly in both:

- source execution
- PyInstaller-frozen execution

This same module also centralizes window icon application across top-level windows and dialogs.

---

## 13. Utilities Layer

### Utility modules

| File | Responsibility |
|---|---|
| `export_filename.py` | Safe, timestamped export filename generation |
| `format_utils.py` | Number, quantity, currency parsing/display, Excel number formats |
| `i18n_utils.py` | Translation service and domain normalization helpers |
| `resource_utils.py` | Resource resolution and window icon setup |
| `string_utils.py` | Username generation and accent stripping |
| `supplier_name_normalization.py` | Soft normalization and match keys for supplier names |
| `user_utils.py` | App data dir, config path, and persisted user identity |
| `validation_utils.py` | Filename sanitization, date conversion, price display, email/website validation |
| `vsm_config.py` | Persisted payment-coefficient helper |
| `window_utils.py` | Centering and dynamic window geometry helpers |

### Current utility themes

#### Formatting and currency

`format_utils.py` now supports configurable currency display and matching Excel numeric formats. Allowed currency codes persisted by settings are:

- `NONE`
- `EUR`
- `USD`
- `GBP`
- `CHF`

#### Validation

`validation_utils.py` is used broadly for:

- filename sanitization
- `dd/mm/yyyy` to DB date conversion
- website validation
- email validation

#### Resource and config paths

`user_utils.py` and `app_paths.py` split responsibilities:

- `user_utils.py` owns config and identity persistence
- `app_paths.py` owns the DataFlow working folder and DB/attachment paths

---

## 14. Execution Flow

### 14.1 Application startup

1. Process starts with `dataflow.py`.
2. DPI-awareness is configured where possible.
3. i18n is initialized.
4. startup cleanup/logging services run.
5. the root Tk window is created.
6. if required, the first-run license dialog collects acceptance and the initial interface language.
7. user identity is loaded or collected.
8. the DataFlow folder structure is ensured.
9. `crea_database_v4()` ensures the schema exists.
10. `MainWindow` builds the dashboard.
11. initial data loads populate RFQ and VSM/Derisking views.
12. `mainloop()` begins.

### 14.2 Main dashboard interaction flow

1. The user changes tab or enters search/filter criteria.
2. `MainWindow` forwards the request to `DashboardController`.
3. The controller selects the correct data pipeline based on the active tab.
4. The relevant dashboard service loads local or aggregated data.
5. Rows and metadata are pushed to the appropriate `tksheet`.
6. Selection metadata determines which actions are enabled.

### 14.3 RFQ operational flow

1. A new RFQ shell is created or an existing RFQ is opened.
2. `ViewRequestWindow` loads RFQ header, details, suppliers, offers, and attachment state.
3. Child windows manage notes, suppliers, purchase orders, attachments, and SQDC.
4. RFQ exports can be generated to Excel or PDF.
5. Changes are persisted through `DatabaseManager`.

### 14.4 VSM and Derisking flow

1. The user opens the relevant VSM or Derisking tab.
2. Local or aggregated data is loaded depending on the username scope.
3. Dialogs build or edit domain objects.
4. VSM save/update operations regenerate monthly impacts transactionally.
5. KPI and export flows consume persisted VSM/Derisking data afterward.

### 14.5 Maintenance flow

1. Settings changes are saved through dedicated settings services.
2. Manual backup copies the DB and available WAL/SHM companions.
3. Auto-backup enforces timestamped retention of up to three backup sets.
4. DataFlow folder migration copies the full user folder and schedules restart.

---

## 15. Design Principles

### 15.1 Conservative extraction over rewrite

The 2.3.0 codebase keeps backward-compatible orchestration in `dataflow.py` while extracting isolated concerns into services. The direction is modularization, not framework replacement.

### 15.2 Local-first and file-system-visible state

The application favors transparent, inspectable state:

- SQLite databases
- file-based attachments when selected
- user-editable text templates for PDF export
- config-driven preferences

### 15.3 Multi-user by aggregation, not by shared server

There is no backend service or central API. Multi-user visibility is achieved by reading sibling databases conservatively and surfacing ownership in the UI.

### 15.4 Domain-specific persistence with translated UI

The code persists domain values in canonical forms and translates them at the UI boundary. This is especially visible for RFQ types, VSM actions, and Derisking statuses.

### 15.5 Service-first handling for non-UI complexity

The most sensitive non-visual responsibilities have been pulled away from widget code:

- backup and migration
- export generation
- KPI calculation
- impact generation
- dashboard ownership/action policies

### 15.6 Minimal external footprint

The application intentionally stays within a compact Python desktop stack:

- Tkinter/ttk
- SQLite
- `tksheet`
- `openpyxl`
- `reportlab`
- `Pillow`

### 15.7 Accuracy over abstraction

The repository still contains some legacy comments and build variants that are not fully harmonized. The operational code, however, consistently reflects a pragmatic design principle: keep behavior explicit and close to the real persisted/application state rather than hiding it behind broad abstractions.
