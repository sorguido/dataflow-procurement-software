# DataFlow 2.2.0
@sorguido

## v2.2.0

DataFlow 2.2.0 – Feature Release

---

## 🆕 Added

- **RFQ PDF export**: new export flow directly from the RFQ window
  - ReportLab-based A4 multi-page document generation
  - Optional persistent company logo for PDF exports
  - External editable language-specific text templates with `{{TABLE}}` placeholder validation

- **Dedicated export services**: Excel export logic extracted from `dataflow.py` into dedicated services
  - RFQ export
  - VSM export
  - Derisking export
  - Timestamped safe export filename generation

- **Supplier name suggestion system**: reusable suggestion and soft-duplicate detection flow for supplier input fields
  - Suggestions built from both RFQ supplier history and Derisking supplier registry
  - Integrated in supplier editing and potential supplier dialogs

- **Settings maintenance services**: dedicated support layer for operational maintenance flows
  - Manual backup bundle copy including DB/WAL/SHM when available
  - Daily auto-backup with timestamped retention
  - DataFlow folder migration validation helpers and controlled restart lifecycle

- **Sheet factory layer**: standalone builders for RFQ, VSM, and Derisking `tksheet` instances

---

## ✨ Improvements

- Dashboard search and advanced filters aligned across RFQ, Saving, Cost Avoidance, and Derisking tabs
- Global search now coexists more consistently with contextual filters and export flows
- Runtime translation coverage consolidated around `tr(...)` and centralized normalization helpers for RFQ type, VSM action, and Derisking status
- Main dialog flows standardized through reusable dialog components for message, confirmation, export language, splash, license, and identity prompts
- KPI Analysis filters extended with rolling period presets (`1M`, `3M`, `12M`, `3Y`, `5Y`, `10Y`, `ALL`) and dedicated Derisking year handling
- Derisking dashboard flow now uses dedicated supplier-sheet population and column auto-sizing logic
- RFQ, VSM, KPI, and Derisking export flows aligned more closely with bilingual UI behaviour

---

## 🛠 Fixes

- Search input now applies defensive sanitization and length validation before query execution
- Dashboard row ownership checks are enforced through dedicated selection-policy helpers before enabling destructive actions
- Notes loading now prefers `json.loads()` with fallback for legacy serialized content
- Attachment drag-and-drop activation is guarded more safely when `tkdnd` / `tkinterdnd2` are unavailable in the runtime
- Restart flow after settings changes is handled through dedicated lifecycle helpers for more reliable re-launch behaviour
- DataFlow folder migration and backup flows include more conservative validation and copy safeguards

---

## ♻️ Refactoring

- Large portions of `dataflow.py` operational logic extracted into dedicated services under `services/`
- Dashboard logic split across controller, search, actions policy, selection policy, RFQ dashboard, VSM dashboard, and Derisking dashboard modules
- Settings logic split into preferences, maintenance, location, and restart lifecycle services
- RFQ PDF export split into dedicated export, logo, and template service modules
- New `ui/sheet_factories.py` centralizes `tksheet` construction for dashboard views
- New utility modules introduced for export filename generation and supplier-name normalization
- `utils/i18n_utils.py` refactored into a centralized translation service with runtime initialization and domain-specific normalization helpers

---

## 📦 Packaging

- Distribution packages updated for version 2.2.0:
  - `AppImage`
  - `.exe`
- `requirements.txt` updated to include runtime dependencies required by the new export and drag-and-drop flows
- Root PyInstaller specs updated to collect the modules and assets required by PDF export and multilingual runtime resources

---

## 🔒 Other

- No breaking changes to the existing RFQ / VSM / Derisking user workflows
- Database evolution remains conservative and migration-oriented for new fields and tables
- Multi-user dashboard aggregation remains read-only for data belonging to other users

---

## v2.1.0

DataFlow 2.1.0 – Feature Release

---

## 🆕 Added

- **VSM module — Value Stream Mapping**: new functional area, independent from the RFQ workflow, for tracking quantified negotiation outcomes
  - Three supported event types: **Saving**, **Cost Avoidance**, **Derisking**
  - VSM Event dialog with dynamic form layout (fields shown/hidden based on event type)
  - Read-only view mode for events belonging to other users (multi-user consistency)
  - OPEX-repetitive flag: Saving and Cost Avoidance events with recurring impact propagate over up to 24 months
  - Payment terms driver (Pagamenti) for Saving events, with configurable monthly opportunity cost coefficient

- **VSM Engine**: automatic monthly impact projection engine
  - Pro-rata calculation for the first month
  - Deterministic regeneration of all impacts on every event update (DELETE–REGENERATE–SAVE pattern)
  - Atomic database transactions: full rollback on any insert failure

- **KPI Analysis window**: dedicated window with aggregated procurement KPIs
  - Four tabs: RFQ, Saving, Cost Avoidance, Derisking
  - Period filters: year selector or custom date range (preset: last 3, 6, 12 months or All)
  - KPI cards with real data from the KPI engine
  - Integrated bar charts (pure Tkinter, no external chart library required)
  - Excel export of the full KPI summary

- **Potential Supplier registry** (Derisking tab): dedicated module for managing supplier qualification
  - `PotentialSupplier` model with name, category, contact details, website, and notes
  - Qualification lifecycle: New → Under Evaluation → Qualified / Rejected
  - Per-user ownership with multi-user visibility
  - Integrated with Derisking VSM events (new supplier field)

- **Three new main dashboard tabs**: Saving, Cost Avoidance, Derisking — directly accessible alongside the existing RFQ tabs

- **Global search bar** (`MainDashboardToolbar`): central search entry with placeholder text, active across all dashboard tabs
  - Multi-field OR search: RFQ number, reference, supplier, part code, description, order number
  - Coexists with advanced filters (global OR + contextual filters AND)
  - Empty search triggers filter reset

- **Supplier Category management**: dialog for creating and managing reusable supplier category labels (used in the Potential Supplier registry)

---

## ✨ Improvements

- Main dashboard filter panel now includes a dedicated VSM sub-frame (user, date from, date to) shown/hidden contextually when switching between RFQ and VSM tabs
- KPI charts adapt to canvas resize; dual-bar charts include legend and axis labels
- VSM event duplication: existing events can be duplicated from the dashboard, preserving all fields
- VSM username filter on the dashboard: populated from multi-database aggregation, consistent with existing RFQ username filter behaviour
- `get_available_years` and `get_available_years_derisking` functions expose distinct year lists for filter combos
- Improved numeric formatting for monetary values in KPI cards and Excel export (Italian locale: dot thousands separator, comma decimal)
- KPI Excel export includes bilingual support (Italian / English) consistent with existing Excel exports

---

## 🛠 Fixes

- Dashboard controller separated from `MainWindow.__init__`, eliminating several cases where filter state was not preserved across tab switches
- VSM filter sub-frame correctly hidden when returning to RFQ tabs and shown when activating VSM tabs
- Defensive validation in `populate_username_filter` prevents crash on variable-length aggregated tuples from multi-database queries
- Logo image rendering guard: zero-dimension images no longer cause `ZeroDivisionError` on startup

---

## ♻️ Refactoring

- `MainWindow` UI construction extracted into `ui/main_dashboard_builder.py` (pure widget builder, no data loading)
- Dashboard orchestration logic extracted into `services/dashboard_controller.py`
- `MainDashboardToolbar` and `CollapsibleFilters` introduced as standalone UI components under `ui/components/`
- VSM persistence, VSM engine, KPI engine, KPI chart data, and KPI Excel export each live in dedicated modules under `services/` and `ui/`
- `models/` package extended with `VSMEvent`, `VSMImpact`, and `PotentialSupplier` dataclasses
- `utils/vsm_config.py` introduced for user-configurable VSM parameters (payment coefficient) stored in `config.ini`
- `utils/validation_utils.py` extended with email and website format validation used by the Potential Supplier dialog
- Test suite extended: `test_vsm_engine.py`, `test_vsm_persistence.py`, `test_vsm_event_model.py`, `test_supplier_category_persistence.py`

---

## 📦 Packaging

- Distribution packages updated for version 2.1.0:
  - `AppImage`
  - `.exe`
- `dataflow.spec` updated to include new modules and assets introduced in this release

---

## 🔒 Other

- No breaking changes to existing RFQ data or database schema for existing tables
- New database tables (`vsm_events`, `vsm_impacts`, `potential_suppliers`, `supplier_categories`) are created transparently on first launch
- No new mandatory external dependencies introduced

---

> This release introduces new functional areas and extends the capabilities of DataFlow beyond RFQ management. Procurement teams can now track the quantified impact of negotiation activities and measure buyer performance through a dedicated KPI layer.
