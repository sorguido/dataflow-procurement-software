# DATAFLOW 2.1.0 — IMPLEMENTAZIONE CONSERVATIVA "GESTISCI CATEGORIE"

## TL;DR

Introduce `supplier_categories` table as a canonical category catalogue. Add a dedicated persistence layer, a `ManageSupplierCategoriesDialog`, and wire everything into the existing `PotentialSupplierDialog`. Conservative: `potential_suppliers.category` stays TEXT, no FK, no schema breakage.

---

## Steps

### Phase 1 — Database (`database_manager.py`)

1. In `create_tables()`, after the `idx_ps_category` index (line ~307) and before `# Commit finale`:
   - `CREATE TABLE IF NOT EXISTS supplier_categories (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE)`
   - Idempotent migration: `INSERT OR IGNORE INTO supplier_categories (name) SELECT DISTINCT category FROM potential_suppliers WHERE category IS NOT NULL AND TRIM(category) != ''`

2. Add 6 new methods to `DatabaseManager` (grouped near `get_distinct_macrocategories`, line ~2690):
   - `get_all_supplier_categories() → list[str]`
   - `ensure_supplier_category_exists(name)` — `INSERT OR IGNORE`
   - `rename_supplier_category(old, new)` — transactional: UPDATE suppliers + swap table row
   - `merge_supplier_categories(source, target)` — transactional: UPDATE suppliers + DELETE source
   - `delete_supplier_category_if_unused(name) → int` — count first, delete only if 0, return count
   - `count_suppliers_by_category(name) → int`

### Phase 2 — Persistence Layer (new file, *parallel with Phase 1*)

3. Create `services/supplier_category_persistence.py`:
   - `CategoryError(Exception)` for business logic violations
   - 6 thin wrapper functions with input validation and readable error messages
   - `rename`: if `new_name` already exists → `CategoryError("La categoria esiste già. Usa la funzione Unisci.")`

### Phase 3 — Management Dialog (new file, *parallel with Phase 2*)

4. Create `ui/dialogs/manage_supplier_categories_dialog.py` — `ManageSupplierCategoriesDialog(tk.Toplevel)`:
   - `self.changes_made = False` → set `True` after any DB write
   - Layout: `LabelFrame` "Categorie" with `Listbox` + scrollbar; three `LabelFrame` sections (Rinomina / Unisci / Elimina); "Chiudi" button at bottom-right
   - Rinomina section: Entry for new name + "Rinomina" button
   - Unisci section: target Combobox (readonly) + "Unisci" button + `SimpleYesNoDialog` confirm
   - Elimina section: dynamic label showing supplier count + "Elimina se non usata" button
   - `_refresh_list()`: reloads DB, repopulates Listbox + merge Combobox after every operation
   - Standard window pattern (same as all other dialogs): `withdraw → set_window_icon → transient → grab_set → center_window → deiconify`

### Phase 4 — Integration in `PotentialSupplierDialog` (*depends on Phase 3*)

5. `ui/dialogs/potential_supplier_dialog.py`:
   - Add import: `get_all_supplier_categories`, `ensure_supplier_category_exists`, `CategoryError`
   - Update `_load_known_categories()` → calls `get_all_supplier_categories` (not `get_distinct_macrocategories`)
   - `_build_ui()` buttons: add "Gestisci Categorie" at `side="left"` (Save/Cancel stay at right)
   - Add `_on_manage_categories()`: opens dialog, `wait_window`, then `_refresh_categories()` if `changes_made`
   - Add `_refresh_categories()`: reloads combo values; keeps current selection if still valid
   - `_on_save()`: after resolving `category`, call `ensure_supplier_category_exists(db, category)` inside existing `with DatabaseManager` block
   - `_apply_read_only()`: disable the "Gestisci Categorie" button

### Phase 5 — Tests (*parallel with Phase 4*)

6. Create `tests/test_supplier_category_persistence.py` (7 test methods, same pattern as existing persistence tests)

---

## Relevant Files

- `database_manager.py` — table insertion at line ~307, 6 new methods at line ~2700
- `services/supplier_category_persistence.py` — **NEW**
- `ui/dialogs/manage_supplier_categories_dialog.py` — **NEW**
- `ui/dialogs/potential_supplier_dialog.py` — 5 targeted changes
- `tests/test_supplier_category_persistence.py` — **NEW**

---

## Verification

1. `python3 -m unittest discover -s tests -q` → 54+ tests OK
2. New Supplier dialog: combo reads from `supplier_categories`; free-text save auto-adds to catalogue
3. Gestisci Categorie: all 7 scenarios from spec pass manually

**Scenario 1 — Migrazione**
- DB esistente con category già presenti in potential_suppliers
- apertura app / init DB
- supplier_categories viene creata
- categorie storiche vengono importate una sola volta, senza duplicati

**Scenario 2 — Nuova categoria da dialog supplier**
- apro New Supplier
- scrivo nuova categoria nel campo libero
- salvo
- la categoria viene salvata nel supplier
- la categoria compare anche nell'anagrafica categorie

**Scenario 3 — Rinomina**
- categoria "Pippo" usata da più supplier
- rinomino in "Plastica"
- tutti i supplier passano a "Plastica"
- "Pippo" sparisce dalla tabella categorie
- "Plastica" resta disponibile

**Scenario 4 — Rinomina verso categoria già esistente**
- esistono "Tornerie" e "Lavorazioni meccaniche"
- provo a rinominare "Tornerie" in "Lavorazioni meccaniche"
- operazione bloccata con messaggio che invita a usare Unisci

**Scenario 5 — Unisci**
- unisco "Tornerie" → "Lavorazioni meccaniche"
- tutti i supplier della prima passano alla seconda
- la prima categoria sparisce
- la seconda resta

**Scenario 6 — Elimina non usata**
- categoria presente in supplier_categories ma senza supplier associati
- eliminazione consentita

**Scenario 7 — Elimina usata**
- categoria associata ad almeno un supplier
- eliminazione bloccata

---

## Decisions

- `potential_suppliers.category` stays TEXT — zero schema risk on existing data
- No `created_at`/`updated_at` on `supplier_categories` (not needed, keep it minimal)
- "Gestisci Categorie" button is disabled only in `read_only=True` mode (categories are global, not per-supplier)
- Dialog stays open after each operation (no auto-close)
- `ManageSupplierCategoriesDialog` does NOT call `wait_window` internally — parent calls it

---

## Not Touched

- KPI / kpi_engine.py / kpi_window.py
- VSMEvent / VSMEventDialog
- Export Excel
- Global search
- Old Derisking cleanup
