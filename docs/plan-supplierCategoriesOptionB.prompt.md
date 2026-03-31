# Plan: Gestisci Categorie — Opzione B (in-memory + Save/Cancel)

## TL;DR

Refactor `ManageSupplierCategoriesDialog` to work fully in memory. No DB writes until the user clicks **Salva**. **Annulla** or window-close discards all changes. A new `apply_category_ops_atomic()` DB method ensures all ops are committed in a single transaction.

---

## Steps

### Phase 1 — `database_manager.py` — new atomic method

1. After `count_suppliers_by_category` (line ~2891), add:

```python
def apply_category_ops_atomic(self, ops: list) -> None:
    """
    Applica una lista di operazioni sulle categorie in un'unica transazione.
    ops: list of dict, each with key 'type':
      {"type": "rename",       "old": str, "new": str}
      {"type": "merge",        "source": str, "target": str}
      {"type": "delete_unused","name": str}
    Per delete_unused, controlla il count al momento dell'applicazione.
    Per rename, BLOCCA se il nome target esiste già (non degrada silenziosamente in merge).
    """
    if not ops:
        return
    try:
        self.cursor.execute("BEGIN")
        for op in ops:
            t = op["type"]
            if t == "rename":
                old, new = op["old"].strip(), op["new"].strip()
                if not old or not new or old == new:
                    continue  # no-op
                # CORREZIONE 1: blocca se new esiste già — rename != merge
                self.cursor.execute(
                    "SELECT 1 FROM supplier_categories WHERE name = ?", (new,)
                )
                if self.cursor.fetchone():
                    self.cursor.execute("ROLLBACK")
                    raise DatabaseError(
                        f"La categoria '{new}' esiste già. Usa la funzione Unisci."
                    )
                self.cursor.execute(
                    "INSERT OR IGNORE INTO supplier_categories (name) VALUES (?)", (new,)
                )
                self.cursor.execute(
                    "UPDATE potential_suppliers SET category = ? WHERE category = ?", (new, old)
                )
                self.cursor.execute(
                    "DELETE FROM supplier_categories WHERE name = ?", (old,)
                )
            elif t == "merge":
                src, tgt = op["source"].strip(), op["target"].strip()
                self.cursor.execute(
                    "INSERT OR IGNORE INTO supplier_categories (name) VALUES (?)", (tgt,)
                )
                self.cursor.execute(
                    "UPDATE potential_suppliers SET category = ? WHERE category = ?", (tgt, src)
                )
                self.cursor.execute(
                    "DELETE FROM supplier_categories WHERE name = ?", (src,)
                )
            elif t == "delete_unused":
                name = op["name"].strip()
                self.cursor.execute(
                    "SELECT COUNT(*) FROM potential_suppliers WHERE category = ?", (name,)
                )
                row = self.cursor.fetchone()
                count = row[0] if row else 0
                if count > 0:
                    self.cursor.execute("ROLLBACK")
                    raise DatabaseError(
                        f"Impossibile eliminare '{name}': ancora usata da {count} fornitore/i."
                    )
                self.cursor.execute(
                    "DELETE FROM supplier_categories WHERE name = ?", (name,)
                )
        self.conn.commit()
    except DatabaseError:
        raise
    except Exception as e:
        try:
            self.cursor.execute("ROLLBACK")
        except Exception:
            pass
        print(f"[DB Manager] Errore apply_category_ops_atomic: {e}")
        raise DatabaseError(str(e)) from e
```

> **Correzione 1 — rename non degrada in merge**: prima dell'INSERT, si verifica che `new` non esista già nel DB. Se esiste: ROLLBACK + `DatabaseError` con messaggio "La categoria esiste già. Usa la funzione Unisci."

### Phase 2 — Full rewrite of `ui/dialogs/manage_supplier_categories_dialog.py`

Replace entire file with this new implementation:

**Imports** (unchanged except remove `rename_supplier_category`, `merge_supplier_categories`, `delete_supplier_category_if_unused`, `count_suppliers_by_category` — keep only `get_all_supplier_categories`, `CategoryError`):

```python
from services.supplier_category_persistence import (
    get_all_supplier_categories,
    CategoryError,
)
```

**`__init__(self, parent, refresh_derisking_cb=None)`**:
- `self.changes_made = False`
- `self._refresh_derisking_cb = refresh_derisking_cb`
- `self._original_categories: list = []`
- `self._working_categories: list = []`
- `self._pending_ops: list = []`
- Call `_load_initial_state()` then `_build_ui()` then `_refresh_list_from_memory()`
- `self.protocol("WM_DELETE_WINDOW", self._on_cancel)`

**`_load_initial_state()`**: open DB, call `get_all_supplier_categories(db)`, store into both `_original_categories` and `_working_categories`.

**`_build_ui()`**: identical to current except:
- Bottom buttons: replace "Chiudi" single button with:
  - `ttk.Button(text=_("❌ Annulla"), command=self._on_cancel, width=12).pack(side="right", padx=(5,0))`
  - `ttk.Button(text=_("💾 Salva"), command=self._on_save, width=12).pack(side="right")`

**`_refresh_list_from_memory(keep_selection=None)`** (replaces `_refresh_list`):
- No DB call — reads `self._working_categories`
- Repopulates `self._listbox` and `self._combo_merge_target`
- Restores selection if `keep_selection` found in list

**`_update_count_label(name)`**:
- If `name not in self._original_categories`: show `"Fornitori: — (verificato al salvataggio)"`
- Otherwise: query DB `db.count_suppliers_by_category(name)`, show `"Fornitori associati: N"`

**New helper: `_apply_rename_in_ops(old_name, new_name)`** (CORREZIONE 2 — normalizzazione completa)

Riposiziona/aggiorna tutte le pending ops che referenziano `old_name`,
garantendo che `_pending_ops` rifletta sempre lo stato logico corrente.

```python
def _apply_rename_in_ops(self, old_name: str, new_name: str):
    # 1. Consolida: se esiste già un rename che produce old_name → aggiorna new
    #    (catena A→B poi B→C diventa A→C, evitando nomi intermedi nel DB)
    found = False
    for op in self._pending_ops:
        if op["type"] == "rename" and op["new"] == old_name:
            op["new"] = new_name
            found = True
            break
    if not found:
        self._pending_ops.append({"type": "rename", "old": old_name, "new": new_name})

    # 2. Aggiorna tutti gli altri riferimenti a old_name nelle ops successive
    for op in self._pending_ops:
        if op["type"] == "merge":
            if op["source"] == old_name:
                op["source"] = new_name
            if op["target"] == old_name:
                op["target"] = new_name
        elif op["type"] == "delete_unused":
            if op["name"] == old_name:
                op["name"] = new_name

    # 3. Rimuovi rename no-op (old == new, es. dopo catena circolare A→B→A)
    self._pending_ops = [
        op for op in self._pending_ops
        if not (op["type"] == "rename" and op["old"] == op["new"])
    ]

    # 4. Pulizia ops stale
    self._prune_stale_ops()
```

**New helper: `_prune_stale_ops()`** (CORREZIONE 3 — pulizia delete_unused incoerenti)

Rimuove `delete_unused` ops i cui nomi non sono più presenti in `_working_categories`
(cioè sono stati consumati da un merge o rinominati via).
NON viene chiamata da `_on_delete` (che rimuove intenzionalmente dalla working list);
viene chiamata da `_apply_rename_in_ops` e da `_on_merge`.

```python
def _prune_stale_ops(self):
    # Un delete_unused è stale se il suo nome è sparito da working_categories
    # senza essere stato aggiunto come destinazione di un rename pendente.
    valid_names = set(self._working_categories)
    # Aggiungi i target di rename pendenti (nomi che "esisteranno" dopo il save)
    for op in self._pending_ops:
        if op["type"] == "rename":
            valid_names.add(op["new"])
    self._pending_ops = [
        op for op in self._pending_ops
        if not (
            op["type"] == "delete_unused"
            and op["name"] not in valid_names
        )
    ]
```

**`_on_rename()`** — in-memory only:
1. Validate: selected, new_name not empty, old != new
2. **Blocco early**: `if new_name in self._working_categories` → `SimpleMessageDialog` con testo  
   `"La categoria esiste già. Usa la funzione Unisci."` + return  
   *(impossibile arrivare all'atomic method con un rename verso categoria esistente)*
3. Update `_working_categories`: replace old_name with new_name, re-sort case-insensitively
4. Call `self._apply_rename_in_ops(old_name, new_name)` — gestisce consolidamento + normalizzazione completa
5. Call `_refresh_list_from_memory(keep_selection=new_name)`

**`_on_merge()`** — in-memory only:
1. Validate: source selected, target chosen, source != target
2. `SimpleYesNoDialog` confirm
3. `_working_categories.remove(source)`
4. Append `{"type": "merge", "source": source, "target": target}`
5. Call `self._prune_stale_ops()` — rimuove eventuali `delete_unused(source)` pendenti
6. Call `_refresh_list_from_memory(keep_selection=target)`

**`_on_delete()`** — in-memory only:
1. Validate: name selected
2. `_working_categories.remove(name)`
3. Append `{"type": "delete_unused", "name": name}`
4. Call `_refresh_list_from_memory()` *(nessuna normalizzazione: l'op è intenzionale)*

**`_on_save()`**:
1. If `_pending_ops` empty → `self.destroy()` return
2. `with DatabaseManager(get_db_path()) as db: db.apply_category_ops_atomic(self._pending_ops)`
3. On `DatabaseError` → `SimpleMessageDialog` + return (do NOT close, let user fix or cancel)
4. On success: `self.changes_made = True`
5. If `self._refresh_derisking_cb`: call it (wrapped in try/except)
6. `self.destroy()`

**`_on_cancel()`**: `self.destroy()` only, no DB write.

### Phase 3 — `ui/dialogs/potential_supplier_dialog.py`

Two targeted changes:

1. Add `refresh_derisking_cb=None` parameter to `__init__`:
   - Store `self._refresh_derisking_cb = refresh_derisking_cb`

2. In `_on_manage_categories()`, pass it to the dialog:
   ```python
   dlg = ManageSupplierCategoriesDialog(self, refresh_derisking_cb=self._refresh_derisking_cb)
   ```

### Phase 4 — `dataflow.py` — two call sites (*parallel with Phase 3*)

Both places that open `PotentialSupplierDialog` need the callback added:

**Site 1** (~line 1899, `_on_supplier_sheet_double_click`):
```python
dlg = PotentialSupplierDialog(
    self.root,
    self.current_username,
    supplier_id=supplier_id,
    read_only=not is_mine,
    refresh_derisking_cb=lambda: self._load_potential_suppliers(self.sheet_derisking),
)
```

**Site 2** (~line 3570, new supplier creation in `on_new_button_click`):
```python
dlg = PotentialSupplierDialog(
    self.root,
    self.current_username,
    refresh_derisking_cb=lambda: self._load_potential_suppliers(self.sheet_derisking),
)
```

---

## Relevant Files

- `database_manager.py` — add `apply_category_ops_atomic()` after `count_suppliers_by_category` (~line 2891)
- `ui/dialogs/manage_supplier_categories_dialog.py` — full rewrite (in-memory state, Save/Cancel)
- `ui/dialogs/potential_supplier_dialog.py` — add `refresh_derisking_cb` param + forward it
- `dataflow.py` — two call sites: add `refresh_derisking_cb=lambda: ...`

---

## Verification

1. `python3 -m unittest discover -s tests -q` → 63+ tests OK
2. Open Gestisci Categorie, rinomina una categoria → listbox aggiornata, DB invariato
3. Clicca Annulla → DB invariato, combo non cambiata
4. Rinomina + Salva → DB aggiornato, combo nel PotentialSupplierDialog refreshata, griglia Derisking refreshata
5. Elimina categoria in uso → errore chiaro al Save, dialog NON si chiude, utente può correggere o Annulla
6. Elimina categoria non in uso → rimossa al Save
7. Merge → sorgente sparisce al Save, tutti i supplier migrati

### Scenario walkthroughs (correzioni)

**Scenario A** — rename "Tornerie" → "Meccanica", "Meccanica" già esiste
- `_on_rename()`: `new_name in _working_categories` → blocco immediato nel dialog
- Messaggio: `"La categoria esiste già. Usa la funzione Unisci."`
- DB invariato, nessuna op aggiunta

**Scenario B** — rename A→B, merge B→C, save
- `_on_rename(A, B)`: working=[...,B,...], ops=[{rename,old:A,new:B}]
- `_on_merge(B, C)`: B rimosso da working, ops=[{rename,old:A,new:B},{merge,src:B,tgt:C}], `_prune_stale_ops()` non rimuove nulla (B non è in delete_unused)
- `_on_save()` → `apply_category_ops_atomic`:
  1. rename A→B (controlla B non esiste nel DB — corretto se B era nuovo)
  2. merge B→C (B appena creato da step 1, merge into C)
- Risultato finale: nessun A, nessun B, tutti i supplier di A ora in C ✓

**Scenario C** — rename A→B, delete_unused B, save
- `_on_rename(A, B)`: ops=[{rename,old:A,new:B}]
- `_on_delete(B)`: B rimosso da working (intenzionale), ops=[{rename,old:A,new:B},{delete_unused,name:B}]
  - (`_prune_stale_ops` NON chiamata qui — l'op è intenzionale)
- `_on_save()` → atomic:
  1. rename A→B
  2. delete_unused B: count check; se 0 → elimina B
- Nessun riferimento incoerente ad A ✓

**Scenario D** — merge A→B, poi tenta delete_unused A
- `_on_merge(A, B)`: A rimosso da working, ops=[{merge,src:A,tgt:B}], `_prune_stale_ops()`: nessun delete_unused da pruning
- A non più in working → non selezionabile → `_on_delete(A)` non raggiungibile dalla UI ✓
- Se per qualsiasi ragione un `delete_unused(A)` fosse già pendente prima del merge:  
  `_prune_stale_ops()` (chiamata in `_on_merge`) lo rimuove perché A non è in working ✓

---

## Decisions

- `apply_category_ops_atomic` usa `isolation_level=None` già impostato nel DB (autocommit mode). Il `BEGIN`/`COMMIT` manuale nella nuova funzione è necessario e corretto.
- **Rename != merge**: il rename blocca se il target esiste già, sia nel dialog (blocco UI) sia in `apply_category_ops_atomic` (guardia DB). Doppia protezione.
- **`_apply_rename_in_ops`**: consolida chained renames (A→B poi B→C → op unica A→C), aggiorna tutti i riferimenti forward nelle ops successive, rimuove no-op (A→A dopo ciclo).
- **`_prune_stale_ops`**: chiamata da `_apply_rename_in_ops` e da `_on_merge`; NON da `_on_delete`. Rimuove delete_unused le cui target non compaiono né in working_categories né come `new` di un rename pendente.
- Se Save fallisce (es. delete_unused su categoria tornata in uso, o rename verso nome appena inserito da altro utente), la transazione fa ROLLBACK completo → DB invariato, dialog resta aperto per correzione/cancel.
- `_update_count_label` continua a queryare il DB per categorie originali; per categorie nuove in sessione mostra "— (verificato al salvataggio)". Corretto e non fuorviante.
- Rimosso "Chiudi" (ambiguo) → solo "💾 Salva" + "❌ Annulla".
