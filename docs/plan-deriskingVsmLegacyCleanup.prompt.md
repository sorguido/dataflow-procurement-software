# Plan: Cleanup Legacy Derisking VSM

## TL;DR
Ripulire il codice legacy residuo dove Derisking era trattato come evento VSM economico. 3 bug attivi, 3 dead-code entries, 2 mark-only. 2 file coinvolti: `services/dashboard_controller.py` e `dataflow.py`.

---

## Mappatura Completa

### 🔴 BUG — comportamento errato (3 bug attivi)

| ID | File | Dove | Problema |
|----|------|------|----------|
| **B1** | `services/dashboard_controller.py` | `clear_vsm_filters()` ~riga 477 | Chiama `_load_vsm_events("Derisking", sheet_derisking)` → carica VSM events nel foglio fornitore, **corrompendo la visualizzazione** ogni volta che l'utente clicca "Reset filtri" con il tab Derisking attivo |
| **B2** | `dataflow.py` | `_VSM_STATUS_TO_TYPE` riga 3255 | Contiene `'vsm_derisking': 'Derisking'` → se l'utente digita nella **ricerca globale** con Derisking tab attivo, `_search_vsm_events(sheet_derisking, …)` carica VSM events e tenta di popolare lo sheet fornitore con righe a struttura sbagliata |
| **B3** | `dataflow.py` | `_export_vsm_excel()` ~riga 4039 | Contiene `'vsm_derisking': 'Derisking'` → se l'utente clicca **Export Excel** sul tab Derisking, carica 0 eventi VSM e genera un file con header legacy (Nuovo Fornitore, Azione, Valore Teorico…) |

---

### 🟡 DEAD CODE — sicuro da rimuovere

| ID | File | Dove | Perché è morto |
|----|------|------|-----------------|
| **D4** | `dataflow.py` | `_delete_vsm_events()` ~riga 2482 | `event_type_map` include `'vsm_derisking'` ma la funzione fa early-return verso `_delete_supplier()` prima di raggiungere quella riga |
| **D5** | `dataflow.py` | `open_new_event()` ~riga 3649 | Stessa struttura: early-return verso `PotentialSupplierDialog` prima dell'`event_type_map` |
| **D6** | `dataflow.py` | `_edit_vsm_event()` ~righe 2390 e 2410 | Due `event_type_map` con `'vsm_derisking'`; funzione mai raggiunta per il tab Derisking (double-click bindato solo a sheet_saving e sheet_cost_avoidance) |

---

### 🔵 DEAD CODE — lasciare per ora con nota

| ID | File | Dove | Motivazione |
|----|------|------|-------------|
| **C7** | `dataflow.py` | `_create_vsm_event_sheet()` branch `else:` | Headers "Nuovo Fornitore" e `_NEW_SUPPLIER_COL_IDX` logic. Mai chiamato: `sheet_derisking` ora usa `_create_supplier_sheet`. Rimozione nel prossimo step. |
| **C8** | `dataflow.py` | `_populate_vsm_sheet()` branch `else:` | Branch non-`use_dual_value` con `event.new_supplier`. Non più raggiungibile per Derisking dopo fix B2. Rimozione nel prossimo step con cleanup header. |

---

### ✅ KEEP — corretto, necessario

- `services/vsm_engine.py` — `VALID_EVENT_TYPES` include "Derisking" + early-return `[]` → guard per vecchi record DB
- `models/vsm_event.py` — `new_supplier`, `action="Derisking"` → compatibilità schema DB
- `database_manager.py` — tutti i `new_supplier` read/write → schema DB esistente
- `ui/dialogs/vsm_event_dialog.py` — sezione Derisking (new_supplier_frame, lbl_derisking_info) → potrebbe servire per leggere vecchi record VSM Derisking; non causa bug
- `dataflow.py` — tutti i guard `if status == 'vsm_derisking': return / _delete_supplier / PotentialSupplierDialog` → routing corretto, necessari
- `dataflow.py` — `if status != 'vsm_derisking':` (Duplica guard) → corretto
- `tests/test_vsm_engine.py`, `tests/test_vsm_event_model.py` → testano comportamento engine su vecchi record
- `services/kpi_*`, `ui/kpi_window.py` → già aggiornati allo step precedente

---

## Steps

### Phase 1 — Bug fixes (B1, B2, B3)

#### 1a. `services/dashboard_controller.py` — `clear_vsm_filters()` (~riga 472)

```python
# PRIMA
for _et, _sh in [
    ("Saving", getattr(self.app, 'sheet_saving', None)),
    ("Cost Avoidance", getattr(self.app, 'sheet_cost_avoidance', None)),
    ("Derisking", getattr(self.app, 'sheet_derisking', None)),
]:
    if _sh is not None:
        self.app._load_vsm_events(_et, _sh)

# DOPO
for _et, _sh in [
    ("Saving", getattr(self.app, 'sheet_saving', None)),
    ("Cost Avoidance", getattr(self.app, 'sheet_cost_avoidance', None)),
]:
    if _sh is not None:
        self.app._load_vsm_events(_et, _sh)
_sh_dr = getattr(self.app, 'sheet_derisking', None)
if _sh_dr is not None:
    self.app._load_potential_suppliers(_sh_dr)
```

#### 1b. `dataflow.py` — `_VSM_STATUS_TO_TYPE` (~riga 3255)

```python
# PRIMA
_VSM_STATUS_TO_TYPE = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    'vsm_derisking': 'Derisking',
}

# DOPO
_VSM_STATUS_TO_TYPE = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    # vsm_derisking excluded: supplier-based tab, not VSM-event-based
}
```

**Effetto**: `_search_vsm_events(sheet_derisking, 'vsm_derisking')` riceve `event_type=None` da `.get()`, fa early-return silenzioso. La ricerca globale non è attiva sul tab Derisking — comportamento accettabile per questa versione.

#### 1c. `dataflow.py` — `_export_vsm_excel()` (~riga 4035)

Aggiungere early-return prima del blocco `status_to_event_type` e rimuovere l'entry `'vsm_derisking'`:

```python
# PRIMA (dentro _export_vsm_excel)
status_to_event_type = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    'vsm_derisking': 'Derisking',
}
event_type = status_to_event_type.get(status, status)

# DOPO
if status == 'vsm_derisking':
    SimpleMessageDialog(
        self.root,
        _("Export non disponibile"),
        _("L'export Excel del tab Derisking sarà disponibile in una versione successiva."),
        "info"
    )
    return

status_to_event_type = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
}
event_type = status_to_event_type.get(status, status)
```

---

### Phase 2 — Dead code entries (D4, D5, D6)

#### 2a. `dataflow.py` — `_delete_vsm_events()`: rimuovere entry dead

```python
# PRIMA
event_type_map = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    'vsm_derisking': 'Derisking'
}

# DOPO
event_type_map = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
}
```

#### 2b. `dataflow.py` — `open_new_event()`: fix docstring + rimuovere entry dead

```python
# docstring (PRIMA)
"""
- VSM (Saving/Cost Avoidance/Derisking): crea nuovo evento VSM
"""

# docstring (DOPO)
"""
- VSM Saving/Cost Avoidance: crea nuovo evento VSM
- Derisking: apre PotentialSupplierDialog (non VSMEventDialog)
"""

# event_type_map (PRIMA)
event_type_map = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    'vsm_derisking': 'Derisking'
}

# event_type_map (DOPO)
event_type_map = {
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
}
```

#### 2c. `dataflow.py` — `_edit_vsm_event()`: rimuovere `'vsm_derisking': 'Derisking'` da entrambe le `event_type_map` (read-only path e edit path)

---

### Phase 3 — Mark ambiguous dead code (C7, C8)

#### 3a. `dataflow.py` — `_create_vsm_event_sheet()` `else:` branch

Aggiungere commento sopra il branch:

```python
        else:
            # NOTE: dead code — Derisking tab now uses _create_supplier_sheet().
            # This branch (event_type=None) is no longer called. Remove in next cleanup step.
            headers = [
                _("Data"), _("Nuovo Fornitore"), _("Descrizione"),
```

#### 3b. `dataflow.py` — `_populate_vsm_sheet()` `else:` branch

Aggiungere commento sopra:

```python
            else:
                # NOTE: dead code — never called for Derisking tab (uses _populate_supplier_sheet).
                # Remove in next cleanup step alongside _create_vsm_event_sheet cleanup.
                row = [
                    event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
```

---

## Relevant Files

- `services/dashboard_controller.py` — `clear_vsm_filters()` ~riga 472
- `dataflow.py` — `_VSM_STATUS_TO_TYPE` ~riga 3255; `_export_vsm_excel` ~riga 4039; `_delete_vsm_events` ~riga 2482; `open_new_event` ~riga 3618+3649; `_edit_vsm_event` ~riga 2390+2410; `_create_vsm_event_sheet` else-branch; `_populate_vsm_sheet` else-branch

**NON toccare:**
- `services/vsm_engine.py`
- `models/vsm_event.py`
- `database_manager.py`
- `ui/dialogs/vsm_event_dialog.py`
- `tests/`
- `services/kpi_*`
- `ui/kpi_window.py`

---

## Verification

1. `python3 -m unittest discover -s tests -q` → 63+ OK
2. Tab Derisking → CRUD fornitori funziona (doppio click → PotentialSupplierDialog)
3. Tab Derisking → Reset filtri (Clear button) → mostra fornitori potenziali (non VSM events corrotti)
4. Tab Derisking → Global search → non corrompe lo sheet (silenzioso, nessuna risposta — accettabile)
5. Tab Derisking → Export Excel → messaggio "non disponibile" invece di export vuoto/errato
6. Tab Saving / Cost Avoidance → invariati (KPI, search, export, CRUD)
7. KPI Derisking window → card supplier-based invariate

---

## Decisions

- `vsm_event_dialog.py` sezione Derisking: **KEEP** — potrebbe servire per editare vecchi record VSM in view Saving/CA; non causa bug
- `_create_vsm_event_sheet` else-branch: **mark as dead**, rimozione nel prossimo step
- `_populate_vsm_sheet` else-branch: **mark as dead**, rimozione nel prossimo step
- Export Derisking: messaggio "coming soon" — implementazione nel prossimo step con `created_at`
- Ricerca globale Derisking tab: disabilitata silenziosamente — scope prossimo step

---

## Limiti residui (prossimo step)

- Ricerca globale non attiva sul tab Derisking
- Export Excel tab Derisking non ancora disponibile
- `created_at` nei `potential_suppliers` ancora da introdurre
- KPI Derisking non filtrabili per tempo
- Branch dead code C7/C8 ancora presenti (marcati con nota)
