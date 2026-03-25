# STEP CLEANUP 1 — ANALISI CODICE VSM LEGACY

**Data Analisi**: 25 marzo 2026  
**Scope**: Identificazione codice VSM non più utilizzato post-integrazione Step 4A-4D.2

---

## CONTESTO

- La UI VSM è stata completamente integrata nel notebook principale (Saving, Cost Avoidance, Derisking)
- Le sheet VSM sono gestite in dataflow.py
- Il menu Actions è stato integrato
- I dati VSM sono caricati direttamente in dataflow.py

---

## FILE ANALIZZATO

**`ui/windows/vsm_management_window.py`** (444 righe)

---

## STATO ATTUALE

### Import Statement
```python
# dataflow.py line 98 (COMMENTATO)
# from ui.windows.vsm_management_window import VSMManagementWindow
```

### Export Module
```python
# ui/windows/__init__.py
# VSMManagementWindow NON presente in __all__
```

### Istanziazione
❌ **NESSUNA** istanza attiva trovata in dataflow.py o altri moduli

### Riferimenti Attivi
❌ **NESSUNO** - Solo riferimenti in:
- Commenti (`# Step 4B: estratto da VSMManagementWindow...`)
- Documentazione (`.md` files in `docs/`)

---

## A) CODICE SAFE DELETE (Basso Rischio)

### Classe Principale
- ✅ `VSMManagementWindow(ttk.Frame)` - linee 29-450
  - **Motivo**: NON più istanziata, import commentato, non in `__all__`
  - **Logica estratta**: Tutto il necessario migrato in dataflow.py

### Metodi UI (Toolbar Interna)
- ✅ `_build_ui()` - linee 57-124
  - Toolbar locale con 5 pulsanti (Nuovo/Modifica/Elimina/KPI/Aggiorna)
  - **Sostituito da**: Menu Actions principale + toolbar globale
  
- ✅ `vsm_notebook` (nested notebook) - linee 106-124
  - Sub-notebook con 3 tab (Saving/Cost Avoidance/Derisking)
  - **Sostituito da**: 3 tab dirette nel notebook principale

### Metodi Sheet Creation
- ✅ `_create_event_sheet()` - linee 126-182
  - **GIÀ ESTRATTO**: `_create_vsm_event_sheet()` in dataflow.py (linee 4275-4335)
  - Logica identica, solo parametri adattati

### Metodi Data Loading
- ✅ `refresh_events()` - linee 201-227
  - **GIÀ ESTRATTO**: `_load_vsm_events()` in dataflow.py (linee 4327-4357)
  
- ✅ `_populate_sheet()` - linee 229-267
  - **GIÀ ESTRATTO**: `_populate_vsm_sheet()` in dataflow.py (linee 4370-4410)

### Metodi Gestione Tab
- ✅ `_on_tab_changed()` - linee 184-192
  - **Non necessario**: Tab gestite direttamente dal notebook principale
  
- ✅ `_on_sheet_double_click()` - linee 194-196
  - **Da re-implementare**: In Step 4D.3+ (binding double-click)
  
- ✅ `_update_buttons_state()` - linee 198-200
  - **Sostituito da**: `update_button_visibility()` in dataflow.py

### Handler CRUD (Toolbar Locale)
- ✅ `on_new_event()` - linee 269-295
  - **Da migrare**: In Step 4D.3+ come `_new_vsm_event()` toolbar
  
- ✅ `on_edit_event()` - linee 297-359
  - **Placeholder esistente**: `_edit_vsm_event()` in dataflow.py (linee 4415-4438)
  
- ✅ `on_delete_event()` - linee 361-430
  - **Placeholder esistente**: `_delete_vsm_events()` in dataflow.py (linee 4440-4471)

### Handler KPI
- ✅ `on_kpi_click()` - linee 432-444
  - **Placeholder**: Funzionalità futura (Step 4F)

---

## B) CODICE POTENZIALMENTE RIUTILIZZABILE

### Logica CRUD Completa

Già estratto in dataflow.py come placeholder (Step 4D.2), da completare in Step 4D.3:

```python
_edit_vsm_event()     # Line 4415 - Solo retrieval event_id
_delete_vsm_events()  # Line 4440 - Solo retrieval event_ids
```

### Pattern da Riutilizzare da VSMManagementWindow

#### 1. Validation Ownership (linee 328-335, 387-395)
```python
if not is_mine:
    messagebox.showerror(
        _("Operazione Non Consentita"),
        _("Puoi modificare solo i tuoi eventi VSM."),
        parent=self
    )
    return
```

#### 2. Dialog VSMEventDialog Integration (linee 272-283, 341-354)
```python
from ui.dialogs.vsm_event_dialog import VSMEventDialog

dialog = VSMEventDialog(
    self,
    current_username=self.current_username,
    event_type=self.current_event_type,
    event_id=event_id  # None per create, int per edit
)
self.wait_window(dialog)

if hasattr(dialog, 'result') and dialog.result:
    self.refresh_events()
    if self.refresh_callback:
        self.refresh_callback()
```

#### 3. Delete Confirmation (linee 405-411)
```python
count = len(events_to_delete)
if not messagebox.askyesno(
    _("Conferma Eliminazione"),
    _("Sei sicuro di voler eliminare {} evento(i) VSM?\nQuesta operazione non può essere annullata.").format(count),
    parent=self
):
    return
```

#### 4. Database Operations (linee 413-420)
```python
from services.vsm_persistence import delete_event_and_impacts

with DatabaseManager(get_db_path()) as db_manager:
    for event_id in events_to_delete:
        delete_event_and_impacts(db_manager, event_id)

messagebox.showinfo(
    _("Successo"),
    _("{} evento(i) VSM eliminato(i) con successo.").format(count),
    parent=self
)

# Refresh
self.refresh_events()
```

---

## C) DIPENDENZE ANCORA ATTIVE (NON Eliminabili)

### Moduli VSM Condivisi
✅ **KEEP** - Usati da VSMEventDialog e altri:

#### 1. `models/vsm_event.py`
- Classe `VSMEvent`
- **Usato da**: database_manager.py, dataflow.py, vsm_event_dialog.py, vsm_persistence.py

#### 2. `models/vsm_impact.py`
- Classe `VSMImpact`
- **Usato da**: database_manager.py, vsm_engine.py, vsm_persistence.py

#### 3. `services/vsm_engine.py`
- `generate_impacts_for_event()`
- **Usato da**: vsm_persistence.py

#### 4. `services/vsm_persistence.py`
- `save_event_with_impacts()`
- `update_event_with_impacts()`
- `delete_event_and_impacts()`
- `get_event_with_impacts()`
- **Usato da**: VSMEventDialog (ancora attivo)

#### 5. `ui/dialogs/vsm_event_dialog.py`
- Classe `VSMEventDialog`
- **USATO ATTIVAMENTE** da placeholder handlers in dataflow.py (Step 4D.3)

#### 6. `database_manager.py` (Metodi VSM)
- `get_all_vsm_events(username)`
- `insert_vsm_event(event)`
- `update_vsm_event(event)`
- `delete_vsm_event(event_id)`
- `get_vsm_event_by_id(event_id)`
- `insert_vsm_impact(impact)`
- `get_vsm_impacts_for_event(event_id)`
- `delete_vsm_impacts_for_event(event_id)`
- **Usato da**: dataflow.py, vsm_persistence.py, vsm_event_dialog.py

---

## SUMMARY REPORT

### File Eliminabili
| File | Righe | Livello Rischio | Note |
|------|-------|-----------------|------|
| `ui/windows/vsm_management_window.py` | 444 | 🟢 **BASSO** | NON istanziato, import commentato, logica estratta |

### Classi Eliminabili
| Classe | Motivo | Sostituzione |
|--------|--------|--------------|
| `VSMManagementWindow` | Import commentato, NON referenziata | Logica integrata in `DataFlowApp` |

### Metodi Già Estratti (Duplicati Safe Delete)
| Metodo Originale | Nuovo Metodo | Location |
|------------------|--------------|----------|
| `_create_event_sheet()` | `_create_vsm_event_sheet()` | dataflow.py:4275 |
| `refresh_events()` | `_load_vsm_events()` | dataflow.py:4327 |
| `_populate_sheet()` | `_populate_vsm_sheet()` | dataflow.py:4370 |

### Import Inutilizzati (Da Rimuovere)
```python
# dataflow.py line 98 - GIÀ COMMENTATO
# from ui.windows.vsm_management_window import VSMManagementWindow
```

### Dipendenze Attive (Keep)
- ✅ `models/vsm_event.py`
- ✅ `models/vsm_impact.py`
- ✅ `services/vsm_engine.py`
- ✅ `services/vsm_persistence.py`
- ✅ `ui/dialogs/vsm_event_dialog.py`
- ✅ `database_manager.py` (metodi VSM)

---

## RACCOMANDAZIONE FINALE

### Livello Rischio: 🟢 BASSO

**VSMManagementWindow è codice morto al 100%**:
- ❌ NON istanziato
- ❌ Import commentato
- ❌ NON in `__all__`
- ❌ Nessun riferimento attivo
- ✅ Logica critica già estratta
- ✅ Test NON dipendono da esso

### Strategia di Eliminazione

#### Pre-requisiti (Prima di eliminare)
1. **Step 4D.3 Completo**: Implementare handler CRUD completi
   - `_edit_vsm_event()` - Integrare logica dialog da VSMManagementWindow.on_edit_event()
   - `_delete_vsm_events()` - Integrare logica conferma + DB da VSMManagementWindow.on_delete_event()
   - `_new_vsm_event()` - Nuovo handler toolbar (riutilizzare VSMManagementWindow.on_new_event())

2. **Test Manuale**: Verificare funzionalità VSM
   - ✅ Visualizzazione dati nelle 3 tab
   - ✅ Selezione righe → Actions button enabled
   - ✅ Edit singolo evento (ownership valida)
   - ✅ Delete multiplo (ownership valida)
   - ✅ Ownership validation funzionante

#### Post-Step 4D.3 (Safe Delete)

**Azione 1**: Eliminare file
```bash
rm ui/windows/vsm_management_window.py
```

**Azione 2**: Rimuovere import commentato (se desiderato)
```python
# dataflow.py line 98
# RIMUOVERE: # from ui.windows.vsm_management_window import VSMManagementWindow
```

**Azione 3**: Update documentazione
- Aggiornare `docs/VSM_MANAGEMENT_WINDOW_ANALYSIS.md` con nota deprecazione
- Eventualmente archiviare analisi storica

#### Rollback Plan (Se problemi)
- File ancora presente in Git history
- Recuperabile con: `git checkout HEAD~1 -- ui/windows/vsm_management_window.py`

---

## VERIFICHE PRE-ELIMINAZIONE

### Checklist Sicurezza
- [ ] Step 4D.3 completato (CRUD handlers funzionanti)
- [ ] Test manuale Edit VSM superato
- [ ] Test manuale Delete VSM superato
- [ ] Test manuale ownership validation superato
- [ ] Nessun riferimento diretto a VSMManagementWindow in codice attivo
- [ ] Import commentato o rimosso
- [ ] Commit Step 4D.3 effettuato (safety net)

### Comando Verifica Riferimenti
```bash
# Verifica nessun import attivo
grep -r "from.*vsm_management_window" --include="*.py" --exclude-dir=__pycache__ .
grep -r "import.*VSMManagementWindow" --include="*.py" --exclude-dir=__pycache__ .

# Verifica nessuna istanza
grep -r "VSMManagementWindow(" --include="*.py" --exclude-dir=__pycache__ .
```

**Output Atteso**: Solo match in:
- `ui/windows/vsm_management_window.py` (il file stesso)
- File `.md` in `docs/` (documentazione)
- Commenti in dataflow.py

---

## NOTE FINALI

### Perché NON Eliminare Subito
1. **Riferimento Pattern**: Utile durante implementazione Step 4D.3
2. **Documentazione Vivente**: Codice esprime pattern meglio di commenti
3. **Safety Net**: Git history non immediato come file fisico

### Quando Eliminare
**Momento Ideale**: Immediatamente dopo completamento e test Step 4D.3

**Trigger**: Quando tutti i placeholder VSM CRUD sono sostituiti con implementazioni complete che riutilizzano i pattern di VSMManagementWindow

---

## TIMELINE

- ✅ **Step 4A**: Tab VSM integrate (completato)
- ✅ **Step 4B**: UI sheet estratta (completato)
- ✅ **Step 4C**: Data loading estratto (completato)
- ✅ **Step 4D.1**: Button state (completato)
- ✅ **Step 4D.2**: Menu Actions VSM (completato)
- 🔄 **Step 4D.3**: CRUD handlers completi (in corso)
- ⏳ **Step Cleanup**: Eliminazione VSMManagementWindow (post-4D.3)
- ⏳ **Step 4E**: Export VSM integration
- ⏳ **Step 4F**: KPI Dashboard
