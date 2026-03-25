# STEP 4D.3 — IMPLEMENTAZIONE CRUD VSM (MODIFICA + ELIMINA)

**Data**: 25 marzo 2026  
**Scope**: Implementazione completa handler Edit e Delete VSM

---

## CONTESTO

- UI VSM già integrata nel notebook principale (Saving, Cost Avoidance, Derisking)
- Pulsante Actions già funzionante (Step 4D.1 + 4D.2)
- Menu Actions già differenziato tra RFQ e VSM
- Placeholder methods già presenti:
  - `_edit_vsm_event()` (line ~4415)
  - `_delete_vsm_events()` (line ~4445)

---

## OBIETTIVO

Implementare la logica reale di:
- **Modifica Evento VSM** (singolo)
- **Eliminazione Eventi VSM** (multipla)

---

## VINCOLI

- ❌ NON creare nuove architetture
- ❌ NON duplicare codice
- ✅ RIUTILIZZARE pattern già presenti in VSMManagementWindow
- ❌ NON toccare logica RFQ
- ✅ Modifiche minime e reversibili

---

## TASK 1 — IMPLEMENTARE `_edit_vsm_event()`

### Riferimento Pattern
**Source**: `ui/windows/vsm_management_window.py` - `on_edit_event()` (lines 297-359)

### Implementazione

**Location**: `dataflow.py` lines ~4415-4443

**Sostituire placeholder con**:

```python
    # ===========================
    # Step 4D.3: VSM CRUD Handlers (implementazione completa)
    # ===========================
    
    def _edit_vsm_event(self):
        """Handler per modifica evento VSM.
        
        Step 4D.3: Implementazione completa con VSMEventDialog.
        Pattern estratto da VSMManagementWindow.on_edit_event().
        """
        sheet, status = self.get_current_tree_and_status()
        if not status.startswith('vsm_'):
            return
        
        # Ottieni selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            messagebox.showwarning(
                _("Nessuna Selezione"),
                _("Seleziona un evento da modificare."),
                parent=self.root
            )
            return
        
        if len(selected_rows) > 1:
            messagebox.showwarning(
                _("Selezione Multipla"),
                _("Seleziona un solo evento per la modifica."),
                parent=self.root
            )
            return
        
        # Ottieni event_id e ownership
        row_idx = selected_rows[0]
        if row_idx >= len(sheet._event_metadata):
            return
        
        metadata = sheet._event_metadata[row_idx]
        event_id = metadata['event_id']
        is_mine = metadata['is_mine']
        
        # Valida ownership
        if not is_mine:
            messagebox.showerror(
                _("Operazione Non Consentita"),
                _("Puoi modificare solo i tuoi eventi VSM."),
                parent=self.root
            )
            return
        
        # Determina event_type da status
        event_type_map = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking'
        }
        event_type = event_type_map.get(status, 'Saving')
        
        # Apri dialog edit
        from ui.dialogs.vsm_event_dialog import VSMEventDialog
        
        try:
            dialog = VSMEventDialog(
                self.root,
                current_username=self.current_username,
                event_type=event_type,
                event_id=event_id  # event_id not None = modalità edit
            )
            self.root.wait_window(dialog)
            
            # Refresh se salvato
            if hasattr(dialog, 'result') and dialog.result:
                self._load_vsm_events(event_type, sheet)
                logger.info(f"Evento VSM {event_id} modificato con successo")
        
        except Exception as e:
            logger.error(f"Errore apertura dialog modifica evento VSM: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile aprire il form: {}").format(e),
                parent=self.root
            )
```

### Logica Implementata

1. **Guard clause**: verifica `status.startswith('vsm_')`
2. **Get selection**: `_get_selected_row_indices(sheet)`
3. **Validation selezione**:
   - Nessuna riga → warning "Nessuna Selezione"
   - Multiple righe → warning "Selezione Multipla"
4. **Get metadata**: `sheet._event_metadata[row_idx]`
5. **Ownership check**: `is_mine` → se False, error "Operazione Non Consentita"
6. **Status mapping**: `'vsm_saving'` → `'Saving'` (etc.)
7. **Open dialog**: `VSMEventDialog(parent=self.root, event_id=event_id)`
8. **Wait and check result**: `self.root.wait_window(dialog)`
9. **Auto-refresh**: se `dialog.result == True`, chiama `_load_vsm_events()`
10. **Exception handling**: messagebox "Errore" + log

### Adattamenti da VSMManagementWindow

| Originale | Adattato | Motivo |
|-----------|----------|--------|
| `self.sheets[self.current_event_type]` | `self.get_current_tree_and_status()` | Context detection centralizzato |
| `parent=self` | `parent=self.root` | DataFlowApp non è Toplevel |
| `self.refresh_events()` | `self._load_vsm_events(event_type, sheet)` | Metodo estratto Step 4C |
| `self.current_event_type` | `event_type_map[status]` | Status string → event_type |
| `self.refresh_callback()` | *rimosso* | Non necessario in main window |

---

## TASK 2 — IMPLEMENTARE `_delete_vsm_events()`

### Riferimento Pattern
**Source**: `ui/windows/vsm_management_window.py` - `on_delete_event()` (lines 361-430)

### Implementazione

**Location**: `dataflow.py` lines ~4445-4475

**Sostituire placeholder con**:

```python
    def _delete_vsm_events(self):
        """Handler per eliminazione eventi VSM.
        
        Step 4D.3: Implementazione completa con conferma e delete_event_and_impacts.
        Pattern estratto da VSMManagementWindow.on_delete_event().
        """
        sheet, status = self.get_current_tree_and_status()
        if not status.startswith('vsm_'):
            return
        
        # Ottieni selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            messagebox.showwarning(
                _("Nessuna Selezione"),
                _("Seleziona uno o più eventi da eliminare."),
                parent=self.root
            )
            return
        
        # Raccolta event_id e validazione ownership
        events_to_delete = []
        for row_idx in selected_rows:
            if row_idx >= len(sheet._event_metadata):
                continue
            
            metadata = sheet._event_metadata[row_idx]
            
            # Valida ownership
            if not metadata['is_mine']:
                messagebox.showerror(
                    _("Operazione Non Consentita"),
                    _("Puoi eliminare solo i tuoi eventi VSM.\nAlcuni eventi selezionati appartengono ad altri utenti."),
                    parent=self.root
                )
                return
            
            events_to_delete.append(metadata['event_id'])
        
        if not events_to_delete:
            return
        
        # Conferma eliminazione
        count = len(events_to_delete)
        if not messagebox.askyesno(
            _("Conferma Eliminazione"),
            _("Sei sicuro di voler eliminare {} evento(i) VSM?\nQuesta operazione non può essere annullata.").format(count),
            parent=self.root
        ):
            return
        
        # Determina event_type da status
        event_type_map = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking'
        }
        event_type = event_type_map.get(status, 'Saving')
        
        # Elimina eventi
        from services.vsm_persistence import delete_event_and_impacts, VSMError
        
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                for event_id in events_to_delete:
                    delete_event_and_impacts(db_manager, event_id)
            
            messagebox.showinfo(
                _("Successo"),
                _("{} evento(i) VSM eliminato(i) con successo.").format(count),
                parent=self.root
            )
            
            # Refresh
            self._load_vsm_events(event_type, sheet)
            logger.info(f"Eliminati {count} eventi VSM con successo")
        
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore eliminazione eventi VSM: {e}")
            messagebox.showerror(
                _("Errore Eliminazione"),
                _("Impossibile eliminare gli eventi:\n{}").format(e),
                parent=self.root
            )
```

### Logica Implementata

1. **Guard clause**: verifica `status.startswith('vsm_')`
2. **Get selection**: `_get_selected_row_indices(sheet)`
3. **Validation selezione**: Nessuna riga → warning "Nessuna Selezione"
4. **Loop ownership check**:
   - Per ogni riga: verifica `metadata['is_mine']`
   - Se anche uno è False → error immediato + early return
   - Accumula `events_to_delete` lista
5. **Validation lista**: se vuota → return
6. **Conferma utente**: `messagebox.askyesno()` con count
   - Se No → return (nessuna azione)
7. **Status mapping**: `'vsm_saving'` → `'Saving'` (etc.)
8. **Import lazy**: `from services.vsm_persistence import delete_event_and_impacts, VSMError`
9. **Database operation**: loop `delete_event_and_impacts()` per ogni event_id
10. **Success feedback**: messagebox "Successo" con count
11. **Auto-refresh**: `_load_vsm_events(event_type, sheet)`
12. **Exception handling**: catch `(DatabaseError, VSMError)` → messagebox "Errore Eliminazione"

### Adattamenti da VSMManagementWindow

| Originale | Adattato | Motivo |
|-----------|----------|--------|
| `self.sheets[self.current_event_type]` | `self.get_current_tree_and_status()` | Context detection centralizzato |
| `parent=self` | `parent=self.root` | DataFlowApp non è Toplevel |
| `self.refresh_events()` | `self._load_vsm_events(event_type, sheet)` | Metodo estratto Step 4C |
| `self.current_event_type` | `event_type_map[status]` | Status string → event_type |
| `self.refresh_callback()` | *rimosso* | Non necessario in main window |

---

## TASK 3 — RIUTILIZZO CODICE ESISTENTE

### Pattern Riutilizzati da VSMManagementWindow

#### 1. Validazione Selezione
```python
if not selected_rows:
    messagebox.showwarning(_("Nessuna Selezione"), ...)
    return

if len(selected_rows) > 1:  # Solo per Edit
    messagebox.showwarning(_("Selezione Multipla"), ...)
    return
```

#### 2. Ownership Check
```python
metadata = sheet._event_metadata[row_idx]
is_mine = metadata['is_mine']

if not is_mine:
    messagebox.showerror(_("Operazione Non Consentita"), ...)
    return
```

#### 3. Dialog Integration
```python
from ui.dialogs.vsm_event_dialog import VSMEventDialog

dialog = VSMEventDialog(
    self.root,
    current_username=self.current_username,
    event_type=event_type,
    event_id=event_id
)
self.root.wait_window(dialog)

if hasattr(dialog, 'result') and dialog.result:
    # Refresh
```

#### 4. Delete Confirmation
```python
count = len(events_to_delete)
if not messagebox.askyesno(
    _("Conferma Eliminazione"),
    _("Sei sicuro di voler eliminare {} evento(i) VSM?\n...").format(count),
    parent=self.root
):
    return
```

#### 5. Database Operations
```python
from services.vsm_persistence import delete_event_and_impacts, VSMError

try:
    with DatabaseManager(get_db_path()) as db_manager:
        for event_id in events_to_delete:
            delete_event_and_impacts(db_manager, event_id)
    
    messagebox.showinfo(_("Successo"), ...)
    self._load_vsm_events(event_type, sheet)

except (DatabaseError, VSMError) as e:
    logger.error(...)
    messagebox.showerror(_("Errore Eliminazione"), ...)
```

---

## TASK 4 — INVARIANTI (OBBLIGATORIO)

### ❌ NON Modificare

- ✅ `update_button_visibility()` - Logica button state invariata
- ✅ `_populate_actions_menu()` - Menu population invariata
- ✅ Logica RFQ - Nessun impatto su RFQ handlers
- ✅ Struttura notebook - Tab structure invariata
- ✅ `_load_vsm_events()` - Data loading invariato
- ✅ `_populate_vsm_sheet()` - Sheet population invariata
- ✅ `_create_vsm_event_sheet()` - UI creation invariata

### ✅ Solo Modifiche

- `_edit_vsm_event()` - Da placeholder a implementazione completa
- `_delete_vsm_events()` - Da placeholder a implementazione completa

### Nessuna Modifica a

- Import esistenti (già presenti in dataflow.py)
- Metadata structure (`sheet._event_metadata`)
- Selection handlers (già presenti Step 4D.1)
- Menu Actions structure (già presente Step 4D.2)

---

## TASK 5 — TEST MANUALE

### Checklist EDIT

- [ ] **Selezione 0 righe** → messagebox warning "Nessuna Selezione"
- [ ] **Selezione 2+ righe** → messagebox warning "Selezione Multipla"
- [ ] **Evento NON mio** → messagebox error "Operazione Non Consentita"
- [ ] **Evento mio + singolo** → VSMEventDialog si apre correttamente
- [ ] **Dialog salva modifiche** → sheet refreshata, dati aggiornati visibili
- [ ] **Dialog annulla** → sheet immutata, nessuna modifica
- [ ] **Exception/error DB** → messagebox error "Impossibile aprire il form"
- [ ] **Tab switch post-edit** → modifiche persistenti

### Checklist DELETE

- [ ] **Selezione 0 righe** → messagebox warning "Nessuna Selezione"
- [ ] **Selezione include evento NON mio** → messagebox error "Operazione Non Consentita" (blocco immediato)
- [ ] **Eventi miei** → messagebox confirmation "Sei sicuro di voler eliminare N evento(i)?"
- [ ] **Conferma No** → nessuna azione, eventi presenti
- [ ] **Conferma Si** → eventi eliminati, messagebox success "N evento(i) VSM eliminato(i)"
- [ ] **Post-delete** → sheet refreshata, eventi spariti
- [ ] **Error DB** → messagebox error "Errore Eliminazione"
- [ ] **Delete multiplo (3+ eventi)** → tutti eliminati correttamente

### Checklist OWNERSHIP

- [ ] **Ownership propria** → Edit e Delete funzionano normalmente
- [ ] **Ownership altrui** → Edit e Delete bloccati con messagebox
- [ ] **Ownership mista (multipla selezione)** → Delete bloccato se almeno uno non è mio

### Checklist INTEGRATION

- [ ] **RFQ tabs** → nessun impatto, funzionamento invariato
- [ ] **Actions button VSM** → enabled/disabled correttamente
- [ ] **Actions menu VSM** → voci Edit/Delete con state corretto
- [ ] **Tab switch** → context detection corretto (Saving/Cost Avoidance/Derisking)
- [ ] **Refresh automatico** → dati aggiornati senza reload manuale

---

## VERIFICHE PRE-COMMIT

### Code Review Checklist

- [ ] Placeholder methods completamente sostituiti
- [ ] Imports lazy presenti (`from ui.dialogs.vsm_event_dialog import...`)
- [ ] Exception handling completo con log
- [ ] Messagebox parent corretti (`parent=self.root`)
- [ ] Status mapping completo (Saving/Cost Avoidance/Derisking)
- [ ] Refresh chiamato con parametri corretti (`event_type`, `sheet`)
- [ ] Nessuna modifica a metodi RFQ
- [ ] Nessuna modifica a update_button_visibility()
- [ ] Nessuna modifica a _populate_actions_menu()

### Syntax Check

```bash
python3 -m py_compile dataflow.py
```

### Runtime Test

```bash
cd /home/guido/Repository/vsm
source .venvLinux/bin/activate
python dataflow.py
```

---

## ROLLBACK PLAN

### Se problemi dopo implementazione

**Opzione 1**: Git revert
```bash
git checkout HEAD~1 -- dataflow.py
```

**Opzione 2**: Ripristino placeholder
Sostituire implementazione completa con placeholder originali (Step 4D.2)

---

## PROSSIMI STEP (Post-4D.3)

### Step 4D.4 (Opzionale)
- Double-click binding per Edit
- Binding `<Double-Button-1>` su VSM sheets

### Step 4E
- Export VSM integration
- Rimuovere guard in `mega_export_excel()`
- Implementare `_export_vsm_to_excel(status)`

### Step 4F
- KPI Dashboard
- Sostituire placeholder `on_kpi_click()`
- Aggregazioni mensili/trimestrali/annuali

### Step Cleanup
- Eliminare `ui/windows/vsm_management_window.py`
- Verificare nessun riferimento attivo
- Update documentazione

---

## SUMMARY

### Modifiche Step 4D.3

| File | Metodo | Righe | Azione |
|------|--------|-------|--------|
| `dataflow.py` | `_edit_vsm_event()` | ~4415-4443 | Sostituire placeholder con implementazione completa |
| `dataflow.py` | `_delete_vsm_events()` | ~4445-4475 | Sostituire placeholder con implementazione completa |

### Pattern Strategy

- ✅ **100% riutilizzo** da VSMManagementWindow
- ✅ **Adattamenti minimali** per context DataFlowApp
- ✅ **Nessuna nuova architettura**
- ✅ **Nessun impatto RFQ**
- ✅ **Reversibile** (git revert safe)

### Expected Behavior Post-Step 4D.3

**VSM Edit**: Singolo evento, ownership check, dialog integration, auto-refresh  
**VSM Delete**: Multipli eventi, ownership check, confirmation, auto-refresh  
**RFQ**: Comportamento invariato al 100%

---

## NOTES

### Import Dependency
- `VSMEventDialog` - lazy import in metodo (evita circular dependency)
- `delete_event_and_impacts` - lazy import in metodo
- `VSMError` - lazy import in metodo
- `DatabaseManager`, `get_db_path()` - già importati in dataflow.py header

### Metadata Structure
```python
sheet._event_metadata = [
    {
        'event_id': int,
        'username': str,
        'is_mine': bool
    },
    ...
]
```

### Status → Event Type Mapping
```python
{
    'vsm_saving': 'Saving',
    'vsm_cost_avoidance': 'Cost Avoidance',
    'vsm_derisking': 'Derisking'
}
```
