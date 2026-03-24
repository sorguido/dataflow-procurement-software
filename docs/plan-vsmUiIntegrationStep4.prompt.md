# Plan: VSM UI Integration (Step 4) — Conservative First Version

**TL;DR**: Integro la UI VSM nella Main Dashboard di DataFlow aggiungendo una nuova tab "VSM" al Notebook esistente. Dentro questa tab, creo un sotto-notebook con 3 tab (Saving, Cost Avoidance, Derisking), ciascuna con tabella eventi + toolbar. Implemento Create/Edit/Delete usando dialog window con form dinamico. Zero regressioni RFQ, approccio minimale e conservativo.

---

## Steps

### Phase 1: Core Infrastructure (Foundation)

1. **Create VSM management module** `ui/windows/vsm_management_window.py`
   - Contenitore principale per la tab VSM
   - Sotto-notebook con 3 tab (Saving, Cost Avoidance, Derisking)
   - Toolbar con pulsanti: "+ New Event", "Edit", "Delete", "KPI"
   - tksheet per visualizzazione eventi (pattern identico a RFQ)
   - *depends on: nulla, può iniziare subito*

2. **Create VSM event dialog** `ui/dialogs/vsm_event_dialog.py`
   - tk.Toplevel window per Create/Edit evento VSM
   - Form dinamico che mostra solo campi pertinenti al tipo evento
   - DateEntry per event_date, Combobox per event_type/action
   - Validazione UI minima (campi obbligatori, range percent_realizzo)
   - *depends on: step 1 completato*

3. **Integrate VSM tab into Main Dashboard** in `dataflow.py`
   - Aggiungere tab "VSM" al self.notebook esistente (dopo tab archiviate)
   - Istanziare VSMManagementWindow dentro la tab
   - Bind tab change handler per refresh quando si switcha su VSM
   - *depends on: step 1 completato*

### Phase 2: CRUD Operations (Business Logic)

4. **Implement Create Event** in `vsm_event_dialog.py`
   - Validare input form
   - Creare VSMEvent dataclass da form
   - Chiamare `save_event_with_impacts(db_manager, event)`
   - Return event_id al chiamante per refresh
   - *depends on: step 2*

5. **Implement Read/List Events** in `vsm_management_window.py`
   - Metodo `refresh_events()` che carica eventi per tab corrente
   - Usare `db_manager.get_all_vsm_events()` filtrato per event_type
   - Popolare tksheet con dati evento (data, tipo, descrizione, valore, username)
   - Aggiungere metadata per tracking event_id e ownership
   - *depends on: step 1*

6. **Implement Edit Event** in `vsm_management_window.py`
   - Verificare selezione singola su tksheet
   - Caricare evento esistente da DB via event_id
   - Aprire VSMEventDialog in modalità edit con dati precaricati
   - Chiamare `update_event_with_impacts(db_manager, event)`
   - Refresh lista dopo salvataggio
   - *depends on: steps 2, 5*

7. **Implement Delete Event** in `vsm_management_window.py`
   - Verificare selezione (singola o multipla)
   - Validare ownership eventi selezionati
   - Conferma utente con messagebox.askyesno
   - Chiamare `delete_event_and_impacts(db_manager, event_id)`
   - Refresh lista dopo eliminazione
   - *depends on: step 5*

### Phase 3: UI Polish (User Experience)

8. **Add KPI button placeholder** in `vsm_management_window.py`
   - Creare pulsante "📊 KPI" nella toolbar
   - Al click mostrare messagebox.showinfo("KPI Analysis", "Feature coming in next step")
   - Preparare hook per modulo futuro senza implementare logica
   - *parallel with any step in Phase 2*

9. **Add validations and error handling**
   - Gestire DatabaseError con messagebox.showerror
   - Gestire VSMError con messagebox appropriati
   - Validazione form: date obbligatoria, campi economici per Saving/Cost Avoidance
   - Disabilitare pulsanti Edit/Delete se nessuna selezione
   - *parallel with Phase 2 steps*

10. **Testing and verification**
    - Test manuale: Create Saving one-shot → compare in lista
    - Test manuale: Create Saving ripetitivo → campo num_mesi visibile
    - Test manuale: Edit evento → dati caricati correttamente
    - Test manuale: Delete evento → scompare da lista
    - Test regressione: tab RFQ funziona invariata
    - *depends on: all previous steps*

---

## Relevant Files

### Files to Create:

- `ui/windows/vsm_management_window.py` — Main VSM tab container with sub-notebook, tables, toolbar
  - Class `VSMManagementWindow(ttk.Frame)` with 3 sub-tabs
  - Methods: `refresh_events()`, `on_new_event()`, `on_edit_event()`, `on_delete_event()`, `on_kpi_click()`
  - tksheet configuration for displaying VSM events (columns: Date, Type, Description, Value, User, Repetitive)

- `ui/dialogs/vsm_event_dialog.py` — Create/Edit dialog for VSM events
  - Class `VSMEventDialog(tk.Toplevel)` with dynamic form
  - Methods: `_build_form()`, `_on_event_type_changed()`, `_validate_and_save()`, `_load_event_data(event_id)`
  - Dynamic field visibility based on selected event_type

### Files to Modify:

- `dataflow.py` (minimal changes, ~10 lines)
  - Line ~3680: Add `self.tab_vsm = VSMManagementWindow(self.notebook, self.current_username, self.refresh_vsm_callback)`
  - Line ~3682: `self.notebook.add(self.tab_vsm, text=_("VSM"))`
  - Add method `refresh_vsm_callback()` to refresh VSM tab on demand

### Files to Use (Read-Only):

- `models/vsm_event.py` — VSMEvent dataclass with all fields
- `models/vsm_impact.py` — VSMImpact dataclass (not exposed in UI)
- `services/vsm_persistence.py` — Functions: `save_event_with_impacts()`, `update_event_with_impacts()`, `delete_event_and_impacts()`, `get_event_with_impacts()`
- `services/vsm_engine.py` — Backend logic for impact generation (UI doesn't call this directly)
- `database_manager.py` — Methods: `get_all_vsm_events()`, `get_vsm_event_by_id()`

---

## Verification

1. **Create Saving one-shot** → event visible in Saving tab, editable, deletable ✓
2. **Create Saving repetitive** → num_mesi_ripetizione field visible ✓
3. **Create Cost Avoidance** → form shows importo_richiesto_iniziale instead of importo_bdg ✓
4. **Create Derisking** → no economic fields forced (only description/reference) ✓
5. **Edit existing event** → form pre-populated, update works, list refreshes ✓
6. **Delete event** → disappears from list, no UI errors ✓
7. **RFQ tab regression test** → RFQ create/edit/delete still works ✓
8. **Multi-user isolation** → only own events editable/deletable ✓
9. **KPI button** → shows placeholder message ✓

---

## Decisions

### Architecture Choice: Nested Notebook Pattern

**Decision**: Add single "VSM" tab to main notebook, with nested sub-notebook inside for Saving/Cost Avoidance/Derisking.

**Rationale**: 
- Conservative approach (only +1 main tab, not +3)
- Clear separation between RFQ and VSM workflows
- Sub-tabs allow filtering by event_type without complex UI logic
- Consistent with user request for dedicated tabs per type

### Form Complexity: Minimal Essential Fields Only

**Decision**: Expose only ~8-10 core fields in Create/Edit dialog.

**Fields selected**:
- Common: event_date, event_type, action, description, reference, buyer, percent_realizzo
- Economic (conditional): importo_bdg, importo_negoziato (Saving) | importo_richiesto_iniziale (Cost Avoidance)
- Distribution: opex_ripetitivo (checkbox), num_mesi_ripetizione (only if opex_ripetitivo=True)
- Derived automatically: username (from session), created_at/updated_at (from DB)

**Excluded fields** (defer to future steps):
- quantita_annua, spending_annuo (not critical for MVP)
- driver, giorni_pagamento_attuali, giorni_pagamento_negoziati (payment terms, defer to enhancement)
- note (can use existing description field)

### Table Columns: Event-Centric View

**Decision**: Show events in table, NOT impacts.

**Columns**:
1. Data Evento (event_date)
2. Tipo (event_type)
3. Azione (action)
4. Descrizione/Riferimento (truncated)
5. Valore Teorico (calculated from importo fields)
6. Realizzo % (percent_realizzo)
7. Ripetitivo (Yes/No icon)
8. Username

### Ownership Validation: Strict by Default

**Decision**: User can only edit/delete own events (same as RFQ pattern).

**Implementation**: Store username in event, check `event.username == current_username` before allowing edit/delete.

---

## Further Considerations (Deferred to Future Steps)

1. **KPI Analysis Module** — Currently placeholder button. Next step should implement:
   - Aggregated metrics by period (monthly/quarterly/annual)
   - Saving vs Cost Avoidance comparison
   - Realized vs Theoretical value charts
   - Export to Excel/CSV

2. **Advanced Filtering** — Current implementation shows all events in tab. Future:
   - Date range filter
   - Buyer filter
   - Reference search
   - Realized value range filter

3. **Bulk Operations** — Current: single or multi-delete only. Future:
   - Bulk edit percent_realizzo
   - Bulk archive/restore
   - Duplicate event with modifications

4. **Impact Visualization** — Current: impacts completely hidden. Future:
   - Drill-down view: event → impacts table
   - Monthly timeline chart for repetitive events
   - Pro-rata coefficient visualization

5. **Form Enhancements** — Current: minimal 8-10 fields. Future:
   - Add remaining fields (quantita_annua, spending_annuo, driver, payment terms)
   - Field dependencies (e.g., driver=Pagamenti → show giorni_pagamento fields)
   - Multi-step wizard for complex events

---

## Notes

- **Zero Breaking Changes**: New code in separate modules, existing RFQ code untouched
- **Database Schema**: Already exists from Step 1-2, no migrations needed
- **Testing**: Manual testing protocol defined, no automated UI tests in scope
- **Localization**: Use existing `_()` wrapper for all strings
- **Icons**: Use emoji in button text (consistent with existing DataFlow style: "📄", "🗑", "📊")
- **Performance**: Pre-filter events by event_type at DB level (avoid loading all events for all tabs)

---

## Implementation Risk Assessment

**Low Risk**:
- New module creation (no conflicts with existing code)
- Using proven patterns (tksheet, tk.Toplevel, ttk.Notebook)
- Backend already tested (15/15 tests pass)

**Medium Risk**:
- Dynamic form complexity (conditional field visibility)
- Mitigation: Start with static form, add dynamics incrementally

**Zero Risk**:
- RFQ regression (no modifications to RFQ code paths)
- Data corruption (backend has atomic transactions verified)

**Dependencies**: Nessuna dipendenza esterna nuova richiesta.

---

**Author**: Plan Mode Agent  
**Date**: 24 marzo 2026  
**Status**: Ready for Implementation
