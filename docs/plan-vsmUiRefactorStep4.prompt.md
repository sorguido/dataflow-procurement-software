# Plan: VSM Step 4 UI Refactor - Integrazione Completa Context-Aware

## TL;DR
Eliminare duplicazioni UI e allineare completamente VSM al modello UX di DataFlow. Trasformare tab VSM con sub-notebook in 3 tab diretti (Saving, Cost Avoidance, Derisking) allo stesso livello di RFQ. Implementare controlli globali context-aware che cambiano comportamento in base al tab attivo. Semplificare form eventi rimuovendo campi ridondanti (Reference, Buyer editabile, dropdown Tipo Evento).

---

## Architettura Esistente (Discovery Completata)

### Toolbar Globale DataFlow
- Pulsanti: `[New RFQ] [Actions ▼] [Export Excel]` + settings/help buttons
- **New RFQ**: `open_new_request_window()` apre ViewRequestWindow
- **Actions**: menu dinamico popolato da `_populate_actions_menu(status, can_delete, ...)` 
  - RFQ Attive: Elimina, Duplica, Archivia
  - RFQ Archiviate: Elimina, Duplica, Riattiva
- **Export Excel**: `mega_export_excel()` context-aware per RFQ (Attive/Archiviate)

### Tab Detection
- `notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)`
- `get_current_tree_and_status()` → ritorna `(sheet, status)` basato su `notebook.index(notebook.select())`
- Attualmente solo per tab 0 (Attive) vs tab 1 (Archiviate)

### RFQ Sheet Pattern
- tksheet widgets con double-click → `on_sheet_double_click()` → apre ViewRequestWindow
- Metadata: `._sheet_rows_metadata` list con `{'is_mine': bool, 'source_file': path}`
- Ownership validation prima di edit/delete

### VSM Attuale (da rifattorizzare)
- `VSMManagementWindow(ttk.Frame)` con toolbar locale + sub-notebook a 3 tab
- Toolbar: Nuovo Evento, Modifica, Elimina, KPI, Aggiorna
- Sub-notebook: Saving, Cost Avoidance, Derisking
- Doppio click: `_on_sheet_double_click()` → `on_edit_event()`
- Metadata: `._event_metadata` list con `(event_id, username, is_mine)`

---

## Decisions (Alignment Completato ✓)

**Q1 - Metadata VSM**: Includere `source_file` per coerenza architetturale con RFQ (anche se non usato attivamente in questo step)

**Q2 - Filtri Globali**: Nascondere completamente i filtri collapsibili quando tab VSM è attivo (UI pulita, filtri VSM custom in step futuro)

**Q3 - Actions Menu VSM**: Solo "Delete" + "Duplicate" (no Archive - richiederebbe campo status nel DB)

**Q4 - Export Excel VSM**: Tutte le colonne - Data, Tipo, Azione, Descrizione, User, Valore Teorico, Realizzo%, Ripetitivo, Driver

**Q5 - KPI Button**: Sempre visibile e attivo, comportamento dinamico basato su tab attivo (placeholder KPI RFQ vs KPI VSM)

---

## Steps (Design Dettagliato)

### **Phase 1: Ristrutturazione Tab VSM** (Breaking changes - modifica dataflow.py)

**1.1 Rimuovere tab VSM container**
- In `dataflow.py` __init__: rimuovere creazione `self.tab_vsm = VSMManagementWindow(...)`
- Rimuovere `self.notebook.add(self.tab_vsm, text=_("VSM"))`

**1.2 Creare 3 tab frame diretti**
```python
self.tab_saving = ttk.Frame(self.notebook)
self.tab_cost_avoidance = ttk.Frame(self.notebook)
self.tab_derisking = ttk.Frame(self.notebook)
```

**1.3 Creare sheet per ogni tab VSM**
- Estrarre logica sheet creation da `VSMManagementWindow._create_event_sheet()`
- Creare metodo `create_vsm_event_sheet(parent, event_type)` in dataflow.py
- Headers: `["Data", "Tipo", "Azione", "Descrizione", "Valore Teorico", "Realizzo %", "Ripetitivo", "Utente"]`
- Column widths: `[100, 120, 120, 300, 120, 80, 80, 120]`
- Readonly: tutte le colonne
- Double-click binding: `self.on_sheet_double_click`
- Metadata storage: `._vsm_event_metadata` list con `{'event_id': int, 'username': str, 'is_mine': bool, 'source_file': str}`

**1.4 Aggiungere tab al notebook**
```python
self.notebook.add(self.tab_saving, text=_("Saving"))
self.notebook.add(self.tab_cost_avoidance, text=_("Cost Avoidance"))
self.notebook.add(self.tab_derisking, text=_("Derisking"))
```

**1.5 Creare sheet widgets**
```python
self.sheet_saving = self.create_vsm_event_sheet(self.tab_saving, "Saving")
self.sheet_cost_avoidance = self.create_vsm_event_sheet(self.tab_cost_avoidance, "Cost Avoidance")
self.sheet_derisking = self.create_vsm_event_sheet(self.tab_derisking, "Derisking")
```

**1.6 Implementare metodo refresh VSM**
- Creare `refresh_vsm_events(event_type)` in dataflow.py
- Logica: chiamare `db_manager.get_all_vsm_events(username=self.current_username)`
- Filtrare per event_type
- Popolare sheet corrispondente con `_populate_vsm_sheet(sheet, events, event_type)`

**Dependencies**: Nessuna (task fondante)

---

### **Phase 2: Context-Aware Tab Detection** (Modifica get_current_tree_and_status)

**2.1 Estendere get_current_tree_and_status()**
- Ritornare tupla estesa: `(sheet, status_or_event_type, tab_category)`
- tab_category: `'rfq'` o `'vsm'`
- Esempio per tab 0 (Attive): `(self.tree_attive, 'attiva', 'rfq')`
- Esempio per tab 2 (Saving): `(self.sheet_saving, 'Saving', 'vsm')`

**2.2 Aggiornare on_tab_changed()**
- Chiamare `update_button_visibility()` (esistente)
- Chiamare nuovo `update_filters_visibility()` → nasconde filtri se tab_category=='vsm'
- Chiamare `clear_selection()` (esistente)

**Dependencies**: Phase 1 completata (tab VSM esistenti)

---

### **Phase 3: Context-Aware Controls** (Modifica toolbar buttons)

**3.1 Pulsante "New RFQ" → dinamico "New Event"**
- In __init__: salvare riferimento `self.btn_new = ttk.Button(...)`
- In `update_button_visibility()`:
  - Se tab_category=='rfq': `self.btn_new.config(text=_("➕ Nuova RdO"))`
  - Se tab_category=='vsm': `self.btn_new.config(text=_("➕ Nuovo Evento"))`

**3.2 Modificare open_new_request_window()**
- Rilevare contesto: `sheet, status_or_event_type, tab_category = self.get_current_tree_and_status()`
- Se tab_category=='rfq': comportamento esistente (apre ViewRequestWindow)
- Se tab_category=='vsm': aprire `VSMEventDialog(self.root, self.current_username, event_type=status_or_event_type)`
- Dopo dialog chiuso: se `dialog.result`: chiamare `refresh_vsm_events(event_type)`

**3.3 Estendere _populate_actions_menu()**
- Rilevare contesto: `sheet, status_or_event_type, tab_category = self.get_current_tree_and_status()`
- Se tab_category=='rfq': comportamento esistente
- Se tab_category=='vsm':
  - Menu: `[Elimina] [Duplica]`
  - Elimina: abilitato se `has_selection AND all_mine`
  - Duplica: abilitato se `exactly_1_selection AND is_mine`

**3.4 Implementare delete_selected_vsm_event()**
- Legge selezione da sheet VSM corrente
- Valida ownership tramite metadata `is_mine`
- Conferma con `messagebox.askyesno`
- Chiama `vsm_persistence.delete_event_and_impacts(db_manager, event_id)` per ogni evento
- Refresh: `refresh_vsm_events(current_event_type)`

**3.5 Implementare duplicate_selected_vsm_event()**
- Legge evento selezionato tramite `get_event_with_impacts(db_manager, event_id)`
- Crea nuovo VSMEvent con:
  - `id=None` (nuovo)
  - `event_date=oggi`
  - `username=self.current_username`
  - Copia tutti gli altri campi (importo, descrizione, etc.)
- Salva con `save_event_with_impacts(db_manager, new_event)`
- Refresh: `refresh_vsm_events(current_event_type)`

**Dependencies**: Phase 2 completata (tab detection esteso)

---

### **Phase 4: Export Excel Context-Aware** (Modifica mega_export_excel)

**4.1 Rilevare contesto in mega_export_excel()**
- All'inizio metodo: `sheet, status_or_event_type, tab_category = self.get_current_tree_and_status()`
- Se tab_category=='rfq': comportamento esistente (invariato)
- Se tab_category=='vsm': chiamare nuovo `export_vsm_events_to_excel(event_type)`

**4.2 Implementare export_vsm_events_to_excel(event_type)**
- Carica tutti eventi: `db_manager.get_all_vsm_events(username=self.current_username)`
- Filtra per event_type
- Crea Excel con colonne: Data, Tipo, Azione, Descrizione, User, Valore Teorico, Realizzo%, Ripetitivo, Driver
- File naming: `VSM_<event_type>_<timestamp>.xlsx`
- Salvare con dialog: `filedialog.asksaveasfilename(defaultextension=".xlsx")`
- Success message

**Dependencies**: Phase 3 completata (context detection funzionante)

---

### **Phase 5: Double-Click Unificato** (Modifica on_sheet_double_click)

**5.1 Estendere on_sheet_double_click(sheet, event)**
- Rilevare tipo sheet: controllare se `sheet in [self.tree_attive, self.tree_archiviate]` → RFQ
- Altrimenti: se `sheet in [self.sheet_saving, self.sheet_cost_avoidance, self.sheet_derisking]` → VSM
- Se RFQ: comportamento esistente (apre ViewRequestWindow)
- Se VSM: aprire VSM edit dialog

**5.2 Handler VSM double-click**
- Leggi row_index da selezione
- Estrai metadata: `metadata = sheet._vsm_event_metadata[row_index]`
- Se not `metadata['is_mine']`: apri in readonly o mostra warning
- Altrimenti: apri `VSMEventDialog(parent, current_username, event_type=event_type, event_id=metadata['event_id'])`
- Dopo dialog chiuso: se `dialog.result`: refresh sheet

**Dependencies**: Phase 1 completata (sheet VSM esistenti con metadata)

---

### **Phase 6: Pulsante KPI** (Aggiunta pulsante globale)

**6.1 Aggiungere pulsante KPI in toolbar**
- In __init__ dopo "Export Excel": 
  ```python
  self.btn_kpi = ttk.Button(frame_top, text=_("📊 KPI"), command=self.open_kpi_window)
  self.btn_kpi.pack(side="left", padx=5)
  ```

**6.2 Implementare open_kpi_window()**
- Rilevare contesto: `sheet, status_or_event_type, tab_category = self.get_current_tree_and_status()`
- Se tab_category=='rfq': 
  - Mostrare messagebox placeholder: "KPI RFQ - Feature in sviluppo"
- Se tab_category=='vsm':
  - Mostrare messagebox placeholder: "KPI VSM - Dashboard Saving/Cost Avoidance/Derisking in sviluppo"
- Pulsante sempre enabled (non context-aware for enable/disable, solo per contenuto)

**Dependencies**: Phase 2 completata (context detection)

---

### **Phase 7: Form VSM Semplificazione** (Modifica VSMEventDialog)

**7.1 Rimuovere dropdown Tipo Evento**
- Rimuovere `self.combo_event_type` (ttk.Combobox)
- Sostituire con label readonly: `ttk.Label(text=f"Tipo Evento: {self.event_type_var.get()}")`
- event_type_var non più Combobox, solo StringVar passato da __init__

**7.2 Rimuovere campo Reference**
- Rimuovere: `self.entry_reference` + label "Riferimento:"
- Non salvare campo reference in VSMEvent (sarà None o "")

**7.3 Sostituire Buyer con User readonly**
- Rimuovere: `self.entry_buyer` (ttk.Entry editabile)
- Aggiungere: `ttk.Label(text=f"User: {self.current_username}")` readonly
- Salvare sempre `VSMEvent.buyer = self.current_username` (campo buyer nel DB mantiene username)

**7.4 Rimuovere "Pagamenti" da driver**
- Modificare Combobox driver values:
  - Prima: `["Prezzo", "Pagamenti", "Volume", "Altro"]`
  - Dopo: `["Prezzo", "Volume", "Altro"]`

**7.5 Bloccare event_type in edit mode**
- In `_load_event_data()`: già caricato event_type da DB
- Assicurarsi che label readonly mostri event_type caricato
- Nessuna possibilità di modifica (era già così con label readonly)

**Dependencies**: Nessuna (task isolato al dialog)

---

### **Phase 8: Filtri Collapsibili Visibility** (Modifica CollapsibleFilters)

**8.1 Implementare update_filters_visibility()**
- In dataflow.py:
  ```python
  def update_filters_visibility(self):
      _, _, tab_category = self.get_current_tree_and_status()
      if tab_category == 'vsm':
          self.collapsible_filters.grid_remove()  # nasconde
      else:
          self.collapsible_filters.grid()  # mostra
  ```

**8.2 Chiamare da on_tab_changed()**
- Aggiungere chiamata: `self.update_filters_visibility()`

**Dependencies**: Phase 2 completata (tab_category detection)

---

### **Phase 9: Cleanup & Deprecation**

**9.1 Rimuovere VSMManagementWindow**
- File: `ui/windows/vsm_management_window.py`
- ❌ NON eliminare il file (potrebbe servire logica)
- ✅ Commentare import in dataflow.py
- ✅ Aggiungere commento deprecation: "# DEPRECATED: VSM tabs now integrated directly in main notebook"

**9.2 Aggiornare imports dataflow.py**
- Rimuovere: `from ui.windows.vsm_management_window import VSMManagementWindow`
- Aggiungere: `from ui.dialogs.vsm_event_dialog import VSMEventDialog`

**Dependencies**: Tutte le phase precedenti completate

---

### **Phase 10: Testing & Verification**

**10.1 Test RFQ invariato**
- [ ] RFQ Attive: double-click apre ViewRequestWindow
- [ ] RFQ Archiviate: double-click apre ViewRequestWindow
- [ ] New RFQ: crea nuova RdO
- [ ] Actions menu RFQ: Elimina, Duplica, Archivia/Riattiva funzionanti
- [ ] Export Excel RFQ: genera file corretto

**10.2 Test VSM CRUD**
- [ ] Tab Saving: crea evento Saving con form semplificato
- [ ] Tab Cost Avoidance: crea evento CA
- [ ] Tab Derisking: crea evento Derisking
- [ ] Double-click evento: apre dialog edit con dati corretti
- [ ] Delete evento: conferma ownership, elimina, refresh sheet
- [ ] Duplicate evento: crea copia con data odierna

**10.3 Test Context Switching**
- [ ] Switch RFQ → VSM: pulsante "Nuovo Evento" appare, filtri scompaiono
- [ ] Switch VSM → RFQ: pulsante "Nuova RdO" appare, filtri riappaiono
- [ ] Actions menu: contenuto cambia tra RFQ e VSM
- [ ] Export Excel: comportamento diverso per RFQ vs VSM
- [ ] KPI button: sempre visibile, placeholder diverso

**10.4 Test Form Semplificato**
- [ ] Tipo Evento: label readonly (non modificabile)
- [ ] User: label readonly con current_username
- [ ] Campo Reference: assente
- [ ] Driver: "Pagamenti" non presente nelle opzioni
- [ ] Edit mode: event_type non modificabile

**Dependencies**: Tutte le phase precedenti completate

---

## Relevant Files (con modifiche dettagliate)

### dataflow.py (MODIFICHE ESTENSIVE)
- **Lines 3590-3620** (toolbar): aggiungere btn_kpi, salvare riferimenti btn_new
- **Lines 3687-3695** (notebook): sostituire tab_vsm con tab_saving/tab_cost_avoidance/tab_derisking
- **Lines 4361** (on_tab_changed): aggiungere update_filters_visibility()
- **Lines 4464** (get_current_tree_and_status): estendere return con tab_category
- **Lines 4384** (_populate_actions_menu): aggiungere branch VSM
- **Lines 5321** (mega_export_excel): aggiungere branch VSM → export_vsm_events_to_excel()
- **Lines 5197** (on_sheet_double_click): aggiungere detection VSM sheets
- **Nuovi metodi**:
  - `create_vsm_event_sheet(parent, event_type)` → crea tksheet per VSM
  - `refresh_vsm_events(event_type)` → carica e popola eventi
  - `_populate_vsm_sheet(sheet, events, event_type)` → riempie sheet con metadata
  - `delete_selected_vsm_event()` → handler delete VSM
  - `duplicate_selected_vsm_event()` → handler duplicate VSM
  - `export_vsm_events_to_excel(event_type)` → export VSM
  - `update_filters_visibility()` → mostra/nasconde filtri
  - `open_kpi_window()` → placeholder KPI

### ui/dialogs/vsm_event_dialog.py (SEMPLIFICAZIONI)
- **Rimuovere**: 
  - Combobox event_type (sostituire con label readonly)
  - Entry reference + label
  - Entry buyer (sostituire con label readonly User)
- **Modificare**:
  - Driver Combobox values: rimuovere "Pagamenti"
  - __init__ signature: event_type passato come string (non modificabile)
  - _build_ui(): layout form semplificato
- **Invariato**:
  - Logica salvataggio (_validate_and_save)
  - Logica caricamento (_load_event_data)
  - Dynamic form (_on_event_type_changed)

### ui/windows/vsm_management_window.py (DEPRECATO)
- **Non eliminare** (logica riutilizzabile)
- **Aggiungere header deprecation notice**
- **Non più importato** in dataflow.py

### database_manager.py (INVARIATO)
- Metodo `get_all_vsm_events(username=None)` già implementato

### services/vsm_persistence.py (INVARIATO)
- Metodi save/update/delete_event_with_impacts già implementati

---

## Verification Checklist

**RFQ Non Impattato**
- [ ] RFQ Attive/Archiviate: double-click apre ViewRequestWindow
- [ ] Actions menu RFQ: Elimina, Duplica, Archivia/Riattiva funzionanti
- [ ] Export Excel RFQ: genera file corretto

**VSM Funzionante**
- [ ] Tab Saving/CA/Derisking: creazione eventi con form semplificato (no Reference, no Buyer editabile)
- [ ] Double-click: apre dialog edit con dati corretti
- [ ] Delete: ownership validation + conferma + refresh
- [ ] Duplicate: crea copia con data odierna

**Context Switching**
- [ ] RFQ→VSM: "Nuovo Evento", filtri nascosti
- [ ] VSM→RFQ: "Nuova RdO", filtri visibili
- [ ] Actions menu: contenuto dinamico (RFQ vs VSM)
- [ ] Export: comportamento diverso (RFQ vs VSM)
- [ ] KPI: sempre visibile, placeholder diverso

---

## Status
**Ready for Implementation** - Piano completato e validato con utente
