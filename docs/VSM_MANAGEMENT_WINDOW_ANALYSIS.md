# VSMManagementWindow - Analisi Completa

**Data Analisi**: 25 marzo 2026  
**File Analizzato**: `ui/windows/vsm_management_window.py` (444 righe)  
**Versione**: Step 4B (post-integrazione tab dirette in DataFlow)

---

## 📑 Indice

1. [Panoramica Architetturale](#1-panoramica-architetturale)
2. [Componenti UI](#2-componenti-ui)
3. [Logica Popolamento Dati](#3-logica-popolamento-dati)
4. [Binding e Comportamento Eventi](#4-binding-e-comportamento-eventi)
5. [Operazioni CRUD](#5-operazioni-crud)
6. [Pattern Architetturali](#6-pattern-architetturali)
7. [Dipendenze Esterne](#7-dipendenze-esterne)
8. [Limitazioni e Considerazioni](#8-limitazioni-e-considerazioni)
9. [Confronto con Integrazione Step 4B](#9-confronto-con-integrazione-step-4b)
10. [Raccomandazioni per Prossimi Step](#10-raccomandazioni-per-prossimi-step)

---

## 1. Panoramica Architetturale

### 1.1 Definizione Classe

```python
class VSMManagementWindow(ttk.Frame):
    """
    Contenitore principale per il modulo VSM.
    Embedded come tab nella Main Dashboard.
    """
```

**Tipo**: `ttk.Frame` - Container embedded (non finestra top-level)  
**Parent**: Notebook principale di DataFlow (`dataflow.py`)  
**Ruolo**: Gestione completa eventi VSM (Value Stream Mapping) con interfaccia multi-tab

### 1.2 Parametri Inizializzazione

```python
def __init__(self, parent, current_username, refresh_callback=None):
```

| Parametro | Tipo | Descrizione |
|-----------|------|-------------|
| `parent` | Widget | Widget parent (notebook principale) |
| `current_username` | str | Username utente corrente per ownership validation |
| `refresh_callback` | callable\|None | Callback opzionale per refresh dashboard principale |

### 1.3 Attributi Principali

| Attributo | Tipo | Scopo |
|-----------|------|-------|
| `self.current_username` | str | Username utente corrente |
| `self.refresh_callback` | callable\|None | Callback per notificare dashboard padre |
| `self.current_event_type` | str | Tipo evento corrente ("Saving"\|"Cost Avoidance"\|"Derisking") |
| `self.sheets` | dict | Mappa `{event_type: Sheet}` per accesso sheet |
| `self.vsm_notebook` | ttk.Notebook | Nested notebook con 3 sub-tab |

---

## 2. Componenti UI

### 2.1 Struttura Layout

```
VSMManagementWindow (ttk.Frame)
├── Toolbar (ttk.Frame) - Orizzontale, top
│   ├── ➕ Nuovo Evento
│   ├── ✏️ Modifica
│   ├── 🗑 Elimina
│   ├── [Separator]
│   ├── 📊 KPI
│   └── 🔄 Aggiorna
│
└── vsm_notebook (ttk.Notebook)
    ├── Tab "Saving"
    │   └── Sheet (tksheet) - 8 colonne
    ├── Tab "Cost Avoidance"
    │   └── Sheet (tksheet) - 8 colonne
    └── Tab "Derisking"
        └── Sheet (tksheet) - 8 colonne
```

### 2.2 Toolbar - Linee 61-98

#### Pulsanti Implementati

| # | Icona | Label | Command | Stato |
|---|-------|-------|---------|-------|
| 1 | ➕ | Nuovo Evento | `on_new_event()` | ✅ Implementato |
| 2 | ✏️ | Modifica | `on_edit_event()` | ✅ Implementato |
| 3 | 🗑 | Elimina | `on_delete_event()` | ✅ Implementato |
| 4 | 📊 | KPI | `on_kpi_click()` | ⚠️ Placeholder |
| 5 | 🔄 | Aggiorna | `refresh_events()` | ✅ Implementato |

**Layout**: Tutti allineati a sinistra con `pack(side="left", padx=5)`

### 2.3 Nested Notebook - Linee 99-118

#### Configurazione Tab

```python
self.vsm_notebook = ttk.Notebook(self)
self.vsm_notebook.pack(fill="both", expand=True, padx=10, pady=5)

# Tab 1: Saving
self.tab_saving = ttk.Frame(self.vsm_notebook)
self.vsm_notebook.add(self.tab_saving, text=_("Saving"))
self.sheets["Saving"] = self._create_event_sheet(self.tab_saving)

# Tab 2: Cost Avoidance (identico pattern)
# Tab 3: Derisking (identico pattern)
```

**Binding Globale**:
```python
self.vsm_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
```
→ Ogni cambio tab ricarica automaticamente i dati

### 2.4 Sheet VSM - Metodo `_create_event_sheet()` (linee 120-182)

#### Configurazione Colonne

| # | Header | Larghezza (px) | Allineamento | Readonly |
|---|--------|----------------|--------------|----------|
| 0 | Data | 100 | center | ✓ |
| 1 | Tipo | 120 | center | ✓ |
| 2 | Azione | 120 | center | ✓ |
| 3 | Descrizione | 300 | **left** | ✓ |
| 4 | Valore Teorico | 120 | center | ✓ |
| 5 | Realizzo % | 90 | center | ✓ |
| 6 | Ripetitivo | 90 | center | ✓ |
| 7 | Utente | 140 | center | ✓ |

#### Proprietà Sheet

```python
sheet = Sheet(
    frame,
    theme="light blue",              # Tema coerente con RFQ sheets
    header_font=("Calibri", 11, "bold"),
    font=("Calibri", 11, "normal"),
    show_header=True,
    show_row_index=False             # No numerazione righe
)
```

#### Configurazioni Applicate

```python
# Larghezze colonne
sheet.set_column_widths([100, 120, 120, 300, 120, 90, 90, 140])

# Allineamento (Descrizione rimane left-aligned di default)
sheet.align_columns(columns=[0, 1, 2, 4, 5, 6, 7], align="center")

# Abilita bindings (enable sorting, selection, etc.)
sheet.enable_bindings()

# Rendi tutte le colonne readonly
for col_idx in range(8):
    sheet.readonly_columns(columns=[col_idx], readonly=True)
```

#### Metadata Storage

```python
sheet._event_metadata = []  # Lista di dict con event_id, username, is_mine
```

**Struttura metadata** (popolata in `_populate_sheet()`):
```python
[
    {
        'event_id': 123,           # ID evento per operazioni CRUD
        'username': 'user@domain', # Owner evento
        'is_mine': True            # Flag ownership per current_user
    },
    ...  # Un dict per ogni riga
]
```

---

## 3. Logica Popolamento Dati

### 3.1 Metodo `refresh_events()` - Linee 201-227

#### Trigger Points

| Evento | Linea di Codice | Note |
|--------|-----------------|------|
| Inizializzazione | 58 | Carica dati all'apertura |
| Click "🔄 Aggiorna" | 96 | Refresh manuale utente |
| Cambio tab | 193 (`_on_tab_changed`) | Auto-refresh tab attiva |
| Post-save evento | 285, 353 | Dopo Create/Edit |
| Post-delete evento | 424 | Dopo Delete |

#### Flusso Logico

```python
def refresh_events(self):
    try:
        # 1. Connessione database
        with DatabaseManager(get_db_path()) as db_manager:
            # 2. Carica TUTTI eventi utente (no filtro iniziale)
            all_events = db_manager.get_all_vsm_events(username=self.current_username)
        
        # 3. Filtra per tipo tab corrente
        filtered_events = [e for e in all_events if e.event_type == self.current_event_type]
        
        # 4. Ottieni sheet corrente dal dictionary
        sheet = self.sheets[self.current_event_type]
        
        # 5. Popola sheet con eventi filtrati
        self._populate_sheet(sheet, filtered_events)
        
        logger.debug(f"Caricati {len(filtered_events)} eventi {self.current_event_type}")
        
    except DatabaseError as e:
        # 6. Error handling
        messagebox.showerror(...)
```

**Caratteristiche**:
- ✅ Carica sempre TUTTI gli eventi utente (no paginazione)
- ✅ Filtro lato client per `event_type`
- ⚠️ No caching (ogni refresh = query DB completa)
- ⚠️ Refresh integrale (no row-level updates)

### 3.2 Metodo `_populate_sheet()` - Linee 229-267

#### Processo di Popolamento

```python
def _populate_sheet(self, sheet, events):
    data_rows = []
    metadata = []
    
    for event in events:
        # 1. CALCOLO VALORE TEORICO
        valore_teorico = event.calculate_theoretical_value()
        
        # 2. FORMATTAZIONE ROW
        row = [
            # Col 0: Data in formato italiano
            event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
            
            # Col 1: Tipo (Saving|Cost Avoidance|Derisking)
            event.event_type,
            
            # Col 2: Azione
            event.action,
            
            # Col 3: Descrizione (TRUNCATED a 50 caratteri!)
            (event.description or event.reference or "")[:50],
            
            # Col 4: Valore teorico formattato con € e separatore migliaia
            f"€ {valore_teorico:,.2f}",
            
            # Col 5: Realizzo percentuale (intero)
            f"{event.percent_realizzo:.0f}%",
            
            # Col 6: Ripetitivo (checkmark o vuoto)
            "✓" if event.opex_ripetitivo else "",
            
            # Col 7: Username owner
            event.username
        ]
        data_rows.append(row)
        
        # 3. METADATA TRACKING
        metadata.append({
            'event_id': event.id,
            'username': event.username,
            'is_mine': event.username == self.current_username
        })
    
    # 4. UPDATE SHEET
    sheet.set_sheet_data(data_rows)      # Sostituisce tutti i dati
    sheet._event_metadata = metadata      # Aggiorna metadata sincronizzati
```

#### Regole di Formattazione

| Colonna | Formato Input | Formato Output | Note |
|---------|---------------|----------------|------|
| Data | `datetime.date` | `"25/03/2026"` | Pattern italiano `dd/mm/yyyy` |
| Tipo | `str` | As-is | No trasformazione |
| Azione | `str` | As-is | No trasformazione |
| Descrizione | `str` | **Max 50 char** | ⚠️ **Truncation loss** |
| Valore Teorico | `float` | `"€ 1,234.56"` | Virgola separatore migliaia |
| Realizzo % | `float` (0-100) | `"85%"` | Arrotondato a intero |
| Ripetitivo | `bool` | `"✓"` o `""` | Simbolo checkmark Unicode |
| Utente | `str` | As-is | No trasformazione |

**⚠️ TRUNCATION WARNING**: Descrizioni oltre 50 caratteri vengono tagliate in visualizzazione. La descrizione completa è visibile solo nel dialog di modifica.

---

## 4. Binding e Comportamento Eventi

### 4.1 Eventi Sheet - Linee 172-176

#### Double-Click Handler

```python
sheet.bind("<Double-Button-1>", lambda event: self._on_sheet_double_click(sheet, event))
```

**Comportamento**:
- **Trigger**: Doppio click su qualsiasi cella del sheet
- **Azione**: Chiama `_on_sheet_double_click()` → Delega a `on_edit_event()`
- **Effetto**: Apre dialog modifica evento (se ownership valida)

**Implementazione** (linee 196-197):
```python
def _on_sheet_double_click(self, sheet, event):
    """Handler per doppio click su sheet (apre edit)."""
    self.on_edit_event()
```

#### Selection Handlers

```python
sheet.extra_bindings("cell_select", lambda event_data: self._update_buttons_state())
sheet.extra_bindings("row_select", lambda event_data: self._update_buttons_state())
```

**Comportamento**:
- **Trigger**: Selezione cella o riga intera
- **Azione**: Chiama `_update_buttons_state()`
- **Stato attuale**: ⚠️ **NON IMPLEMENTATO** (linee 198-200)

```python
def _update_buttons_state(self):
    """Aggiorna stato pulsanti in base alla selezione (future enhancement)."""
    # Per ora non implementato, ma lascia hook per future
    pass
```

**Future Enhancement Potenziale**:
- Disabilitare "Modifica" se selezione multipla
- Disabilitare "Elimina" se nessuna selezione
- Disabilitare "Modifica/Elimina" se `is_mine=False` per selezione corrente

### 4.2 Eventi Notebook - Linea 118

#### Tab Change Handler

```python
self.vsm_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
```

**Implementazione** (linee 188-195):
```python
def _on_tab_changed(self, event=None):
    """Handler per cambio tab VSM."""
    # 1. Ottieni index tab corrente (0, 1, 2)
    current_tab_idx = self.vsm_notebook.index(self.vsm_notebook.select())
    
    # 2. Mappa a nome tipo evento
    tab_names = ["Saving", "Cost Avoidance", "Derisking"]
    self.current_event_type = tab_names[current_tab_idx]
    
    # 3. Refresh automatico dati tab attiva
    self.refresh_events()
```

**Comportamento**:
- ✅ **Auto-refresh**: Ogni cambio tab ricarica dati automaticamente
- ✅ **Context switching**: Aggiorna `current_event_type` per operazioni successive
- ⚠️ **Performance**: Ricarica completa anche se tab già visitata (no caching)

---

## 5. Operazioni CRUD

### 5.1 CREATE - `on_new_event()` (linee 269-290)

#### Flusso Operativo

```python
def on_new_event(self):
    # Import lazy per evitare circular dependencies
    from ui.dialogs.vsm_event_dialog import VSMEventDialog
    
    try:
        # 1. Apri dialog in modalità CREATE
        dialog = VSMEventDialog(
            self,
            current_username=self.current_username,
            event_type=self.current_event_type,  # Preimpostato da tab corrente
            event_id=None                         # None = CREATE mode
        )
        self.wait_window(dialog)                  # Blocca fino a chiusura
        
        # 2. Check se salvato con successo
        if hasattr(dialog, 'result') and dialog.result:
            # 3. Refresh sheet corrente
            self.refresh_events()
            
            # 4. Notifica dashboard padre (se callback presente)
            if self.refresh_callback:
                self.refresh_callback()
    
    except Exception as e:
        logger.error(f"Errore apertura dialog nuovo evento: {e}", exc_info=True)
        messagebox.showerror(...)
```

**Caratteristiche**:
- ✅ `event_type` preimpostato da tab corrente (user-friendly)
- ✅ Import lazy di `VSMEventDialog` (evita circular imports)
- ✅ `wait_window()` = dialogo modale bloccante
- ✅ Refresh automatico post-save
- ✅ Propagazione callback a dashboard padre

### 5.2 UPDATE - `on_edit_event()` (linee 292-365)

#### Flusso di Validazione

```python
def on_edit_event(self):
    # 1. Ottieni sheet corrente
    sheet = self.sheets[self.current_event_type]
    
    # 2. VALIDAZIONE SELEZIONE
    selected_rows = list(sheet.get_selected_rows())
    
    # 2a. Nessuna selezione
    if not selected_rows:
        messagebox.showwarning(
            _("Nessuna Selezione"),
            _("Seleziona un evento da modificare.")
        )
        return
    
    # 2b. Selezione multipla non supportata
    if len(selected_rows) > 1:
        messagebox.showwarning(
            _("Selezione Multipla"),
            _("Seleziona un solo evento per la modifica.")
        )
        return
    
    # 3. ESTRAZIONE METADATA
    row_idx = selected_rows[0]
    if row_idx >= len(sheet._event_metadata):  # Safety check
        return
    
    metadata = sheet._event_metadata[row_idx]
    event_id = metadata['event_id']
    is_mine = metadata['is_mine']
    
    # 4. OWNERSHIP VALIDATION
    if not is_mine:
        messagebox.showerror(
            _("Operazione Non Consentita"),
            _("Puoi modificare solo i tuoi eventi VSM.")
        )
        return
    
    # 5. APRI DIALOG EDIT
    from ui.dialogs.vsm_event_dialog import VSMEventDialog
    
    try:
        dialog = VSMEventDialog(
            self,
            current_username=self.current_username,
            event_type=self.current_event_type,
            event_id=event_id  # event_id presente = EDIT mode
        )
        self.wait_window(dialog)
        
        # 6. REFRESH POST-SAVE
        if hasattr(dialog, 'result') and dialog.result:
            self.refresh_events()
            if self.refresh_callback:
                self.refresh_callback()
    
    except Exception as e:
        logger.error(f"Errore apertura dialog modifica evento: {e}", exc_info=True)
        messagebox.showerror(...)
```

**Validazioni Implementate**:
1. ✅ **No selection check**: Almeno 1 riga selezionata
2. ✅ **Single selection only**: Max 1 riga (no edit massivo)
3. ✅ **Ownership validation**: Solo eventi propri (no admin override)
4. ✅ **Index bounds check**: Verifica `row_idx < len(metadata)`

**Limitazioni**:
- ❌ **No multi-edit**: Anche se logica simile possibile, non implementata
- ❌ **No admin override**: Nessun bypass per ruoli privilegiati

### 5.3 DELETE - `on_delete_event()` (linee 367-433)

#### Flusso Multi-Delete

```python
def on_delete_event(self):
    # 1. Ottieni sheet corrente
    sheet = self.sheets[self.current_event_type]
    
    # 2. VALIDAZIONE SELEZIONE
    selected_rows = list(sheet.get_selected_rows())
    
    if not selected_rows:
        messagebox.showwarning(
            _("Nessuna Selezione"),
            _("Seleziona uno o più eventi da eliminare.")
        )
        return
    
    # 3. BATCH VALIDATION + OWNERSHIP CHECK
    events_to_delete = []
    for row_idx in selected_rows:
        if row_idx >= len(sheet._event_metadata):
            continue  # Skip invalid indices
        
        metadata = sheet._event_metadata[row_idx]
        
        # 3a. Ownership check per ogni riga
        if not metadata['is_mine']:
            messagebox.showerror(
                _("Operazione Non Consentita"),
                _("Puoi eliminare solo i tuoi eventi VSM.\n"
                  "Alcuni eventi selezionati appartengono ad altri utenti.")
            )
            return  # ⚠️ ABORT COMPLETO se anche solo 1 non-mine
        
        events_to_delete.append(metadata['event_id'])
    
    # 4. CONFERMA ELIMINAZIONE
    count = len(events_to_delete)
    if not messagebox.askyesno(
        _("Conferma Eliminazione"),
        _("Sei sicuro di voler eliminare {} evento(i) VSM?\n"
          "Questa operazione non può essere annullata.").format(count)
    ):
        return
    
    # 5. ELIMINAZIONE BATCH
    try:
        with DatabaseManager(get_db_path()) as db_manager:
            for event_id in events_to_delete:
                # Eliminazione cascading (evento + impacts correlati)
                delete_event_and_impacts(db_manager, event_id)
        
        messagebox.showinfo(
            _("Successo"),
            _("{} evento(i) VSM eliminato(i) con successo.").format(count)
        )
        
        # 6. REFRESH POST-DELETE
        self.refresh_events()
        if self.refresh_callback:
            self.refresh_callback()
    
    except (DatabaseError, VSMError) as e:
        logger.error(f"Errore eliminazione eventi VSM: {e}")
        messagebox.showerror(
            _("Errore Eliminazione"),
            _("Impossibile eliminare gli eventi:\n{}").format(e)
        )
```

**Caratteristiche Batch Delete**:
- ✅ **Multi-selection supportata**: Può eliminare N eventi in una volta
- ✅ **Atomic ownership check**: Se anche solo 1 evento non-mine → abort TUTTO
- ✅ **Conferma count-aware**: Mostra quanti eventi saranno eliminati
- ✅ **Cascading delete**: Via `delete_event_and_impacts()` elimina anche impacts correlati
- ✅ **Transactional**: `DatabaseManager` context manager garantisce rollback su errore

**Safety Features**:
1. **Conferma esplicita**: `askyesno()` dialog con warning "non può essere annullata"
2. **ALL-or-NOTHING**: Se ownership fail → nessun delete (no partial)
3. **Error handling**: Catch `DatabaseError` e `VSMError` separatamente

### 5.4 KPI - `on_kpi_click()` (linee 435-444)

```python
def on_kpi_click(self):
    """Handler per pulsante KPI (placeholder per step futuro)."""
    messagebox.showinfo(
        _("KPI Analysis"),
        _("Funzionalità KPI Analysis in arrivo nel prossimo step.\n\n"
          "Prossime funzionalità:\n"
          "• Aggregazioni mensili/trimestrali/annuali\n"
          "• Confronto Saving vs Cost Avoidance\n"
          "• Grafici Valore Teorico vs Effettivo\n"
          "• Export Excel/CSV")
    )
```

**Stato**: ⚠️ **PLACEHOLDER** - Nessuna implementazione logica

**Future Features Pianificate**:
1. Aggregazioni temporali (mensile/trimestrale/annuale)
2. Confronto cross-type (Saving vs Cost Avoidance vs Derisking)
3. Grafici valore teorico vs effettivo (realizzo %)
4. Export dati (Excel/CSV)

---

## 6. Pattern Architetturali

### 6.1 Ownership Model

#### Implementazione

```python
# In _populate_sheet()
metadata.append({
    'event_id': event.id,
    'username': event.username,
    'is_mine': event.username == self.current_username  # Flag ownership
})
```

#### Enforcement Points

| Operazione | Validazione | Comportamento se `is_mine=False` |
|------------|-------------|----------------------------------|
| **View** | ❌ No check | Visibile a tutti gli eventi utente |
| **Create** | ❌ N/A | Sempre permesso (nuovo evento = auto-owned) |
| **Edit** | ✅ Strict | Error: "Puoi modificare solo i tuoi eventi" |
| **Delete** | ✅ Strict | Error: "Puoi eliminare solo i tuoi eventi" |

**Caratteristiche**:
- ✅ **Row-level authorization**: Check per-row via metadata
- ✅ **User-scoped visibility**: `get_all_vsm_events(username)` filtra solo eventi utente
- ❌ **No admin override**: Nessun bypass per role="admin"
- ❌ **No delegation**: Non esiste concetto di "shared ownership"

### 6.2 Event Type Isolation

#### State Management

```python
self.current_event_type = "Saving"  # Default iniziale

# In _on_tab_changed():
tab_names = ["Saving", "Cost Avoidance", "Derisking"]
self.current_event_type = tab_names[current_tab_idx]
```

#### Isolation Enforcement

| Metodo | Filtro Applicato |
|--------|------------------|
| `refresh_events()` | `[e for e in all_events if e.event_type == self.current_event_type]` |
| `on_new_event()` | `VSMEventDialog(..., event_type=self.current_event_type)` |
| `on_edit_event()` | Opera su sheet di `self.current_event_type` |
| `on_delete_event()` | Opera su sheet di `self.current_event_type` |

**Implicazioni**:
- ✅ **No cross-contamination**: Impossibile editare evento Saving mentre si è in tab Cost Avoidance
- ✅ **Context-aware operations**: Ogni operazione conosce event_type corrente
- ⚠️ **Redundancy in DB**: Stesso dato potrebbe essere copiato manualmente cross-type (no validazione)

### 6.3 Data Flow Pattern

```
┌──────────────────────────────────────────────────────────┐
│                    USER INTERACTION                       │
└────────────────────┬─────────────────────────────────────┘
                     │
                     ▼
        ┌────────────────────────┐
        │   on_new_event()       │
        │   on_edit_event()      │
        │   on_delete_event()    │
        │   refresh_events()     │
        └────────┬───────────────┘
                 │
                 ▼
        ┌────────────────────────┐
        │   DatabaseManager      │
        │   (context manager)    │
        └────────┬───────────────┘
                 │
                 ▼
        ┌────────────────────────┐
        │ get_all_vsm_events()   │  ← Carica TUTTI eventi utente
        │ save_event()           │
        │ delete_event_and_      │
        │   impacts()            │
        └────────┬───────────────┘
                 │
                 ▼
        ┌────────────────────────┐
        │  Filter by event_type  │  ← Filtro lato client
        └────────┬───────────────┘
                 │
                 ▼
        ┌────────────────────────┐
        │  _populate_sheet()     │  ← Formattazione + metadata
        └────────┬───────────────┘
                 │
                 ▼
        ┌────────────────────────┐
        │  sheet.set_sheet_data()│  ← Refresh completo UI
        └────────────────────────┘
```

**Caratteristiche**:
- ✅ **Single source of truth**: DB è autorità
- ✅ **Stateless UI**: Ogni refresh ricostruisce da zero
- ⚠️ **No incremental updates**: Sempre full refresh (anche per singola modifica)
- ⚠️ **No client-side caching**: Ogni operazione = query DB

### 6.4 Callback Pattern

```python
# In __init__:
self.refresh_callback = refresh_callback

# In operazioni CRUD:
if self.refresh_callback:
    self.refresh_callback()
```

**Uso Corrente**:
- Chiamato dopo ogni operazione CUD (Create/Update/Delete)
- Permette a `DataFlow` (dashboard padre) di aggiornare UI globale

**Limitazioni**:
- ❌ **No event propagation**: Callback non riceve info su cosa è cambiato
- ❌ **Blind refresh**: Dashboard deve rifare query complete per sapere cosa aggiornare
- ❌ **No granularity**: Stesso callback per create/edit/delete (no distinzione)

---

## 7. Dipendenze Esterne

### 7.1 Database Layer

#### DatabaseManager

```python
from database_manager import DatabaseManager, DatabaseError
```

**Metodi Utilizzati**:

| Metodo | Uso | Parametri | Ritorno |
|--------|-----|-----------|---------|
| `get_all_vsm_events()` | Lettura | `username` (str) | `List[VSMEvent]` |
| `save_event()` | Create/Update | `event` (VSMEvent) | `void` (raise on error) |

**Pattern Context Manager**:
```python
with DatabaseManager(get_db_path()) as db_manager:
    events = db_manager.get_all_vsm_events(username=self.current_username)
```
→ Auto-commit su successo, auto-rollback su eccezione

#### VSM Persistence

```python
from services.vsm_persistence import delete_event_and_impacts, VSMError
```

**Metodo `delete_event_and_impacts()`**:
- **Scopo**: Eliminazione cascading evento + impacts correlati
- **Firma**: `delete_event_and_impacts(db_manager: DatabaseManager, event_id: int)`
- **Comportamento**:
  1. Elimina tutti `vsm_impacts` correlati (via FK `event_id`)
  2. Elimina evento VSM principale
  3. Raise `VSMError` su fallimento

### 7.2 UI Components

#### VSMEventDialog

```python
from ui.dialogs.vsm_event_dialog import VSMEventDialog
```

**Import Pattern**: ⚠️ **Lazy import** (dentro metodi, non top-level)  
**Motivo**: Evita circular dependencies (VSMEventDialog potrebbe importare VSMManagementWindow)

**Modalità Utilizzo**:

**CREATE Mode**:
```python
dialog = VSMEventDialog(
    parent=self,
    current_username=self.current_username,
    event_type=self.current_event_type,  # Preimpostato
    event_id=None                         # None = CREATE
)
```

**EDIT Mode**:
```python
dialog = VSMEventDialog(
    parent=self,
    current_username=self.current_username,
    event_type=self.current_event_type,
    event_id=123                          # ID esistente = EDIT
)
```

**Result Handling**:
```python
self.wait_window(dialog)  # Blocca fino a chiusura

if hasattr(dialog, 'result') and dialog.result:
    # Dialog salvato con successo
    self.refresh_events()
```

### 7.3 Domain Models

#### VSMEvent

```python
from models.vsm_event import VSMEvent
```

**Metodo Utilizzato**:
```python
valore_teorico = event.calculate_theoretical_value()
```

**Calcolo Valore Teorico** (implementato in `models.vsm_event.py`):
```python
def calculate_theoretical_value(self) -> float:
    """
    Calcola il valore teorico dell'evento VSM.
    
    Formula:
    - Se unique: valore_singolo
    - Se ripetitivo: valore_singolo * frequenza_annua
    """
    base_value = self.valore_singolo or 0.0
    
    if self.opex_ripetitivo:
        frequenza = self.opex_freq_annua or 12  # Default mensile
        return base_value * frequenza
    else:
        return base_value
```

**Attributi VSMEvent Rilevanti**:
| Attributo | Tipo | Descrizione |
|-----------|------|-------------|
| `id` | int | Primary key |
| `event_date` | date | Data evento |
| `event_type` | str | "Saving"\|"Cost Avoidance"\|"Derisking" |
| `action` | str | Azione VSM |
| `description` | str | Descrizione lunga |
| `reference` | str | Riferimento corto |
| `valore_singolo` | float | Valore unitario |
| `opex_ripetitivo` | bool | Flag evento ripetitivo |
| `opex_freq_annua` | int | Frequenza annuale (se ripetitivo) |
| `percent_realizzo` | float | % realizzo (0-100) |
| `username` | str | Owner evento |

### 7.4 Utility Modules

#### i18n_utils

```python
from utils.i18n_utils import _
```

**Uso**: Wrapper per gettext internazionalizzazione
```python
text = _("Testo da tradurre")  # Cerca traduzione in locale/*.po
```

**Lingue Supportate** (basato su `locale/` directory):
- Italiano (IT)
- Inglese (EN)

#### app_paths

```python
from services.app_paths import get_db_path
```

**Uso**: Ottiene path assoluto database SQLite corrente
```python
db_path = get_db_path()  # Es: "/home/user/.vsm/database.db"
```

---

## 8. Limitazioni e Considerazioni

### 8.1 Funzionalità Mancanti

#### 8.1.1 Ricerca e Filtri

**Stato**: ❌ **NON IMPLEMENTATO**

**Impatto**:
- Nessuna search bar per cercare eventi per descrizione/azione/data
- Nessun filtro dropdown per data range, utente, stato
- Con molti eventi (100+), navigazione difficoltosa

**Workaround Utente**:
- Ordinamento colonne (click header) - built-in tksheet
- Ctrl+F per ricerca nativa (se supportata da sheet)

#### 8.1.2 Export Dati

**Stato**: ❌ **NON IMPLEMENTATO** (KPI button è placeholder)

**Funzionalità Attese** (non presenti):
- Export Excel (.xlsx) con formattazione
- Export CSV per analisi esterna
- Export PDF per reportistica
- Grafici/dashboards aggregati

**Implicazioni**:
- Utenti devono copy-paste manualmente da sheet
- No backup agevole dati VSM
- No integrazione con BI tools esterni

#### 8.1.3 Bulk Operations

**Stato**: ⚠️ **PARZIALE** (solo delete)

| Operazione | Supporto Multi-Selection | Implementato |
|------------|--------------------------|--------------|
| **View** | ✓ (selezione permessa) | ✓ |
| **Edit** | ✗ (solo singola) | ✓ Single only |
| **Delete** | ✓ (batch) | ✓ |
| **Duplicate** | ✗ | ✗ |
| **Move to Type** | ✗ | ✗ |

**Mancanze**:
- No bulk edit (es: aggiornare `percent_realizzo` su N eventi)
- No duplicate evento template
- No cambio event_type batch

### 8.2 Limitazioni UI

#### 8.2.1 Truncamento Descrizione

**Problema** (linea 242):
```python
(event.description or event.reference or "")[:50]  # ⚠️ HARD-CUT
```

**Impatto**:
- Descrizioni lunghe perse in visualizzazione
- No tooltip hover per vedere testo completo
- Utente deve aprire dialog edit per leggere tutto

**Miglioramenti Possibili**:
1. Tooltip on hover con testo completo
2. Colonna espandibile (drag width)
3. Truncate con `...` suffix per segnalare troncamento
4. Render HTML con `<abbr title="full text">short</abbr>`

#### 8.2.2 Button State Management

**Problema** (linee 198-200):
```python
def _update_buttons_state(self):
    """Aggiorna stato pulsanti in base alla selezione (future enhancement)."""
    pass  # ⚠️ NON IMPLEMENTATO
```

**Impatto**:
- Pulsanti sempre abilitati (anche senza selezione valida)
- Utente può cliccare "Modifica" senza selezione → warning dialog
- Esperienza utente subottimale (feedback reattivo tardivo)

**Behavior Atteso** (non implementato):
- **Modifica**: Disabilitato se `len(selected_rows) != 1`
- **Elimina**: Disabilitato se `len(selected_rows) == 0`
- **Modifica/Elimina**: Disabilitato se `is_mine=False` per selezione corrente

#### 8.2.3 Selezione Persa su Refresh

**Problema**:
```python
sheet.set_sheet_data(data_rows)  # ⚠️ Reset completo, perde selezione
```

**Impatto**:
- Dopo edit/delete, selezione corrente persa
- Utente deve ri-selezionare manualmente per operazioni successive
- Workflow ripetitivo rallentato

**Workaround Possibile**:
1. Salvare `selected_rows` pre-refresh
2. Dopo `set_sheet_data()`, ri-selezionare stessi indici (se ancora validi)
3. Scroll automatico a riga modificata/creata

### 8.3 Limitazioni Performance

#### 8.3.1 No Caching

**Problema**:
```python
# Ogni refresh carica TUTTI eventi utente
all_events = db_manager.get_all_vsm_events(username=self.current_username)
```

**Impatto**:
- Query DB completa ogni volta (anche per tab già visitata)
- Con 1000+ eventi, latenza percepibile
- Spreco risorse se eventi cambiano raramente

**Miglioramenti Possibili**:
1. Cache in-memory con invalidation su CUD
2. Lazy loading per tab (carica solo al primo accesso)
3. Paginazione (100 eventi per pagina)
4. Background refresh async (no block UI)

#### 8.3.2 Full Refresh Always

**Problema**:
```python
sheet.set_sheet_data(data_rows)  # ⚠️ Sostituisce TUTTE le righe
```

**Impatto**:
- Anche per edit singolo, ricarica intero sheet
- Flash visivo su refresh (sheet cleared momentaneamente)
- Scroll position non preservata

**Incremental Update Possibile** (non implementato):
```python
# Per edit singolo:
sheet.set_cell_data(row_idx, col_idx, new_value)
# Per delete singolo:
sheet.delete_row(row_idx)
# Per insert singolo:
sheet.insert_row(new_row_data, idx=0)  # Top di sheet
```

#### 8.3.3 Refresh Cascading

**Problema**:
```python
# In on_edit_event():
self.refresh_events()              # Refresh VSMManagementWindow
if self.refresh_callback:
    self.refresh_callback()        # Refresh DataFlow intero!
```

**Impatto**:
- Ogni modifica VSM triggera refresh globale dashboard
- Possibile refresh anche di tab RFQ non modificate
- Overhead crescente con complessità dashboard

**Ottimizzazione Possibile**:
- Callback granulare: `refresh_callback(scope='vsm', event_type='Saving')`
- DataFlow decide selettivamente cosa refreshare
- Publisher-Subscriber pattern per event-driven updates

### 8.4 Problematiche Sicurezza/Validazione

#### 8.4.1 No Soft Delete

**Problema**:
```python
delete_event_and_impacts(db_manager, event_id)  # ⚠️ HARD DELETE
```

**Rischi**:
- Eliminazione permanente immediata
- No undo possibile
- Perdita dati per errori accidentali

**Best Practice Mancante**:
- Soft delete: `deleted_at` timestamp field
- Eventi "eliminati" filtrati da queries ma recuperabili
- Garbage collection periodica per cleanup definitivo

#### 8.4.2 No Audit Trail

**Problema**: Nessun logging modifiche

**Mancanze**:
- Chi ha modificato cosa e quando
- Storia valori precedenti (no versioning)
- No compliance per audit governance

**Implementazione Tipica** (assente):
```sql
CREATE TABLE vsm_audit_log (
    id INTEGER PRIMARY KEY,
    event_id INTEGER,
    operation TEXT,  -- 'create', 'update', 'delete'
    changed_by TEXT,
    changed_at TIMESTAMP,
    old_values TEXT,  -- JSON
    new_values TEXT   -- JSON
);
```

#### 8.4.3 No Rate Limiting

**Problema**: Nessun controllo spam operations

**Rischi**:
- Utente può creare 1000 eventi in loop
- Batch delete 5000 eventi (no conferma count-based scale)
- Possibile DoS accidentale DB

**Protezioni Mancanti**:
- Max N eventi creati per minuto
- Conferma speciale per batch >100 delete
- Throttling queries ripetitive

---

## 9. Confronto con Integrazione Step 4B

### 9.1 Architettura Originale (VSMManagementWindow)

```
DataFlow (main window)
└── Notebook principale
    └── Tab "VSM" (container)
        └── VSMManagementWindow (ttk.Frame)
            ├── Toolbar (5 pulsanti)
            └── Nested Notebook
                ├── Tab "Saving"
                ├── Tab "Cost Avoidance"
                └── Tab "Derisking"
```

**Caratteristiche**:
- ✅ UI completa e funzionale
- ✅ Tutte operazioni CRUD implementate
- ✅ Nested notebook (2 livelli: main → VSM → subtabs)
- ⚠️ Toolbar dedicata (non integrata con toolbar principale DataFlow)
- ⚠️ Logica duplicata per ogni subtab (3 sheet simili)

### 9.2 Architettura Step 4B (Integrazione Diretta)

```
DataFlow (main window)
└── Notebook principale (flat)
    ├── Tab "RdO Attive"
    ├── Tab "RdO Archiviate"
    ├── Tab "Saving" (direct)
    ├── Tab "Cost Avoidance" (direct)
    └── Tab "Derisking" (direct)
```

**Modifiche Applicate** ([dataflow.py](dataflow.py)):

#### 9.2.1 UI Extraction (linee 3683-3694)

```python
# Step 4A: Creazione tab dirette
self.tab_saving = ttk.Frame(self.notebook)
self.tab_cost_avoidance = ttk.Frame(self.notebook)
self.tab_derisking = ttk.Frame(self.notebook)
self.notebook.add(self.tab_saving, text=_("Saving"))
self.notebook.add(self.tab_cost_avoidance, text=_("Cost Avoidance"))
self.notebook.add(self.tab_derisking, text=_("Derisking"))

# Step 4B: Riutilizzo UI struttura sheet
self.sheet_saving = self._create_vsm_event_sheet(self.tab_saving)
self.sheet_cost_avoidance = self._create_vsm_event_sheet(self.tab_cost_avoidance)
self.sheet_derisking = self._create_vsm_event_sheet(self.tab_derisking)
```

#### 9.2.2 Metodo Estratto (linee 4265-4330)

```python
def _create_vsm_event_sheet(self, parent):
    """
    ESTRATTO da VSMManagementWindow._create_event_sheet() (Step 4B).
    """
    # ✅ Stessa struttura 8 colonne
    # ✅ Stesse larghezze [100, 120, 120, 300, 120, 90, 90, 140]
    # ✅ Stesso allineamento (center except Descrizione)
    # ✅ Stesso readonly config
    # ✅ Stessa metadata structure: sheet._event_metadata = []
    
    # ❌ RIMOSSI binding toolbar-specific:
    #    - sheet.bind("<Double-Button-1>", ...)
    #    - sheet.extra_bindings("cell_select", ...)
    #    - sheet.extra_bindings("row_select", ...)
```

#### 9.2.3 Context Detection (linee 4538-4552)

```python
def get_current_tree_and_status(self):
    tab_index = self.notebook.index(self.notebook.select())
    if tab_index == 0:
        return (self.tree_attive, 'attiva')
    elif tab_index == 1:
        return (self.tree_archiviate, 'archiviata')
    # Step 4B: VSM tabs
    elif tab_index == 2:
        return (self.sheet_saving, 'vsm_saving')
    elif tab_index == 3:
        return (self.sheet_cost_avoidance, 'vsm_cost_avoidance')
    elif tab_index == 4:
        return (self.sheet_derisking, 'vsm_derisking')
```

#### 9.2.4 Logic Guards (3 locations)

```python
# update_button_visibility() - linea 4435
if sheet is None or status.startswith('vsm_'):
    # Disabilita Actions button per VSM (logica non ancora connessa)
    self.btn_actions.config(state="disabled")

# search_requests() - linea 4762
if tree is None or status.startswith('vsm_'):
    # Skip search per VSM (no search implementata)
    return

# mega_export_excel() - linea 5424
if current_tree is None or status.startswith('vsm_'):
    # Export non disponibile per VSM (KPI export non implementato)
    messagebox.showinfo("Export Non Disponibile", ...)
```

### 9.3 Cosa Manca in Step 4B (Intentional)

| Funzionalità | VSMManagementWindow | Step 4B DataFlow | Motivo |
|--------------|---------------------|------------------|--------|
| **Toolbar dedicata** | ✅ 5 pulsanti | ❌ Non integrata | Toolbar principale DataFlow context-agnostic |
| **Data loading** | ✅ `refresh_events()` | ❌ Sheet vuoti | Step 4C - Data Population |
| **CRUD handlers** | ✅ Completo | ❌ No handlers | Step 4D - Event Handlers |
| **Double-click edit** | ✅ Binding | ❌ No binding | Step 4D |
| **Selection handlers** | ⚠️ Stub | ❌ No binding | Step 4D (o skip) |
| **Tab change refresh** | ✅ Auto-refresh | ❌ No handler | Step 4C |
| **Export/KPI** | ⚠️ Placeholder | ❌ No integration | Step 4E/4F |

### 9.4 Vantaggi Integrazione Diretta

#### 9.4.1 UX Improvements
- ✅ **Flat navigation**: 1 click per accedere a VSM tab (vs 2 click nested)
- ✅ **Visual consistency**: Tutte le tab allo stesso livello (RFQ + VSM)
- ✅ **Breadcrumb naturale**: Tab name basta per context (no nested hierarchy)

#### 9.4.2 Code Simplification
- ✅ **No container overhead**: Rimosso `VSMManagementWindow` wrapper class
- ✅ **Single notebook**: Logica tab change unificata in `get_current_tree_and_status()`
- ✅ **Reusable method**: `_create_vsm_event_sheet()` riutilizzato 3 volte (DRY principle)

#### 9.4.3 Toolbar Unification Potential
- 🔄 **Future**: Toolbar DataFlow può diventare context-aware
  - "Nuovo" → Apre VSMEventDialog se VSM tab attiva
  - "Export" → Export VSM KPI se VSM tab attiva
  - "Actions" → Menu VSM-specific se VSM tab attiva

### 9.5 Svantaggi/Trade-offs

#### 9.5.1 Loss of Encapsulation
- ⚠️ Logica VSM ora sparsa in `dataflow.py` (file già 5400+ righe)
- ⚠️ No più modulo self-contained (testing più difficile)
- ⚠️ Accoppiamento DataFlow ↔ VSM aumentato

#### 9.5.2 Backward Compatibility
- ⚠️ `VSMManagementWindow` diventa **DEPRECATED** (ma ancora presente)
- ⚠️ Codice duplicato temporaneamente (Step 4B usa estratto, originale rimane)
- 🔄 **Future cleanup**: Rimuovere `VSMManagementWindow` una volta Step 4D completato

#### 9.5.3 Incremental Migration Risk
- ⚠️ Step 4B attuale: **UI-only** (no logica) = tab VSM non funzionali
- ⚠️ Rischio regressione se step intermedi abbandonati
- 🔄 **Mitigation**: Tag git per ogni step, rollback agevole

---

## 10. Raccomandazioni per Prossimi Step

### 10.1 Step 4C - Data Population (PRIORITÀ ALTA)

#### Obiettivo
Popolare i 3 sheet VSM con dati reali da database.

#### Tasks
1. **Estrarre `refresh_events()`** da VSMManagementWindow
   - Adattare per ricevere `event_type` come parametro
   - Chiamare da `_load_vsm_events(event_type, sheet)`

2. **Estrarre `_populate_sheet()`** da VSMManagementWindow
   - Reusabile as-is (sheet + events → formatted rows)
   - Verificare metadata sync con sheet correct

3. **Initial Load**
   - Chiamare durante `__init__` per ogni tab
   - O lazy load al primo accesso (performance migliore)

4. **Tab Change Handler**
   - Bind `<<NotebookTabChanged>>` in DataFlow
   - Refresh sheet VSM corrente quando tab attiva

#### Code Esempio
```python
def _load_vsm_events(self, event_type, sheet):
    """Carica eventi VSM per un tipo specifico."""
    try:
        with DatabaseManager(get_db_path()) as db_manager:
            all_events = db_manager.get_all_vsm_events(username=self.current_username)
        
        filtered_events = [e for e in all_events if e.event_type == event_type]
        self._populate_vsm_sheet(sheet, filtered_events)
        
    except DatabaseError as e:
        logger.error(f"Errore caricamento VSM {event_type}: {e}")
        messagebox.showerror(_("Errore Database"), ...)

def _populate_vsm_sheet(self, sheet, events):
    """Popolamento sheet VSM (estratto da VSMManagementWindow)."""
    # ... copy logic from VSMManagementWindow._populate_sheet() ...
```

### 10.2 Step 4D - Event Handlers (PRIORITÀ ALTA)

#### Obiettivo
Connettere operazioni CRUD alla toolbar principale DataFlow (context-aware).

#### Tasks
1. **Context Detection Enhancement**
   ```python
   def get_current_context(self):
       tree, status = self.get_current_tree_and_status()
       if status.startswith('vsm_'):
           return 'vsm', status.split('_')[1]  # ('vsm', 'saving')
       else:
           return 'rfq', status  # ('rfq', 'attiva')
   ```

2. **Nuovo Button Integration**
   ```python
   def on_btn_nuovo_click(self):
       context, subtype = self.get_current_context()
       if context == 'vsm':
           self._vsm_create_event(event_type=subtype.title())
       else:
           self.open_request_window()  # RFQ flow esistente
   ```

3. **Actions Menu Context-Aware**
   ```python
   def show_actions_menu(self):
       context, _ = self.get_current_context()
       if context == 'vsm':
           menu = self._build_vsm_actions_menu()  # Modifica, Elimina, KPI
       else:
           menu = self._build_rfq_actions_menu()  # Menu esistente
       menu.post(event.x_root, event.y_root)
   ```

4. **Double-Click Binding**
   ```python
   # In _create_vsm_event_sheet():
   sheet.bind("<Double-Button-1>", lambda e: self._on_vsm_sheet_double_click(sheet, e))
   
   def _on_vsm_sheet_double_click(self, sheet, event):
       # Ottieni row selezionata
       # Estrai event_id da metadata
       # Apri VSMEventDialog in edit mode
   ```

#### Estratti da VSMManagementWindow da Riutilizzare
- `on_new_event()` → `_vsm_create_event()`
- `on_edit_event()` → `_vsm_edit_event()`
- `on_delete_event()` → `_vsm_delete_event()`

### 10.3 Step 4E - Export Integration (PRIORITÀ MEDIA)

#### Obiettivo
Aggiungere export VSM a `mega_export_excel()`.

#### Tasks
1. **Rimuovere Guard**
   ```python
   # Da:
   if status.startswith('vsm_'):
       messagebox.showinfo("Export Non Disponibile", ...)
       return
   
   # A:
   if status.startswith('vsm_'):
       self._export_vsm_to_excel(status)
       return
   ```

2. **Export VSM Sheet Structure**
   ```python
   def _export_vsm_to_excel(self, status):
       # Ottieni event_type da status
       event_type = status.split('_')[1].title()  # 'vsm_saving' → 'Saving'
       
       # Carica eventi
       events = self._get_vsm_events(event_type)
       
       # Crea Excel con openpyxl
       wb = Workbook()
       ws = wb.active
       ws.title = event_type
       
       # Headers (9 colonne: 8 sheet + 1 extra "Valore Effettivo")
       ws.append(["Data", "Tipo", "Azione", "Descrizione Completa", 
                  "Valore Teorico", "Realizzo %", "Valore Effettivo", 
                  "Ripetitivo", "Utente"])
       
       # Rows
       for event in events:
           valore_teorico = event.calculate_theoretical_value()
           valore_effettivo = valore_teorico * (event.percent_realizzo / 100)
           ws.append([
               event.event_date.strftime("%d/%m/%Y"),
               event.event_type,
               event.action,
               event.description or event.reference,  # ⚠️ FULL text, no truncate!
               valore_teorico,
               event.percent_realizzo,
               valore_effettivo,  # ✅ Calcolato (non in sheet UI)
               "Sì" if event.opex_ripetitivo else "No",
               event.username
           ])
       
       # Formattazione
       # ... borders, bold headers, number formats, etc.
       
       # Save dialog
       filepath = filedialog.asksaveasfilename(
           defaultextension=".xlsx",
           filetypes=[("Excel", "*.xlsx")],
           initialfile=f"VSM_{event_type}_{date.today().isoformat()}.xlsx"
       )
       if filepath:
           wb.save(filepath)
           messagebox.showinfo("Successo", f"Esportato {len(events)} eventi.")
   ```

#### Features Extra
- **Multi-sheet export**: 1 workbook con 3 sheets (Saving, Cost Avoidance, Derisking)
- **Summary sheet**: Aggregazioni (totale teorico/effettivo per tipo)
- **Grafici embedded**: Chart valore teorico vs effettivo per mese

### 10.4 Step 4F - KPI Dashboard (PRIORITÀ BASSA)

#### Obiettivo
Implementare funzionalità KPI (attualmente placeholder).

#### Features Proposte
1. **Aggregazioni Temporali**
   - Filtro data range (da/a)
   - Aggregazione mensile/trimestrale/annuale
   - Tabella pivot: Tipo × Mese × Somma(Valore Teorico/Effettivo)

2. **Confronto Cross-Type**
   - Bar chart: Saving vs Cost Avoidance vs Derisking
   - Stacked chart: Teorico vs Effettivo per tipo

3. **Realizzo Analysis**
   - Media `percent_realizzo` per tipo
   - Distribuzione realizzo (histogram)
   - Eventi sotto-performing (<50% realizzo)

4. **Trend Analysis**
   - Line chart: Valore effettivo nel tempo
   - Forecast lineare (se sample size sufficiente)

#### UI Proposta
- **Opzione 1**: Dialog modale KPI (come attuale placeholder)
- **Opzione 2**: Sidebar collapsible in DataFlow
- **Opzione 3**: Tab dedicata "VSM Analytics" (5th tab after Derisking)

### 10.5 Refactoring & Cleanup (PRIORITÀ DOPO 4D)

#### 10.5.1 Rimozione VSMManagementWindow
```python
# TODO: Dopo Step 4D completato, eliminare:
# - ui/windows/vsm_management_window.py (444 righe)
# - Import in dataflow.py (linea 97, ora commentato)
# - Tests specifici VSMManagementWindow (se presenti)
```

#### 10.5.2 Code Organization
**Problema**: `dataflow.py` sta diventando monolitico (5400+ righe).

**Soluzioni**:
1. **Extract VSM Module**
   ```python
   # Creare: ui/components/vsm_integration.py
   class VSMIntegration:
       def __init__(self, parent_dataflow):
           self.dataflow = parent_dataflow
       
       def create_tabs(self, notebook): ...
       def load_events(self, event_type): ...
       def on_create_event(self): ...
       def on_edit_event(self): ...
       def on_delete_event(self): ...
   
   # In dataflow.py:
   self.vsm = VSMIntegration(self)
   self.vsm.create_tabs(self.notebook)
   ```

2. **Mixins per Feature Isolation**
   ```python
   class DataFlowVSMMixin:
       """Mixin per funzionalità VSM integrate."""
       def _create_vsm_event_sheet(self, parent): ...
       def _load_vsm_events(self, event_type, sheet): ...
       # ... altri metodi VSM
   
   class DataFlow(tk.Tk, DataFlowVSMMixin, DataFlowRFQMixin):
       pass
   ```

#### 10.5.3 Testing
**Gap Attuale**: Nessun test automatico menzionato.

**Test Suite Proposto**:
```python
# tests/test_vsm_integration.py
class TestVSMIntegration(unittest.TestCase):
    def test_sheet_creation(self):
        """Verifica _create_vsm_event_sheet() crea 8 colonne."""
        ...
    
    def test_metadata_sync(self):
        """Verifica metadata allineati con rows dopo populate."""
        ...
    
    def test_context_detection(self):
        """Verifica get_current_tree_and_status() per VSM tabs."""
        ...
    
    def test_ownership_enforcement(self):
        """Verifica edit/delete bloccato per eventi non-mine."""
        ...
```

---

## 11. Conclusioni

### 11.1 Stato Attuale VSMManagementWindow

**✅ Componente Maturo**:
- UI completa con 3 tab isolate per event type
- CRUD operations funzionali con ownership validation
- Metadata tracking sincronizzato con sheet rows
- Error handling robusto (DatabaseError, VSMError)
- i18n support completo

**⚠️ Limitazioni Evidenti**:
- No ricerca/filtri
- No export dati (KPI placeholder)
- Truncamento descrizione perdita informazione
- Refresh completo sempre (no incremental)
- No caching (performance concern con molti eventi)
- Nested navigation (UX subottimale)

### 11.2 Step 4B - Stato Integrazione

**✅ Obiettivi Raggiunti**:
- Flat tab structure (Saving/Cost Avoidance/Derisking dirette)
- UI extraction senza riscrittura (DRY principle)
- Context detection aggiornato (`get_current_tree_and_status()`)
- Guards preventive per logica non ancora connessa

**🔄 Lavoro Incrementale Successivo**:
- **Step 4C**: Data loading (popolamento sheet)
- **Step 4D**: Event handlers (CRUD integration)
- **Step 4E**: Export integration
- **Step 4F**: KPI dashboard

### 11.3 Approccio Conservativo Validato

**Principio Followed**: "Riutilizzare, non reinventare"

**Vantaggi Strategici**:
1. ✅ **Low-risk migration**: UI provata, logica testata → extraction only
2. ✅ **Rollback easy**: Ogni step committable separatamente
3. ✅ **No big-bang**: Incrementalità permette testing frequente
4. ✅ **Code preservation**: VSMManagementWindow rimane reference fino a Step 4D

**Lesson Learned**:
- Tentativo iniziale Step 4 completo (10 fasi) → troppo ambizioso → regress risk
- Approccio corretto: Step 4A (tab only) → Step 4B (UI only) → Step 4C (data only) → etc.
- Granularità fine-grained = migliore controllo qualità

---

## 12. Appendice A - Mappatura File Sorgenti

| File | Righe | Responsabilità | Dipendenze Chiave |
|------|-------|----------------|-------------------|
| `ui/windows/vsm_management_window.py` | 444 | VSM UI originale (deprecated post-4D) | DatabaseManager, VSMEventDialog |
| `dataflow.py` | 5400+ | Main application window + RFQ + VSM integrato | Tutti i moduli (monolitico) |
| `ui/dialogs/vsm_event_dialog.py` | ? | Form CRUD evento VSM | VSMEvent model |
| `models/vsm_event.py` | ? | Domain model VSM | - |
| `database_manager.py` | ? | Database access layer | SQLite |
| `services/vsm_persistence.py` | ? | VSM-specific DB operations | DatabaseManager |

---

## 13. Appendice B - Glossario Termini VSM

| Termine | Definizione |
|---------|-------------|
| **VSM** | Value Stream Mapping - Metodologia Lean per mapping flussi valore |
| **Saving** | Risparmio monetario diretto (riduzione costi) |
| **Cost Avoidance** | Evitamento costi futuri (prevenzione) |
| **Derisking** | Riduzione rischi operativi/finanziari |
| **Valore Teorico** | Potenziale risparmio/evitamento max (100% realizzo) |
| **Realizzo %** | Percentuale effettiva realizzata del valore teorico |
| **Valore Effettivo** | `Valore Teorico × (Realizzo % / 100)` |
| **OPEX Ripetitivo** | Operational Expenditure ricorrente (monthly/quarterly) |
| **Freq Annua** | Frequenza annuale evento ripetitivo (12 = mensile, 4 = trimestrale) |
| **Ownership** | Proprietà evento (solo creator può edit/delete) |

---

**Fine Documento - Versione 1.0**  
*Generato: 25 marzo 2026*  
*Autore: GitHub Copilot (Claude Sonnet 4.5)*  
*Target Audience: Dev Team VSM Project*
