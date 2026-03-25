# STEP 4C — CARICAMENTO DATI VSM (ESTRAZIONE SENZA RISCRITTURA)

## CONTESTO
- Le tab VSM (Saving, Cost Avoidance, Derisking) sono già state create in dataflow.py
- Ogni tab ha già una tksheet inizializzata
- Il file vsm_management_window.py contiene già la logica funzionante di caricamento dati e popolamento sheet

## OBIETTIVO
Riutilizzare la logica esistente di VSMManagementWindow per popolare le nuove sheet in dataflow.py.

## REGOLE FONDAMENTALI
- NON riscrivere la logica
- NON semplificare
- NON cambiare comportamento
- SOLO estrarre e adattare il minimo indispensabile
- NON toccare la logica RFQ esistente

## MODIFICHE DA APPLICARE

### MODIFICA 1: Import VSMEvent Model

**File**: `dataflow.py`  
**Posizione**: Dopo riga 123 (dopo import `ui.dialogs.common_dialogs`)

```python
from ui.dialogs.common_dialogs import (
    LanguagePrompt,
    NewRdOTypeDialog,
    UserIdentityDialog,
    CopyProgressWindow,
    SplashScreen
)

# Import modelli VSM (Step 4C)
from models.vsm_event import VSMEvent

# Esegui pulizia all'avvio
```

---

### MODIFICA 2: Metodi di Caricamento Dati VSM

**File**: `dataflow.py`  
**Posizione**: Dopo `_create_vsm_event_sheet()` (~riga 4330, dopo il commento NOTA sui metodi sort)

```python
    # NOTA: I metodi sort_treeview_column e update_sort_indicators sono stati rimossi
    # perché tksheet ha funzionalità di ordinamento integrate che si abilitano automaticamente
    # con enable_bindings(). L'utente può cliccare sugli header delle colonne per ordinare.

    def _load_vsm_events(self, event_type, sheet):
        """
        Carica eventi VSM per un tipo specifico.
        
        ESTRATTO da VSMManagementWindow.refresh_events() (Step 4C).
        
        Args:
            event_type: Tipo evento ("Saving"|"Cost Avoidance"|"Derisking")
            sheet: Widget tksheet da popolare
        """
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                # Carica tutti eventi per utente corrente
                all_events = db_manager.get_all_vsm_events(username=self.current_username)
            
            # Filtra per event_type specificato
            filtered_events = [e for e in all_events if e.event_type == event_type]
            
            # Popola sheet
            self._populate_vsm_sheet(sheet, filtered_events)
            
            logger.debug(f"Caricati {len(filtered_events)} eventi VSM {event_type}")
            
        except DatabaseError as e:
            logger.error(f"Errore caricamento eventi VSM {event_type}: {e}")
            messagebox.showerror(
                _("Errore Database"),
                _("Impossibile caricare gli eventi VSM: {}\n").format(e),
                parent=self
            )
    
    def _populate_vsm_sheet(self, sheet, events):
        """
        Popola un tksheet con lista di eventi VSM.
        
        ESTRATTO da VSMManagementWindow._populate_sheet() (Step 4C).
        
        Args:
            sheet: Widget tksheet da popolare
            events: Lista di VSMEvent
        """
        data_rows = []
        metadata = []
        
        for event in events:
            # Calcola valore teorico
            valore_teorico = event.calculate_theoretical_value()
            
            # Formatta row
            row = [
                event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
                event.event_type,
                event.action,
                (event.description or event.reference or "")[:50],  # Truncate
                f"€ {valore_teorico:,.2f}",
                f"{event.percent_realizzo:.0f}%",
                "✓" if event.opex_ripetitivo else "",
                event.username
            ]
            data_rows.append(row)
            
            # Metadata per ownership e event_id
            metadata.append({
                'event_id': event.id,
                'username': event.username,
                'is_mine': event.username == self.current_username
            })
        
        # Aggiorna sheet
        sheet.set_sheet_data(data_rows)
        sheet._event_metadata = metadata

    def _get_selected_row_indices(self, sheet):
```

---

### MODIFICA 3: Caricamento Iniziale Dati VSM

**File**: `dataflow.py`  
**Posizione**: Dopo creazione sheet VSM (~riga 3694)

**TROVA**:
```python
        # Step 4B: Riutilizzo UI VSM esistente (estratta da VSMManagementWindow)
        # Crea sheet VSM per ogni tab usando la stessa struttura della UI originale
        self.sheet_saving = self._create_vsm_event_sheet(self.tab_saving)
        self.sheet_cost_avoidance = self._create_vsm_event_sheet(self.tab_cost_avoidance)
        self.sheet_derisking = self._create_vsm_event_sheet(self.tab_derisking)
        
        footer_frame = ttk.Frame(self.root); footer_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
```

**SOSTITUISCI CON**:
```python
        # Step 4B: Riutilizzo UI VSM esistente (estratta da VSMManagementWindow)
        # Crea sheet VSM per ogni tab usando la stessa struttura della UI originale
        self.sheet_saving = self._create_vsm_event_sheet(self.tab_saving)
        self.sheet_cost_avoidance = self._create_vsm_event_sheet(self.tab_cost_avoidance)
        self.sheet_derisking = self._create_vsm_event_sheet(self.tab_derisking)
        
        # Step 4C: Caricamento iniziale dati VSM (estratto da VSMManagementWindow.refresh_events)
        # Popola ogni sheet con i dati correnti dell'utente
        self._load_vsm_events("Saving", self.sheet_saving)
        self._load_vsm_events("Cost Avoidance", self.sheet_cost_avoidance)
        self._load_vsm_events("Derisking", self.sheet_derisking)
        
        footer_frame = ttk.Frame(self.root); footer_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
```

---

## RIEPILOGO STEP 4C

### Modifiche Applicate
1. ✅ Import `VSMEvent` model (per coerenza e type hints)
2. ✅ Metodo `_load_vsm_events(event_type, sheet)` - **ESTRATTO** da `VSMManagementWindow.refresh_events()`
3. ✅ Metodo `_populate_vsm_sheet(sheet, events)` - **ESTRATTO** da `VSMManagementWindow._populate_sheet()`
4. ✅ Caricamento iniziale dati per le 3 tab VSM

### Adattamenti Minimi
- `refresh_events()` → `_load_vsm_events(event_type, sheet)` - riceve parametri invece di usare `self.current_event_type` e `self.sheets[...]`
- `_populate_sheet()` → `_populate_vsm_sheet()` - identico, usa `self.current_username` già disponibile
- Nessun cambio logico, solo parametrizzazione

### Logica Riutilizzata
- ✅ Query DB identica: `get_all_vsm_events(username=current_username)`
- ✅ Filtro identico: `[e for e in all_events if e.event_type == event_type]`
- ✅ Formattazione identica: Data dd/mm/yyyy, Valore €X,XXX.XX, Realizzo XX%, Ripetitivo ✓
- ✅ Metadata identici: `{event_id, username, is_mine}`
- ✅ Truncazione descrizione identica: `[:50]`

### Risultato Atteso
- Le 3 tab VSM (Saving, Cost Avoidance, Derisking) mostrano dati reali all'apertura
- Comportamento identico a `VSMManagementWindow` originale
- Zero regressioni su logica RFQ esistente

### NON Implementato (Intentional)
- ❌ Refresh su cambio tab (Step 4D)
- ❌ CRUD handlers (Step 4D)
- ❌ Double-click bindings (Step 4D)
- ❌ Export/KPI (Step 4E/4F)

---

## SOURCE REFERENCE

### VSMManagementWindow.refresh_events() (Linee 201-227)
```python
def refresh_events(self):
    """
    Ricarica gli eventi VSM per la tab corrente.
    Filtra per event_type corrispondente alla tab attiva.
    """
    try:
        with DatabaseManager(get_db_path()) as db_manager:
            # Carica tutti eventi per utente corrente
            all_events = db_manager.get_all_vsm_events(username=self.current_username)
        
        # Filtra per event_type della tab corrente
        filtered_events = [e for e in all_events if e.event_type == self.current_event_type]
        
        # Ottieni sheet corrente
        sheet = self.sheets[self.current_event_type]
        
        # Popola sheet
        self._populate_sheet(sheet, filtered_events)
        
        logger.debug(f"Caricati {len(filtered_events)} eventi {self.current_event_type}")
        
    except DatabaseError as e:
        logger.error(f"Errore caricamento eventi VSM: {e}")
        messagebox.showerror(
            _("Errore Database"),
            _("Impossibile caricare gli eventi VSM: {}\n").format(e),
            parent=self
        )
```

### VSMManagementWindow._populate_sheet() (Linee 229-267)
```python
def _populate_sheet(self, sheet, events):
    """
    Popola un tksheet con lista di eventi VSM.
    
    Args:
        sheet: Widget tksheet da popolare
        events: Lista di VSMEvent
    """
    data_rows = []
    metadata = []
    
    for event in events:
        # Calcola valore teorico
        valore_teorico = event.calculate_theoretical_value()
        
        # Formatta row
        row = [
            event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
            event.event_type,
            event.action,
            (event.description or event.reference or "")[:50],  # Truncate
            f"€ {valore_teorico:,.2f}",
            f"{event.percent_realizzo:.0f}%",
            "✓" if event.opex_ripetitivo else "",
            event.username
        ]
        data_rows.append(row)
        
        # Metadata per ownership e event_id
        metadata.append({
            'event_id': event.id,
            'username': event.username,
            'is_mine': event.username == self.current_username
        })
    
    # Aggiorna sheet
    sheet.set_sheet_data(data_rows)
    sheet._event_metadata = metadata
```

---

## CHECKLIST IMPLEMENTAZIONE

- [ ] Import `VSMEvent` model aggiunto
- [ ] Metodo `_load_vsm_events()` creato
- [ ] Metodo `_populate_vsm_sheet()` creato
- [ ] Caricamento iniziale delle 3 tab implementato
- [ ] Test manuale: verificare dati VSM visibili all'apertura
- [ ] Test manuale: verificare metadata sincronizzati
- [ ] Verificare nessuna regressione RFQ (tab Attive/Archiviate)
- [ ] Verificare logs: `logger.debug(f"Caricati {N} eventi VSM {type}")`

---

## TESTING PLAN

### Test Case 1: Caricamento Iniziale
1. Avviare DataFlow
2. Verificare tab "Saving" mostra eventi Saving dell'utente
3. Verificare tab "Cost Avoidance" mostra eventi Cost Avoidance
4. Verificare tab "Derisking" mostra eventi Derisking
5. Verificare formattazione colonne corretta (date, valori €, %)

### Test Case 2: Metadata Sync
1. Selezionare una riga in sheet VSM
2. Verificare `sheet._event_metadata[row_idx]` contiene `event_id`, `username`, `is_mine`
3. Verificare `is_mine=True` per eventi propri, `False` per altri

### Test Case 3: No Regression RFQ
1. Verificare tab "RdO Attive" funziona normalmente
2. Verificare tab "RdO Archiviate" funziona normalmente
3. Verificare filtri RFQ funzionano
4. Verificare operazioni RFQ (nuovo, modifica, delete) funzionano

### Test Case 4: Error Handling
1. Simulare errore DB (es: DB file mancante)
2. Verificare messagebox error mostrato
3. Verificare app non crasha

---

## NEXT STEPS (Future)

### Step 4D - Event Handlers
- Estrarre `on_new_event()`, `on_edit_event()`, `on_delete_event()`
- Integrare con toolbar principale (context-aware)
- Aggiungere double-click binding per edit

### Step 4E - Export Integration
- Rimuovere guard in `mega_export_excel()`
- Implementare export VSM (9 colonne con valore effettivo)

### Step 4F - KPI Dashboard
- Implementare funzionalità KPI (attualmente placeholder)
- Aggregazioni temporali, grafici, confronti cross-type
