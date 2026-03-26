# PIANO RETTIFICATO - DRIVER "PAGAMENTI" VSM

## A. Cosa Confermo del Piano Originario

### ✓ Validazioni Architetturali (Corrette)
1. **Database Schema Completo**: Tabella `vsm_events` contiene tutti i campi necessari per Pagamenti (spending_annuo, giorni_pagamento_attuali, giorni_pagamento_negoziati, driver)
2. **Motore VSM Driver-Agnostic**: `services/vsm_engine.py` completamente riutilizzabile, nessuna modifica necessaria
3. **Layer Persistenza Stabile**: `database_manager.py` e `vsm_persistence.py` già pronti, zero modifiche
4. **Pattern UI Dinamica**: Metodo `_on_event_type_changed()` in `vsm_event_dialog.py` è template perfetto per `_on_driver_changed()`
5. **Separazione Driver**: Conferma alternatività rigida tra Prezzo e Pagamenti (no coesistenza campi)

### ✓ Nomenclatura Campi (Definitiva)
- `spending_annuo` (REAL) → spesa annuale
- `giorni_pagamento_attuali` (INTEGER) → termini pagamento baseline
- `giorni_pagamento_negoziati` (INTEGER) → termini pagamento negoziati
- **NON usare** "bdg" per Pagamenti (riservato a driver Prezzo)

### ✓ Approccio Incrementale (Confermato)
- FASE 1: Backend calcolo (separato da UI)
- FASE 2: UI dinamica
- FASE 3: Testing e stabilizzazione

### ✓ Gestione Delta Negativo (Corretta)
- Se `delta_giorni < 0`: risultato negativo (peggioramento)
- NON bloccare calcolo, NON forzare a zero
- Rappresenta loss implicita da peggioramento condizioni

---

## B. Cosa Correggo/Modifico

### ❌ ELIMINATO: Coefficiente Hardcoded in constants.py
**Errore Originale**: Piano proponeva `VSM_PAGAMENTI_COEFFICIENT = 2.0` in `constants.py`

**Problema**: 
- Valore arbitrario hardcoded
- Modificabile solo ricompilando
- `constants.py` contiene solo costanti UI/layout, non business logic

**Correzione**: Vedi sezione C (coefficiente configurabile)

---

### ❌ CORRETTO: Formula senza % Realizzo per Pagamenti
**Errore Originale**: Piano citava `calculate_effective_value()` applicando `percent_realizzo` anche a Pagamenti

**Problema**:
- % Realizzo NON applicabile a driver Pagamenti
- Pagamenti sono o accordati o non lo sono (binario)
- Concetto fuzzy di realizzo valido solo per Prezzo

**Correzione**:
```python
def calculate_theoretical_value(self) -> float:
    if self.driver == "Prezzo":
        # Formula esistente
        if self.importo_bdg is not None and self.importo_negoziato is not None:
            return self.importo_bdg - self.importo_negoziato
        return 0.0
    
    elif self.driver == "Pagamenti":
        # Formula senza percent_realizzo
        coeff = self._get_pagamenti_coefficient()  # Da implementare
        if (self.spending_annuo is not None and 
            self.giorni_pagamento_attuali is not None and 
            self.giorni_pagamento_negoziati is not None):
            
            delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
            saving_percentuale = (delta_giorni / 30) * coeff
            return self.spending_annuo * saving_percentuale
        return 0.0
    
    else:
        return 0.0

def calculate_effective_value(self) -> float:
    """Valore effettivo: solo Prezzo usa percent_realizzo, Pagamenti no."""
    theoretical = self.calculate_theoretical_value()
    
    # percent_realizzo SOLO per driver Prezzo
    if self.driver == "Prezzo":
        return theoretical * (self.percent_realizzo / 100.0)
    else:
        # Pagamenti, Volume, Altro: valore teorico = effettivo
        return theoretical
```

**Nota Formula**: Definire chiaramente unità coefficiente:
- Se `coeff = 0.005` → è già 0.5%, formula diretta: `spending_annuo * (delta_giorni/30) * coeff`
- Se `coeff = 0.5` → è 0.5%, formula: `spending_annuo * (delta_giorni/30) * (coeff/100)`

**Raccomandazione**: Usare formato 0.005 (percentuale decimale) per evitare conversioni ambigue.

---

### ❌ CORRETTO: UI senza % Realizzo per Pagamenti
**Errore Originale**: Piano non specificava rimozione esplicita campo % Realizzo

**Correzione**:
Quando `driver == "Pagamenti"`:
- Campo `entry_percent_realizzo` → `grid_remove()` (nascosto)
- Label `lbl_percent_realizzo` → `grid_remove()` (nascosta)
- NON leggere valore percent_realizzo durante salvataggio
- Salvare sempre `percent_realizzo = 100` per coerenza DB (ignorato nel calcolo)

```python
def _on_driver_changed(self, event=None):
    driver = self.combo_driver.get()
    
    if driver == "Prezzo":
        # Mostra campi Prezzo
        self.lbl_importo_bdg.grid(...)
        self.entry_importo_bdg.grid(...)
        self.lbl_importo_negoziato.grid(...)
        self.entry_importo_negoziato.grid(...)
        self.lbl_percent_realizzo.grid(...)  # Visibile per Prezzo
        self.entry_percent_realizzo.grid(...)
        
        # Nascondi campi Pagamenti
        self.lbl_spending_annuo.grid_remove()
        self.entry_spending_annuo.grid_remove()
        self.lbl_giorni_attuali.grid_remove()
        self.entry_giorni_attuali.grid_remove()
        self.lbl_giorni_negoziati.grid_remove()
        self.entry_giorni_negoziati.grid_remove()
        
    elif driver == "Pagamenti":
        # Nascondi campi Prezzo
        self.lbl_importo_bdg.grid_remove()
        self.entry_importo_bdg.grid_remove()
        self.lbl_importo_negoziato.grid_remove()
        self.entry_importo_negoziato.grid_remove()
        self.lbl_percent_realizzo.grid_remove()  # NASCOSTO per Pagamenti
        self.entry_percent_realizzo.grid_remove()
        
        # Mostra campi Pagamenti
        self.lbl_spending_annuo.grid(...)
        self.entry_spending_annuo.grid(...)
        self.lbl_giorni_attuali.grid(...)
        self.entry_giorni_attuali.grid(...)
        self.lbl_giorni_negoziati.grid(...)
        self.entry_giorni_negoziati.grid(...)
```

---

### ❌ CORRETTO: Salvataggio Pulito (No Record Ibridi)
**Errore Originale**: Piano non garantiva pulizia campi non pertinenti

**Correzione**:
Durante salvataggio in `_validate_and_save()`:

```python
def _prepare_event_data(self):
    """Prepara dati evento puliti in base a driver."""
    driver = self.combo_driver.get()
    
    # Campi comuni
    event_data = {
        'event_date': self.entry_date.get_date(),
        'username': self.current_username,
        'event_type': self.event_type_var.get(),
        'action': self.action_var.get(),
        'description': self.text_description.get("1.0", "end-1c"),
        'driver': driver,
        'opex_ripetitivo': self.opex_ripetitivo_var.get(),
        'note': self.text_note.get("1.0", "end-1c") if hasattr(self, 'text_note') else "",
    }
    
    if driver == "Prezzo":
        # Solo campi Prezzo
        event_data.update({
            'importo_bdg': float(self.entry_importo_bdg.get()),
            'importo_negoziato': float(self.entry_importo_negoziato.get()),
            'percent_realizzo': float(self.entry_percent_realizzo.get()),
            # Campi Pagamenti a NULL
            'spending_annuo': None,
            'giorni_pagamento_attuali': None,
            'giorni_pagamento_negoziati': None,
        })
        
    elif driver == "Pagamenti":
        # Solo campi Pagamenti
        event_data.update({
            'spending_annuo': float(self.entry_spending_annuo.get()),
            'giorni_pagamento_attuali': int(self.entry_giorni_attuali.get()),
            'giorni_pagamento_negoziati': int(self.entry_giorni_negoziati.get()),
            'percent_realizzo': 100.0,  # Default fisso, ignorato nel calcolo
            # Campi Prezzo a NULL/zero
            'importo_bdg': None,
            'importo_negoziato': None,
        })
    
    return event_data
```

**Obiettivo**: Record DB puliti, no ambiguità, campi non pertinenti a NULL.

---

## C. Gestione Coefficiente Configurabile

### Soluzione Proposta: Tabella `vsm_settings` (Minima e Conservativa)

**Approccio**:
1. Creare tabella `vsm_settings` nel database
2. Popolare con valore default al primo avvio
3. Metodi get/set per leggere/scrivere coefficiente
4. **NON** sporcare `dataflow.py`
5. **NON** hardcodare valore nel codice

### Implementazione

#### 1. Schema Database (`database_manager.py`)
```python
def _initialize_vsm_tables(self):
    """Inizializza tabelle VSM inclusa configurazione."""
    
    # ... tabelle vsm_events, vsm_impacts esistenti ...
    
    # Nuova tabella settings
    self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS vsm_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT NOT NULL,
            description TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Popola valore default se non esiste
    self.cursor.execute('''
        INSERT OR IGNORE INTO vsm_settings (setting_key, setting_value, description)
        VALUES ('pagamenti_coefficient', '0.005', 'Coefficiente costo opportunità per driver Pagamenti (0.005 = 0.5% ogni 30 giorni)')
    ''')
    
    self.conn.commit()
```

#### 2. Metodi Accesso Settings (`database_manager.py`)
```python
def get_vsm_setting(self, key: str, default: str = None) -> str:
    """
    Recupera un setting VSM dal database.
    
    Args:
        key: Chiave del setting
        default: Valore di default se non trovato
        
    Returns:
        str: Valore del setting
    """
    try:
        self.cursor.execute(
            'SELECT setting_value FROM vsm_settings WHERE setting_key = ?',
            (key,)
        )
        result = self.cursor.fetchone()
        return result[0] if result else default
    except Exception as e:
        logger.error(f"Errore lettura setting {key}: {e}")
        return default

def set_vsm_setting(self, key: str, value: str, description: str = None) -> bool:
    """
    Imposta un setting VSM nel database.
    
    Args:
        key: Chiave del setting
        value: Valore da salvare
        description: Descrizione opzionale
        
    Returns:
        bool: True se successo
    """
    try:
        self.cursor.execute('''
            INSERT OR REPLACE INTO vsm_settings (setting_key, setting_value, description, updated_at)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ''', (key, value, description))
        self.conn.commit()
        return True
    except Exception as e:
        logger.error(f"Errore scrittura setting {key}: {e}")
        return False
```

#### 3. Metodo Helper in `vsm_event.py`
```python
def _get_pagamenti_coefficient(self) -> float:
    """
    Recupera il coefficiente Pagamenti dal database.
    
    Returns:
        float: Coefficiente (default 0.005 se non configurato)
    """
    try:
        from database_manager import DatabaseManager
        from services.app_paths import get_db_path
        
        db_path = get_db_path()
        db_manager = DatabaseManager(db_path)
        coeff_str = db_manager.get_vsm_setting('pagamenti_coefficient', '0.005')
        db_manager.close()
        
        return float(coeff_str)
    except Exception:
        # Fallback safe se DB non disponibile
        return 0.005
```

### Vantaggi Soluzione
- ✓ Coefficiente configurabile runtime (no ricompilazione)
- ✓ Struttura minimalista (una sola tabella generica)
- ✓ Valore default sensato (0.5% mensile)
- ✓ Estendibile per futuri settings VSM
- ✓ NON inquina `dataflow.py` o `constants.py`
- ✓ Isolato in zona VSM del database

### Future Enhancement (Opzionale)
Creare una UI settings VSM per modificare coefficiente:
- Menu → Impostazioni VSM
- Form con campo numerico per coefficiente
- Salvataggio tramite `set_vsm_setting()`

---

## D. Implementazione UI Pagamenti senza % Realizzo

### Widget da Creare (in `_build_ui()`)
```python
# === SEZIONE DATI PAGAMENTI (nascosta di default) ===
# Memorizza riferimenti widget per show/hide dinamico

# Spending Annuo
self.lbl_spending_annuo = ttk.Label(
    self.economic_frame, 
    text=_("Spending Annuo (€): *")
)
self.entry_spending_annuo = ttk.Entry(self.economic_frame, width=20)

# Termini Pagamento Attuali
self.lbl_giorni_attuali = ttk.Label(
    self.economic_frame,
    text=_("Termini Pagamento Attuali (giorni): *")
)
self.entry_giorni_attuali = ttk.Entry(self.economic_frame, width=20)

# Termini Pagamento Negoziati
self.lbl_giorni_negoziati = ttk.Label(
    self.economic_frame,
    text=_("Termini Pagamento Negoziati (giorni): *")
)
self.entry_giorni_negoziati = ttk.Entry(self.economic_frame, width=20)

# Posizionamento iniziale (nascosti)
# NON fare .grid() ora, sarà fatto dinamicamente da _on_driver_changed()
```

### Binding su Combobox Driver
```python
# In _build_ui(), dopo creazione combo_driver:
self.combo_driver.bind("<<ComboboxSelected>>", self._on_driver_changed)
```

### Validazioni Specifiche
```python
elif driver == "Pagamenti":
    # Validazioni Pagamenti
    if not self.entry_spending_annuo.get().strip():
        messagebox.showerror("Errore", _("Spending Annuo richiesto per driver Pagamenti"))
        return False
    
    if not self.entry_giorni_attuali.get().strip():
        messagebox.showerror("Errore", _("Termini Pagamento Attuali richiesti"))
        return False
    
    if not self.entry_giorni_negoziati.get().strip():
        messagebox.showerror("Errore", _("Termini Pagamento Negoziati richiesti"))
        return False
    
    try:
        spending = float(self.entry_spending_annuo.get())
        giorni_att = int(self.entry_giorni_attuali.get())
        giorni_neg = int(self.entry_giorni_negoziati.get())
        
        if spending <= 0:
            messagebox.showerror("Errore", _("Spending Annuo deve essere positivo"))
            return False
        
        if giorni_att < 0 or giorni_neg < 0:
            messagebox.showerror("Errore", _("Giorni di pagamento non possono essere negativi"))
            return False
        
        # Opzionale: warning se delta negativo (peggioramento)
        if giorni_neg < giorni_att:
            risposta = messagebox.askyesno(
                "Attenzione",
                f"Termini negoziati ({giorni_neg}) inferiori a termini attuali ({giorni_att}).\n"
                f"Questo genera un impatto negativo (peggioramento).\n\n"
                f"Confermi di voler procedere?",
                icon='warning'
            )
            if not risposta:
                return False
        
    except ValueError:
        messagebox.showerror("Errore", _("Valori numerici non validi per campi Pagamenti"))
        return False
```

---

## E. Separazione Netta Prezzo/Pagamenti

### Garanzie Implementative

1. **UI Mutualmente Esclusiva**: Solo campi pertinenti visibili
   - `_on_driver_changed()` usa `grid()` / `grid_remove()` per toggle visibilità
   - Mai mostrare importi e giorni contemporaneamente

2. **Validazione Driver-Specific**: Branch espliciti in `_validate_and_save()`
   - `if driver == "Prezzo"`: valida solo importi
   - `elif driver == "Pagamenti"`: valida solo spending/giorni

3. **Salvataggio Pulito**: Campi non pertinenti a NULL
   - `_prepare_event_data()` popola solo campi del driver attivo
   - Campi altri driver: `None` o valori default ignorati

4. **Calcolo Isolato**: Branch separati in `calculate_theoretical_value()`
   - Prezzo: `importo_bdg - importo_negoziato`
   - Pagamenti: `spending_annuo * (delta_giorni/30) * coeff`
   - Zero overlap, zero ambiguità

5. **Motore VSM**: Driver-agnostic per design
   - Riceve solo valore calcolato, non conosce campi sottostanti
   - Nessuna logica condizionale su driver

---

## F. Prevenzione Ambiguità Dati Salvati

### Strategia: NULL Enforcement

**Problema**: Record DB con valori residui in campi non pertinenti al driver

**Soluzione**:
```python
def _prepare_event_data(self):
    """Prepara dati puliti basati su driver."""
    driver = self.combo_driver.get()
    
    # Base comune
    data = {...}
    
    if driver == "Prezzo":
        data.update({
            # Campi ATTIVI
            'importo_bdg': float(self.entry_importo_bdg.get()),
            'importo_negoziato': float(self.entry_importo_negoziato.get()),
            'percent_realizzo': float(self.entry_percent_realizzo.get()),
            
            # Campi INATTIVI → NULL esplicito
            'spending_annuo': None,
            'giorni_pagamento_attuali': None,
            'giorni_pagamento_negoziati': None,
        })
    
    elif driver == "Pagamenti":
        data.update({
            # Campi ATTIVI
            'spending_annuo': float(self.entry_spending_annuo.get()),
            'giorni_pagamento_attuali': int(self.entry_giorni_attuali.get()),
            'giorni_pagamento_negoziati': int(self.entry_giorni_negoziati.get()),
            
            # Campi INATTIVI → NULL esplicito
            'importo_bdg': None,
            'importo_negoziato': None,
            'percent_realizzo': 100.0,  # Valore fisso ignorato
        })
    
    return data
```

### Controllo Coerenza (Opzionale ma Raccomandato)
```python
def validate_driver_consistency(event: VSMEvent) -> bool:
    """Verifica coerenza campi rispetto a driver."""
    if event.driver == "Prezzo":
        # Prezzo richiede importi, non giorni
        if event.importo_bdg is None or event.importo_negoziato is None:
            return False
        if event.spending_annuo is not None or event.giorni_pagamento_attuali is not None:
            logger.warning(f"Evento {event.id}: campi Pagamenti non nulli con driver Prezzo")
            
    elif event.driver == "Pagamenti":
        # Pagamenti richiede spending/giorni, non importi
        if event.spending_annuo is None or event.giorni_pagamento_attuali is None:
            return False
        if event.importo_bdg is not None or event.importo_negoziato is not None:
            logger.warning(f"Evento {event.id}: campi Prezzo non nulli con driver Pagamenti")
    
    return True
```

---

## G. Sequenza Implementazione RETTIFICATA

### FASE 1: Configurazione Database (30-45 min)
**File**: `database_manager.py`

1. Aggiungere creazione tabella `vsm_settings` in `_initialize_vsm_tables()`
2. Popolare with default `pagamenti_coefficient = 0.005`
3. Implementare `get_vsm_setting()` e `set_vsm_setting()`
4. Test manuale: verificare tabella creata e valore default presente

**Deliverable**: Infrastruttura configurazione pronta, NO codice hardcoded

---

### FASE 2: Backend Calcolo (45 min - 1 ora)
**File**: `models/vsm_event.py`

1. Implementare `_get_pagamenti_coefficient()` helper method
2. Modificare `calculate_theoretical_value()` con branch `elif driver == "Pagamenti"`
   - Formula: `spending_annuo * (delta_giorni / 30) * coeff`
   - NO divisione `/100` se coeff già in forma decimale (0.005)
3. Modificare `calculate_effective_value()` per escludere `percent_realizzo` da Pagamenti
4. Test manuale in Python REPL

**Test Case**:
```python
from models.vsm_event import VSMEvent
from datetime import datetime

event = VSMEvent(
    id=None,
    event_date=datetime.now(),
    username="test_user",
    event_type="Saving",
    driver="Pagamenti",
    spending_annuo=120000.0,  # 120k€
    giorni_pagamento_attuali=30,
    giorni_pagamento_negoziati=60,  # +30 giorni
    opex_ripetitivo=True
)

theoretical = event.calculate_theoretical_value()
# Atteso: 120000 * (30/30) * 0.005 = 120000 * 1 * 0.005 = 600€
effective = event.calculate_effective_value()
# Atteso: 600€ (no percent_realizzo per Pagamenti)

print(f"Teorico: {theoretical}€")
print(f"Effettivo: {effective}€")
assert theoretical == 600.0
assert effective == 600.0
```

**Deliverable**: Calcolo Pagamenti funzionante, NO impatto su UI

---

### FASE 3: UI Dinamica (2-3 ore)
**File**: `ui/dialogs/vsm_event_dialog.py`

#### Step 3.1: Creazione Widget Pagamenti
- Aggiungere label e entry per spending_annuo, giorni_attuali, giorni_negoziati
- NON posizionare con `.grid()` subito
- Salvare riferimenti in `self.lbl_*` e `self.entry_*`

#### Step 3.2: Implementare `_on_driver_changed()`
- Logica show/hide basata su combobox driver
- `grid()` / `grid_remove()` per toggle visibilità
- Nascondere % Realizzo quando driver = Pagamenti

#### Step 3.3: Binding su Combobox
```python
self.combo_driver.bind("<<ComboboxSelected>>", self._on_driver_changed)
```

#### Step 3.4: Validazioni in `_validate_and_save()`
- Branch `if driver == "Prezzo"`: valida importi
- Branch `elif driver == "Pagamenti"`: valida spending/giorni
- Warning opzionale se delta_giorni < 0

#### Step 3.5: Salvataggio Pulito in `_prepare_event_data()`
- Popolare solo campi pertinenti al driver
- Campi non pertinenti → `None`
- `percent_realizzo = 100.0` fisso per Pagamenti (ignorato)

#### Step 3.6: Gestione Modalità EDIT
```python
# In _load_event_data(), dopo popolamento campi:
def _load_event_data(self):
    # ... caricamento evento ...
    # ... popolamento tutti i campi ...
    
    # Chiamare _on_driver_changed() per mostrare campi appropriati
    self._on_driver_changed()
```

**Testing FASE 3**:
1. CREATE: seleziona Prezzo → vedi campi importi
2. CREATE: seleziona Pagamenti → vedi campi spending/giorni, NO % Realizzo
3. CREATE: salva evento Pagamenti con spending=100000, giorni_att=30, giorni_neg=60
4. EDIT: riapri evento Pagamenti → verifica campi popolati e visibili
5. Dashboard: verifica calcolo saving corretto
6. EDIT: riapri evento Prezzo esistente → verifica NO regressione

**Deliverable**: UI completa e funzionante

---

### FASE 4: Testing e Stabilizzazione (1-2 ore)
**File**: `tests/test_vsm_pagamenti.py` (nuovo)

#### Test Case Calcolo
```python
def test_pagamenti_delta_positivo():
    """Dilazione aumentata → saving positivo."""
    event = VSMEvent(
        driver="Pagamenti",
        spending_annuo=120000.0,
        giorni_pagamento_attuali=30,
        giorni_pagamento_negoziati=60
    )
    assert event.calculate_theoretical_value() == 600.0

def test_pagamenti_delta_negativo():
    """Dilazione ridotta → impatto negativo."""
    event = VSMEvent(
        driver="Pagamenti",
        spending_annuo=120000.0,
        giorni_pagamento_attuali=60,
        giorni_pagamento_negoziati=30
    )
    assert event.calculate_theoretical_value() == -600.0

def test_pagamenti_delta_zero():
    """Nessuna variazione → zero saving."""
    event = VSMEvent(
        driver="Pagamenti",
        spending_annuo=120000.0,
        giorni_pagamento_attuali=45,
        giorni_pagamento_negoziati=45
    )
    assert event.calculate_theoretical_value() == 0.0

def test_pagamenti_no_realizzo():
    """percent_realizzo ignorato per Pagamenti."""
    event = VSMEvent(
        driver="Pagamenti",
        spending_annuo=120000.0,
        giorni_pagamento_attuali=30,
        giorni_pagamento_negoziati=60,
        percent_realizzo=50  # Dovrebbe essere ignorato
    )
    theoretical = event.calculate_theoretical_value()
    effective = event.calculate_effective_value()
    assert theoretical == effective  # NO applicazione realizzo

def test_prezzo_regression():
    """Verifica NO regressione su driver Prezzo."""
    event = VSMEvent(
        driver="Prezzo",
        importo_bdg=1000.0,
        importo_negoziato=800.0,
        percent_realizzo=80
    )
    theoretical = event.calculate_theoretical_value()
    effective = event.calculate_effective_value()
    assert theoretical == 200.0
    assert effective == 160.0  # 200 * 0.8
```

#### Test Integrazione End-to-End
1. CREATE evento Pagamenti via UI
2. Verifica salvataggio DB pulito (campi Prezzo a NULL)
3. Generazione impatti mensili via motore VSM
4. Verifica valori impatti corretti (pro-rata, riverbero)
5. Dashboard: verifica visualizzazione corretta

**Deliverable**: Test suite completa, verifica stabilità

---

## H. Checklist Pre-Release RETTIFICATA

### Backend
- [ ] Tabella `vsm_settings` creata in database
- [ ] Valore default `pagamenti_coefficient = 0.005` popolato
- [ ] Metodi `get_vsm_setting()` / `set_vsm_setting()` implementati
- [ ] Helper `_get_pagamenti_coefficient()` in `vsm_event.py`
- [ ] Branch `elif driver == "Pagamenti"` in `calculate_theoretical_value()`
- [ ] Formula corretta: `spending_annuo * (delta_giorni/30) * coeff` (NO /100)
- [ ] `calculate_effective_value()` esclude percent_realizzo per Pagamenti
- [ ] Gestione campi NULL corretta

### UI
- [ ] Widget spending_annuo, giorni_attuali, giorni_negoziati creati
- [ ] Metodo `_on_driver_changed()` implementato
- [ ] Binding su combobox driver funzionante
- [ ] Logica show/hide campi corretta (mutualmente esclusiva)
- [ ] Campo % Realizzo nascosto quando driver = Pagamenti
- [ ] Validazioni condizionali implementate
- [ ] Warning opzionale delta_giorni < 0 implementato
- [ ] Metodo `_prepare_event_data()` con NULL enforcement
- [ ] Gestione modalità EDIT corretta

### Testing
- [ ] Test calcolo delta positivo/zero/negativo
- [ ] Test percent_realizzo ignorato per Pagamenti
- [ ] Test regressione driver Prezzo (NO breaking changes)
- [ ] Test UI CREATE evento Pagamenti
- [ ] Test UI EDIT evento Pagamenti
- [ ] Test integrazione end-to-end: CREATE → impatti → dashboard
- [ ] Validazione con almeno 3 scenari reali

### Documentazione
- [ ] Commento in codice su formula Pagamenti
- [ ] Documentazione significato coefficiente (0.005 = 0.5% mensile)
- [ ] Commento su esclusione percent_realizzo per Pagamenti
- [ ] README aggiornato se necessario

---

## I. Risposte a Domande Aperte

### 1. Coefficiente: Formato e Gestione
**Risposta**: 
- Formato: **0.005** (percentuale decimale)
- Significato: 0.5% costo opportunità ogni 30 giorni
- Storage: Tabella `vsm_settings`, chiave `pagamenti_coefficient`, valore TEXT convertito a float
- Default: 0.005 popolato automaticamente al primo avvio
- Modificabile: Via future UI settings o manualmente in DB

### 2. Delta Negativo: Gestione UI
**Risposta**:
- Comportamento: Calcolo procede normalmente, risultato negativo
- UI: Warning opzionale `messagebox.askyesno()` per conferma consapevole
- NON bloccare salvataggio (peggioramento è scenario legittimo)
- Dashboard: Impatti negativi visualizzati correttamente (già gestito)

### 3. Driver Volume/Altro: Priorità
**Risposta**:
- **NON implementare** nella release corrente
- Focus solo su Pagamenti
- Branch `else: return 0.0` in `calculate_theoretical_value()`
- Future enhancement richiede analisi separata

### 4. Opex Ripetitivo per Pagamenti
**Risposta**:
- **Default**: opex_ripetitivo = True per Pagamenti (suggerito)
- Motivo: Saving da dilazione pagamento è ricorrente multi-mese
- Configurabile: Utente può modificare checkbox (non forzato)
- Motore VSM: Gestisce già opex_ripetitivo (24 mesi riverbero)

### 5. Spending Annuo: Definizione
**Risposta**:
- Importo totale spesa annuale con **quel fornitore specifico**
- NON categoria merceologica generica
- Utente inserisce valore contrattuale annuale stimato
- Label UI: "Spending Annuo (€): *" con tooltip opzionale "Spesa annuale stimata con il fornitore"

---

## J. Riepilogo Correzioni Funzionali

| # | Punto Criticato | Correzione Applicata |
|---|----------------|----------------------|
| 1 | Coefficiente hardcoded in constants.py | Tabella `vsm_settings`, valore configurabile runtime |
| 2 | Nomenclatura campi "bdg" per Pagamenti | Confermato: spending_annuo, giorni_pagamento_attuali, giorni_pagamento_negoziati |
| 3 | % Realizzo applicato a Pagamenti | `calculate_effective_value()` modificato: realizzo solo per Prezzo |
| 4 | Alternatività Prezzo/Pagamenti ambigua | UI mutualmente esclusiva, NULL enforcement in salvataggio |
| 5 | Salvataggio ambiguo (record ibridi) | `_prepare_event_data()` popola solo campi driver attivo, altri a NULL |
| 6 | Formula con conversioni implicite | Coefficiente formato 0.005 (decimale), formula diretta NO /100 |
| 7 | Delta negativo bloccato | Calcolo procede, risultato negativo, warning UI opzionale |
| 8 | Duplicazione logica motore VSM | Confermato: motore riutilizzato as-is, driver-agnostic |

---

**Fine Piano Rettificato**

**Prossima Azione**: Attendere approvazione utente prima di procedere con implementazione FASE 1 (configurazione database).
