# Piano Implementazione Driver "Pagamenti" - Modulo VSM

## Contesto
Il modulo VSM gestisce eventi di risparmio/cost avoidance/de-risking con diversi driver di calcolo. Attualmente il driver "Prezzo" è completamente implementato (calcolo saving basato su differenza tra importo budget e importo negoziato). Il driver "Pagamenti" esiste nel modello dati ma manca la logica di calcolo e la UI dinamica per gestire i campi specifici (spending annuo, giorni di pagamento attuali/negoziati).

## A. Cosa Esiste Già

### 1. Database Schema
- Tabella `vsm_events` contiene tutti i campi necessari:
  - `driver` (TEXT): valori ["Prezzo", "Pagamenti", "Volume", "Altro"]
  - `spending_annuo` (REAL): spesa annuale per calcolo Pagamenti
  - `giorni_pagamento_attuali` (INTEGER): giorni di pagamento baseline
  - `giorni_pagamento_negoziati` (INTEGER): giorni di pagamento negoziati
  - `importo_bdg` / `importo_negoziato` (REAL): per driver Prezzo

### 2. Modello Dati
File: `models/vsm_event.py`
- Dataclass `VSMEvent` ha tutti i campi necessari
- Metodo `calculate_theoretical_value()` esiste ma implementa solo formula Prezzo
- Metodo `calculate_effective_value()` applica percent_realizzo, completamente riutilizzabile

### 3. Persistenza
File: `database_manager.py`
- `insert_vsm_event()`: salva tutti i campi inclusi quelli per Pagamenti
- `update_vsm_event()`: aggiorna tutti i campi
- `get_vsm_event_by_id()`: carica evento completo con tutti i campi
- Layer persistenza completamente pronto, nessuna modifica necessaria

### 4. Motore VSM
File: `services/vsm_engine.py`
- `generate_impacts_for_event()`: genera impatti mensili con riverbero
- `_calculate_distribution_months()`: gestisce opex_ripetitivo (24 mesi) vs one-shot
- `_distribute_value()`: distribuzione matematica con pro-rata primo mese
- Motore è driver-agnostic, utilizza solo il valore teorico/effettivo calcolato
- Completamente riutilizzabile senza modifiche

### 5. UI Parziale
File: `ui/dialogs/vsm_event_dialog.py`
- Combobox `driver` presente con valori corretti
- Pattern `_on_event_type_changed()` mostra/nasconde campi dinamicamente (riutilizzabile)
- Widget esistenti: importo_bdg, importo_negoziato, descrizione, etc.

## B. Cosa Manca

### 1. Formula di Calcolo
`models/vsm_event.py` → `calculate_theoretical_value()`
- Manca branch per `driver == "Pagamenti"`
- Formula proposta:
  ```python
  if self.driver == "Pagamenti":
      delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
      saving_percentuale = (delta_giorni / 30) * coefficiente
      return self.spending_annuo * (saving_percentuale / 100)
  ```

### 2. Coefficiente Configurabile
`constants.py` o nuovo file settings
- Manca costante `VSM_PAGAMENTI_COEFFICIENT` (valore proposto: 2.0)
- Rappresenta costo opportunità del capitale (es. 2% mensile)

### 3. UI Dinamica
`ui/dialogs/vsm_event_dialog.py`
- Mancano widget specifici per Pagamenti:
  - Entry `spending_annuo` (€)
  - Entry `giorni_pagamento_attuali` (giorni)
  - Entry `giorni_pagamento_negoziati` (giorni)
- Manca metodo `_on_driver_changed()` per show/hide campi in base a driver selezionato
- Manca binding su combobox driver: `combo_driver.bind("<<ComboboxSelected>>", self._on_driver_changed)`

### 4. Validazioni Condizionali
`ui/dialogs/vsm_event_dialog.py` → `_validate_and_save()`
- Validare campi appropriati in base a driver:
  - Se `driver == "Prezzo"`: validare importo_bdg e importo_negoziato
  - Se `driver == "Pagamenti"`: validare spending_annuo, giorni_attuali, giorni_negoziati
- Prevenire salvataggio dati incoerenti

## C. Cosa È Riutilizzabile

### 1. Pattern UI Dinamica
Metodo `_on_event_type_changed()` mostra il pattern perfetto:
```python
def _on_event_type_changed(self, event=None):
    event_type = self.combo_event_type.get()
    if event_type in ["Saving", "Cost Avoidance"]:
        # mostra campi economici
    else:
        # nascondi campi economici
```
Stesso pattern applicabile a `_on_driver_changed()`

### 2. Motore VSM Completo
- `vsm_engine.py` è completamente driver-agnostic
- Genera impatti mensili chiamando solo `event.calculate_theoretical_value()`
- Nessuna modifica necessaria al motore

### 3. Grid Layout Manager
Sistema posizionamento widget con `rowconfigure(weight=1)` e `sticky="nsew"` già utilizzato per campo Description, riutilizzabile per nuovi campi

### 4. Sistema Persistenza
Layer database completo, testato, stabile. Zero modifiche necessarie.

## D. Cosa NON Va Riutilizzato

### 1. NON Riusare Campi € per Giorni
- `importo_bdg` / `importo_negoziato` sono per driver Prezzo (€)
- `giorni_pagamento_attuali` / `giorni_pagamento_negoziati` sono per driver Pagamenti (giorni)
- Mantenere separazione netta per evitare confusione semantica

### 2. NON Assumere Driver Unico per Evento
- Un evento ha UN solo driver attivo
- UI deve essere mutualmente esclusiva: o campi Prezzo visibili, o campi Pagamenti visibili
- Non mostrare entrambi contemporaneamente

## E. File da Toccare

### 1. `constants.py` (PRIORITÀ ALTA)
```python
# VSM Settings
VSM_PAGAMENTI_COEFFICIENT = 2.0  # Coefficiente costo opportunità (% mensile)
```

### 2. `models/vsm_event.py` (PRIORITÀ ALTA)
Modificare metodo `calculate_theoretical_value()`:
```python
def calculate_theoretical_value(self) -> float:
    if self.driver == "Prezzo":
        # Formula esistente
        if self.importo_bdg is not None and self.importo_negoziato is not None:
            return self.importo_bdg - self.importo_negoziato
        return 0.0
    
    elif self.driver == "Pagamenti":
        # Nuova formula
        from constants import VSM_PAGAMENTI_COEFFICIENT
        if (self.spending_annuo is not None and 
            self.giorni_pagamento_attuali is not None and 
            self.giorni_pagamento_negoziati is not None):
            
            delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
            saving_percentuale = (delta_giorni / 30) * VSM_PAGAMENTI_COEFFICIENT
            return self.spending_annuo * (saving_percentuale / 100)
        return 0.0
    
    else:
        # Volume, Altro: logica futura
        return 0.0
```

### 3. `ui/dialogs/vsm_event_dialog.py` (PRIORITÀ MEDIA)
**Sezione A: Creazione Widget**
```python
# In _build_ui(), dopo campi Prezzo, aggiungere sezione Pagamenti:
pagamenti_label = ttk.Label(bottom_frame, text="Dati Pagamenti:", style="TLabel")
pagamenti_label.grid(row=next_row, column=0, sticky="w", pady=(10, 5))

# Spending Annuo
ttk.Label(bottom_frame, text="Spending Annuo (€):").grid(row=next_row+1, column=0, sticky="w", padx=(20, 5))
self.entry_spending_annuo = ttk.Entry(bottom_frame, width=20)
self.entry_spending_annuo.grid(row=next_row+1, column=1, sticky="w", padx=5)

# Giorni Pagamento Attuali
ttk.Label(bottom_frame, text="Giorni Pagamento Attuali:").grid(row=next_row+2, column=0, sticky="w", padx=(20, 5))
self.entry_giorni_attuali = ttk.Entry(bottom_frame, width=20)
self.entry_giorni_attuali.grid(row=next_row+2, column=1, sticky="w", padx=5)

# Giorni Pagamento Negoziati
ttk.Label(bottom_frame, text="Giorni Pagamento Negoziati:").grid(row=next_row+3, column=0, sticky="w", padx=(20, 5))
self.entry_giorni_negoziati = ttk.Entry(bottom_frame, width=20)
self.entry_giorni_negoziati.grid(row=next_row+3, column=1, sticky="w", padx=5)

# Binding su driver combobox
self.combo_driver.bind("<<ComboboxSelected>>", self._on_driver_changed)
```

**Sezione B: Metodo _on_driver_changed**
```python
def _on_driver_changed(self, event=None):
    """Mostra/nasconde campi in base al driver selezionato."""
    driver = self.combo_driver.get()
    
    if driver == "Prezzo":
        # Mostra campi importi, nascondi campi pagamenti
        self.importo_bdg_label.grid()
        self.entry_importo_bdg.grid()
        self.importo_negoziato_label.grid()
        self.entry_importo_negoziato.grid()
        
        self.entry_spending_annuo.grid_remove()
        self.entry_giorni_attuali.grid_remove()
        self.entry_giorni_negoziati.grid_remove()
        # Nascondi anche le label corrispondenti
        
    elif driver == "Pagamenti":
        # Nascondi campi importi, mostra campi pagamenti
        self.importo_bdg_label.grid_remove()
        self.entry_importo_bdg.grid_remove()
        self.importo_negoziato_label.grid_remove()
        self.entry_importo_negoziato.grid_remove()
        
        self.entry_spending_annuo.grid()
        self.entry_giorni_attuali.grid()
        self.entry_giorni_negoziati.grid()
        # Mostra anche le label corrispondenti
        
    else:
        # Volume, Altro: nascondi tutto (per ora)
        # Logica futura
        pass
```

**Sezione C: Validazioni in _validate_and_save**
```python
def _validate_and_save(self):
    driver = self.combo_driver.get()
    
    if driver == "Prezzo":
        # Validazioni esistenti per importi
        if not self.entry_importo_bdg.get().strip():
            messagebox.showerror("Errore", "Importo Budget richiesto per driver Prezzo")
            return
        # ... altre validazioni
        
    elif driver == "Pagamenti":
        # Nuove validazioni per pagamenti
        if not self.entry_spending_annuo.get().strip():
            messagebox.showerror("Errore", "Spending Annuo richiesto per driver Pagamenti")
            return
        if not self.entry_giorni_attuali.get().strip():
            messagebox.showerror("Errore", "Giorni Pagamento Attuali richiesto per driver Pagamenti")
            return
        if not self.entry_giorni_negoziati.get().strip():
            messagebox.showerror("Errore", "Giorni Pagamento Negoziati richiesto per driver Pagamenti")
            return
        
        try:
            spending = float(self.entry_spending_annuo.get())
            giorni_att = int(self.entry_giorni_attuali.get())
            giorni_neg = int(self.entry_giorni_negoziati.get())
            
            if spending <= 0:
                messagebox.showerror("Errore", "Spending Annuo deve essere positivo")
                return
            if giorni_att < 0 or giorni_neg < 0:
                messagebox.showerror("Errore", "Giorni di pagamento non possono essere negativi")
                return
        except ValueError:
            messagebox.showerror("Errore", "Valori numerici non validi per campi Pagamenti")
            return
    
    # Procedi con salvataggio...
```

### 4. File NON da Toccare
- `database_manager.py`: layer persistenza completo
- `services/vsm_engine.py`: motore completo e driver-agnostic
- `services/vsm_persistence.py`: nessuna modifica necessaria

## F. Rischi e Mitigazioni

### Rischio 1: Confusione UI tra € e Giorni
**Impatto**: Alto - utente potrebbe inserire € nei campi giorni o viceversa
**Mitigazione**: 
- Label chiarissime con unità di misura ("€", "giorni")
- Validazioni strict (int per giorni, float per €)
- UI mutualmente esclusiva: solo campi pertinenti visibili

### Rischio 2: Regressione su Driver Prezzo
**Impatto**: Critico - rompere logica esistente che funziona
**Mitigazione**:
- Branch `if/elif` esplicito in `calculate_theoretical_value()`
- Test regressione: verificare che eventi Prezzo esistenti mantengano calcolo corretto
- Commit atomici: backend separato da UI

### Rischio 3: Coefficiente Hardcoded
**Impatto**: Medio - difficile da modificare senza ricompilazione
**Mitigazione**:
- Usare costante in `constants.py` (modificabile)
- Documentare significato coefficiente (costo opportunità capitale)
- Future: considerare configurazione runtime in tabella settings

### Rischio 4: Delta Giorni Negativo
**Impatto**: Basso - saving negativo semanticamente valido (peggioramento)
**Mitigazione**:
- Permettere delta negativo (giorni_negoziati < giorni_attuali)
- Risultato sarà importo negativo (de-saving), gestito correttamente dal motore VSM
- Validare solo non-negatività dei singoli valori giorni

### Rischio 5: Campi NULL in Calcolo
**Impatto**: Medio - potenziale NoneType error
**Mitigazione**:
- Check espliciti `if field is not None` prima di calcolo
- Return 0.0 se campi mancanti (safe default)
- Validazioni UI prevengono salvataggio con campi vuoti

## G. Proposta Tecnica MVP (Minimum Viable Product)

### FASE 1: Backend (Calcolo) - Implementabile SUBITO
**Scope**: Implementare logica di calcolo senza toccare UI
**File**: `constants.py`, `models/vsm_event.py`
**Deliverable**: 
- Costante `VSM_PAGAMENTI_COEFFICIENT = 2.0` in constants.py
- Branch `elif self.driver == "Pagamenti"` in `calculate_theoretical_value()`
- Formula: `spending_annuo * ((delta_giorni/30) * coeff / 100)`

**Testing FASE 1**:
```python
# Test in Python REPL o test_vsm_engine.py
from models.vsm_event import VSMEvent
from datetime import date

event = VSMEvent(
    id=None,
    event_type="Saving",
    driver="Pagamenti",
    spending_annuo=120000.0,  # 120k€ annui
    giorni_pagamento_attuali=30,
    giorni_pagamento_negoziati=60,  # +30 giorni
    data_inizio=date.today(),
    opex_ripetitivo=True,
    percent_realizzo=100
)

theoretical = event.calculate_theoretical_value()
# Atteso: 120000 * ((30/30) * 2 / 100) = 120000 * 0.02 = 2400€
print(f"Saving teorico: {theoretical}€")
```

**Vantaggi FASE 1**:
- Zero impatto su UI esistente
- Testabile immediatamente con codice Python
- Commit atomico, facile revert se problemi

### FASE 2: UI Dinamica - Richiede FASE 1 Completata
**Scope**: Implementare show/hide campi in base a driver
**File**: `ui/dialogs/vsm_event_dialog.py`
**Deliverable**:
- Widget entry per spending_annuo, giorni_attuali, giorni_negoziati
- Metodo `_on_driver_changed()` con logica show/hide
- Binding su combobox driver
- Validazioni condizionali in `_validate_and_save()`

**Testing FASE 2**:
1. Aprire VSMEventDialog in modalità CREATE
2. Selezionare driver "Prezzo" → verificare campi importi visibili
3. Selezionare driver "Pagamenti" → verificare campi pagamenti visibili, importi nascosti
4. Inserire dati validi per Pagamenti e salvare
5. Riaprire evento in modalità EDIT → verificare campi popolati correttamente
6. Verificare calcolo saving nel dashboard VSM

**Vantaggi FASE 2**:
- UI intuitiva e mutualmente esclusiva
- Pattern riutilizzabile per driver futuri (Volume, Altro)
- Validazioni prevengono data corruption

### FASE 3: Testing e Stabilizzazione - Raccomandato ma Non Bloccante
**Scope**: Test unitari e integrazione
**File**: `tests/test_vsm_engine.py`, nuovo `tests/test_vsm_pagamenti.py`
**Deliverable**:
- Test `calculate_theoretical_value()` con delta positivo/negativo/zero
- Test regressione driver Prezzo (events esistenti)
- Test integrazione end-to-end: CREATE evento Pagamenti → genera impatti → verifica valori mensili
- Test edge cases: spending=0, giorni=0, campi NULL

**Test Cases**:
```python
def test_pagamenti_delta_positivo():
    # Dilazione maggiore → saving positivo
    event = VSMEvent(..., giorni_attuali=30, giorni_negoziati=60)
    assert event.calculate_theoretical_value() > 0

def test_pagamenti_delta_negativo():
    # Dilazione minore → de-saving (negativo)
    event = VSMEvent(..., giorni_attuali=60, giorni_negoziati=30)
    assert event.calculate_theoretical_value() < 0

def test_pagamenti_delta_zero():
    # Nessuna variazione → zero saving
    event = VSMEvent(..., giorni_attuali=30, giorni_negoziati=30)
    assert event.calculate_theoretical_value() == 0.0

def test_prezzo_regression():
    # Verifica che eventi Prezzo esistenti mantengano calcolo corretto
    event = VSMEvent(..., driver="Prezzo", importo_bdg=1000, importo_negoziato=800)
    assert event.calculate_theoretical_value() == 200.0
```

## H. Sequenza Implementazione Consigliata

### Step 1: Backend Calcolo (30 min)
1. Aggiungere `VSM_PAGAMENTI_COEFFICIENT = 2.0` in `constants.py`
2. Modificare `calculate_theoretical_value()` in `models/vsm_event.py`
3. Test manuale in Python REPL

### Step 2: UI Dinamica (2-3 ore)
1. Creare widget entry per campi Pagamenti in `_build_ui()`
2. Implementare `_on_driver_changed()` con logica show/hide
3. Aggiungere binding su `combo_driver`
4. Implementare validazioni condizionali in `_validate_and_save()`
5. Test manuale UI: create/edit eventi Pagamenti

### Step 3: Testing (1-2 ore)
1. Scrivere test unitari per `calculate_theoretical_value()`
2. Test regressione driver Prezzo
3. Test integrazione end-to-end
4. Validare con dati reali

### Step 4: Documentazione (30 min)
1. Commentare formula in `calculate_theoretical_value()`
2. Documentare significato coefficiente in `constants.py`
3. Aggiornare README se necessario

## I. Note Implementative Aggiuntive

### Gestione EDIT Mode
Quando si carica un evento esistente in modalità EDIT:
- Determinare driver dall'evento caricato
- Chiamare `_on_driver_changed()` per mostrare campi appropriati
- Popolare i widget con valori esistenti

```python
# In _build_ui() dopo popolamento campi
if self.mode == "EDIT":
    # Popola campi...
    self._on_driver_changed()  # Mostra/nascondi in base a driver caricato
```

### Label con Unità di Misura
Utilizzare label esplicite per evitare confusioni:
- "Importo Budget (€):" - NON solo "Importo Budget:"
- "Giorni Pagamento Attuali:" - NON "Pagamento Attuali"
- "Spending Annuo (€/anno):" - chiarire se importo annuale

### Placeholder e Tooltip
Considerare aggiunta tooltip per spiegare campi complessi:
```python
# Esempio con tooltip
from tkinter import ttk
spending_entry = ttk.Entry(...)
# Tooltip: "Importo totale spesa annuale con questo fornitore"
```

### Separazione Visiva
Suggerito inserire separator tra sezione Prezzo e sezione Pagamenti:
```python
separator = ttk.Separator(bottom_frame, orient="horizontal")
separator.grid(row=..., column=0, columnspan=2, sticky="ew", pady=10)
```

## J. Checklist Pre-Release

### Backend
- [ ] Costante `VSM_PAGAMENTI_COEFFICIENT` definita in `constants.py`
- [ ] Branch `elif driver == "Pagamenti"` implementato in `calculate_theoretical_value()`
- [ ] Formula matematica verificata con esempi numerici
- [ ] Gestione campi NULL implementata (return 0.0 se mancanti)
- [ ] Test manuale con evento mock eseguito

### UI
- [ ] Widget entry per `spending_annuo`, `giorni_attuali`, `giorni_negoziati` creati
- [ ] Metodo `_on_driver_changed()` implementato
- [ ] Binding su `combo_driver` aggiunto
- [ ] Logica show/hide campi funzionante (Prezzo vs Pagamenti)
- [ ] Validazioni condizionali implementate in `_validate_and_save()`
- [ ] Label con unità di misura chiare (€, giorni)
- [ ] Gestione modalità EDIT corretta (caricamento valori + show/hide appropriati)

### Testing
- [ ] Test `calculate_theoretical_value()` con delta positivo/zero/negativo
- [ ] Test regressione driver "Prezzo" (eventi esistenti mantengono calcolo)
- [ ] Test UI manuale: CREATE evento Pagamenti
- [ ] Test UI manuale: EDIT evento Pagamenti
- [ ] Test end-to-end: evento Pagamenti → impatti mensili → dashboard
- [ ] Validazione con almeno 3 scenari reali

### Documentazione
- [ ] Commento in codice su formula Pagamenti
- [ ] Documentazione significato coefficiente
- [ ] README aggiornato se necessario

## K. Domande Aperte per l'Utente

1. **Coefficiente 2.0%**: Il valore proposto (2% mensile, ~24% annuo) è corretto per il vostro business? Dovrebbe essere configurabile?

2. **Delta Negativo**: Un delta giorni negativo (peggioramento dilazione) dovrebbe generare warning esplicito in UI o è comportamento atteso?

3. **Driver Volume/Altro**: Priorità implementazione futura? Logica di calcolo prevista?

4. **Opex Ripetitivo**: Per driver Pagamenti, il saving dovrebbe sempre essere considerato ripetitivo (24 mesi) o configurabile?

5. **Spending Annuo**: Deve essere importo totale con fornitore o specifico per la categoria merceologica negoziata?

---

**Fine Piano Implementazione**
