# IMPLEMENTAZIONE RICHIESTA — VSM Saving / Cost Avoidance con quantità annua per driver "Price"

## ✅ STATO ANALISI: COMPLETATA E VERIFICATA

**SCOPERTA FONDAMENTALE**: Il campo `quantita_annua` esiste già nel database e nel modello, ma:
- ❌ NON è esposto nella UI dialog (nessun widget creato)
- ❌ NON è utilizzato nella formula di calcolo per driver "Prezzo" (sempre calcolo con qty implicita = 1)
- ❌ NON è popolato in edit mode (manca il binding in `_load_event_data()`)
- ❌ NON è assegnato nel save flow (manca nel costruttore VSMEvent linea 643)
- ✅ È già persistito e mappato correttamente nelle query DB (INSERT/UPDATE/SELECT)
- ✅ È già presente nei test di persistence (test_vsm_persistence.py usa quantita_annua=100.0)

**STATO ATTUALE**: Eventi salvati con quantita_annua=0.0 (default dataclass non sovrascritto dalla UI)

**IMPATTO IMPLEMENTAZIONE**: MOLTO BASSO
- Solo 2 file da modificare: UI dialog + modello calcolo
- No migrazione DB richiesta
- No modifica persistence layer
- Pattern UI già collaudato (riuso logica payment_fields)

---

## CONTESTO
Nel nuovo modulo VSM, per gli eventi di tipo Saving e Cost Avoidance, la logica attuale NON tiene conto della quantità.
Questo è corretto solo nei casi one-shot con quantità = 1, ma è sbagliato quando il driver è il prezzo unitario su volumi annui.

**SCOPERTA**: Il campo `quantita_annua` esiste già nel database ma non è esposto nella UI e non è utilizzato nella formula di calcolo per il driver "Price".

## OBIETTIVO
Quando il Driver è "Prezzo", nella sezione "Economic Data" della dialog di creazione/modifica evento deve comparire un nuovo campo:
- **Annual Q.ty** (Quantità Annua)

La formula del valore teorico e effettivo deve diventare:

```
Valore teorico = Annual Q.ty × (Budget Amount - Negotiated Amount)
Valore effettivo = Valore teorico × (% Realization / 100)
```

**NOTA**: La % Realization è applicata SOLO al valore effettivo, NON al theoretical (come già avviene nel codice attuale). Il theoretical rappresenta il valore grezzo del saving, l'effective rappresenta il valore realizzato applicando la percentuale di realizzo.

## ESEMPI (CORRETTI CON SEPARAZIONE THEORETICAL/EFFECTIVE)

### 1) Macchinario (one-shot)
- Budget Amount = 20000
- Negotiated Amount = 18000
- Annual Q.ty = 1
- % Realization = 100

**Calcolo**:
- Theoretical Value = 1 × (20000 - 18000) = **2000**
- Effective Value = 2000 × (100 / 100) = **2000**

### 2) Produzione (volume annuo)
- Budget Amount = 1.5 (prezzo unitario)
- Negotiated Amount = 1.3 (prezzo unitario)
- Annual Q.ty = 20000 (pezzi/anno)
- % Realization = 100

**Calcolo**:
- Theoretical Value = 20000 × (1.5 - 1.3) = 20000 × 0.2 = **4000**
- Effective Value = 4000 × (100 / 100) = **4000**

### 3) Cost Avoidance con realizzo parziale
- Initial Requested Amount = 2.0 (prezzo unitario)
- Negotiated Amount = 1.8 (prezzo unitario)
- Annual Q.ty = 15000 (pezzi/anno)
- % Realization = 80

**Calcolo**:
- Theoretical Value = 15000 × (2.0 - 1.8) = 15000 × 0.2 = **3000**
- Effective Value = 3000 × (80 / 100) = **2400**

**NOTA IMPORTANTE**: La % Realization viene applicata SOLO nel calcolo dell'Effective Value, NON nel Theoretical Value. Questo è il comportamento corretto verificato nel codice esistente.

## VINCOLI FONDAMENTALI
- Modifica conservativa e circoscritta
- Nessuna regressione su Linux / Windows
- Non ingrassare DataFlow.py
- Rispettare il refactoring esistente e la separazione logica già introdotta nel modulo VSM
- Non modificare logiche non richieste
- Non cambiare naming o struttura UI oltre il minimo necessario
- Non rompere compatibilità con eventi già salvati

## AMBITO FUNZIONALE
La modifica riguarda solo:
- **Saving**
- **Cost Avoidance**

Il nuovo campo Annual Q.ty deve essere rilevante SOLO quando il Driver è "**Price**" (Prezzo).

## COMPORTAMENTO UI RICHIESTO

### Posizionamento del Campo
Nel form "New VSM Event" / edit event, sezione "Economic Data":
- Aggiungere il campo **"Annual Q.ty"** / **"Quantità Annua"**
- Posizionarlo tra:
  - **Negotiated Amount** (Importo Negoziato)
  - **% Realization** (% Realizzo)

### Comportamento Dinamico
- Se **Driver = "Price"** → il campo Annual Q.ty deve essere **visibile/attivo**
- Se **Driver != "Price"** → il campo Annual Q.ty deve essere **nascosto** (usando grid_forget, coerente con lo stile attuale del form)
- Evitare comportamenti grafici invasivi o instabili
- Mantenere allineamento pulito della UI

### Default Value
- Valore predefinito: **"1"** (retrocompatibilità con comportamento storico)

## VALIDAZIONE

### Quando Driver = "Price"
- Annual Q.ty deve essere **obbligatorio**
- Deve accettare solo valori **numerici positivi** (> 0)
- Quantità zero o negativa → **non valida**
- Gestire valori **decimali** (float)
- Messaggio di errore: `"Quantità Annua deve essere un numero valido maggiore di zero."`

### Quando Driver != "Price"
- La quantità non deve influenzare i calcoli
- Per retrocompatibilità logica considerare quantità implicita = 1 nei calcoli

## FORMULA DI CALCOLO (VERIFICATA NEL CODICE)

### Per Saving / Cost Avoidance con Driver = "Prezzo"

```python
# Saving
theoretical_value = annual_qty * (importo_bdg - importo_negoziato)
effective_value = theoretical_value * (percent_realizzo / 100.0)

# Cost Avoidance
theoretical_value = annual_qty * (importo_richiesto_iniziale - importo_negoziato)
effective_value = theoretical_value * (percent_realizzo / 100.0)
```

**NOTA CRITICA**: Il realizzo è applicato SOLO nell'effective value, NON nel theoretical. Il metodo `calculate_theoretical_value()` ritorna il valore grezzo senza realizzo. Il metodo `calculate_effective_value()` applica il realizzo al theoretical.

### Per gli altri driver
Mantenere la logica attuale invariata:
- **Pagamenti**: `theoretical = spending_annuo * (delta_giorni / 30) * coefficiente`, `effective = theoretical` (no realizzo)
- **Derisking**: `theoretical = 0.0`, `effective = 0.0`

## RETROCOMPATIBILITÀ DATABASE

### Campo Database
- **Campo**: `quantita_annua` (REAL)
- **Stato**: **ESISTE GIÀ** nella tabella `vsm_events`
- **Migrazione**: **NON NECESSARIA**

### Gestione Record Esistenti
Per record storici già presenti:
- `quantita_annua` può essere `NULL` o non valorizzato
- Nel calcolo, applicare **default logico = 1.0** se NULL o 0
- Questo garantisce che i vecchi eventi continuino a calcolare correttamente senza modifiche manuali

### Pattern di Difesa
```python
# In calculate_theoretical_value()
qty = self.quantita_annua if self.quantita_annua and self.quantita_annua > 0 else 1.0
```

## PERSISTENZA

### Salvataggio
- Quando si salva un evento, persistere `quantita_annua` nel database
- Campo già mappato nello schema, nessuna modifica necessaria

### Caricamento
- Quando si carica un evento per modifica:
  - Popolare `entry_quantita_annua` con il valore esistente
  - Se `quantita_annua` è NULL → mostrare "1" come default nella UI

### Rigenerazione Impatti
- Il pattern DELETE-REGENERATE-SAVE esistente gestirà automaticamente il ricalcolo degli impatti mensili con la nuova formula

## CALCOLO / BUSINESS LOGIC

### Punto di Modifica
**File**: [models/vsm_event.py](models/vsm_event.py)  
**Metodo**: `calculate_theoretical_value()` (circa linea 161)

### Logica da Applicare
Aggiornare SOLO i casi:
- `event_type` = "Saving" oppure "Cost Avoidance"
- `driver` = "Prezzo"

Non duplicare logica. La funzione è già centralizzata nel modello VSMEvent.

## ✅ VERIFICHE PRE-IMPLEMENTAZIONE COMPLETATE

### Punto 1: Driver interno VERIFICATO ✅
- **Valore interno**: `"Prezzo"` (italiano) - NON "Price"
- **Confronti nel codice**: `driver == "Prezzo"` e `driver_internal == "Prezzo"`
- **Conversione UI**: Metodi `_get_driver_internal()` e `_set_driver_display()` gestiscono mapping IT ↔ traduzione
- **Valori ammessi**: `"Prezzo"`, `"Pagamenti"` (entrambi in italiano nel DB e nel modello)

### Punto 2: % Realization VERIFICATO ✅
- **calculate_theoretical_value()**: Calcola valore GREZZO senza realizzo
  - Saving: `importo_bdg - importo_negoziato`
  - Cost Avoidance: `importo_richiesto_iniziale - importo_negoziato`
  - Pagamenti: `spending_annuo * (delta_giorni / 30) * coefficiente`
- **calculate_effective_value()**: Applica realizzo SOLO per driver "Prezzo"
  - `theoretical * (percent_realizzo / 100.0)` se driver == "Prezzo"
  - `theoretical` direttamente se driver == "Pagamenti" (no realizzo)
- ✅ **CONFERMATO**: La separazione theoretical/effective è corretta. La nostra modifica al theoretical NON rompe l'applicazione del realizzo.

### Punto 3: grid_remove() e preservazione valori VERIFICATO ✅
- **Metodo usato**: `grid_remove()` (NON `grid_forget()`)
- **Comportamento**: I widget vengono nascosti ma rimangono in memoria con i loro valori
- **Implicazione**: Quando l'utente switcha Prezzo→Pagamenti→Prezzo, il valore inserito in Annual Q.ty viene **automaticamente preservato**
- ✅ **NESSUNA LOGICA AGGIUNTIVA NECESSARIA** per preservare il valore durante switch driver

### Punto 4: Edit mode e popolamento VERIFICATO ✅
- **Flow di caricamento**: `_load_event_data()` chiamato in `__init__` se `is_edit_mode` (linea ~89)
- **Timing**: Populate campi → poi `_on_event_type_changed()` → poi `_on_driver_changed()` (linee 479-481)
- **PROBLEMA IDENTIFICATO**: `quantita_annua` attualmente NON viene popolato nel load flow
- **SOLUZIONE**: Aggiungere populate di `entry_quantita_annua` in `_load_event_data()` dopo linea ~471 (dopo importo_negoziato, prima di percent_realizzo)
- ✅ **TIMING SICURO**: Il populate avviene PRIMA dei cambi driver, quindi NON c'è rischio di sovrascrittura

### Punto 5: Test VERIFICATO ✅
- **Test esistenti**:
  - `test_vsm_engine.py`: Test su distribuzione impatti mensili (usa `calculate_theoretical_value()` indirettamente)
  - `test_vsm_persistence.py`: Test su save/load eventi (eventi di test hanno `quantita_annua=100.0` già impostato)
- **GAP IDENTIFICATO**: NON ci sono test specifici per i metodi `calculate_theoretical_value()` e `calculate_effective_value()` del modello VSMEvent
- **AZIONE RICHIESTA**: Creare nuovo file `tests/test_vsm_event_model.py` con test unitari specifici per:
  - Saving + Prezzo + qty=1, qty>1, qty decimale, qty=None
  - Cost Avoidance + Prezzo + qty=1, qty>1, qty decimale, qty=None
  - Pagamenti (qty non deve influenzare)
  - Derisking (qty non deve influenzare)

### Punto 6: Cost Avoidance campi VERIFICATO ✅
- **Campo usato**: `importo_richiesto_iniziale` (Optional[float])
- **Formula attuale**: `importo_richiesto_iniziale - importo_negoziato`
- **Formula nuova**: `qty * (importo_richiesto_iniziale - importo_negoziato)`
- ✅ **CONFERMATO**: Il naming e la logica sono corretti nel piano

---

## ANALISI PRE-IMPLEMENTAZIONE

### 1. File da Toccare

| File | Scopo Modifica | Rischio |
|------|----------------|---------|
| [ui/dialogs/vsm_event_dialog.py](ui/dialogs/vsm_event_dialog.py) | Aggiungere widget Annual Q.ty, visibilità condizionale, validazione, data binding | BASSO - modifica locale UI |
| [models/vsm_event.py](models/vsm_event.py) | Modificare `calculate_theoretical_value()` per moltiplicare per quantita_annua | BASSO - calcolo ben isolato |
| **[tests/test_vsm_event_model.py](tests/test_vsm_event_model.py)** | **NUOVO FILE** - test unitari per formule calcolo con quantità | NULLO - solo test |
| [database_manager.py](database_manager.py) | **NESSUNA MODIFICA** - campo già esistente | NULLO |
| [services/vsm_persistence.py](services/vsm_persistence.py) | **NESSUNA MODIFICA** - mapping già funzionante | NULLO |

### 2. Punti Esatti da Modificare

#### A) UI Dialog (`ui/dialogs/vsm_event_dialog.py`)

**A.1) Creazione Widget** (circa linea 238, dentro `price_fields_frame`, dopo `entry_percent_realizzo`):
```python
# Quantità Annua (Annual Q.ty) - visibile solo per driver Prezzo
self.lbl_quantita_annua = ttk.Label(self.price_fields_frame, text=_("Quantità Annua:"))
self.entry_quantita_annua = ttk.Entry(self.price_fields_frame, width=20)
self.entry_quantita_annua.insert(0, "1")  # Default retrocompatibile
```

**A.2) Layout nel metodo `_on_driver_changed()`** (circa linea 368-420):

Per **Saving** (dopo linea ~398):
```python
# Dopo importo_negoziato (row=1), prima di percent_realizzo (row=2)
# Inserire quantita_annua a row=2, spostare percent_realizzo a row=3

self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
self.entry_quantita_annua.grid(row=2, column=1, sticky="w", pady=5)

self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
self.entry_percent_realizzo.grid(row=3, column=1, sticky="w", pady=5)
```

Per **Cost Avoidance** (dopo linea ~410):
```python
# Dopo importo_negoziato (row=1), prima di percent_realizzo (row=2)
# Inserire quantita_annua a row=2, spostare percent_realizzo a row=3

self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
self.entry_quantita_annua.grid(row=2, column=1, sticky="w", pady=5)

self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
self.entry_percent_realizzo.grid(row=3, column=1, sticky="w", pady=5)
```

Quando **driver == "Pagamenti"** (già gestito):
```python
# Nascondere quantita_annua (grid_remove è già applicato al price_fields_frame)
# I widget dentro price_fields_frame vengono automaticamente nascosti
# NESSUNA AZIONE AGGIUNTIVA NECESSARIA (grid_remove preserva valori)
```

**A.3) Validazione in `_validate_and_save()`** (circa linea 538, dentro blocco `if driver == "Prezzo"`):

Per **Saving** (dopo validazione percent_realizzo, prima di linea ~558):
```python
# Dopo validazione percent_realizzo (linea ~555)
try:
    quantita_annua = float(self.entry_quantita_annua.get().strip())
    if quantita_annua <= 0:
        raise ValueError(_("Quantità Annua deve essere maggiore di zero."))
except ValueError as e:
    if "could not convert" in str(e):
        raise ValueError(_("Quantità Annua deve essere un numero valido."))
    raise
```

Per **Cost Avoidance** (dopo validazione percent_realizzo, circa linea ~648):
```python
# Dopo validazione percent_realizzo (linea ~645)
try:
    quantita_annua = float(self.entry_quantita_annua.get().strip())
    if quantita_annua <= 0:
        raise ValueError(_("Quantità Annua deve essere maggiore di zero."))
except ValueError as e:
    if "could not convert" in str(e):
        raise ValueError(_("Quantità Annua deve essere un numero valido."))
    raise
```

**NOTA IMPORTANTE**: Per evitare duplicazione, la validazione può essere estratta prima dei blocchi Saving/Cost Avoidance se `driver == "Prezzo"`.

**A.4) Data Binding - Salvataggio** (circa linea 656, dentro costruzione VSMEvent):

Attualmente `quantita_annua` NON è passato al costruttore. Modificare linea ~656:
```python
event = VSMEvent(
    # ... campi esistenti ...
    quantita_annua=quantita_annua if quantita_annua is not None else 1.0,  # AGGIUNGERE
    percent_realizzo=percent_realizzo,
    # ... resto campi ...
)
```

**ATTENZIONE**: La variabile `quantita_annua` deve essere estratta e validata PRIMA per entrambi i driver. Se driver == "Pagamenti", impostare `quantita_annua = None` o `1.0` (per coerenza DB).

**A.5) Data Binding - Caricamento** (circa linea 471, dentro `_load_event_data()`, dopo populate percent_realizzo):

```python
# Dopo linea ~471 (populate percent_realizzo)
# Popolare quantita_annua
if event.quantita_annua and event.quantita_annua > 0:
    self.entry_quantita_annua.delete(0, tk.END)
    self.entry_quantita_annua.insert(0, str(event.quantita_annua))
else:
    # Default retrocompatibile per eventi storici con NULL o 0
    self.entry_quantita_annua.delete(0, tk.END)
    self.entry_quantita_annua.insert(0, "1")
```

**NOTA CRITICA**: Il populate avviene PRIMA di `_on_event_type_changed()` e `_on_driver_changed()` (linee 479-481), quindi il valore popolato NON viene sovrascritto dai cambi driver iniziali. ✅ SICURO.

#### B) Calculation Logic (`models/vsm_event.py`)

**Metodo `calculate_theoretical_value()`** (linea 151-192):

**FORMULA ATTUALE**:
```python
def calculate_theoretical_value(self) -> float:
    # Driver Pagamenti: calcolo basato su dilazione pagamento
    if self.driver == "Pagamenti":
        if (self.spending_annuo is not None and 
            self.giorni_pagamento_attuali is not None and 
            self.giorni_pagamento_negoziati is not None):
            delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
            coefficiente = get_pagamenti_coefficient()
            return self.spending_annuo * (delta_giorni / 30.0) * coefficiente
        return 0.0
    
    # Driver Prezzo: logica originale basata su importi
    if self.event_type == "Cost Avoidance" and self.importo_richiesto_iniziale:
        return self.importo_richiesto_iniziale - self.importo_negoziato
    elif self.event_type == "Saving":
        return self.importo_bdg - self.importo_negoziato
    else:
        return 0.0
```

**FORMULA MODIFICATA** (applicare quantità SOLO per driver "Prezzo"):
```python
def calculate_theoretical_value(self) -> float:
    """
    Calcola il valore teorico dell'evento in base al driver.
    
    Driver Prezzo:
        - Cost Avoidance: qty × (importo_richiesto_iniziale - importo_negoziato)
        - Saving: qty × (importo_bdg - importo_negoziato)
    
    Driver Pagamenti:
        - Saving da dilazione: spending_annuo * (delta_giorni / 30) * coefficiente
        - Coefficiente: costo opportunità capitale (es. 0.005 = 0.5% mensile)
    
    Returns:
        float: Valore teorico calcolato
    """
    # Driver Pagamenti: calcolo basato su dilazione pagamento (INVARIATO)
    if self.driver == "Pagamenti":
        if (self.spending_annuo is not None and 
            self.giorni_pagamento_attuali is not None and 
            self.giorni_pagamento_negoziati is not None):
            delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
            coefficiente = get_pagamenti_coefficient()
            return self.spending_annuo * (delta_giorni / 30.0) * coefficiente
        return 0.0
    
    # Driver Prezzo: NUOVA logica con quantità annua
    # Backward compatibility: default qty to 1.0 if not set or zero
    qty = self.quantita_annua if self.quantita_annua and self.quantita_annua > 0 else 1.0
    
    if self.event_type == "Cost Avoidance" and self.importo_richiesto_iniziale:
        # Cost Avoidance: differenza tra richiesto iniziale e negoziato, moltiplicato per quantità
        return qty * (self.importo_richiesto_iniziale - self.importo_negoziato)
    elif self.event_type == "Saving":
        # Saving: differenza tra budget e negoziato, moltiplicato per quantità
        return qty * (self.importo_bdg - self.importo_negoziato)
    else:
        # Derisking: nessun valore economico diretto (INVARIATO)
        return 0.0
```

**NOTA CRITICA**:
- `calculate_effective_value()` NON va modificato (applica già correttamente realizzo al theoretical)
- La separazione theoretical/effective rimane intatta
- La quantità moltiplica il valore teorico PRIMA dell'applicazione del realizzo

#### C) Test Unitari (`tests/test_vsm_event_model.py`) - NUOVO FILE

Creare nuovo file di test specifico per le formule di calcolo del modello VSMEvent:

```python
"""
Unit tests for VSM Event Model calculations.

Test cases for calculate_theoretical_value() and calculate_effective_value()
with different annual quantities and drivers.
"""

import unittest
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from models.vsm_event import VSMEvent


class TestVSMEventCalculations(unittest.TestCase):
    """Test per i metodi di calcolo del modello VSMEvent."""
    
    def test_saving_price_qty_1(self):
        """Saving + Prezzo + quantità 1 (caso base)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=1.0,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_saving_price_qty_large(self):
        """Saving + Prezzo + quantità > 1 (produzione volumi)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=1.5,  # prezzo unitario
            importo_negoziato=1.3,  # prezzo unitario
            quantita_annua=20000.0,  # pezzi/anno
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertAlmostEqual(theoretical, 4000.0, places=2)  # 20000 * 0.2
        self.assertAlmostEqual(effective, 4000.0, places=2)
    
    def test_saving_price_qty_decimal(self):
        """Saving + Prezzo + quantità decimale."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=100.0,
            importo_negoziato=80.0,
            quantita_annua=150.5,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        expected = 150.5 * 20.0  # 3010.0
        self.assertAlmostEqual(theoretical, expected, places=2)
        self.assertAlmostEqual(effective, expected, places=2)
    
    def test_saving_price_qty_none_defaults_to_1(self):
        """Saving + Prezzo + quantità None deve usare default 1.0."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=None,  # NULL nel DB
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        # Deve comportarsi come qty=1
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_saving_price_qty_zero_defaults_to_1(self):
        """Saving + Prezzo + quantità 0 deve usare default 1.0."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=0.0,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        # Deve comportarsi come qty=1
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_cost_avoidance_price_qty_large(self):
        """Cost Avoidance + Prezzo + quantità > 1."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Cost Avoidance",
            driver="Prezzo",
            importo_richiesto_iniziale=2.0,
            importo_negoziato=1.8,
            quantita_annua=15000.0,
            percent_realizzo=80.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        expected_theoretical = 15000.0 * 0.2  # 3000.0
        expected_effective = 3000.0 * 0.8  # 2400.0
        
        self.assertAlmostEqual(theoretical, expected_theoretical, places=2)
        self.assertAlmostEqual(effective, expected_effective, places=2)
    
    def test_pagamenti_driver_qty_ignored(self):
        """Driver Pagamenti: quantità NON deve influenzare calcolo."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Pagamenti",
            spending_annuo=100000.0,
            giorni_pagamento_attuali=30,
            giorni_pagamento_negoziati=60,
            quantita_annua=999.0,  # Deve essere ignorato
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        
        # Formula: spending * (delta / 30) * coeff
        # 100000 * (30 / 30) * 0.005 = 500
        self.assertAlmostEqual(theoretical, 500.0, places=2)
    
    def test_derisking_qty_ignored(self):
        """Derisking: quantità NON deve influenzare calcolo (sempre 0)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Derisking",
            driver="Prezzo",
            quantita_annua=999.0,  # Deve essere ignorato
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertEqual(theoretical, 0.0)
        self.assertEqual(effective, 0.0)


if __name__ == '__main__':
    unittest.main()
```

**COVERAGE TEST**:
- ✅ Saving + Prezzo + qty=1, qty>1, qty decimale, qty=None, qty=0
- ✅ Cost Avoidance + Prezzo + qty>1 con realizzo parziale
- ✅ Pagamenti (qty ignorato)
- ✅ Derisking (qty ignorato)

### 3. Strategia di Retrocompatibilità

| Aspetto | Strategia | Verificato |
|---------|-----------|------------|
| **DB Schema** | Campo `quantita_annua REAL` già presente nella tabella `vsm_events` (linea 215 database_manager.py) | ✅ |
| **Mapping Persistence** | Campo già mappato in INSERT/UPDATE/SELECT query (linee 1849, 1864, 1893, 1908, 2074, 2098) | ✅ |
| **Record Esistenti** | Attualmente `quantita_annua` salvato come 0.0 (default dataclass) perché UI non lo popola. Con modifica: NULL o 0 → default a 1.0 nel calcolo | ✅ |
| **UI Default** | Inizializzare entry a "1" per coerenza con comportamento storico e default retrocompatibile | ✅ |
| **Caricamento Edit Mode** | Se `quantita_annua` è None o 0 → mostrare "1" nella UI, calcolare con 1.0. Se > 0 → mostrare valore effettivo | ✅ |
| **Calcolo Defensive** | `qty = self.quantita_annua if self.quantita_annua and self.quantita_annua > 0 else 1.0` protegge da NULL/0/negative | ✅ |
| **Impact Regeneration** | DELETE-REGENERATE-SAVE (linea 43 services/vsm_persistence.py) ricalcola automaticamente impatti con nuova formula | ✅ |
| **grid_remove() preserva valori** | All'interno della stessa sessione dialog, switch driver Prezzo↔Pagamenti preserva valore Annual Q.ty inserito | ✅ |

### 4. Formula Applicata (VERIFICATA NEL CODICE)

**Per Driver "Prezzo"** (Saving / Cost Avoidance):
```
⚠️ CORREZIONE IMPORTANTE: % Realization NON è nel theoretical, solo nell'effective!

Theoretical Value = Annual Q.ty × (Price_Budget - Price_Negotiated)
Effective Value = Theoretical Value × (% Realization / 100)
```

**Dettaglio per tipo evento**:
- **Saving**: 
  - Theoretical = `qty × (importo_bdg - importo_negoziato)`
  - Effective = `theoretical × (percent_realizzo / 100.0)`
  
- **Cost Avoidance**: 
  - Theoretical = `qty × (importo_richiesto_iniziale - importo_negoziato)`
  - Effective = `theoretical × (percent_realizzo / 100.0)`

**Per Driver "Pagamenti"** (Saving / Cost Avoidance) - INVARIATA:
```
Theoretical Value = Spending Annuo × (Δ Giorni / 30) × Coefficiente
Effective Value = Theoretical Value (no realization % applicata - pagamenti deterministici)
```

Dove:
- Δ Giorni = `giorni_pagamento_negoziati - giorni_pagamento_attuali`
- Coefficiente = costo opportunità del capitale (default 0.005 = 0.5% mensile)

**Per "Derisking"** (tutti i driver) - INVARIATA:
```
Theoretical Value = 0.0
Effective Value = 0.0
```

**NOTA ARCHITETTURALE**:
- `calculate_theoretical_value()`: calcolo base senza realizzo
- `calculate_effective_value()`: applica realizzo al theoretical (solo per driver Prezzo)
- Questa separazione garantisce che modifiche al theoretical non rompano l'applicazione del realizzo

### 5. Rischi di Regressione (VERIFICATI NEL CODICE)

| Rischio | Probabilità | Mitigazione | Verificato |
|---------|-------------|-------------|------------|
| **Calcolo errato per eventi Pagamenti** | NULLA | Formula modificata SOLO per driver "Prezzo", blocco `if self.driver == "Pagamenti"` non toccato (linea 161-173 vsm_event.py) | ✅ |
| **Eventi Derisking influenzati** | NULLA | Derisking ritorna sempre 0.0, nessuna modifica al return statement (linea 189-190 vsm_event.py) | ✅ |
| **Vecchi eventi calcolati male** | NULLA | Default `qty=1.0` se None/0 preserva comportamento storico IDENTICO (test backward compatibility aggiunto) | ✅ |
| **UI instabile su switch driver** | NULLA | Usa `grid_remove()` (NON `grid_forget()`), pattern già testato per `payment_fields_frame` (linea 386-420 vsm_event_dialog.py) | ✅ |
| **Validazione troppo rigida** | NULLA | Validazione solo quando driver="Prezzo", float supporta decimali, controllo `> 0` appropriato | ✅ |
| **NULL pointer in calcolo** | NULLA | Defensive check `if quantita_annua and quantita_annua > 0` previene crash, default safe a 1.0 | ✅ |
| **Impatti mensili errati** | NULLA | Engine non modificato, usa `event.calculate_theoretical_value()` e `calculate_effective_value()` (linea 313-314 vsm_engine.py) | ✅ |
| **Valori preservati durante switch** | NULLA | `grid_remove()` preserva widget e valori in memoria (non distrugge), testabile manualmente | ✅ |
| **Edit mode sovrascrive quantità** | NULLA | Populate avviene PRIMA di `_on_event_type_changed()` e `_on_driver_changed()` (linee 479-481 vsm_event_dialog.py) | ✅ |
| **DB persistence rotto** | NULLA | Campo `quantita_annua` già mappato in INSERT/UPDATE/SELECT (linee 1849-2165 database_manager.py) | ✅ |

**Rischio Complessivo**: **NULLO / MOLTO BASSO**

**Motivi**:
1. Campo DB già esiste (nessuna migrazione schema)
2. Persistence layer già mappa il campo (nessuna modifica query)
3. Modifica calcolo isolata in un solo metodo per un solo driver
4. Pattern UI già collaudato (riuso logica payment_fields)
5. Backward compatibility garantita da default qty=1.0
6. Test coverage completa (model + engine + persistence)

---

## IMPLEMENTAZIONE (PIANO VERIFICATO)

### Phase 1: UI Enhancement (ui/dialogs/vsm_event_dialog.py)

**Step 1.1: Create Widget** (linea ~238 dopo `entry_percent_realizzo`)
```python
# Quantità Annua (visibile solo per driver Prezzo)
self.lbl_quantita_annua = ttk.Label(self.price_fields_frame, text=_("Quantità Annua:"))
self.entry_quantita_annua = ttk.Entry(self.price_fields_frame, width=20)
self.entry_quantita_annua.insert(0, "1")  # Default retrocompatibile
```

**Step 1.2: Layout Logic in `_on_driver_changed()`** (linea 368-420)

Modificare layout per **Saving** (circa linea 390-398):
```python
if event_type == "Saving":
    # Row 0: Importo a Budget
    self.lbl_importo_bdg.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_importo_bdg.grid(row=0, column=1, sticky="w", pady=5)
    
    # Row 1: Importo Negoziato
    self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_importo_negoziato.grid(row=1, column=1, sticky="w", pady=5)
    
    # Row 2: Quantità Annua (NUOVO)
    self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_quantita_annua.grid(row=2, column=1, sticky="w", pady=5)
    
    # Row 3: % Realizzo (spostato da row=2 a row=3)
    self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_percent_realizzo.grid(row=3, column=1, sticky="w", pady=5)
    
    # Hide Cost Avoidance specific field
    self.lbl_importo_richiesto.grid_remove()
    self.entry_importo_richiesto.grid_remove()
```

Modificare layout per **Cost Avoidance** (circa linea 400-418):
```python
elif event_type == "Cost Avoidance":
    # Row 0: Importo Richiesto Iniziale
    self.lbl_importo_richiesto.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_importo_richiesto.grid(row=0, column=1, sticky="w", pady=5)
    
    # Row 1: Importo Negoziato
    self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_importo_negoziato.grid(row=1, column=1, sticky="w", pady=5)
    
    # Row 2: Quantità Annua (NUOVO)
    self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_quantita_annua.grid(row=2, column=1, sticky="w", pady=5)
    
    # Row 3: % Realizzo (spostato da row=2 a row=3)
    self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
    self.entry_percent_realizzo.grid(row=3, column=1, sticky="w", pady=5)
    
    # Hide Saving specific field
    self.lbl_importo_bdg.grid_remove()
    self.entry_importo_bdg.grid_remove()
```

**NOTA**: `grid_remove()` preserva automaticamente i valori dei widget quando nascosti. Nessuna logica aggiuntiva necessaria per preservare Annual Q.ty durante switch driver.

**Step 1.3: Validation in `_validate_and_save()`** (linea 504-650)

**Approccio**: Validare `quantita_annua` UNA SOLA VOLTA all'inizio del blocco `if event_type == "Saving"`, PRIMA del blocco `if driver == "Prezzo"`.

Posizione: circa linea 540 (subito dopo `driver = self._get_driver_internal()`):
```python
# Inizializza quantita_annua a None (sarà valorizzato solo per driver Prezzo)
quantita_annua = None

if event_type == "Saving":
    # Validazione quantità annua per driver Prezzo
    if driver == "Prezzo":
        # Valida quantità annua
        try:
            quantita_annua = float(self.entry_quantita_annua.get().strip())
            if quantita_annua <= 0:
                raise ValueError(_("Quantità Annua deve essere maggiore di zero."))
        except ValueError as e:
            if "could not convert" in str(e):
                raise ValueError(_("Quantità Annua deve essere un numero valido."))
            raise
        
        # Valida importi (codice esistente)
        try:
            importo_bdg = float(self.entry_importo_bdg.get().strip())
        # ... resto validazione esistente ...
```

Stesso pattern per `elif event_type == "Cost Avoidance"` (circa linea 627):
```python
elif event_type == "Cost Avoidance":
    # Validazione quantità annua per driver Prezzo
    if driver == "Prezzo":
        # Valida quantità annua
        try:
            quantita_annua = float(self.entry_quantita_annua.get().strip())
            if quantita_annua <= 0:
                raise ValueError(_("Quantità Annua deve essere maggiore di zero."))
        except ValueError as e:
            if "could not convert" in str(e):
                raise ValueError(_("Quantità Annua deve essere un numero valido."))
            raise
    
    # Valida importi (codice esistente)
    try:
        importo_richiesto_iniziale = float(self.entry_importo_richiesto.get().strip())
    # ... resto validazione esistente ...
```

**Step 1.4: Data Binding - Save** (linea 656, costruzione VSMEvent)

Attualmente il costruttore NON passa `quantita_annua`. Modificare:
```python
event = VSMEvent(
    id=self.event_id,
    event_date=datetime.combine(event_date, datetime.min.time()),
    username=self.current_username,
    buyer=buyer,
    event_type=event_type,
    action=action,
    description=description,
    reference=reference,
    importo_bdg=importo_bdg if importo_bdg is not None else 0.0,
    importo_negoziato=importo_negoziato if importo_negoziato is not None else 0.0,
    importo_richiesto_iniziale=importo_richiesto_iniziale,
    quantita_annua=quantita_annua if quantita_annua is not None else 1.0,  # AGGIUNGERE
    percent_realizzo=percent_realizzo,
    driver=driver,
    spending_annuo=spending_annuo if spending_annuo is not None else 0.0,
    giorni_pagamento_attuali=giorni_pagamento_attuali,
    giorni_pagamento_negoziati=giorni_pagamento_negoziati,
    opex_ripetitivo=opex_ripetitivo
)
```

**Step 1.5: Data Binding - Load** (linea 471, dentro `_load_event_data()`)

Aggiungere DOPO il populate di `percent_realizzo` (linea ~471):
```python
# Dopo populate percent_realizzo
self.entry_percent_realizzo.delete(0, tk.END)
self.entry_percent_realizzo.insert(0, str(event.percent_realizzo))

# Popolare quantita_annua (NUOVO)
if event.quantita_annua and event.quantita_annua > 0:
    self.entry_quantita_annua.delete(0, tk.END)
    self.entry_quantita_annua.insert(0, str(event.quantita_annua))
else:
    # Default retrocompatibile per eventi storici con NULL o 0
    self.entry_quantita_annua.delete(0, tk.END)
    self.entry_quantita_annua.insert(0, "1")

# Poi continua con driver e campi Pagamenti (codice esistente linea ~473+)
if event.driver:
    # ...
```

**TIMING VERIFICATO**: Il populate avviene PRIMA di `_on_event_type_changed()` e `_on_driver_changed()` (linee 479-481), quindi il valore NON viene sovrascritto. ✅ SICURO.

### Phase 2: Calculation Formula Update (models/vsm_event.py)

**Step 2.1: Add Defensive Default**
- All'inizio di `calculate_theoretical_value()`, aggiungere:
  ```python
  qty = self.quantita_annua if self.quantita_annua and self.quantita_annua > 0 else 1.0
  ```

**Step 2.2: Update Price Driver Formula**
- Per `driver == "Prezzo"`:
  - Saving: `return qty * (self.importo_bdg - self.importo_negoziato)`
  - Cost Avoidance: `return qty * (self.importo_richiesto_iniziale - self.importo_negoziato)`
- Non toccare calcoli per Pagamenti o Derisking

### Phase 3: Database & Persistence Verification

**Step 3.1: Verify DB Schema**
- Confermare che `quantita_annua REAL` esiste in `vsm_events` table
- **Risultato**: Campo già presente, nessuna migrazione necessaria

**Step 3.2: Verify Persistence Mapping**
- Verificare che `VSMEvent.to_dict()` e `from_dict()` gestiscono `quantita_annua`
- **Risultato**: Mapping automatico via dataclass fields, già funzionante

**Step 3.3: Handle NULL Values**
- Se necessario, aggiungere logica in `from_dict()` per convertire NULL → 1.0
- Alternativamente, gestire nel calcolo stesso (già fatto in Step 2.1)

### Phase 4: Testing & Validation

**Step 4.1: Unit Tests**
- Aggiungere test case in [tests/test_vsm_engine.py](tests/test_vsm_engine.py) per nuova formula
- Test con qty=1, qty>1, qty decimale
- Test backward compatibility con qty=None

**Step 4.2: Manual Testing Checklist**
1. ✓ New Saving + Driver Price + Q.ty = 1
2. ✓ New Saving + Driver Price + Q.ty > 1
3. ✓ New Cost Avoidance + Driver Price + Q.ty > 1
4. ✓ Driver diverso da Price → campo nascosto, comportamento invariato
5. ✓ Apertura e modifica di vecchi eventi già salvati
6. ✓ Verifica calcolo corretto del valore teorico
7. ✓ Verifica impatti mensili riflettono il valore moltiplicato
8. ✓ Validazione: Q.ty = 0, negativa, non numerica → errore
9. ✓ Validazione: Q.ty decimale (e.g., 1500.5) → accettato
10. ✓ Switch driver Prezzo ↔ Pagamenti → campo compare/scompare

---

## FILE COINVOLTI (VERIFICATI NEL CODICE REALE)

### Modifiche Necessarie

1. **[ui/dialogs/vsm_event_dialog.py](ui/dialogs/vsm_event_dialog.py)** - 5 modifiche puntuali
   - **Linea ~238**: Creare widget `lbl_quantita_annua` e `entry_quantita_annua` dentro `price_fields_frame`
   - **Linea 390-398**: Aggiornare layout Saving in `_on_driver_changed()` - aggiungere row=2 per quantità, spostare realizzo a row=3
   - **Linea 400-418**: Aggiornare layout Cost Avoidance in `_on_driver_changed()` - aggiungere row=2 per quantità, spostare realizzo a row=3
   - **Linea ~540 e ~627**: Aggiungere validazione `quantita_annua` in `_validate_and_save()` per entrambi event types quando driver="Prezzo"
   - **Linea ~656**: Aggiungere parametro `quantita_annua` nel costruttore VSMEvent
   - **Linea ~471**: Aggiungere populate `entry_quantita_annua` in `_load_event_data()` (edit mode)

2. **[models/vsm_event.py](models/vsm_event.py)** - 1 modifica chirurgica
   - **Linea 174-192**: Modificare blocco driver "Prezzo" in `calculate_theoretical_value()` per moltiplicare per qty (con default 1.0)

3. **[tests/test_vsm_event_model.py](tests/test_vsm_event_model.py)** - NUOVO FILE
   - Creare file con 8 test cases per verificare formule calcolo con quantità variabili

### Nessuna Modifica Necessaria (VERIFICATO ✅)

4. **[database_manager.py](database_manager.py)**
   - Campo `quantita_annua REAL` già presente nello schema (linea 215)
   - Già mappato in INSERT (linee 1849, 1864), UPDATE (linee 1893, 1908), SELECT (linee 2074, 2098, 2128, 2165)

5. **[services/vsm_persistence.py](services/vsm_persistence.py)**
   - Mapping automatico via VSMEvent dataclass già funzionante
   - Pattern DELETE-REGENERATE-SAVE gestisce automaticamente ricalcolo impatti

6. **[services/vsm_engine.py](services/vsm_engine.py)**
   - Engine usa `event.calculate_theoretical_value()` e `calculate_effective_value()` (linea 313-314)
   - Nessuna modifica necessaria, usa automaticamente i nuovi valori calcolati

---

## ATTENZIONE SUL SIGNIFICATO DEI CAMPI

### Per driver "Price"
- **Budget Amount** e **Negotiated Amount** rappresentano **prezzo per unità**
- **Annual Q.ty** rappresenta il **volume annuo** (quantità di pezzi/unità)
- Il prodotto fornisce il valore economico totale annuo

### Considerazione UX Futura
Le label attuali "Importo a Budget" e "Importo Negoziato" non indicano esplicitamente che sono prezzi unitari quando Annual Q.ty > 1.

**Opzioni** (per fase separata):
- Aggiungere tooltip: "Prezzo unitario"
- Aggiungere suffisso label: "Importo a Budget (per unità)"
- Lasciare invariato (l'utente capisce dal contesto)

**Decisione**: Non rinominare ora. Prima implementare la logica corretta. Eventuali miglioramenti semantici UI si valuteranno in una fase separata.

---

## FORMULA DI ESEMPIO COMPLETA

```python
# models/vsm_event.py - calculate_theoretical_value()

def calculate_theoretical_value(self) -> float:
    """Calcola il valore teorico dell'evento basato su driver e dati economici."""
    
    # Backward compatibility: default quantity to 1.0 if not set
    qty = self.quantita_annua if self.quantita_annua and self.quantita_annua > 0 else 1.0
    
    if self.driver == "Prezzo":
        if self.event_type == "Cost Avoidance":
            # NEW: multiply by annual quantity
            return qty * (self.importo_richiesto_iniziale - self.importo_negoziato)
        elif self.event_type == "Saving":
            # NEW: multiply by annual quantity
            return qty * (self.importo_bdg - self.importo_negoziato)
        else:  # Derisking
            return 0.0
    
    elif self.driver == "Pagamenti":
        # UNCHANGED: Payment driver logic
        delta_giorni = self.giorni_pagamento_negoziati - self.giorni_pagamento_attuali
        coefficiente = get_pagamenti_coefficient()
        return self.spending_annuo * (delta_giorni / 30.0) * coefficiente
    
    return 0.0
```

---

## OUTPUT RICHIESTO

### Pre-Implementation Checklist
- [x] Identificati file da modificare (2 file principali)
- [x] Individuati punti esatti di modifica
- [x] Definita strategia di retrocompatibilità (default qty=1.0)
- [x] Formula matematica validata
- [x] Analisi rischi di regressione completata (rischio MOLTO BASSO)
- [x] Verificato che campo DB già esiste (no migration)

### Post-Implementation Report (da completare dopo implementazione)

**File Modificati:**
- [ ] ui/dialogs/vsm_event_dialog.py
- [ ] models/vsm_event.py

**Campi Aggiunti (UI):**
- [ ] Label "Quantità Annua" / "Annual Q.ty"
- [ ] Entry widget per input numerico (float)

**Logica Applicata:**
- [ ] Formula: `qty × (budget - negotiated)` per Saving/Price
- [ ] Formula: `qty × (initial_requested - negotiated)` per Cost Avoidance/Price
- [ ] Default qty = 1.0 per retrocompatibilità
- [ ] Validazione: float, > 0, solo per driver Prezzo

**Comportamento UI:**
- [ ] Campo visibile solo con Driver = "Prezzo"
- [ ] Campo nascosto con Driver = "Pagamenti" (grid_forget)
- [ ] Default value = "1" in form
- [ ] Posizione: tra Importo Negoziato e % Realizzo

**Compatibilità Record Esistenti:**
- [ ] Eventi con quantita_annua NULL → calcolo usa 1.0
- [ ] Eventi con quantita_annua valorizzato → usa valore esistente
- [ ] Modifica vecchi eventi → campo popolato con valore storico o 1
- [ ] Salvataggio → persiste quantita_annua correttamente

---

## MINI CHECKLIST TEST MANUALI

### Test 1: Saving + Price + Q.ty = 1 (Macchinario)
- [ ] Creare nuovo evento: Saving, Driver Prezzo
- [ ] Budget Amount = 20000, Negotiated = 18000, Q.ty = 1, Realizzo = 100%
- [ ] Salvare evento
- [ ] Verificare impatti mensili: valore teorico = 2000, valore effettivo = 2000
- [ ] Se OPEX Ripetitivo: verificare distribuzione 24 mesi con pro-rata primo mese

### Test 2: Saving + Price + Q.ty > 1 (Produzione)
- [ ] Creare nuovo evento: Saving, Driver Prezzo
- [ ] Budget Amount = 1.5, Negotiated = 1.3, Q.ty = 20000, Realizzo = 100%
- [ ] Salvare evento
- [ ] Verificare impatti mensili: valore teorico totale = 4000
- [ ] Se one-shot: valore concentrato nel mese dell'evento

### Test 3: Cost Avoidance + Price + Q.ty > 1
- [ ] Creare nuovo evento: Cost Avoidance, Driver Prezzo
- [ ] Initial Requested = 2.0, Negotiated = 1.8, Q.ty = 15000, Realizzo = 80%
- [ ] Salvare evento
- [ ] Verificare: valore teorico = 3000, valore effettivo = 2400

### Test 4: Driver ≠ Price → Comportamento Invariato
- [ ] Creare evento Saving, Driver Pagamenti
- [ ] Verificare: campo Annual Q.ty NON visibile
- [ ] Spending Annuo = 100000, Giorni Attuali = 30, Giorni Negoziati = 60
- [ ] Salvare e verificare calcolo basato su payment terms (no qty involved)

### Test 5: Apertura Vecchi Eventi (Backward Compatibility)
- [ ] Se esistono eventi pre-implementazione, aprirli in edit
- [ ] Verificare: campo Annual Q.ty mostra "1" come default
- [ ] Verificare: calcolo rimane identico rispetto a prima (qty implicita = 1)
- [ ] Salvare senza modifiche → quantita_annua = 1.0 ora persistita
- [ ] Verificare: impatti mensili non cambiano rispetto a prima

### Test 6: Validazione Quantità
- [ ] Tentare salvataggio con Q.ty = "" (vuoto) → errore atteso
- [ ] Tentare salvataggio con Q.ty = 0 → errore atteso
- [ ] Tentare salvataggio con Q.ty = -100 → errore atteso
- [ ] Tentare salvataggio con Q.ty = "abc" → errore atteso
- [ ] Tentare salvataggio con Q.ty = 1500.5 (decimale) → successo atteso

### Test 7: Switch Driver
- [ ] Creare evento Saving, Driver Prezzo, impostare Q.ty = 5000
- [ ] Cambiare combo Driver da "Prezzo" a "Pagamenti"
- [ ] Verificare: campo Annual Q.ty scompare (grid_forget)
- [ ] Cambiare combo Driver da "Pagamenti" a "Prezzo"
- [ ] Verificare: campo Annual Q.ty ricompare con valore preservato (5000)

### Test 8: Cross-Platform (Linux / Windows)
- [ ] Testare su Linux: UI rendering corretto, nessun crash
- [ ] Testare su Windows: UI rendering corretto, nessun crash
- [ ] Verificare: comportamento identico su entrambe le piattaforme

---

## TECHNICAL NOTES

### UI Framework Context
- **Tkinter/ttk** - Standard Python GUI
- **Pattern**: `ttk.Entry` per input numerici
- **Validation**: Manuale in `_validate_and_save()` con try/except float conversion
- **Localization**: Usare `_("Quantità Annua:")` per i18n support (EN/IT)

### Data Model Context
- **VSMEvent** dataclass con 21 campi
- **quantita_annua** (float) è già presente nel modello
- **from_dict()** / **to_dict()** fanno serializzazione automatica
- **calculate_theoretical_value()** e **calculate_effective_value()** sono i metodi da aggiornare

### Persistence Context
- **Atomicity**: TRANSACTION + rollback su errore
- **Impact Regeneration**: DELETE old impacts + INSERT new impacts
- **Multi-user**: username field tracked in ogni operazione

### Calculation Context
- **Engine** [services/vsm_engine.py]: distribuisce valori mensili, NON calcola valori teorici
- **Model** [models/vsm_event.py]: calcola valori teorici/effettivi, engine li distribuisce
- **Separation**: Engine non modifica la formula, usa output del modello

---

## 📋 SUMMARY ESECUTIVO

### File da Modificare
- ✏️ **2 file**: `ui/dialogs/vsm_event_dialog.py` + `models/vsm_event.py`
- ➕ **1 nuovo test file**: `tests/test_vsm_event_model.py`
- ✅ **0 migration DB**: campo già esistente
- ✅ **0 modifica persistence**: mapping già funzionante

### Modifiche Principali
1. **UI** - Aggiungere widget "Quantità Annua" visibile solo per driver "Prezzo"
2. **Calcolo** - Moltiplicare differenza prezzo per quantità annua in `calculate_theoretical_value()`
3. **Test** - 8 test cases per formule con quantità variabili

### Backward Compatibility
- Default qty = 1.0 se NULL/0 → eventi storici calcolano come prima
- `grid_remove()` preserva valori → switch driver non perde dati inseriti
- Edit mode sicuro → populate prima dei trigger dinamici

### Risk Assessment
**NULLO / MOLTO BASSO** - Modifica circoscritta, campo già esistente, pattern UI collaudato, test coverage completa

### Ready for Implementation
✅ **SÌ** - Tutte le verifiche completate, piano dettagliato e sicuro

---

## NEXT STEPS

1. ✅ **Analisi completata** - Plan approvato
2. ⏳ **Implementazione**:
   - Modificare [ui/dialogs/vsm_event_dialog.py](ui/dialogs/vsm_event_dialog.py)
   - Modificare [models/vsm_event.py](models/vsm_event.py)
3. ⏳ **Verifica funzionale**:
   - Test manuali checklist
   - Verifica retrocompatibilità
4. ⏳ **Report finale**:
   - Riepilogo modifiche
   - Test eseguiti
   - Note per deploy

---

## ✅ RIEPILOGO VERIFICHE RICHIESTE (TUTTI I PUNTI CONFERMATI)

### 1. Valore interno driver VERIFICATO ✅
**Richiesta**: Capire se il confronto corretto è "Price" oppure "Prezzo"  
**Risposta**: 
- ✅ Valore interno: **"Prezzo"** (italiano)
- ✅ Confronti nel codice: `driver == "Prezzo"` e `driver_internal == "Prezzo"`
- ✅ Persistenza DB: "Prezzo" o "Pagamenti" (sempre italiano)
- ✅ UI tradotta: `_("Prezzo")` mostra "Prezzo" in IT, "Price" in EN (solo display, non interno)
- ✅ Metodi conversione: `_get_driver_internal()` ritorna sempre valore italiano

**Implicazione**: Tutti i confronti nel codice devono usare `"Prezzo"`, non `"Price"`.

### 2. Applicazione % Realization VERIFICATO ✅
**Richiesta**: Controllare sia calculate_theoretical_value() sia calculate_effective_value()  
**Risposta**:
- ✅ `calculate_theoretical_value()` (linea 151-192): Calcola valore GREZZO senza realizzo
  - Saving: `importo_bdg - importo_negoziato`
  - Cost Avoidance: `importo_richiesto_iniziale - importo_negoziato`
  - Pagamenti: `spending_annuo * (delta_giorni / 30) * coefficiente`
- ✅ `calculate_effective_value()` (linea 194-215): Applica realizzo SOLO per driver "Prezzo"
  - Se driver == "Prezzo": `theoretical * (percent_realizzo / 100.0)`
  - Se driver == "Pagamenti": `theoretical` (no realizzo - pagamenti deterministici)
- ✅ **Separazione corretta**: theoretical e effective NON vanno fuori sync
- ✅ **Nostra modifica**: Tocca SOLO theoretical (moltiplica per qty), effective continua ad applicare realizzo correttamente

**Implicazione**: La modifica alla formula theoretical è safe, non rompe l'applicazione del realizzo.

### 3. Comportamento UI grid_remove() VERIFICATO ✅
**Richiesta**: Campo deve scomparire quando driver != Price, ma valore deve essere preservato se utente torna a Price  
**Risposta**:
- ✅ Codice usa `grid_remove()` NON `grid_forget()` (linea 386-387)
- ✅ `grid_remove()`: Nasconde widget ma li mantiene in memoria con i loro valori
- ✅ **Automatico**: Nessuna logica aggiuntiva necessaria, il valore inserito è preservato quando si switcha driver
- ✅ Pattern già testato: `payment_fields_frame` usa stesso meccanismo grid_remove()

**Implicazione**: Quando utente switcha Prezzo→Pagamenti→Prezzo, il valore di Annual Q.ty rimane quello inserito. Funziona out-of-the-box.

### 4. Edit mode e popolamento VERIFICATO ✅
**Richiesta**: Popolamento quantita_annua deve avvenire dopo creazione widget e senza essere sovrascritto  
**Risposta**:
- ✅ **Timing**: `_load_event_data()` chiamato in `__init__` dopo `_build_ui()` (linea 89)
- ✅ **Ordine load**: Populate campi → `_on_event_type_changed()` → `_on_driver_changed()` (linee 479-481)
- ✅ **Momento populate**: Prima dei trigger dinamici, quindi NON viene sovrascritto
- ✅ **PROBLEMA IDENTIFICATO**: Attualmente quantita_annua NON viene popolato in `_load_event_data()` (linea 430-481)
- ✅ **SOLUZIONE**: Aggiungere populate dopo linea 471 (dopo percent_realizzo, prima di driver)

**Implicazione**: Sicuro aggiungere populate a linea ~472, nessun rischio di sovrascrittura da trigger successivi.

### 5. Test VERIFICATO ✅
**Richiesta**: Test mirati su modello VSMEvent per formula, non solo engine  
**Risposta**:
- ✅ **Test esistenti**: Solo `test_vsm_engine.py` (distribuzione) e `test_vsm_persistence.py` (save/load)
- ✅ **GAP**: Nessun test specifico per `calculate_theoretical_value()` e `calculate_effective_value()`
- ✅ **AZIONE**: Creare `tests/test_vsm_event_model.py` con 8 test cases:
  - qty=1, qty>1, qty decimale, qty=None, qty=0 (backward compatibility)
  - Saving + Cost Avoidance separati
  - Pagamenti e Derisking (qty deve essere ignorato)
  - Verifica applicazione realizzo solo per Prezzo

**Implicazione**: Test coverage completa per isolare i comportamenti della formula.

### 6. Cost Avoidance - Campi reali VERIFICATO ✅
**Richiesta**: Confermare campi reali usati nel modello attuale  
**Risposta**:
- ✅ Campo usato: `importo_richiesto_iniziale` (Optional[float]) (linea 61 vsm_event.py)
- ✅ Formula attuale Cost Avoidance: `importo_richiesto_iniziale - importo_negoziato` (linea 185-187)
- ✅ UI field: `entry_importo_richiesto` con label "Importo Richiesto Iniziale: *" (linea 233-234 dialog)
- ✅ Validazione: float conversion obbligatoria per Cost Avoidance (linea 627+ dialog)

**Implicazione**: Il naming nel piano è corretto, nessuna assunzione errata.

---

## ✅ TUTTE LE VERIFICHE COMPLETATE - PRONTO PER IMPLEMENTAZIONE CONSERVATIVA

**Punti chiave confermati**:
1. ✅ Driver interno è "Prezzo" (non "Price") - tutti i confronti devono usare "Prezzo"
2. ✅ Realizzo applicato solo in effective, theoretical è grezzo - separazione corretta
3. ✅ grid_remove() preserva valori automaticamente - nessuna logica extra per preservare
4. ✅ Edit mode timing sicuro per populate senza sovrascritture - populate prima dei trigger
5. ✅ Test model specifici da aggiungere per coverage completa - nuovo file test_vsm_event_model.py
6. ✅ Cost Avoidance usa importo_richiesto_iniziale - confermato nel codice

**Confidence level**: MOLTO ALTO  
**Risk level**: MOLTO BASSO  
**Ready to implement**: ✅ SÌ

---

## 🔧 NOTE TECNICHE PER IMPLEMENTAZIONE

### Ordine Consigliato delle Modifiche
1. **Prima**: Modificare `models/vsm_event.py` (calcolo) → test rapido con script Python
2. **Seconda**: Creare `tests/test_vsm_event_model.py` → run unit tests per verificare formule
3. **Terza**: Modificare `ui/dialogs/vsm_event_dialog.py` (UI) → test manuale con applicazione
4. **Quarta**: Test end-to-end manuali con checklist completa

### Attenzioni Critiche Durante Implementazione

**UI Dialog**:
- ⚠️ Creare widget DOPO `entry_percent_realizzo` (circa linea 238) ma PRIMA di payment_fields_frame
- ⚠️ Usare `grid_remove()` NON `grid_forget()` per nascondere campo (coerente con pattern esistente)
- ⚠️ Aggiungere populate quantita_annua in `_load_event_data()` DOPO percent_realizzo (linea ~471)
- ⚠️ Inizializzare variabile `quantita_annua = None` all'inizio di `_validate_and_save()` (linea ~540)
- ⚠️ Validare PRIMA del blocco validazione importi, per fallire fast

**Modello Calcolo**:
- ⚠️ Aggiungere riga `qty = ...` PRIMA del blocco `if self.event_type` (dopo blocco Pagamenti, linea ~174)
- ⚠️ NON modificare `calculate_effective_value()` - applica già correttamente realizzo
- ⚠️ NON toccare blocco driver "Pagamenti" (linee 161-173)  
- ⚠️ Mantenere commenti esistenti e aggiungere nota "moltiplicato per quantità"

**Test**:
- ⚠️ Test pagamenti deve verificare che qty=999 NON influenzi calcolo (formula completamente diversa)
- ⚠️ Test backward compatibility (qty=None, qty=0) sono critici per eventi storici

### Validazione Conservativa

**Pattern da usare** (già usato per altri campi):
```python
try:
    quantita_annua = float(self.entry_quantita_annua.get().strip())
    if quantita_annua <= 0:
        raise ValueError(_("Quantità Annua deve essere maggiore di zero."))
except ValueError as e:
    if "could not convert" in str(e):
        raise ValueError(_("Quantità Annua deve essere un numero valido."))
    raise
```

**Coerente con**:
- Validazione `importo_bdg` (linea ~543-546)
- Validazione `percent_realizzo` (linea ~555-561)
- Validazione `spending_annuo` (linea ~575-580)

### Localizzazione i18n

**Label**: `_("Quantità Annua:")` 
- IT: "Quantità Annua:"
- EN: "Annual Q.ty:" (sarà nella traduzione locale/en/LC_MESSAGES/messages.po)

**Messaggi errore**:
- `_("Quantità Annua deve essere un numero valido.")`
- `_("Quantità Annua deve essere maggiore di zero.")`

**NOTA**: Dopo implementazione, eseguire `compile_translations.py` per aggiornare file .mo

---

## QUESTIONS FOR REVIEW

1. **Label clarification**: Should we update Budget/Negotiated labels to indicate "per unit" when Annual Q.ty is visible?
   - **Suggestion**: Keep current labels, add as separate UX ticket

2. **Historical recalculation**: Should we provide a tool to recalculate impacts for existing events?
   - **Suggestion**: Not needed - impacts regenerate on edit, default qty=1 preserves behavior

3. **Display precision**: Should we show more decimal places in calculated values when working with large quantities?
   - **Suggestion**: Keep current format, precision is maintained internally

4. **Field naming**: Should internal field be renamed from `quantita_annua` to something more specific?
   - **Suggestion**: NO - field already exists in DB and model, renaming creates migration complexity

---

## RISK ASSESSMENT: ✅ VERY LOW

This is a **conservative, well-isolated change**:
- Leverages existing DB field
- Touches only 2 files with localized changes
- Maintains full backward compatibility
- Uses established UI patterns
- No impact on other drivers or event types
- Calculation logic is centralized and testable
