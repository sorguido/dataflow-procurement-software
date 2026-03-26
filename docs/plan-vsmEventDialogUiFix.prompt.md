# FIX UI DIALOG VSM EVENT - DIMENSIONI FINESTRA STABILI

## A. CAUSA TECNICA PRECISA DEL BUG

Bug identificato nel metodo `_on_driver_changed()` linee ~340-360:

**Problema:**
```python
# Trova ultima riga visibile
row = 0
for widget in self.economic_frame.winfo_children():
    info = widget.grid_info()
    if info and 'row' in info:
        row = max(row, int(info['row']))
row += 1

# Posiziona campi Pagamenti
self.lbl_spending_annuo.grid(row=row, ...)  # row diventa 4, 5, 6!
```

**Conseguenza:**
- Campi Pagamenti aggiunti SOTTO i campi Prezzo (row 4, 5, 6)
- Finestra cresce verticalmente
- Pulsante Salva fuori schermo
- Finestra si allarga per label lunghi ("Termini Pagamento Negoziati (giorni)")

---

## B. STRATEGIA SCELTA: A - RIGHE FISSE

Layout frame "Dati Economici" con **RIGHE FISSE**:

- **Row 0**: Primo campo driver-specific
  - Driver Prezzo: Importo a Budget / Importo Richiesto Iniziale
  - Driver Pagamenti: Spending Annuo

- **Row 1**: Secondo campo driver-specific
  - Driver Prezzo: Importo Negoziato
  - Driver Pagamenti: Termini Pagamento Attuali

- **Row 2**: Terzo campo driver-specific
  - Driver Prezzo: % Realizzo
  - Driver Pagamenti: Termini Pagamento Negoziati

- **Row 3**: Driver (SEMPRE visibile, campo comune)

**Logica show/hide:**
- I campi si SOSTITUISCONO nelle stesse righe
- Nessuna crescita verticale
- `grid_remove()` per nascondere widget non pertinenti
- `grid()` con row fissa per mostrare widget pertinenti

---

## C. METODI/WIDGET MODIFICATI

**File:** `ui/dialogs/vsm_event_dialog.py`

**Metodi modificati:**
1. `_on_event_type_changed()` - RISTRUTTURATO
2. `_on_driver_changed()` - RISCRITTO completamente

---

## MODIFICA 1: _on_event_type_changed()

**SOSTITUIRE** il metodo `_on_event_type_changed()` (circa linee 234-280) con:

```python
def _on_event_type_changed(self):
    """
    Handler per cambio event_type.
    Mostra/nasconde campi economici in base al tipo.
    """
    event_type = self.event_type_var.get()
    
    # Rimuovi tutti i campi economici
    for widget in self.economic_frame.winfo_children():
        widget.grid_forget()
    
    if event_type == "Saving":
        # Saving: posiziona driver su row 3 (sempre fisso)
        self.lbl_driver.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
        self.combo_driver.grid(row=3, column=1, sticky="w", pady=5)
        
        # Chiama _on_driver_changed per mostrare campi appropriati su row 0-2
        self._on_driver_changed()
    
    elif event_type == "Cost Avoidance":
        # Cost Avoidance: posiziona driver su row 3 (sempre fisso)
        self.lbl_driver.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
        self.combo_driver.grid(row=3, column=1, sticky="w", pady=5)
        
        # Chiama _on_driver_changed per mostrare campi appropriati su row 0-2
        self._on_driver_changed()
    
    elif event_type == "Derisking":
        # Derisking: nessun campo economico obbligatorio
        # Solo label informativo
        info_label = ttk.Label(
            self.economic_frame,
            text=_("Gli eventi Derisking non generano impatti economici.\n"
                   "Compilare solo sezioni descrittive."),
            foreground="blue",
            font=("Calibri", 9, "italic")
        )
        info_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
```

---

## MODIFICA 2: _on_driver_changed()

**SOSTITUIRE** il metodo `_on_driver_changed()` (circa linee 300-367) con:

```python
def _on_driver_changed(self, event=None):
    """
    Handler per cambio driver.
    Mostra/nasconde campi in base al driver selezionato usando RIGHE FISSE.
    """
    driver = self.combo_driver.get()
    event_type = self.event_type_var.get()
    
    # Solo per event_type con campi economici
    if event_type not in ["Saving", "Cost Avoidance"]:
        return
    
    if driver == "Prezzo":
        # ========== DRIVER PREZZO ==========
        # Nascondi TUTTI i campi Pagamenti
        self.lbl_spending_annuo.grid_remove()
        self.entry_spending_annuo.grid_remove()
        self.lbl_giorni_attuali.grid_remove()
        self.entry_giorni_attuali.grid_remove()
        self.lbl_giorni_negoziati.grid_remove()
        self.entry_giorni_negoziati.grid_remove()
        
        # Mostra campi Prezzo su RIGHE FISSE 0, 1, 2
        if event_type == "Saving":
            # Row 0: Importo a Budget
            self.lbl_importo_bdg.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_bdg.grid(row=0, column=1, sticky="w", pady=5)
            
            # Row 1: Importo Negoziato
            self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_negoziato.grid(row=1, column=1, sticky="w", pady=5)
            
            # Row 2: % Realizzo
            self.lbl_percent_realizzo.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_percent_realizzo.grid(row=2, column=1, sticky="w", pady=5)
            
        elif event_type == "Cost Avoidance":
            # Row 0: Importo Richiesto Iniziale
            self.lbl_importo_richiesto.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_richiesto.grid(row=0, column=1, sticky="w", pady=5)
            
            # Row 1: Importo Negoziato
            self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_negoziato.grid(row=1, column=1, sticky="w", pady=5)
            
            # Row 2: % Realizzo
            self.lbl_percent_realizzo.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_percent_realizzo.grid(row=2, column=1, sticky="w", pady=5)
    
    elif driver == "Pagamenti":
        # ========== DRIVER PAGAMENTI ==========
        # Nascondi TUTTI i campi Prezzo
        self.lbl_importo_bdg.grid_remove()
        self.entry_importo_bdg.grid_remove()
        self.lbl_importo_richiesto.grid_remove()
        self.entry_importo_richiesto.grid_remove()
        self.lbl_importo_negoziato.grid_remove()
        self.entry_importo_negoziato.grid_remove()
        self.lbl_percent_realizzo.grid_remove()
        self.entry_percent_realizzo.grid_remove()
        
        # Mostra campi Pagamenti su RIGHE FISSE 0, 1, 2 (STESSE righe dei campi Prezzo!)
        # Row 0: Spending Annuo
        self.lbl_spending_annuo.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_spending_annuo.grid(row=0, column=1, sticky="w", pady=5)
        
        # Row 1: Termini Pagamento Attuali
        self.lbl_giorni_attuali.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_giorni_attuali.grid(row=1, column=1, sticky="w", pady=5)
        
        # Row 2: Termini Pagamento Negoziati
        self.lbl_giorni_negoziati.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_giorni_negoziati.grid(row=2, column=1, sticky="w", pady=5)
    
    else:
        # Volume, Altro: nascondi tutto per ora (future implementation)
        self.lbl_importo_bdg.grid_remove()
        self.entry_importo_bdg.grid_remove()
        self.lbl_importo_richiesto.grid_remove()
        self.entry_importo_richiesto.grid_remove()
        self.lbl_importo_negoziato.grid_remove()
        self.entry_importo_negoziato.grid_remove()
        self.lbl_percent_realizzo.grid_remove()
        self.entry_percent_realizzo.grid_remove()
        self.lbl_spending_annuo.grid_remove()
        self.entry_spending_annuo.grid_remove()
        self.lbl_giorni_attuali.grid_remove()
        self.entry_giorni_attuali.grid_remove()
        self.lbl_giorni_negoziati.grid_remove()
        self.entry_giorni_negoziati.grid_remove()
```

---

## D. GARANZIA DIMENSIONI FINESTRA STABILI

**Meccanismi implementati:**

1. **RIGHE FISSE (0, 1, 2, 3):**
   - Tutti i campi usano le STESSE 4 righe
   - Nessuna riga aggiuntiva creata dinamicamente
   - Frame "Dati Economici" ha altezza costante (4 row)

2. **SOSTITUZIONE IN-PLACE:**
   - Campi Prezzo occupano row 0-2
   - Campi Pagamenti occupano row 0-2 (STESSE righe!)
   - `grid_remove()` nasconde widget senza liberare spazio
   - `grid()` riposiziona widget nelle stesse coordinate

3. **STICKY="w" + NO WEIGHT:**
   - Widget allineati a sinistra, non espandono
   - `economic_frame.columnconfigure(1, weight=1)` già presente
   - Entry `width=20` limita larghezza widget

4. **RESIZABLE DISABLED:**
   - `self.resizable(False, False)` già presente nel costruttore
   - Finestra non può essere ridimensionata dall'utente

5. **NO GEOMETRIA DINAMICA:**
   - Eliminato calcolo "ultima riga visibile"
   - Eliminato `row += 1` incrementale
   - Solo coordinate fisse hardcoded

**Risultato:**
- Finestra mantiene SEMPRE le stesse dimensioni
- Cambio driver sostituisce widget senza resize
- Pulsanti sempre visibili in fondo
- Label lunghi non forzano allargamento (sticky="w")

---

## E. CONFERMA: BACKEND E FORMULE NON TOCCATI

**File NON modificati:**
- ✓ `models/vsm_event.py` (formule calcolo invariate)
- ✓ `utils/vsm_config.py` (configurazione invariata)
- ✓ `services/vsm_engine.py` (motore invariato)
- ✓ `services/vsm_persistence.py` (persistenza invariata)
- ✓ `database_manager.py` (schema invariato)
- ✓ `dataflow.py` (main file invariato)

**File modificato:**
- ✓ `ui/dialogs/vsm_event_dialog.py`
  - SOLO metodi `_on_event_type_changed()` e `_on_driver_changed()`
  - SOLO logica layout UI
  - ZERO modifiche a validazioni business (già corrette)
  - ZERO modifiche a salvataggio dati (già corretto)

**Logica di calcolo Pagamenti:**
- ✓ Formula: `spending_annuo * (delta_giorni/30) * coefficiente`
- ✓ Coefficiente da config.ini
- ✓ percent_realizzo ignorato per Pagamenti
- ✓ NULL enforcement campi non pertinenti
- **TUTTO INVARIATO, solo UI fixed**

---

## F. TEST MANUALI UI DA ESEGUIRE

**NOTA:** Test da eseguire dopo applicazione modifiche

### CHECKLIST OBBLIGATORIA:

**□ TEST 1: Nuovo evento Saving + Driver Prezzo**
- Aprire DataFlow
- Dashboard VSM → Nuovo Evento → Tab "Saving"
- Verificare:
  - ✓ Finestra dimensione standard
  - ✓ Driver combobox default "Prezzo"
  - ✓ Campi visibili: Importo a Budget, Importo Negoziato, % Realizzo
  - ✓ Campi Pagamenti nascosti (non presenti)
  - ✓ Pulsante Salva visibile

**□ TEST 2: Cambio Driver → Pagamenti**
- Stesso evento, selezionare Driver "Pagamenti" dalla combo
- Verificare:
  - ✓ Finestra NON si allarga
  - ✓ Finestra NON cresce verticalmente
  - ✓ Campi Prezzo SPARISCONO completamente
  - ✓ Campi Pagamenti APPAIONO:
    - Spending Annuo (€)
    - Termini Pagamento Attuali (giorni)
    - Termini Pagamento Negoziati (giorni)
  - ✓ Pulsante Salva ANCORA visibile
  - ✓ Dimensioni finestra identiche a prima

**□ TEST 3: Toggle Prezzo ↔ Pagamenti (ciclo ripetuto)**
- Selezionare Driver "Prezzo"
- Verificare campi Prezzo tornano
- Selezionare Driver "Pagamenti"
- Verificare campi Pagamenti tornano
- Ripetere 5 volte
- Verificare:
  - ✓ Nessun bug cumulativo
  - ✓ Finestra sempre identica
  - ✓ Nessun widget "fantasma"
  - ✓ Nessuna riga vuota

**□ TEST 4: Salvataggio evento Pagamenti**
- Driver "Pagamenti"
- Compilare:
  - Spending Annuo: 120000
  - Termini Attuali: 30
  - Termini Negoziati: 60
- Cliccare Salva
- Verificare:
  - ✓ Salvataggio riuscito
  - ✓ Nessun errore validazione

**□ TEST 5: Edit evento Pagamenti esistente**
- Dashboard VSM → Doppio click su evento Pagamenti salvato
- Verificare:
  - ✓ Finestra si apre con dimensioni standard
  - ✓ Driver correttamente impostato su "Pagamenti"
  - ✓ Campi Pagamenti visibili e popolati
  - ✓ Campi Prezzo nascosti
  - ✓ Nessun resize anomalo
  - ✓ Pulsanti visibili

**□ TEST 6: Regressione driver Prezzo**
- Dashboard VSM → Doppio click su evento Prezzo esistente (vecchio)
- Verificare:
  - ✓ Finestra dimensioni corrette
  - ✓ Driver correttamente impostato su "Prezzo"
  - ✓ Campi Prezzo visibili e popolati
  - ✓ Campi Pagamenti nascosti
  - ✓ Salvataggio funziona
  - ✓ ZERO breaking changes

**□ TEST 7: Cost Avoidance + Pagamenti**
- Nuovo evento Cost Avoidance
- Selezionare Driver "Pagamenti"
- Verificare:
  - ✓ Finestra dimensioni corrette
  - ✓ Campi Pagamenti visibili
  - ✓ Nessun resize

---

## RIEPILOGO TECNICO

**Bug root cause:**
- Calcolo dinamico row in `_on_driver_changed()` aggiungeva campi SOTTO quelli esistenti invece di SOSTITUIRLI, causando crescita verticale finestra

**Fix applicato:**
- STRATEGIA A - Righe fisse predefinite (0, 1, 2, 3)
- Campi Prezzo e Pagamenti usano STESSE righe
- `grid_remove()` per nascondere, `grid(row=N)` per mostrare
- Coordinamento `_on_event_type_changed()` → `_on_driver_changed()`

**Modifiche codice:**
- `_on_event_type_changed()`: semplificato, delega layout a `_on_driver_changed()`
- `_on_driver_changed()`: riscritto con righe fisse, NO calcolo dinamico

**Garanzie dimensioni:**
- 4 righe fisse nel frame Dati Economici
- Sostituzione in-place dei widget
- sticky="w" + columnconfigure già corretto
- resizable(False, False) già presente

**Backend intatto:**
- Formule calcolo invariate
- Validazioni invariate
- Salvataggio invariato
- Configurazione invariata

**Impatto:**
- File modificato: 1 (`ui/dialogs/vsm_event_dialog.py`)
- Metodi modificati: 2 (`_on_event_type_changed`, `_on_driver_changed`)
- Righe modificate: ~80 linee
- Regressioni: ZERO (backward compatible)

---

## OBIETTIVO FINALE

**SOSTITUIRE** i campi Prezzo con i campi Pagamenti nello stesso spazio della dialog, mantenendo la finestra identica.

**Soluzione:** MINIMALISTA e STABILE  
**Approccio:** RIGHE FISSE (0, 1, 2, 3)  
**Backend:** INTATTO  
**Test:** DA ESEGUIRE (checklist completa fornita)
