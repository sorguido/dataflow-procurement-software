# Proposta Tecnica: Pulsante Actions ▾

**Data:** 22 marzo 2026  
**Versione:** 1.0  
**Obiettivo:** Redesign UX Main Dashboard - Sostituzione pulsanti operativi con dropdown compatto

---

## 1. Comprensione dell'obiettivo

### Contesto
Stiamo lavorando al redesign della Main Dashboard di DataFlow, un'app desktop Python/Tkinter.

**Elementi già implementati e da mantenere:**
- ✅ Global Search centrale
- ✅ Placeholder UX corretto nella search bar
- ✅ Enter collegato alla ricerca esistente
- ✅ Enter su campo vuoto = "Clear Filters"
- ✅ Filtri avanzati collassabili
- ✅ Layout principale stabile

**Soluzione scartata:**
- ❌ Contextual toolbar dinamica sopra/in prossimità del notebook
  - Problemi di layout
  - Fragile in Tkinter/ttk.Notebook
  - Esteticamente insoddisfacente
  - Rapporto complessità/resa troppo basso

### Nuova direzione
Sostituire la contextual toolbar con una **soluzione desktop più pulita, semplice e robusta**:

**Pulsante compatto: `Actions ▾`**

Questa sarà l'alternativa alle azioni contestuali "volanti".

### Comportamento desiderato

#### Stato senza selezione valida
- Il pulsante `Actions ▾` deve essere **disabilitato**

#### Stato con selezione valida
- Il pulsante `Actions ▾` deve essere **abilitato**

#### Apertura menu
Quando cliccato, il menu deve mostrare azioni contestuali appropriate:

**Nel tab Attive:**
- Delete
- Duplicate
- Archive

**Nel tab Archiviate:**
- Delete
- Duplicate
- Reactivate

---

## 2. Strategia proposta

### Approccio conservativo in 2 fasi

#### Fase 1 (Step corrente)
- Sostituire i 3 pulsanti operativi esistenti con `Actions ▾`
- Widget Tkinter nativo (Menubutton)
- Enable/disable dinamico tramite `update_button_visibility()` esistente
- Menu contestuale popolato dinamicamente in base al tab attivo
- **NO collegamento azioni** (stub methods per ora)

#### Fase 2 (Step successivo)
- Collegare le azioni esistenti al menu
- Test completo workflow

### Principi guida
- ✅ Niente toolbar contestuali dinamiche
- ✅ Niente overlay
- ✅ Niente layout shift
- ✅ Niente "tastoni" grandi e invasivi
- ✅ Pattern desktop classico e pulito
- ✅ Riuso totale della logica già esistente

---

## 3. Widget/menu consigliato per `Actions ▾`

### Soluzione: `ttk.Menubutton` + `tk.Menu`

Pattern desktop classico e robusto:

```python
# Pulsante con freccia dropdown
self.btn_actions = ttk.Menubutton(
    frame_top, 
    text="Actions ▾",
    state="disabled"
)

# Menu popup con azioni contestuali
self.actions_menu = tk.Menu(self.btn_actions, tearoff=0)
self.btn_actions.config(menu=self.actions_menu)
```

### Perché questo widget

| Caratteristica | Vantaggio |
|---------------|-----------|
| Nativo Tkinter/ttk | Zero dipendenze esterne |
| Stato `disabled` supportato | Enable/disable automatico |
| Menu dinamico | Popolabile/ripopolabile al volo |
| Pattern desktop familiare | UX intuitiva per utenti |
| Gestione automatica | Nessun posizionamento manuale |
| Robusto | Testato e stabile in Tkinter |

---

## 4. Moduli/file coinvolti

### File da modificare

#### 1. `dataflow.py` (linee ~3566-3620)
**Sezione:** `create_main_window()` - costruzione toolbar `frame_top`

**Modifiche:**
- Aggiungere `btn_actions` (Menubutton)
- Rimuovere/nascondere:
  - `btn_delete_rdo`
  - `btn_duplicate_rdo`
  - `btn_archive_rdo`
  - `btn_reactivate`

#### 2. `dataflow.py` (linee ~4342-4372)
**Sezione:** `update_button_visibility()`

**Modifiche:**
- Estendere per gestire enable/disable di `btn_actions`
- Aggiungere chiamata a `_populate_actions_menu()`
- Rimuovere logica vecchi pulsanti

#### 3. `dataflow.py` (nuovo metodo)
**Sezione:** Nuovo metodo helper `_populate_actions_menu()`

**Scopo:**
- Popolare dinamicamente il menu in base a:
  - Tab corrente (Attive/Archiviate)
  - Capacità utente (può eliminare, duplicare, etc.)

### File da NON toccare

| File | Motivo |
|------|--------|
| `database_manager.py` | Business logic intatta |
| `ui/components/main_dashboard_toolbar.py` | Global Search separata |
| `ui/components/collapsible_filters.py` | Filtri non modificati |
| Metodi esistenti | `delete_selected_request`, `duplicate_selected_request`, etc. |

---

## 5. Punto di aggancio alla logica di selezione esistente

### Riuso completo di `update_button_visibility()`

La logica esistente **già calcola tutto ciò che serve**:

```python
def update_button_visibility(self):
    # --- LOGICA GIÀ PRESENTE (da riusare) ---
    sheet, status = self.get_current_tree_and_status()
    selected_rows_indices = self._get_selected_row_indices(sheet)
    has_sel = bool(selected_rows_indices)
    num_selected = len(selected_rows_indices)
    all_mine = self._check_if_all_selected_are_mine(sheet, selected_rows_indices)
    
    can_delete = has_sel and all_mine
    can_duplicate = (num_selected == 1) and all_mine
    can_change_status = has_sel and all_mine
    
    # --- DA AGGIUNGERE ---
    # Abilita btn_actions se c'è almeno una selezione valida
    can_act = has_sel and all_mine
    self.btn_actions.config(state="normal" if can_act else "disabled")
    
    # Ripopola il menu in base al tab corrente
    self._populate_actions_menu(status, can_delete, can_duplicate, can_change_status)
```

### Eventi trigger già cablati

| Evento | Handler | Chiama |
|--------|---------|--------|
| Cambio selezione nel tree | (evento interno sheet) | `update_button_visibility()` |
| Cambio tab | `on_tab_changed()` | `update_button_visibility()` |
| Refresh data | `refresh_data()` | `update_button_visibility()` |

**Conclusione:** Zero modifiche alla logica di tracking necessarie.

---

## 6. Primo step minimo consigliato

### Step 1A: Aggiungere struttura UI del pulsante

**File:** `dataflow.py` (linee ~3566-3620)  
**Sezione:** Costruzione pulsanti in `frame_top`

```python
# 2-4. Rimossi (sostituiti da Actions dropdown)
# self.btn_delete_rdo = ...
# self.btn_duplicate_rdo = ...
# self.btn_archive_rdo = ...

# 2. Actions dropdown (sostituisce Delete/Duplicate/Archive/Reactivate)
self.btn_actions = ttk.Menubutton(
    frame_top,
    text="Actions ▾",
    state="disabled"
)
self.btn_actions.pack(side="left", padx=(0, 10))

# Menu popup per azioni contestuali
self.actions_menu = tk.Menu(self.btn_actions, tearoff=0)
self.btn_actions.config(menu=self.actions_menu)

# 6. Reactivate rimosso (ora dentro Actions menu)
# self.btn_reactivate = ...
```

### Step 1B: Aggiungere logica enable/disable

**File:** `dataflow.py` (linee ~4342-4372)  
**Metodo:** `update_button_visibility()`

```python
def update_button_visibility(self):
    """Aggiorna stato pulsante Actions in base a selezione"""
    sheet, status = self.get_current_tree_and_status()
    selected_rows_indices = self._get_selected_row_indices(sheet)
    has_sel = bool(selected_rows_indices)
    num_selected = len(selected_rows_indices)
    
    all_mine = self._check_if_all_selected_are_mine(sheet, selected_rows_indices) if has_sel else False
    
    # Abilita Actions solo se c'è selezione valida (tutte mie)
    can_act = has_sel and all_mine
    self.btn_actions.config(state="normal" if can_act else "disabled")
    
    # Popola menu in base a tab e capacità
    can_delete = can_act
    can_duplicate = (num_selected == 1) and all_mine
    can_change_status = can_act
    
    self._populate_actions_menu(status, can_delete, can_duplicate, can_change_status)
```

### Step 1C: Aggiungere metodo helper per popolare menu

**File:** `dataflow.py`  
**Posizione:** Nuovo metodo nella classe `DataFlowApp`

```python
def _populate_actions_menu(self, status, can_delete, can_duplicate, can_change_status):
    """Popola il menu Actions in base al tab corrente e capacità utente.
    
    Args:
        status: 'attiva' o 'archiviata'
        can_delete: bool, se può eliminare
        can_duplicate: bool, se può duplicare (1 sola selezione)
        can_change_status: bool, se può archiviare/riattivare
    """
    # Pulisci menu esistente
    self.actions_menu.delete(0, 'end')
    
    # Azioni comuni a entrambi i tab
    self.actions_menu.add_command(
        label=_("🗑 Elimina"),
        command=lambda: print("TODO: delete"),  # STUB per Step 1
        state="normal" if can_delete else "disabled"
    )
    
    self.actions_menu.add_command(
        label=_("🔁 Duplica"),
        command=lambda: print("TODO: duplicate"),  # STUB per Step 1
        state="normal" if can_duplicate else "disabled"
    )
    
    self.actions_menu.add_separator()
    
    # Azione specifica per tab
    if status == 'attiva':
        self.actions_menu.add_command(
            label=_("📦 Archivia"),
            command=lambda: print("TODO: archive"),  # STUB per Step 1
            state="normal" if can_change_status else "disabled"
        )
    else:  # archiviata
        self.actions_menu.add_command(
            label=_("↩️ Riattiva"),
            command=lambda: print("TODO: reactivate"),  # STUB per Step 1
            state="normal" if can_change_status else "disabled"
        )
```

---

## 7. Rischi da evitare

### ❌ Non fare

| Rischio | Perché evitarlo |
|---------|----------------|
| Duplicare logica selezione | Riusa `_get_selected_row_indices()` e `_check_if_all_selected_are_mine()` |
| Modificare metodi business logic | `delete_selected_request` etc. rimangono intatti |
| Usare `tk.OptionMenu` | Non supporta stato disabled per singole voci |
| Gestire eventi click manualmente | Usa `Menubutton` nativo, gestisce tutto Tkinter |
| Hardcodare stringhe | Usa sempre `_()` per i18n |
| Toccare la Global Search | Toolbar separata, zero interferenza |

### ✅ Principi da rispettare

| Principio | Applicazione |
|-----------|--------------|
| **Conservatività** | Minime modifiche, massimo riuso |
| **Robustezza** | Widget nativi, zero posizionamento manuale |
| **Semplicità** | 1 pulsante, 1 menu, logica lineare |
| **Testabilità** | Separazione UI (Step 1) da wiring (Step 2) |

---

## 8. Output atteso Step 1

### Funzionalità implementate

- ✅ Pulsante `Actions ▾` visibile in toolbar
- ✅ Menu dropdown con voci appropriate (Delete, Duplicate, Archive/Reactivate)
- ✅ Enable/disable automatico basato su selezione
- ✅ Voci menu enable/disable appropriate
  - Duplica solo con 1 selezione
  - Tutto disabilitato se RfQ altrui
- ✅ **Print stub** invece di azioni reali (per verifica UX)

### Checklist test Step 1

| Test | Risultato atteso |
|------|------------------|
| Senza selezione | Pulsante disabilitato |
| Con selezione valida | Pulsante abilitato |
| Tab Attive | Menu mostra "Archivia" |
| Tab Archiviate | Menu mostra "Riattiva" |
| Più di 1 riga selezionata | "Duplica" disabilitato |
| Selezione contiene RfQ altrui | Tutte le voci disabilitate |

---

## 9. Step successivi

### Step 2: Wiring azioni

**Obiettivo:** Collegare i command del menu ai metodi esistenti

```python
# Sostituire gli stub con i metodi reali
self.actions_menu.add_command(
    label=_("🗑 Elimina"),
    command=self.delete_selected_request,  # ← Metodo esistente
    state="normal" if can_delete else "disabled"
)

self.actions_menu.add_command(
    label=_("🔁 Duplica"),
    command=self.duplicate_selected_request,  # ← Metodo esistente
    state="normal" if can_duplicate else "disabled"
)

# etc...
```

### Step 3: Testing completo

- Test funzionali su tutte le azioni
- Test edge cases (selezioni multiple, RfQ altrui, etc.)
- Test i18n (traduzioni corrette)
- Test layout responsive

### Step 4: Cleanup

- Rimuovere completamente i vecchi pulsanti dal codice
- Aggiornare documentazione
- Commit e tag release

---

## 10. Vincoli rispettati

### Vincoli tecnici

| Vincolo | Status |
|---------|--------|
| NON introdurre nuove librerie | ✅ Solo Tkinter nativo |
| NON cambiare stack | ✅ Stessa tecnologia |
| NON modificare business logic | ✅ Riuso metodi esistenti |
| NON duplicare metodi esistenti | ✅ Zero duplicazione |
| NON toccare database_manager.py | ✅ Non modificato |
| NON rifare layout generale | ✅ Solo toolbar top |
| NON toccare search_requests() | ✅ Non toccato |
| NON riprendere contextual toolbar | ✅ Soluzione diversa |

---

## 11. Conclusioni

### Vantaggi della soluzione proposta

1. **Riduzione complessità UI**
   - Da 4 pulsanti separati → 1 pulsante compatto
   - Meno clutter visivo
   - Più spazio per altri elementi

2. **Pattern desktop familiare**
   - Dropdown menu = standard UX
   - Intuitivo per utenti desktop
   - Robusto e testato

3. **Massimo riuso codice**
   - Zero duplicazione logica
   - Metodi esistenti intatti
   - Tracking selezione invariato

4. **Implementazione conservativa**
   - Modifiche minime e localizzate
   - Testabile in modo incrementale
   - Facilmente reversibile se necessario

5. **Estendibilità futura**
   - Facile aggiungere nuove azioni al menu
   - Menu dinamico = massima flessibilità

### Sforzo stimato

| Step | Complessità | Tempo stimato |
|------|-------------|---------------|
| Step 1A (UI) | Bassa | 15 min |
| Step 1B (Enable/disable) | Bassa | 10 min |
| Step 1C (Populate menu) | Media | 20 min |
| Step 2 (Wiring) | Bassa | 10 min |
| Step 3 (Testing) | Media | 30 min |
| **Totale** | | **~1.5 ore** |

---

## 12. Prossimi passi

### Azione immediata richiesta

**Decisione:** Procedere con implementazione Step 1?

- [ ] **Sì** → Implementare Step 1A, 1B, 1C
- [ ] **No** → Chiarire/modificare proposta

### Supporto necessario

- Conferma approccio tecnico
- Eventuali requisiti UX aggiuntivi
- Priorità implementazione

---

**Fine documento**
