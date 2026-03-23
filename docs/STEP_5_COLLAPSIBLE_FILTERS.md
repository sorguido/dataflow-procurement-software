# Step 5: Filtri Collassabili - Implementazione Finale

## Problema Risolto

I filtri collassabili occupavano spazio verticale anche quando nascosti, lasciando un gap vuoto tra la toolbar e il notebook.

## Modifiche Implementate

### 1. `ui/components/collapsible_filters.py` - Compattamento automatico

**Aggiunto in `__init__`:**
```python
# Assicura che il wrapper si ridimensioni in base al contenuto
# Quando filters_frame è nascosto, il wrapper si compatta a zero
self.pack_propagate(True)
```

**Modificato `expand()`:**
```python
def expand(self):
    """Mostra i filtri di ricerca.
    
    Il filters_frame è già figlio del wrapper, quindi un semplice pack()
    lo rende visibile senza problemi di posizionamento.
    Aggiunge pady per spacing quando espanso.
    """
    if not self._is_expanded:
        self._is_expanded = True
        # Padding verticale solo quando espanso
        self.filters_frame.pack(fill="x", padx=0, pady=(0, 5))
```

**Invariato `collapse()`:**
```python
def collapse(self):
    """Nasconde i filtri di ricerca.
    
    Usa pack_forget() sul filters_frame (che è figlio diretto).
    Il wrapper rimane sempre packed nella stessa posizione.
    """
    if self._is_expanded:
        self._is_expanded = False
        self.filters_frame.pack_forget()
```

### 2. `dataflow.py` - Rimosso padding dal wrapper

**Prima:**
```python
self.collapsible_filters.pack(fill="x", padx=10, pady=(0, 5))
```

**Ora:**
```python
# Pack senza pady: quando collapsed, lo spazio si compatta completamente
self.collapsible_filters.pack(fill="x", padx=10)
```

## Spiegazione Tecnica

### Problema Precedente
- Il wrapper aveva `pady=(0, 5)` → 5px di padding bottom **sempre presente**
- Anche quando `filters_frame` era nascosto, quel padding rimaneva
- Risultato: spazio vuoto quando collapsed

### Soluzione Implementata

**Principio chiave: `pack_propagate(True)`**
- Tkinter ridimensiona automaticamente un Frame in base al contenuto
- Quando `filters_frame` è nascosto (pack_forget), il wrapper **non ha più figli visibili**
- Con `pack_propagate(True)`, il wrapper si **compatta a zero altezza**

**Gestione padding:**
- Wrapper: `pack(fill="x", padx=10)` → **nessun pady**
- filters_frame: `pack(fill="x", pady=(0, 5))` → **pady solo quando visible**

### Risultato

**Collapsed:**
- `filters_frame` nascosto → wrapper height=0px → zero spazio occupato
- Il notebook si espande immediatamente sotto la toolbar

**Expanded:**
- `filters_frame` visibile con `pady=(0, 5)` → spacing corretto
- Il wrapper si ridimensiona automaticamente per contenere i filtri

## Layout Finale

### Stato Collapsed
```
┌─────────────────────────────────────┐
│  Global Search Toolbar              │
├─────────────────────────────────────┤
│ [CollapsibleFilters: height=0px]    │ ← Nessuno spazio
├─────────────────────────────────────┤
│  Notebook (espanso)                 │
│                                     │
```

### Stato Expanded
```
┌─────────────────────────────────────┐
│  Global Search Toolbar              │
├─────────────────────────────────────┤
│ ┌─────────────────────────────────┐ │
│ │ 🔍 Filtri di Ricerca            │ │
│ │ [tutti i widget dei filtri]     │ │
│ └─────────────────────────────────┘ │
│           ↓ pady=(0, 5)             │
├─────────────────────────────────────┤
│  Notebook                           │
│                                     │
```

## Gerarchia Widget

```
root
├── MainDashboardToolbar (sempre visible)
├── CollapsibleFilters (wrapper sempre packed)
│   └── filters_frame (toggle show/hide)
│       └── tutti i widget dei filtri
└── Notebook (sempre visible)
```

## Comportamento Toggle

1. Click su "⌄ Advanced Filters" nella toolbar
2. `MainWindow.toggle_filters()` chiamato
3. `CollapsibleFilters.toggle()` eseguito
4. Se collapsed → `expand()`:
   - `filters_frame.pack(fill="x", pady=(0, 5))`
   - Wrapper si espande automaticamente (pack_propagate)
   - Chevron cambia a "⌃"
5. Se expanded → `collapse()`:
   - `filters_frame.pack_forget()`
   - Wrapper si compatta a zero altezza
   - Chevron cambia a "⌄"

## Vincoli Rispettati

✅ NON modificata la struttura generale del layout  
✅ NON modificata la logica dei filtri  
✅ NON modificato `search_requests()`  
✅ NON usati hack di posizionamento (`pack(before=...)`)  
✅ Wrapper mai rimosso dal layout (posizione stabile)  
✅ True Container Pattern: reparenting reale  

## Testing

Per verificare il corretto funzionamento:

1. Avvia l'applicazione: `python dataflow.py`
2. Clicca su "⌄ Advanced Filters"
3. Verifica che i filtri si espandano con spacing corretto
4. Clicca di nuovo su "⌃ Advanced Filters"
5. Verifica che lo spazio si compatti completamente
6. Il notebook deve espandersi immediatamente senza gap

---

**Data implementazione**: 21 marzo 2026  
**Versione**: DataFlow v2.1.0 - Main Dashboard Redesign  
**Step**: 5 (Collapsible Filters)
