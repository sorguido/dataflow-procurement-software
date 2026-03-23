# Conversione Layout Dashboard da Pack() a Grid()

## Problema: Pack() lascia gap verticali

### Causa Tecnica

**Pack() è un geometry manager sequenziale:**
- Impila i widget in ordine (top, bottom, left, right)
- Calcola lo spazio basandosi sulla sequenza di tutti i widget
- Anche con `pack_forget()` e `pack_propagate(True)`, può lasciare margini residui
- Rimuovere/aggiungere dinamicamente widget crea inconsistenze geometriche

**Il problema specifico con i filtri collassabili:**
1. `CollapsibleFilters` wrapper viene packed
2. `filters_frame` interno viene packed/unpacked
3. Quando `filters_frame` viene rimosso, il wrapper rimane con altezza minima
4. Pack() continua a riservare spazio per padding e margini anche del widget rimosso
5. Risultato: **gap verticale** tra toolbar e notebook quando collapsed

## Soluzione: Grid Layout

### Perché Grid() è più robusto

**Grid() mantiene una struttura a griglia stabile:**
- Non è sequenziale, ma basato su righe e colonne
- `grid_remove()` rimuove completamente il widget dalla griglia **senza lasciare gap**
- Mantiene i parametri grid (row, column, sticky, etc.) per ripristino facile
- Più prevedibile per show/hide dinamici
- Le altre righe si compattano automaticamente quando una riga viene rimossa

**Vantaggi specifici:**
- ✅ Zero gap quando widget rimosso
- ✅ Ripristino preciso con `grid()` senza parametri
- ✅ Layout responsive automatico
- ✅ Configurazione esplicita di row/column weight

## Modifiche Implementate

### 1. `ui/components/collapsible_filters.py`

**Strategia: Grid-based show/hide**
- Il wrapper stesso viene rimosso/ripristinato dalla griglia
- `filters_frame` rimane sempre packed **dentro** il wrapper
- `expand()` → `self.grid()` con parametri salvati
- `collapse()` → `self.grid_remove()`

**Codice aggiornato:**

```python
class CollapsibleFilters(ttk.Frame):
    def __init__(self, parent, label_text="Search Filters"):
        super().__init__(parent)
        
        self._is_expanded = False
        self._grid_config = {}  # Parametri grid salvati
        
        # Pack filters_frame UNA volta sola dentro il wrapper
        # Rimane sempre visible internamente
        self.filters_frame = ttk.LabelFrame(self, text=f"🔍 {label_text}", padding=(10, 5))
        self.filters_frame.pack(fill="x", padx=0, pady=0)
    
    def set_grid_config(self, **kwargs):
        """Salva configurazione grid per ripristino dopo grid_remove()."""
        self._grid_config = kwargs
    
    def expand(self):
        """Mostra i filtri: ripristina il wrapper nella griglia."""
        if not self._is_expanded:
            self._is_expanded = True
            if self._grid_config:
                self.grid(**self._grid_config)  # Usa parametri salvati
    
    def collapse(self):
        """Nasconde i filtri: rimuove il wrapper dalla griglia."""
        if self._is_expanded:
            self._is_expanded = False
            self.grid_remove()  # Rimuove senza lasciare gap
```

**Differenza chiave con pack():**
- **Prima (pack):** `filters_frame.pack_forget()` → wrapper rimane, gap visibile
- **Ora (grid):** `self.grid_remove()` → wrapper rimosso, zero gap

### 2. `dataflow.py` - Layout principale

**Struttura grid:**

| Row | Widget | Weight | Sticky | Descrizione |
|-----|--------|--------|--------|-------------|
| 0 | `frame_top` | 0 | ew | Toolbar pulsanti |
| 1 | `main_dashboard_toolbar` | 0 | ew | Global Search |
| 2 | `collapsible_filters` | 0 | ew | Filtri (dinamico) |
| 3 | `notebook` | 1 | nsew | Tabelle RFQ |

**Configurazione root:**
```python
self.root.grid_rowconfigure(3, weight=1)  # Notebook si espande
self.root.grid_columnconfigure(0, weight=1)  # Colonna si espande
```

**Codice modificato:**

```python
# ===== CONVERSIONE LAYOUT PRINCIPALE A GRID() =====

# Configurazione griglia
self.root.grid_rowconfigure(3, weight=1)  # Row notebook espandibile
self.root.grid_columnconfigure(0, weight=1)  # Colonna espandibile

# Row 0: Toolbar pulsanti
frame_top.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

# Row 1: Global Search Toolbar
self.main_dashboard_toolbar = MainDashboardToolbar(self.root, self)
self.main_dashboard_toolbar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

# Row 2: Filtri collassabili
self.collapsible_filters = CollapsibleFilters(self.root, label_text=_("Filtri di Ricerca"))
self.collapsible_filters.set_grid_config(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))
self.collapsible_filters.collapse()  # Default nascosto

# Row 3: Notebook
self.notebook = ttk.Notebook(self.root)
self.notebook.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
```

**Footer (nota):**
Il footer usa `pack(side="bottom")` ed è indipendente dalla griglia principale - lasciato invariato.

## Comportamento Finale

### Collapsed (default)
```
┌─────────────────────────────────┐
│ Row 0: Frame Top                │
├─────────────────────────────────┤
│ Row 1: Global Search Toolbar    │
├─────────────────────────────────┤  ← Row 2 RIMOSSA (grid_remove)
│ Row 3: Notebook (weight=1)      │  ← Si espande fino in fondo
│                                 │
│                                 │
└─────────────────────────────────┘
```

### Expanded
```
┌─────────────────────────────────┐
│ Row 0: Frame Top                │
├─────────────────────────────────┤
│ Row 1: Global Search Toolbar    │
├─────────────────────────────────┤
│ Row 2: Collapsible Filters      │  ← Ripristinato con grid()
│        🔍 Filtri di Ricerca     │
│        [widget dei filtri]      │
├─────────────────────────────────┤
│ Row 3: Notebook (weight=1)      │  ← Si compatta
│                                 │
└─────────────────────────────────┘
```

## Toggle Flow

**Click su "⌄ Advanced Filters":**
1. `MainDashboardToolbar._on_toggle_filters()` chiamato
2. `MainWindow.toggle_filters()` eseguito
3. `CollapsibleFilters.toggle()` invocato
4. Se collapsed:
   - `expand()` → `self.grid(row=2, column=0, sticky="ew", ...)`
   - Grid inserisce automaticamente il wrapper in row 2
   - Row 3 (notebook) si compatta verso il basso
   - Chevron cambia: ⌄ → ⌃
5. Se expanded:
   - `collapse()` → `self.grid_remove()`
   - Grid rimuove il wrapper dalla row 2 **senza lasciare gap**
   - Row 3 si espande verso l'alto riempiendo lo spazio
   - Chevron cambia: ⌃ → ⌄

## Confronto: Pack vs Grid

| Aspetto | Pack() (prima) | Grid() (ora) |
|---------|----------------|--------------|
| **Tipo** | Sequenziale | A griglia |
| **Show/hide** | pack_forget() sul contenuto | grid_remove() sul wrapper |
| **Gap residuo** | ❌ Sì, margini/padding rimangono | ✅ No, rimozione completa |
| **Ripristino** | pack() con tutti i parametri | grid() senza parametri |
| **Robustezza** | ❌ Instabile per dinamici | ✅ Progettato per dinamici |
| **Prevedibilità** | ❌ Calcolo sequenziale complesso | ✅ Struttura esplicita |

## Vincoli Rispettati

✅ **NON modificati:**
- Widget interni dei filtri (Entry, ComboBox, DateEntry, etc.)
- Logica di ricerca (`search_requests()`)
- Database manager
- Business logic
- Tabelle RFQ (usano ancora pack/grid internamente)
- Footer (usa pack side="bottom")

✅ **Modificato SOLO:**
- Layout verticale principale del container dashboard
- Metodi expand()/collapse() di CollapsibleFilters
- Posizionamento di 4 widget principali (frame_top, toolbar, filters, notebook)

## Testing

Per verificare la risoluzione:

1. Avvia l'applicazione: `python dataflow.py`
2. I filtri sono nascosti di default ✓
3. Clicca su "⌄ Advanced Filters"
4. **Verifica:** Filtri si espandono, nessun gap sopra/sotto ✓
5. **Verifica:** Notebook si compatta verso il basso ✓
6. Clicca su "⌃ Advanced Filters"
7. **Verifica:** Filtri scompaiono completamente ✓
8. **Verifica:** Notebook si espande immediatamente **SENZA GAP** ✓
9. Ripeti toggle più volte: nessun artefatto visivo ✓

---

**Data implementazione:** 21 marzo 2026  
**Versione:** DataFlow v2.1.0 - Main Dashboard Redesign  
**Step:** 5 (Collapsible Filters) - Fix definitivo con Grid Layout
