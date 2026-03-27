# Global Search — Analisi Pre-Implementazione VSM

**Data:** 27 marzo 2026  
**Obiettivo:** Estendere la Global Search Bar, attualmente dedicata al modulo RFQ, anche al modulo VSM, seguendo un approccio conservativo, modulare e reversibile.

---

## 1. File coinvolti nella Global Search attuale

| File | Ruolo |
|------|-------|
| `ui/components/main_dashboard_toolbar.py` | Widget search bar, placeholder, binding `<Return>` |
| `ui/components/collapsible_filters.py` | Pannello filtri avanzati collassabile |
| `dataflow.py` | Orchestrazione: variabili, costruzione query, routing risultati |
| `database_manager.py` | Layer SQL: query RFQ, aggregazione multi-DB |

---

## 2. Classi, funzioni e metodi per ogni fase

### 2.1 Input search bar

**`MainDashboardToolbar`** (`ui/components/main_dashboard_toolbar.py`)
- `_setup_ui()` — crea `self.search_entry` (tk.Entry, width=60, font 12pt)
- `_on_focus_in()` / `_on_focus_out()` — gestione placeholder grigio
- `_set_placeholder()` / `_clear_placeholder()` — testo placeholder: `"Search anything... (RFQ, Supplier, Code, Project...)"`
- Binding `<Return>` → `_on_search()`

### 2.2 Esecuzione ricerca

**`MainDashboardToolbar._on_search()`** (`ui/components/main_dashboard_toolbar.py` ~L145)
1. Se placeholder attivo → esce
2. Se input vuoto → chiama `main_window.clear_filters()`
3. Altrimenti: scrive `main_window.search_vars['global'].set(text)` → chiama `main_window.search_requests()`

**`MainWindow.search_requests()`** (`dataflow.py` L5467) — costruisce ed esegue query SQL

### 2.3 Costruzione risultati

**`search_requests()`** (`dataflow.py` L5518–5785):
- Base query: `SELECT DISTINCT` da `richieste_offerta` + LEFT JOIN su `dettagli_richiesta` e `richiesta_fornitori`
- Global search: blocco OR su 6 campi (`id_richiesta`, `riferimento`, `nome_fornitore`, `codice_materiale`, `descrizione_materiale`, `numeri_ordine`)
- Filtri avanzati: clausole AND su 10 campi + range date

**`DatabaseManager.search_richieste_advanced()`** (`database_manager.py` L1419) — ricerca locale parametrizzata

**`DatabaseManager.get_all_richieste_aggregated()`** (`database_manager.py` L1161) — UNION multi-DB via SQLite ATTACH

**`DatabaseManager.check_richiesta_detail_criteria()`** (`database_manager.py` L1505) — verifica criteri dettaglio su DB remoti

### 2.4 Visualizzazione risultati

- **`MainWindow.update_treeview(tree, results)`** (`dataflow.py`) — popola il widget `ttk.Treeview` (tab RFQ)
- **`_populate_vsm_sheet(sheet, events)`** (`dataflow.py` L4493) — popola widget `tksheet.Sheet` (tab VSM) — **separato, non toccato da `search_requests`**

### 2.5 Navigazione / click sul risultato

- **RFQ:** doppio click su Treeview → apertura dettaglio RFQ
- **VSM:** doppio click su Sheet → `_on_vsm_sheet_double_click(sheet, event)` (`dataflow.py` L4445) → apertura dialog edit VSM

---

## 3. La logica attuale è hardcoded sul modulo RFQ?

**Sì, in un punto critico e in modo esplicito.**

`dataflow.py`, L5471:

```python
def search_requests(self):
    tree, status = self.get_current_tree_and_status()

    # Skip search for VSM tabs (not yet implemented)
    if tree is None or status.startswith('vsm_'):
        return          # ← EXIT IMMEDIATO PER TUTTI I TAB VSM
```

Tutto il corpo di `search_requests()` (circa 320 righe) è SQL specifico per lo schema RFQ (`richieste_offerta`, `dettagli_richiesta`, `richiesta_fornitori`). Non esiste nessuna branch VSM.

---

## 4. Dipendenze e assunzioni specifiche di RFQ

| Elemento | Dove | Natura |
|----------|------|--------|
| Tabelle SQL `richieste_offerta`, `dettagli_richiesta`, `richiesta_fornitori` | `search_requests()`, `search_richieste_advanced()` | Hardcoded nel SQL |
| `search_vars` dict con chiavi `num`, `ref`, `forn`, `cod`, `desc`, `ord`, `cod_grezzo`, `dis_grezzo`, `mat_cl` | `dataflow.py` L3668 | Nomi di campi RFQ |
| `search_tipo` (ComboBox Fornitura piena / Conto lavoro) | `dataflow.py` L3668 | Specifico RFQ |
| `date_entries` per `data_emissione` / `data_scadenza` | `dataflow.py` | Campi RFQ |
| `update_treeview()` — output su `ttk.Treeview` | `dataflow.py` | Widget RFQ |
| Logica aggregazione multi-DB (`get_all_richieste_aggregated`) | `database_manager.py` | Architettura RFQ |
| Placeholder text: `"Search anything... (RFQ, Supplier, Code, Project...)"` | `main_dashboard_toolbar.py` L20 | Testo specifico RFQ |

---

## 5. Formato dei risultati della ricerca oggi

**RFQ (ricerca locale):**
```
tuple(id_richiesta, tipo_rdo, data_emissione, data_scadenza, riferimento, username)
```

**RFQ (aggregati multi-DB):**
```
tuple(id, tipo_rdo, data_emissione, data_scadenza, riferimento, username, stato, is_mine, source_file)
```

**VSM (da `get_all_vsm_events()`):**  
Lista di oggetti `VSMEvent` (dataclass), non tuple — con 23 campi tra cui: `id`, `event_date`, `buyer`, `event_type`, `action`, `description`, `reference`, `importo_bdg`, `importo_negoziato`, `driver`, `new_supplier`, `note`, `username`.

---

## 6. La ricerca legge da DB, da widget UI o da strutture intermedie?

- **RFQ locale:** legge direttamente dal DB SQLite tramite query parametrizzata
- **RFQ aggregato:** carica tutto in memoria via `get_all_richieste_aggregated()`, poi filtra in-memory con check Python + query di dettaglio su DB remoti
- **VSM (caricamento normale):** `_load_vsm_events()` → `get_all_vsm_events()` → lista `VSMEvent` in memoria → `_populate_vsm_sheet()`
- **VSM (search attuale):** non implementato — la ricerca non tocca né il DB né le strutture in memoria

---

## 7. Punto migliore per introdurre un layer astratto

L'architettura esistente suggerisce tre opzioni, in ordine crescente di invasività:

### Opzione A — Intervento solo in `search_requests()` *(minimo invasivo, raccomandato)*

Sostituire il `return` a L5471 con una branch che chiama un metodo dedicato `_search_vsm_events(tree, status)`. Nessun altro file toccato.

```
search_requests()
    ├── if vsm_* → _search_vsm_events(tree, status)   ← NEW
    └── else     → logica RFQ esistente (invariata)
```

### Opzione B — SearchProvider pattern *(modulare, più file)*

Introduce un'astrazione `BaseSearchProvider` con sottoclassi `RFQSearchProvider` e `VSMSearchProvider`. Richiede refactor di `search_requests()` e nuovi file. Più pulita ma maggior rischio di regressione su RFQ.

### Opzione C — Metodo separato `search_vsm_events()` in `dataflow.py`

Un metodo autonomo parallelo a `search_requests()`, chiamato nella stessa toolbar. Nessun tocco alla logica RFQ esistente.

**Raccomandazione:** Opzione A come primo step, eventualmente Opzione C se il metodo diventa complesso.

---

## 8. Modifiche minime e sicure necessarie

### Per mantenere RFQ invariato

Nessuna modifica ai metodi esistenti.  
L'unica modifica a `search_requests()` è sostituire:

```python
if tree is None or status.startswith('vsm_'):
    return
```

con:

```python
if tree is None:
    return
if status.startswith('vsm_'):
    self._search_vsm_events(tree, status)
    return
```

### Per aggiungere VSM (step successivo)

1. **Nuovo metodo `_search_vsm_events(sheet, status)`** in `dataflow.py`:
   - Legge `search_vars['global'].get()` (già disponibile)
   - Carica `get_all_vsm_events()` (già esiste in `database_manager.py`)
   - Filtra in-memory sugli `VSMEvent` (campi candidati: `description`, `reference`, `buyer`, `driver`, `action`, `new_supplier`, `note`)
   - Chiama `_populate_vsm_sheet(sheet, filtered_events)`

2. **Nessuna modifica al toolbar** (`main_dashboard_toolbar.py`) — già generico, già funziona su tutti i tab

3. **Nessuna modifica al DB layer** (`database_manager.py`) — `get_all_vsm_events()` è già sufficiente

4. **Aggiornamento testo placeholder** (opzionale, basso rischio):  
   `"Search anything... (RFQ, Supplier, Code, Project...)"` → aggiungere es. `"... VSM, Buyer..."` o gestirlo dinamicamente in base al tab attivo

---

## 9. Rischi di regressione

| Rischio | Probabilità | Mitigazione |
|---------|-------------|-------------|
| Modifica a `search_requests()` rompe RFQ | Media | Toccare solo il blocco early-return, nessun tocco al corpo esistente |
| `clear_filters()` chiama `refresh_data()` che ricarica tutto — in tab VSM potrebbe non ricaricare gli eventi VSM | Bassa | Verificare se `refresh_data()` ha già un path per VSM |
| `_has_active_search_filters()` non considera il tab corrente, potrebbe triggerare ricerca RFQ anche da tab VSM | Media | Già protetto dall'early return in `search_requests()` — il fix all'early return mantiene la protezione |
| Placeholder non aggiornato crea confusione UX per utenti VSM | Bassa | Cosmetic, non funzionale |
| `_populate_vsm_sheet()` chiamata con dati filtrati invece che completi potrebbe rompere calcoli aggregati (KPI, totali) | **Alta** | **Verificare se `_populate_vsm_sheet` calcola aggregati dai dati passati o da una sorgente separata prima di implementare** |
| Widget VSM usa `tksheet.Sheet` (non `Treeview`) — firma di `update_treeview()` incompatibile | Nessuno | `_populate_vsm_sheet()` esiste già e ha la firma giusta |
| Tab VSM non ha searchvar specifiche (tipo evento, buyer, range data) | Nessuno (fase 1) | Filtro solo sul campo `global`, nessuna UI aggiuntiva richiesta in step 1 |

---

## 10. Elenco file da toccare in un futuro step

| File | Modifica necessaria | Invasività |
|------|---------------------|------------|
| `dataflow.py` | Sostituire early-return in `search_requests()` L5471; aggiungere `_search_vsm_events()` | Bassa — solo aggiunta di codice |
| `ui/components/main_dashboard_toolbar.py` | Aggiornamento testo placeholder (opzionale) | Bassissima |
| `database_manager.py` | Nessuna — `get_all_vsm_events()` è già sufficiente | Nessuna |
| `ui/components/collapsible_filters.py` | Nessuna | Nessuna |

---

## 11. Proposta di architettura minima (step 1)

```
MainDashboardToolbar._on_search()
    │  (già generico, nessuna modifica)
    ▼
search_vars['global'].set(text)
    │
    ▼
MainWindow.search_requests()
    ├── if status.startswith('vsm_')
    │       │
    │       ▼
    │   _search_vsm_events(sheet, status)        ← NEW (solo aggiunta in dataflow.py)
    │       │
    │       ├─ get_all_vsm_events(username)      ← già esiste in database_manager.py
    │       ├─ filter in-memory su VSMEvent      ← nuovo loop Python ~15 righe
    │       └─ _populate_vsm_sheet(sheet, filtered_events)   ← già esiste
    │
    └── else (RFQ)
            └── logica attuale INVARIATA
```

**Totale righe di codice nuovo stimato:** ~40–50 righe, tutte in `dataflow.py`.  
**Nessuna modifica a codice esistente** — solo aggiunta.

---

## Appendice — Struttura dati chiave

### `search_vars` dictionary (RFQ)

```python
self.search_vars = {
    'global':     tk.StringVar(),  # Ricerca globale su 6 campi (unico campo riutilizzabile per VSM)
    'num':        tk.StringVar(),  # Numero RdO (id_richiesta)
    'ref':        tk.StringVar(),  # Riferimento progetto
    'forn':       tk.StringVar(),  # Fornitore (nome_fornitore)
    'cod':        tk.StringVar(),  # Codice materiale
    'desc':       tk.StringVar(),  # Descrizione materiale
    'ord':        tk.StringVar(),  # Numero ordine
    'cod_grezzo': tk.StringVar(),  # Codice grezzo
    'dis_grezzo': tk.StringVar(),  # Disegno grezzo
    'mat_cl':     tk.StringVar(),  # Materiale conto lavoro
}
```

L'unica chiave utile per VSM in step 1 è `'global'`.

### Schema tabella `vsm_events`

```sql
CREATE TABLE vsm_events (
    event_id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    username                    TEXT NOT NULL,
    event_date                  TEXT,
    buyer                       TEXT,
    event_type                  TEXT,   -- Saving / Cost Avoidance / Derisking
    action                      TEXT,   -- Negoziazione / Derisking
    description                 TEXT,
    reference                   TEXT,
    importo_bdg                 REAL,
    importo_negoziato           REAL,
    importo_richiesto_iniziale  REAL,
    quantita_annua              REAL,
    percent_realizzo            REAL,
    driver                      TEXT,   -- Prezzo / Pagamenti
    giorni_pagamento_attuali    INTEGER,
    giorni_pagamento_negoziati  INTEGER,
    spending_annuo              REAL,
    opex_ripetitivo             INTEGER NOT NULL DEFAULT 0,
    note                        TEXT,
    created_at                  TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at                  TEXT DEFAULT CURRENT_TIMESTAMP,
    payments_rate               REAL,
    new_supplier                TEXT DEFAULT ''
)
```

**Campi ricercabili via global search in step 1:**
`description`, `reference`, `buyer`, `driver`, `action`, `event_type`, `new_supplier`, `note`
