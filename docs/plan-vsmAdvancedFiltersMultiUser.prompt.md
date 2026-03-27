# Piano Implementazione — Advanced Filters VSM con scope utente multi-user

**Data:** 27 marzo 2026
**Obiettivo:** Estendere il modulo VSM con filtro utente coerente con RFQ, aggregazione multi-DB, metadata read-only per riga.

---

## Esito Fase 0

| Componente | Stato | Note |
|-----------|-------|------|
| `_event_metadata` con `is_mine` per riga | ✅ GIÀ IMPLEMENTATO | Popolato in `_populate_vsm_sheet()` |
| Validazione ownership in CRUD (edit, delete, duplicate, double-click) | ✅ GIÀ IMPLEMENTATO | Tutti i 4 handler controllano `is_mine` |
| `get_all_vsm_events(username)` | ✅ GIÀ IMPLEMENTATO | Solo DB locale |
| `get_all_vsm_events_aggregated()` | ❌ MANCANTE | Da creare in `database_manager.py` |
| Filtro utente UI nei tab VSM | ❌ MANCANTE | Da aggiungere |
| `read_only` mode in `VSMEventDialog` | ❌ MANCANTE | Da aggiungere |
| Caricamento VSM rispetta scope utente | ❌ PARZIALE | `_load_vsm_events()` hardcoded su `current_username` |
| Global Search VSM rispetta scope utente | ❌ PARZIALE | `_search_vsm_events()` usa sempre `current_username` |

---

## File da modificare

| File | Modifica |
|------|---------|
| `database_manager.py` | Aggiungere `get_all_vsm_events_aggregated()` |
| `dataflow.py` | Aggiungere `vsm_username_filter_var`, `populate_vsm_username_filter()`, `_get_active_vsm_username_filter()`, aggiornare `_load_vsm_events()`, `_search_vsm_events()`, `_populate_vsm_sheet()`, `clear_filters()` |
| `ui/dialogs/vsm_event_dialog.py` | Aggiungere parametro `read_only=False` |

---

## Step 1 — `database_manager.py`: `get_all_vsm_events_aggregated()`

**Posizione:** subito dopo `get_all_vsm_events()` (L2223).

**Firma:**
```python
def get_all_vsm_events_aggregated(self, my_db_full_path: str, username: str = None) -> list:
```

**Ritorna:** `list of (VSMEvent, is_mine: bool, source_file: str)`

**Pattern:** identico a `get_all_richieste_aggregated()`:
1. Carica DB locale con `get_all_vsm_events(username=username)` → `is_mine=True`, `source_file='local'`
2. `glob("**/dataflow_db_*.db")` dalla root condivisa
3. Per ogni DB sibling: apertura con `sqlite3.connect("file:...?mode=ro", uri=True)` (sola lettura, no migration)
4. `PRAGMA table_info(vsm_events)` per gestire schema migration (colonne `payments_rate`, `new_supplier` opzionali)
5. Query con `WHERE username = ?` se username filtro attivo
6. Append `(VSMEvent, False, source_file)` per ogni riga esterna

**Requisito critico:** connessione read-only tramite URI SQLite (`?mode=ro`) — nessuna scrittura, nessuna migrazione schema sui DB altrui.

---

## Step 2 — `dataflow.py`: variabili filtro utente VSM

**Posizione:** in `__init__` / zone di inizializzazione, vicino a `self.username_filter_var = None`.

Aggiungere:
```python
self.vsm_username_filter_var = None
self.vsm_user_filter_combo = None
```

---

## Step 3 — `dataflow.py`: UI filtro utente nei tab VSM

**Posizione:** nel blocco di creazione tab VSM (attorno a L3717), dopo la creazione delle sheet, aggiungere una toolbar leggera (Frame) sopra ciascuna sheet — oppure un singolo controllo condiviso visibile solo quando si è su tab VSM.

**Approccio più conservativo (raccomandato):** singola ComboBox condivisa tra tutti i tab VSM, mostrata/nascosta al cambio tab tramite `on_tab_changed()`.

**Struttura UI:**
```
vsm_filter_frame (ttk.Frame, row=... sopra notebook oppure integrato nella toolbar)
  │
  ├── ttk.Label("Utente:")
  └── ttk.Combobox (textvariable=vsm_username_filter_var, state="readonly")
        → binding <<ComboboxSelected>> → _on_vsm_username_filter_changed()
```

**Nota:** non modificare `collapsible_filters.py` — è già usato solo per RFQ. Il filtro VSM è più semplice e può essere un Frame dedicato.

---

## Step 4 — `dataflow.py`: helper scope utente VSM

```python
def _get_active_vsm_username_filter(self):
    """Analogo a _get_active_username_filter() per il modulo VSM."""
    if not self.vsm_username_filter_var:
        return None
    value = self.vsm_username_filter_var.get().strip()
    if not value or value == self.all_users_placeholder:
        return None
    return value.lower()

def populate_vsm_username_filter(self):
    """Popola la ComboBox utenti VSM da tutti i DB aggregati."""
    # Carica tutti gli eventi aggregati (senza filtro username)
    # Estrae username unici
    # Aggiunge all_users_placeholder come prima voce
    # Mantiene il valore corrente se ancora valido
```

---

## Step 5 — `dataflow.py`: aggiornare `_load_vsm_events()`

**Attuale:**
```python
all_events = db_manager.get_all_vsm_events(username=self.current_username)
```

**Nuovo:**
```python
vsm_username_filter = self._get_active_vsm_username_filter()

if vsm_username_filter is None:
    # "Tutti gli utenti": usa aggregazione multi-DB
    all_data = db_manager.get_all_vsm_events_aggregated(get_db_path(), username=None)
    all_events = [ev for ev, _is_mine, _src in all_data]
    # Passa anche is_mine e source_file a _populate_vsm_sheet via lista enriched
elif vsm_username_filter == self.current_username.lower():
    # Ottimizzazione locale: cerca solo nel DB locale
    all_events = db_manager.get_all_vsm_events(username=self.current_username)
else:
    # Altro utente: usa aggregazione multi-DB con filtro username
    all_data = db_manager.get_all_vsm_events_aggregated(get_db_path(), username=vsm_username_filter)
    all_events = [ev for ev, _is_mine, _src in all_data]
```

**Problema:** `_populate_vsm_sheet()` riceve attualmente una lista di `VSMEvent`, non tuple con metadata. Per supportare `is_mine` da DB aggregati bisogna adeguare il passaggio dei metadata.

**Soluzione conservativa:** passare un dizionario opzionale `metadata_override` oppure passare una lista di tuple `(VSMEvent, is_mine, source_file)` a `_populate_vsm_sheet()` con retrocompatibilità.

**Approccio più pulito:** `_populate_vsm_sheet()` accetta `extra_metadata: list[dict] | None = None`. Se fornito, usa quei metadata invece di calcolare `is_mine` da `event.username == self.current_username`.

---

## Step 6 — `dataflow.py`: aggiornare `_populate_vsm_sheet()`

**Modifica alla riga che calcola `is_mine`:**

```python
# Attuale:
metadata.append({
    'event_id': event.id,
    'username': event.username,
    'is_mine': event.username == self.current_username
})

# Nuovo (con extra_metadata):
if extra_metadata and i < len(extra_metadata):
    is_mine = extra_metadata[i].get('is_mine', event.username == self.current_username)
    source_file = extra_metadata[i].get('source_file', 'local')
else:
    is_mine = event.username == self.current_username
    source_file = 'local'

metadata.append({
    'event_id': event.id,
    'username': event.username,
    'is_mine': is_mine,
    'source_file': source_file,
})
```

La firma diventa:
```python
def _populate_vsm_sheet(self, sheet, events, event_type=None, extra_metadata=None):
```

---

## Step 7 — `dataflow.py`: aggiornare `_search_vsm_events()`

**Attuale:**
```python
all_events = db_manager.get_all_vsm_events(username=self.current_username)
```

**Nuovo:** rispetta lo scope utente VSM attivo, identico al pattern di `_load_vsm_events()`.

```python
vsm_username_filter = self._get_active_vsm_username_filter()

if vsm_username_filter is None or vsm_username_filter != self.current_username.lower():
    # Tutti gli utenti o altro utente: aggregazione
    raw = db_manager.get_all_vsm_events_aggregated(get_db_path(), username=vsm_username_filter or None)
    all_events = [ev for ev, _is_mine, _src in raw]
    extra_meta = [{'is_mine': im, 'source_file': src} for _, im, src in raw]
else:
    # Utente corrente: locale
    all_events = db_manager.get_all_vsm_events(username=self.current_username)
    extra_meta = None
```

---

## Step 8 — `ui/dialogs/vsm_event_dialog.py`: modalità read-only

**Aggiungere parametro `read_only=False` a `__init__`:**

```python
def __init__(self, parent, current_username, event_type="Saving", event_id=None, read_only=False):
```

**Se `read_only=True`:**
- Titolo: `_("Visualizza Evento VSM")`
- Tutti i widget Entry/Combobox/DateEntry → `state="disabled"`
- Pulsante "💾 Salva" nascosto o disabilitato
- Pulsante "❌ Annulla" → diventa "✖ Chiudi"
- `_validate_and_save()` non raggiungibile

**Uso in `_edit_vsm_event()`:**
```python
if not is_mine:
    # Apri in read-only invece di mostrare errore
    dialog = VSMEventDialog(
        self.root,
        current_username=self.current_username,
        event_type=event_type,
        event_id=event_id,
        read_only=True
    )
    self.root.wait_window(dialog)
    return
```

---

## Step 9 — `dataflow.py`: aggiornare `on_tab_changed()` e `clear_filters()`

**`on_tab_changed()`:** mostrare/nascondere il frame filtro VSM in base al tab attivo.

**`clear_filters()`:** aggiungere reset del filtro VSM utente (solo se il tab corrente è VSM):
```python
if self.vsm_username_filter_var:
    self.vsm_username_filter_var.set(self.current_username or self.all_users_placeholder)
```

**`populate_vsm_username_filter()`:** chiamato dopo ogni caricamento dati VSM (come `populate_username_filter()` per RFQ).

---

## Flusso finale post-implementazione

```
[Tab VSM attivo]
    │
    ├── Filtro utente VSM (ComboBox)
    │       │
    │       ├── "Tutti gli utenti" → aggregazione multi-DB (is_mine vario)
    │       ├── utente corrente   → DB locale (is_mine=True per tutti)
    │       └── altro utente      → aggregazione + filtro (is_mine=False per tutti)
    │
    ▼
_load_vsm_events(event_type, sheet)
    │
    ├── _get_active_vsm_username_filter()  ← scope utente attivo
    ├── get_all_vsm_events() o get_all_vsm_events_aggregated()
    └── _populate_vsm_sheet(sheet, events, extra_metadata=[...])
            │
            └── metadata per riga: {event_id, username, is_mine, source_file}

[Doppio click su riga]
    │
    ├── is_mine=True  → VSMEventDialog(read_only=False) [edit normale]
    └── is_mine=False → VSMEventDialog(read_only=True)  [sola lettura]

[Global Search VSM]
    │
    └── _search_vsm_events() legge _get_active_vsm_username_filter()
        filtra dentro lo scope già determinato dal filtro UI
```

---

## Principio architetturale rispettato

- Il filtro utente VSM determina **sempre** il dataset mostrato
- I dati di altri utenti restano read-only (tre livelli: metadata riga, dialog, pulsanti CRUD)
- La futura Global Search VSM multi-user può leggere `_get_active_vsm_username_filter()` senza logica parallela
- Nessun if/elif sparso fuori dai punti di dispatch (`_load_vsm_events`, `_search_vsm_events`)
- Logica RFQ invariata

---

## Ordine di implementazione consigliato

1. `database_manager.py` → `get_all_vsm_events_aggregated()` (layer fondamentale)
2. `ui/dialogs/vsm_event_dialog.py` → parametro `read_only` (prerequisito per step 5)
3. `dataflow.py` → variabili + `_get_active_vsm_username_filter()` + `populate_vsm_username_filter()`
4. `dataflow.py` → aggiornare `_populate_vsm_sheet()` con `extra_metadata`
5. `dataflow.py` → aggiornare `_load_vsm_events()`
6. `dataflow.py` → aggiornare `_search_vsm_events()`
7. `dataflow.py` → UI filtro utente VSM + `on_tab_changed()` + `clear_filters()`
8. Verifica finale: nessuna regressione RFQ, read-only funzionante, scope coerente
