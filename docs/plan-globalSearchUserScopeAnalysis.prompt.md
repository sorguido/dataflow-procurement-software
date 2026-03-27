# Report Tecnico — Global Search RFQ e Scope Utente degli Advanced Filters

---

## 1. File coinvolti

| File | Ruolo |
|------|-------|
| `dataflow.py` | Tutta la logica: filtro username, `search_requests()`, `populate_username_filter()`, `update_treeview()`, `_check_if_all_selected_are_mine()` |
| `database_manager.py` | `get_all_richieste_aggregated()` — produce il campo `is_mine` e `source_file` per ogni riga |

---

## 2. Funzioni e metodi coinvolti

| Metodo | File | Ruolo |
|--------|------|-------|
| `populate_username_filter()` | dataflow.py L3844 | Popola la ComboBox utenti da tutti i DB aggregati |
| `_get_active_username_filter()` | dataflow.py L3896 | Restituisce lo scope utente attivo come stringa lowercase o `None` |
| `search_requests()` | dataflow.py L5526 | Applica global search + filtri; routing locale vs aggregato |
| `update_treeview()` | dataflow.py L5320 | Inserisce `is_mine` + `source_file` nei metadati di ogni riga |
| `_check_if_all_selected_are_mine()` | dataflow.py L4989 | Blocca Edit/Delete se anche solo una riga ha `is_mine=False` |
| `get_all_richieste_aggregated()` | database_manager.py L1161 | UNION multi-DB; aggiunge `is_mine` (TRUE locale, FALSE esterno) e `source_file` |

---

## 3. Come viene gestito oggi il filtro username

**Popolamento lista utenti (`populate_username_filter`, L3844):**
- Chiama `get_all_richieste_aggregated()` che scandisce ricorsivamente la cartella condivisa tramite `glob`
- Estrae gli username unici da tutti i risultati UNION (indice `[5]`)
- Aggiunge `all_users_placeholder` come prima voce
- La ComboBox è `state="readonly"` → nessun input manuale possibile
- Binding: `<<ComboboxSelected>>` → `refresh_data()` (ricarica dati rispettando il nuovo filtro)

**Estrazione filtro attivo (`_get_active_username_filter`, L3896):**
```
None             → se ComboBox = all_users_placeholder  ("Tutti gli utenti")
username.lower() → se ComboBox = utente specifico
```

**Tre scenari di scope:**

| Scenario | `username_filter` | Modalità ricerca |
|----------|-------------------|-----------------|
| Nessun filtro / "Tutti gli utenti" | `None` | Aggregata multi-DB |
| Utente corrente selezionato | `"mionome"` (= `current_username`) | **Locale ottimizzato** (solo DB locale) |
| Altro utente selezionato | `"altroutente"` | Aggregata multi-DB + filtro in-memory per username |

---

## 4. Come viene gestita oggi la Global Search RFQ

In `search_requests()` il `username_filter` viene letto **sempre per primo** (L5537), prima di costruire qualsiasi query o filtro. Il flusso è:

```
username_filter = _get_active_username_filter()     ← scope utente
                     │
          ┌──────────┴─────────────┐
          │ search_local_only?     │
    utente corrente             tutti / altro utente
          │                        │
    DB locale + OR clause    aggregazione multi-DB
    (username già nella SQL) + filtro username in-memory
    + global search inclusa  + global search su source_db
```

**Nel path locale** (L5720-L5740): il `username_filter` viene incorporato in `clauses` direttamente nella SQL (`LOWER(COALESCE(ro.username, '')) = ?`). La global search OR si aggiunge come ulteriore clausola AND — dentro lo stesso scope.

**Nel path aggregato** (L5743-L5855):
1. `get_all_richieste_aggregated()` carica TUTTO (tutti gli utenti, tutti i DB)
2. Primo filtro in-memory: `row[6] != status` → escludi stato sbagliato
3. **Secondo filtro in-memory: `username_filter`** → escludi righe di altri utenti se il filtro è attivo
4. Terzo filtro: tipo RdO
5. Quarto filtro: global search on-the-fly sul `source_db_path` specifico tramite query di dettaglio
6. Quinta fase: filtri standard su campi dettaglio via `check_richiesta_detail_criteria()`

La global search **non bypassa mai il filtro username**: il filtro username viene applicato *prima* che la riga venga valutata per la global search.

---

## 5. La Global Search eredita davvero lo scope utente?

**Sì, in entrambi i path.**

- **Path locale:** lo scope è implicito — si interroga solo il DB corrente, e la SQL include già `LOWER(COALESCE(ro.username, '')) = ?`
- **Path aggregato:** lo scope è applicato in-memory al passo 2 (L5760-L5762), prima che venga eseguita qualsiasi query di dettaglio per la global search

Quindi: se il filtro Advanced punta ad un altro utente, `search_requests()` mostra solo le RFQ di quell'utente, anche tramite global search. Se il filtro è "Tutti gli utenti", vengono mostrate le RFQ di tutti (come atteso).

---

## 6. Il comportamento read-only è preservato?

**Sì, a più livelli indipendenti.**

**Livello 1 — Metadati di riga (`update_treeview`, L5354-L5367):**
Ogni riga del treeview riceve metadati `{'is_mine': bool, 'source_file': str/path}` estratti da `req[7]` e `req[8]` (campi prodotti da `get_all_richieste_aggregated`). Non importa se la riga è arrivata da global search o da caricamento normale: i metadati viaggiano sempre con la riga.

**Livello 2 — Apertura dettaglio (doppio click):**
`ViewRequestWindow` viene aperto con `read_only=not is_mine` — se `is_mine=False`, la finestra è in sola lettura.

**Livello 3 — Pulsanti Edit/Delete (`_check_if_all_selected_are_mine`, L4989):**
Blocca le operazioni distruttive se anche una sola riga selezionata ha `is_mine=False`. Default a `False` per sicurezza in caso di metadati mancanti.

**Conservazione dei metadati nella global search:**
Nel path aggregato, i risultati della ricerca sono elementi dell'array `all_results` — tuple complete con indici `[7]` e `[8]` intatti. `update_treeview()` li legge invariati. **Nessun rischio di bypass.**

---

## 7. Valutazione finale

**COERENTE**

La Global Search RFQ rispetta il principio architetturale stabilito:
> "Gli Advanced Filters definiscono lo scope dati attivo. La Global Search deve sempre operare dentro quello scope, senza bypassarlo."

- Il filtro username viene applicato **prima** della global search, non dopo
- I metadati `is_mine`/`source_file` sopravvivono intatti attraverso tutto il pipeline di ricerca
- Le protezioni read-only operano sui metadati di riga, non sul contesto della ricerca — sono quindi indipendenti dal percorso con cui la riga è stata prodotta

---

## Implicazione per il modulo VSM

L'attuale `_search_vsm_events()` non ha uno scope utente parametrizzabile dagli Advanced Filters, perché:

1. Gli Advanced Filters RFQ non hanno un controllo "Utente" che abbia significato per i tab VSM
2. `get_all_vsm_events(username=self.current_username)` — filtra già sempre per utente corrente, identico al comportamento di default RFQ

**Questa è la differenza architetturale rilevante:** il modulo RFQ supporta la visualizzazione in read-only dei dati altrui (selezionando un altro utente nella ComboBox); il modulo VSM in V1 è limitato al solo utente corrente.

Quando in futuro si vorrà estendere lo scope VSM multi-utente, il pattern da replicare è:
- aggiungere `_get_active_username_filter()` nel dispatch di `_search_vsm_events()`
- passarlo a `get_all_vsm_events(username=…)` (o a una futura versione aggregata)
- propagare `is_mine` nei metadati di riga di `_populate_vsm_sheet()` — il meccanismo è già presente (`metadata['is_mine']` esiste in L4566)
