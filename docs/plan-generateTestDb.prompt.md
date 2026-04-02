# Piano: Generate Full Test Database for DataFlow

## A. SCHEMA E TABELLE COINVOLTE

**Database**: SQLite3, file per-utente. Driver: `sqlite3` stdlib. Nessun ORM. WAL mode. Il file di test verrà creato separatamente: `test_dataflow_full.db`.

| Tabella | Scopo | In scope |
|---|---|---|
| `richieste_offerta` | Header RFQ (id year-based) | ✓ |
| `dettagli_richiesta` | Righe materiali legate a RFQ | ✓ |
| `richiesta_fornitori` | Associazione RFQ ↔ fornitore | ✓ |
| `offerte_ricevute` | Prezzo unitario fornitore per riga | ✓ |
| `vsm_events` | Saving / Cost Avoidance / Derisking | ✓ |
| `vsm_impacts` | Distribuzione mensile impatti (auto) | ✓ (auto-generate) |
| `potential_suppliers` | Anagrafica fornitori derisking | fuori scope |
| `fornitori`, `allegati_richiesta`, `utenti` | — | fuori scope |

---

## B. RELAZIONI E VINCOLI

```
richieste_offerta (1) ──→ (N) dettagli_richiesta       via id_richiesta (FK soft)
richieste_offerta (1) ──→ (N) richiesta_fornitori      via id_richiesta (FK soft)
dettagli_richiesta (1) ──→ (N) offerte_ricevute        via id_dettaglio (FK soft)
vsm_events (1) ──→ (N) vsm_impacts                     via event_id (FK dichiarata)
vsm_events.username ──→ utenti.username                FK dichiarata ma NON ENFORCED
```

**Nota critica**: `DatabaseManager.connect()` NON imposta `PRAGMA foreign_keys=ON`. I vincoli FK sono presenti nello schema ma mai applicati a runtime. La stringa username `'gsoraru'` funziona senza che esista un record nella tabella `utenti`.

**richiesta_fornitori.nome_fornitore**: stringa libera, **nessuna FK** sulla tabella `fornitori` RFQ.

---

## C. CAMPI OBBLIGATORI DA POPOLARE

**richieste_offerta**:
| Campo | Vincolo | Valore/Range ammesso |
|---|---|---|
| `stato` | NOT NULL DEFAULT | `'attiva'` \| `'archiviata'` |
| `tipo_rdo` | NOT NULL DEFAULT | `'Fornitura piena'` \| `'Conto lavoro'` |
| tutti gli altri | nullable | — |

**richiesta_fornitori**: entrambe le colonne obbligatorie (composite PK).

**offerte_ricevute**: `id_dettaglio` + `nome_fornitore` obbligatori (composite PK). `prezzo_unitario` nullable ma deve essere popolato.

**vsm_events**:
| Campo | Vincolo |
|---|---|
| `username` | NOT NULL |
| `opex_ripetitivo` | NOT NULL DEFAULT 0 |
| tutti gli altri | nullable |

**vsm_impacts**: `event_id`, `username`, `anno`, `mese`, `tipo_valore`, `valore_teorico`, `valore_effettivo` — tutti NOT NULL.

---

## D. REGOLA REALE DI GESTIONE NUMERI E DATE

### Numeri — due regole diverse, non una sola

| Contesto | Tipo colonna | Formato da usare nel DB |
|---|---|---|
| `offerte_ricevute.prezzo_unitario` | **VARCHAR** | `"123,4500"` — virgola decimale (es. `f"{price:.4f}".replace('.', ',')`) |
| `vsm_events` campi economici (`importo_bdg`, `importo_negoziato`, ecc.) | **REAL** | Python `float` standard — es. `10000.0`. La virgola è **solo layer UI**: `parse_float_from_comma_string` decodifica l'input UI → float → SQLite. Lo script scrive float direttamente → nessuna virgola necessaria nel DB |
| `dettagli_richiesta.quantita` | **VARCHAR** | Stringa intera: `"100"` |

**La regola applicativa "virgola decimale" riguarda esclusivamente i prezzi VARCHAR in `offerte_ricevute`. I campi REAL di `vsm_events` accettano Python float direttamente.**

### Date

| Colonna | Formato |
|---|---|
| `richieste_offerta.data_emissione`, `data_scadenza` | `"YYYY-MM-DD"` (ISO) |
| `dettagli_richiesta.data_consegna_richiesta` | `"YYYY-MM-DD"` |
| `vsm_events.event_date` | `datetime` object passato a `save_event_with_impacts()` |

---

## E. STRATEGIA PIÙ SICURA PER GENERARE IL DB DI TEST

Il progetto **ha già** `generate_test_db.py` (seed pipeline esistente con 30 RFQ). La strategia è estenderlo:

1. **`DatabaseManager(DB_PATH)` + `db.create_tables()`** → schema autoritative, incluse `vsm_events`, `vsm_impacts`, `potential_suppliers`. Nessuna duplicazione schema.

2. **SQL diretto (raw cursor)** per le tabelle RFQ — stesso pattern del script esistente, collaudato.

3. **`save_event_with_impacts()` da `services/vsm_persistence.py`** per gli eventi VSM → gestisce atomicamente l'inserimento evento + generazione distribuzione mensile impatti tramite l'engine. Per Derisking: lista impacts vuota, comportamento corretto e confermato dai test esistenti.

---

## F. FILE ESATTI DA MODIFICARE/CREARE

| File | Azione | Motivo |
|---|---|---|
| `generate_test_db.py` | **MODIFICARE** | Seed pipeline esistente da estendere (500 RFQ multi-anno + 300 eventi VSM). DB_PATH → `test_dataflow_full.db` per non sovrascrivere il DB precedente. |

**Nessun altro file da toccare.** Schema, engine, persistence layer: invariati.

---

## G. RISCHI / PUNTI DI ATTENZIONE

1. **RFQ ID collision**: IDs year-based `YY%100 * 100000 + seq` → 100.000 slot per anno, utilizziamo max ~72/anno → **sicuro**.

2. **utenti FK**: non enforced → nessun blocco. Username `'gsoraru'` funziona senza record in `utenti`.

3. **Derisking events**: nessun campo economico (NULL). `generate_impacts_for_event()` ritorna `[]` → `save_event_with_impacts()` inserisce solo l'evento, nessun impact. **Comportamento atteso e corretto**.

4. **vsm_events.event_date**: `save_event_with_impacts()` accetta `VSMEvent.event_date` come oggetto `datetime` → generare date come `datetime(year, month, day)`.

5. **Import dal root**: `from database_manager import DatabaseManager`, `from services.vsm_persistence import save_event_with_impacts`, `from models.vsm_event import VSMEvent` — tutti funzionano dal root del progetto.

6. **Saving driver Pagamenti vs Prezzo**: per i test usare preferenzialmente driver `'Prezzo'` (più semplice: richiede solo `importo_bdg` + `importo_negoziato`), con un 20% driver `'Pagamenti'` per varietà.

---

## H. PIANO DI IMPLEMENTAZIONE MINIMALE

1. `DB_PATH = 'test_dataflow_full.db'` — file separato, nessun overwrite.
2. Schema via `DatabaseManager(DB_PATH).create_tables()`.
3. **500 RFQ** in loop `for year in range(2020, 2027)`: ~71-72/anno, `id_richiesta = (year % 100) * 100000 + i + 1`, date ISO random nel range annuale, mix `tipo_rdo`.
4. Per ogni RFQ: 3 fornitori in `richiesta_fornitori`, 1-3 righe in `dettagli_richiesta`, prezzi in `offerte_ricevute` (format `"NNN,NNNN"`).
5. **300 VSM events** (100+100+100) con `save_event_with_impacts()`: distribuiti nei 7 anni, dati economici realistici con Python float, `event_date=datetime(...)`.
6. Verifiche count stampate a fine script.

---

## Query di verifica finali

```sql
-- Count RFQ
SELECT COUNT(*) AS rfq_totali FROM richieste_offerta;

-- Fornitori per RFQ (deve essere 3 per ogni RFQ)
SELECT id_richiesta, COUNT(*) AS n_fornitori
FROM richiesta_fornitori
GROUP BY id_richiesta
HAVING n_fornitori != 3;
-- deve restituire 0 righe

-- Count VSM events per tipo
SELECT event_type, COUNT(*) FROM vsm_events GROUP BY event_type;

-- Distribuzione RFQ per anno
SELECT SUBSTR(data_emissione, 1, 4) AS anno, COUNT(*) FROM richieste_offerta GROUP BY anno ORDER BY anno;

-- Distribuzione VSM events per anno
SELECT SUBSTR(event_date, 1, 4) AS anno, COUNT(*) FROM vsm_events GROUP BY anno ORDER BY anno;

-- Campi obbligatori non nulli
SELECT COUNT(*) FROM richieste_offerta WHERE stato IS NULL OR tipo_rdo IS NULL;
SELECT COUNT(*) FROM vsm_events WHERE username IS NULL;

-- Record orfani
SELECT COUNT(*) FROM dettagli_richiesta dr
LEFT JOIN richieste_offerta ro ON dr.id_richiesta = ro.id_richiesta
WHERE ro.id_richiesta IS NULL;

SELECT COUNT(*) FROM offerte_ricevute o
LEFT JOIN dettagli_richiesta dr ON o.id_dettaglio = dr.id_dettaglio
WHERE dr.id_dettaglio IS NULL;

SELECT COUNT(*) FROM vsm_impacts vi
LEFT JOIN vsm_events ve ON vi.event_id = ve.event_id
WHERE ve.event_id IS NULL;
```
