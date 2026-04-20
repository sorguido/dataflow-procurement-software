# BUG 2 — Analisi tecnica e piano di azione

## 1) Root Cause Analysis

### Sintesi
Il comportamento osservato e' coerente con un **mismatch di contesto DB** nel flusso di apertura dettaglio evento VSM in ambiente multi-DB aggregato.

La vista aggregata mostra correttamente eventi provenienti da database diversi, ma il dettaglio evento viene caricato sempre dal DB locale. Quando l'`event_id` selezionato appartiene a un DB esterno, il lookup locale fallisce con `VSMError: Evento <id> non trovato`.

### Evidenza nel codice
1. **Aggregazione multi-DB preserva la sorgente**
- `database_manager.py:2304-2402` (`get_all_vsm_events_aggregated`) restituisce tuple `(VSMEvent, is_mine, source_file)`.
- `services/vsm_dashboard_service.py:15-29` mantiene `source_file` in `extra_meta`.

2. **Il metadata di sorgente arriva fino alla griglia VSM**
- `dataflow.py:1759-1772` salva per ogni riga: `event_id`, `is_mine`, `source_file`.

3. **In apertura dialog il `source_file` viene perso**
- `dataflow.py:1851-1868` legge `event_id` e `is_mine`, ma non usa `source_file`.
- `VSMEventDialog` viene aperto con `event_id` soltanto.

4. **Il dialog ricarica dal DB locale fisso**
- `ui/dialogs/vsm_event_dialog.py:509-510`: `with DatabaseManager(get_db_path()) as db_manager:`.
- `services/vsm_persistence.py:247-250`: `get_event_with_impacts` cerca per `event_id` nel DB aperto; se assente -> `VSMError("Evento <id> non trovato")`.

5. **`event_id` non e' globale cross-DB**
- `database_manager.py:203-205`: `event_id INTEGER PRIMARY KEY AUTOINCREMENT` (scope locale al singolo file SQLite).

### Punto preciso di rottura
Il punto di rottura e' tra:
- metadata riga aggregata (`source_file`) disponibile in `dataflow.py`, e
- chiamata a `VSMEventDialog` che non riceve/usa quella sorgente.

Il lookup viene quindi eseguito su DB sbagliato.

### Conseguenza tecnica aggiuntiva (critica)
Oltre all'errore "non trovato", esiste rischio di **collisione silente**:
- se due DB diversi hanno lo stesso `event_id`, il caricamento locale puo' aprire un evento diverso da quello cliccato.

---

## 2) Architettura attuale (semplificata)

```text
UI VSM list load
  -> service_get_vsm_dataset()
    -> get_all_vsm_events_aggregated()
      -> [(event, is_mine, source_file), ...]
  -> _populate_vsm_sheet()
    -> sheet._event_metadata[row] = {event_id, is_mine, source_file}

Double click / Edit
  -> _edit_vsm_event()
    -> legge metadata (event_id, is_mine)
    -> apre VSMEventDialog(event_id=... , read_only=...)

Dialog load
  -> _load_event_data()
    -> DatabaseManager(get_db_path())   # sempre locale
    -> get_event_with_impacts(event_id)
```

Conclusione: il pipeline dati di lista e' multi-source, il pipeline dettaglio e' single-source locale.

---

## 3) Impatti

### UX
- L'utente vede eventi aggregati ma non riesce ad aprirli in sola lettura.
- Errore percepito come incoerenza grave del prodotto (lista dice "esiste", dettaglio dice "non trovato").
- Messaggio misto EN/IT riduce chiarezza nel troubleshooting.

### Dati
- Nessuna corruzione diretta nel caso errore "not found".
- Rischio alto di lettura record errato in caso collisione `event_id` tra DB diversi (integrita' informativa compromessa).

### Rischio regressione
- Alto in modalita' multi-utente/multi-DB.
- Area sensibile: apertura record read-only cross-user (funzione documentata come comportamento atteso).

---

## 4) Piano di Azione (step-by-step, conservativo)

## Fix minimo raccomandato (circoscritto)

1. **Propagare la sorgente DB dal metadata UI al dialog**
- In `dataflow.py` (`_edit_vsm_event`), leggere anche `source_file` da `sheet._event_metadata[row_idx]`.
- Passare `source_db_path` al dialog quando disponibile.

2. **Aggiungere parametro opzionale al dialog senza rompere API esistente**
- In `VSMEventDialog.__init__`, introdurre parametro opzionale `source_db_path=None`.
- Default invariato per i flussi esistenti.

3. **Usare la sorgente reale solo nel load evento**
- In `_load_event_data`, scegliere DB target:
  - `source_db_path` se valorizzato e diverso da `local`;
  - altrimenti `get_db_path()`.
- Aprire in lettura quando il contesto e' read-only (coerente multi-DB).

4. **Mantenere invariata la logica di save/update/delete**
- Nessun cambio su persistenza write.
- Nessun cambio architetturale, nessuna nuova dipendenza, nessuna migrazione schema.

5. **Aggiungere logging diagnostico minimo**
- Loggare `event_id`, `source_db_path`, `read_only` in apertura dialog (debug level).

### Perche' questo fix e' minimale
- Tocca solo il passaggio di contesto DB in apertura dettaglio.
- Non cambia modello dati, query aggregate, policy ownership, filtri, export.
- Reversibile in pochi file.

## Alternative (non raccomandate come prima scelta)

### Alternativa A — Fallback "scan tutti i DB" se not found locale
Pro:
- Non richiede passare `source_db_path` dal chiamante.

Contro:
- Ambigua con collisioni `event_id`.
- Piu' costosa e meno deterministica.
- Introduce comportamento implicito difficile da testare.

### Alternativa B — Introdurre identificatore globale evento
Pro:
- Risolve strutturalmente l'ambiguita'.

Contro:
- Cambio architetturale/schemi piu' ampio.
- Fuori dal vincolo "fix minimo".

---

## 5) Piano di Test Manuale

## Setup
- Ambiente con almeno 2 DB (`A.db`, `B.db`) nella stessa cartella condivisa.
- Utente loggato su DB A.
- Presenza di eventi VSM in entrambi i DB.

## Test 1 — Riproduzione bug attuale (pre-fix)
1. Impostare filtro VSM su "All users".
2. Selezionare evento proveniente da DB B (non `is_mine`).
3. Doppio click / Edit.

Atteso pre-fix:
- Popup errore `Unable to load event: Evento <id> non trovato`.

## Test 2 — Apertura read-only evento esterno (post-fix)
1. Ripetere i passi del Test 1.

Atteso post-fix:
- Evento si apre correttamente.
- Dialog in sola lettura.
- Nessun tentativo di save disponibile.

## Test 3 — Regressione zero su evento locale proprietario
1. Selezionare evento locale `is_mine=True`.
2. Aprire in edit, modificare e salvare.

Atteso:
- Flusso invariato rispetto a oggi.

## Test 4 — Collisione ID cross-DB
1. Preparare evento con stesso `event_id` in DB A e DB B.
2. Da vista aggregata aprire evento DB B.

Atteso post-fix:
- Si apre il record di DB B (sorgente corretta), non l'omonimo locale.

## Test 5 — Ricerca/ordinamento + apertura
1. Applicare global search VSM + eventuale sort colonna.
2. Aprire record non locale.

Atteso post-fix:
- Apertura coerente con riga visualizzata.

---

## 6) Strategia di Rollback

Rollback semplice e immediato (nessuna migrazione DB):

1. Ripristinare i file toccati dal fix (tipicamente):
- `dataflow.py`
- `ui/dialogs/vsm_event_dialog.py`
- eventuali file `locale/*/LC_MESSAGES/dataflow.po` solo se aggiornati

2. Comando operativo (esempio):
- `git restore dataflow.py ui/dialogs/vsm_event_dialog.py locale/en/LC_MESSAGES/dataflow.po locale/it/LC_MESSAGES/dataflow.po`

3. Rieseguire i test manuali Test 1/Test 3 per verificare ritorno al comportamento precedente.

---

## 7) Nota i18n (secondaria)

Il messaggio misto EN/IT nasce da composizione di:
- wrapper UI tradotto: `tr("Unable to load event:\n{}")`
- eccezione backend hardcoded in italiano: `"Evento <id> non trovato"`.

### Allineamento consigliato (minimo)
- Uniformare le stringhe di errore VSM su msgid coerenti con `tr(...)`.
- Evitare testo utente hardcoded nelle eccezioni backend; mantenere backend tecnico e UI responsabile della localizzazione.

Risultato atteso:
- messaggi completamente EN o completamente IT in base alla lingua attiva.

---

## Conclusione operativa

La root cause e' un problema di **propagazione contesto sorgente DB** tra lista aggregata e dialog dettaglio.

Il fix consigliato e' locale, reversibile e a basso rischio: passare `source_file` fino a `_load_event_data` e usare quel DB per il lookup evento.

---

## 8) Confronto con il flusso RFQ

### Evidenza tecnica concreta (codice)
1. **RFQ aggregato conserva sempre la sorgente DB**
- `database_manager.py:1239-1252, 1285-1296, 1327-1337`: `get_all_richieste_aggregated()` produce righe con `is_mine` + `source_file`.
- `services/rfq_dashboard_service.py:44-74`: `build_rfq_sheet_payload()` trasferisce `is_mine` e `source_file` in `metadata_rows`.
- `dataflow.py:2498-2511`: `update_treeview()` salva i metadata in `sheet._sheet_rows_metadata`.

2. **RFQ propaga il contesto sorgente fino all'apertura dettaglio**
- `dataflow.py:2907-2911`: su doppio click legge `metadata['is_mine']` e `metadata['source_file']`.
- `dataflow.py:2918-2923`: apre `ViewRequestWindow(..., source_db_path=source_file if not is_mine else None)`.

3. **RFQ usa realmente il DB sorgente nel dettaglio**
- `ui/windows/view_request_window.py:66-83`: `__init__` imposta `self.db_path` da `source_db_path` (se valido), altrimenti locale.
- `ui/windows/view_request_window.py:340-347`: `_get_db_manager()` usa `DatabaseManager(self.db_path, read_only=self.read_only)`.
- `ui/windows/view_request_window.py:536-539, 607-608, 728-737`: le query di dettaglio usano `_get_db_manager()` / `self.db_path`, non un path fisso locale.

4. **VSM non propaga la sorgente nel path di apertura**
- `dataflow.py:1767-1772`: la lista VSM salva correttamente `source_file` nei metadata riga.
- `dataflow.py:1851-1868`: `_edit_vsm_event()` usa `event_id` e `is_mine`, ma non passa `source_file` a `VSMEventDialog`.
- `ui/dialogs/vsm_event_dialog.py:509-510`: `_load_event_data()` apre sempre `DatabaseManager(get_db_path())`.

### Confronto RFQ vs VSM (punti richiesti)
1. **Identificatore record**
- RFQ: `id_richiesta` e' `INTEGER PRIMARY KEY AUTOINCREMENT` (`database_manager.py:152`), quindi locale al DB.
- VSM: `event_id` e' `INTEGER PRIMARY KEY AUTOINCREMENT` (`database_manager.py:203-205`), quindi locale al DB.

2. **Presenza `source_file/source_db_path`**
- RFQ: presente in dataset aggregato (`source_file`) e propagato a `source_db_path` in apertura dettaglio.
- VSM: presente in metadata lista (`source_file`), ma non propagato al dialog.

3. **Passaggio contesto lista -> dettaglio**
- RFQ: SI, esplicito e verificabile (`_sheet_rows_metadata` -> `ViewRequestWindow(... source_db_path=...)`).
- VSM: NO nel flusso attuale (`_event_metadata` contiene `source_file`, ma `VSMEventDialog` non lo riceve).

4. **Lookup record in apertura**
- RFQ: lookup su coppia implicita `(id_richiesta, db_path)` perche' il DB scelto e' `self.db_path`.
- VSM: lookup su `event_id` nel solo DB locale (`get_db_path()`), quindi mismatch in aggregato.

5. **Rischio collisione ID cross-DB**
- RFQ: rischio strutturale teorico presente (ID locali), ma mitigato nel dettaglio dall'uso del DB sorgente.
- VSM: rischio presente e non mitigato nel flusso attuale; oltre al "not found", puo' aprire record sbagliato in caso collisione.

### Risposte dirette alle domande
1. **Perche' nei tab RFQ il problema non si manifesta?**  
Perche' RFQ propaga e usa il contesto DB sorgente fino al dettaglio (`source_file` -> `source_db_path` -> `self.db_path`).

2. **RFQ usa un meccanismo equivalente di propagazione contesto?**  
Si'. Esiste gia' ed e' operativo nel flusso di apertura RdO.

3. **Oppure RFQ evita a monte il problema con pipeline diversa?**  
Usa entrambe le cose: pipeline simile (lista aggregata + metadata), ma con passaggio di contesto completo fino al dettaglio, che evita il mismatch.

4. **La soluzione proposta per BUG 2 VSM e' coerente con RFQ?**  
Si'.

5. **In che senso preciso?**  
Replica lo stesso principio gia' usato in RFQ: usare l'ID nel DB di origine della riga selezionata, non nel DB locale fisso.

6. **Differenze architetturali da evidenziare (senza refactor)?**  
Differenza dimostrata: RFQ dispone gia' di `source_db_path` nel costruttore della finestra dettaglio; VSM no.  
Non risultano dal codice altre differenze necessarie per giustificare un refactor extra-scope.

### Conclusione netta
**fix VSM coerente con RFQ**.
