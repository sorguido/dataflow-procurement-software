# Analisi Global Search RFQ e piano conservativo di estensione ai campi RFQ

## 1. Titolo
Analisi tecnica del comportamento attuale della Global Search RFQ e piano conservativo per estendere la ricerca rapida a tutti i campi RFQ rilevanti.

## 2. Obiettivo
Obiettivo del documento:
- mappare con precisione il comportamento attuale della Global Search RFQ;
- identificare file/funzioni coinvolti e campi realmente interrogati;
- evidenziare differenze tra dati UI, dati modello e dati interrogati;
- proporre una strategia minima, reversibile e conservativa per estendere la ricerca globale.

Nota: questo documento non implementa modifiche al codice.

## 3. Stato attuale osservato
La Global Search RFQ è attiva nel flusso dashboard e usa `search_vars['global']` come input condiviso.

Evidenze principali:
- L'entry point UI della Global Search è il tasto `Enter` nella toolbar (`MainDashboardToolbar._on_search`), che scrive in `search_vars['global']` e chiama `search_requests()`.
  - Riferimento: `ui/components/main_dashboard_toolbar.py:166-220`.
- Il pulsante `Search` dei filtri avanzati invoca direttamente `app.search_requests`.
  - Riferimento: `ui/main_dashboard_builder.py:270-273`.
- La logica RFQ di filtro è in `DashboardController.search_requests`.
  - Riferimento: `services/dashboard_controller.py:134`.
- La Global Search RFQ attuale usa OR su 6 campi:
  - `id_richiesta`, `riferimento`, `nome_fornitore`, `codice_materiale`, `descrizione_materiale`, `numeri_ordine`.
  - Riferimento: `services/dashboard_controller.py:212-221`, `339-370`.

Comportamento matching osservato:
- case-insensitive: sì, tramite `LOWER(...) LIKE LOWER(?)` in SQL o `.lower()` in memoria;
- substring: sì, tramite `%query%` o `query in string`;
- normalizzazione: parziale (trim input, rimozione blacklist caratteri `[';"\\`<>]`, max length 100);
- concatenazione campi: no, OR su clausole separate.

## 4. File coinvolti
File direttamente coinvolti nella pipeline RFQ/search:

1. `ui/components/main_dashboard_toolbar.py`
- Handler evento Enter della Global Search (`_on_search`).

2. `ui/main_dashboard_builder.py`
- Definizione `search_vars`, filtri avanzati RFQ, pulsanti Search/Clear.

3. `services/dashboard_controller.py`
- Orchestrazione ricerca RFQ/VSM (`search_requests`) e combinazione filtri.

4. `services/dashboard_search_service.py`
- Helper `has_active_search_filters` (rilevazione stato filtri).

5. `dataflow.py`
- Wrapper/dispatch (`search_requests`, `_has_active_search_filters`, `get_current_tree_and_status`, `update_treeview`).

6. `database_manager.py`
- Schema RFQ (`richieste_offerta`, `dettagli_richiesta`, `richiesta_fornitori`),
- aggregazione multi-DB (`get_all_richieste_aggregated`),
- ricerca avanzata locale (`search_richieste_advanced`),
- check filtri dettaglio (`check_richiesta_detail_criteria`).

7. `services/rfq_dashboard_service.py`
- Mapping tuple RFQ -> payload tabella (`build_rfq_sheet_payload`).

8. `ui/sheet_factories.py`
- Colonne effettivamente mostrate nella tabella RFQ.

## 5. Flusso attuale della Global Search
Flusso RFQ (tab `attiva` / `archiviata`):

1. Input utente
- Enter sulla search bar globale: scrive `search_vars['global']` e invoca `search_requests()`.
  - `ui/components/main_dashboard_toolbar.py:197-220`.

2. Dispatch tab corrente
- `get_current_tree_and_status()` determina se il tab è RFQ o VSM.
  - `dataflow.py:2315-2331`.
- Se RFQ: `DashboardController.search_requests` segue ramo RFQ.
  - `services/dashboard_controller.py:140-147`.

3. Costruzione criteri
- `crit = {k: v.get().strip() ...}` su `search_vars`.
- Sanitizzazione e validazione (max 100, rimozione caratteri blacklist).
  - `services/dashboard_controller.py:151-193`.

4. Combinazione logica
- Base AND: stato (tab), tipo, username, filtri standard, date.
- Se `global` presente: aggiunge blocco OR multi-campo.
  - `services/dashboard_controller.py:198-247`.

5. Esecuzione ricerca (2 modalità)
- `search_local_only=True` (utente corrente):
  - con global: SQL diretto (perché `search_richieste_advanced` non gestisce `global`);
  - senza global: `db_manager.search_richieste_advanced(...)`.
  - `services/dashboard_controller.py:288-311`.
- modalità aggregata multi-DB:
  - carica testata RFQ da `get_all_richieste_aggregated`;
  - applica filtri in memoria;
  - per campi dettaglio interroga il DB sorgente della singola RFQ.
  - `services/dashboard_controller.py:313-438`.

6. Render tabella
- `update_treeview` usa `build_rfq_sheet_payload` e mostra 6 colonne dashboard.
  - `dataflow.py:2354-2373`, `services/rfq_dashboard_service.py:44-88`.

## 6. Campi attualmente ricercabili
### 6.1 Global Search RFQ (attuale)
Campi inclusi (OR):
1. `ro.id_richiesta`
2. `ro.riferimento`
3. `rf.nome_fornitore`
4. `dr.codice_materiale`
5. `dr.descrizione_materiale`
6. `ro.numeri_ordine`

Riferimenti:
- Locale SQL: `services/dashboard_controller.py:214-221`.
- Aggregato (memoria + SQL dettaglio): `services/dashboard_controller.py:345-370`.

### 6.2 Filtri avanzati RFQ (non global)
Campi testuali dedicati:
- `num`, `ref`, `forn`, `cod`, `desc`, `ord`, `cod_grezzo`, `dis_grezzo`, `mat_cl`.
- Riferimento: `services/dashboard_controller.py:233-243`.

Campi strutturali:
- tipo RFQ (`search_tipo`), username, date emissione/scadenza.
- Riferimento: `services/dashboard_controller.py:200-207`, `244-247`.

## 7. Campi RFQ disponibili ma esclusi
### 7.1 Campi disponibili nel modello RFQ (schema)
Tabelle principali:
- `richieste_offerta`: `id_richiesta`, `data_emissione`, `data_scadenza`, `riferimento`, `note_generali`, `stato`, `numeri_ordine`, `tipo_rdo`, `note_formattate`, `username`.
- `dettagli_richiesta`: `id_dettaglio`, `id_richiesta`, `codice_materiale`, `descrizione_materiale`, `quantita`, `disegno`, `data_consegna_richiesta`, `codice_grezzo`, `disegno_grezzo`, `materiale_conto_lavoro`.
- `richiesta_fornitori`: `id_richiesta`, `nome_fornitore`.

Riferimento: `database_manager.py:152-156`.

### 7.2 Campi esclusi dalla Global Search (oggi)
Esclusi dalla global ma presenti nel modello RFQ:
- `ro.tipo_rdo`
- `ro.username`
- `ro.note_generali`
- `ro.note_formattate`
- `dr.disegno`
- `dr.quantita` (numericamente semantico, memorizzato come testo)
- `dr.data_consegna_richiesta`
- `dr.codice_grezzo` (presente solo nei filtri avanzati)
- `dr.disegno_grezzo` (presente solo nei filtri avanzati)
- `dr.materiale_conto_lavoro` (presente solo nei filtri avanzati)

### 7.3 Differenze tabella UI vs modello vs campi interrogati
- Colonne mostrate in dashboard RFQ: solo 6 (`RfQ No`, `RfQ Type`, `Issue Date`, `Expiry Date`, `Reference`, `User`).
  - `ui/sheet_factories.py:43`, `services/rfq_dashboard_service.py:79-86`.
- Dati disponibili nel modello: più ampi (testata + dettagli + fornitori).
- Dati interrogati da Global Search: sottoinsieme di 6 campi.

### 7.4 Tipologia campi (utili per estensione)
- Testuali: quasi tutti i campi RFQ citati sopra.
- Numerici: `id_richiesta` (INTEGER), `quantita` con semantica numerica ma tipo DB `VARCHAR`.
- Nullabili: larga parte dei campi non `NOT NULL` in schema (`riferimento`, date, note, numeri_ordine, campi dettaglio, username storico).

## 8. Interazione con filtri avanzati
Combinazione logica attuale RFQ:
- `(global OR su campi global)` AND `(filtri avanzati testuali/date/tipo/utente)`.

Evidenze:
- ramo SQL locale: blocco OR inserito nelle `clauses` e poi concatenato con `AND` al resto.
  - `services/dashboard_controller.py:212-247`.
- ramo aggregato: applicazione sequenziale con `continue` (equivalente logico ad AND tra blocchi).
  - `services/dashboard_controller.py:326-434`.

Nota importante:
- `search_richieste_advanced` non gestisce il campo `global`; per questo nel ramo locale con global si usa SQL diretto.
  - `services/dashboard_controller.py:299-304`, `database_manager.py:1497-1577`.

## 9. Rischi tecnici e regressivi
1. Divergenza locale vs aggregato
- Oggi la global è duplicata in due rami (SQL locale e filtro aggregato con query dettaglio). Estendere campi in un ramo solo genera inconsistenze.

2. Divergenza Global Search vs export filtrato
- `mega_export_excel` non applica `crit['global']` nel ramo filtri attivi (applica `num/ref/date/detail`, non global).
  - `dataflow.py:2921-2991`.
- Rischio: risultati esportati diversi da quelli visualizzati quando si usa Global Search.

3. Prestazioni in aggregato multi-DB
- Estendere la global a molti campi aumenta il numero di query di dettaglio per RFQ nel ramo aggregato.

4. Rumore su campi testuali lunghi
- Campi note (`note_generali`, `note_formattate`) possono produrre match molto ampi/non intuitivi.

5. Gestione null/formati
- Campo `quantita` non ha typing forte numerico in DB; confronto come testo può avere effetti inattesi (es. separatori).

## 10. Strategia conservativa consigliata
Strategia minima, additive, reversibile:

1. Preservare entry point e UX attuali
- Nessun cambio a toolbar, pulsanti, layout o comportamento Enter.

2. Introdurre whitelist esplicita campi Global Search RFQ
- Definire una sola whitelist canonica dei campi global ricercabili (no autodiscovery).
- Distinguere campi:
  - `head_fields` disponibili già in tuple aggregata o in query base;
  - `detail_fields` da verificare via JOIN nel DB sorgente.

3. Normalizzazione robusta del valore cercato
- Input query: trim + lower + gestione stringa vuota.
- Valori campo: normalizzare `None -> ""`, numeri -> `str(...)`, stringhe -> lower.
- In SQL usare `COALESCE(CAST(... AS TEXT), '')` dove opportuno.

4. Allineare i due percorsi di ricerca
- Applicare la stessa whitelist sia in:
  - ramo locale SQL;
  - ramo aggregato (match in memoria + query dettaglio).
- Evitare refactor massivo; modifica puntuale in `services/dashboard_controller.py` (e solo helper minimo, se serve).

5. Mantenere combinazione logica attuale
- Conservare esattamente: `(global OR)` + `AND` con filtri avanzati.

6. Rollback semplice
- Whitelist centralizzata permette rollback rapido ai 6 campi attuali.

## 11. Piano implementativo proposto a step
1. Inventario e freeze campi
- Definire elenco finale "tutti i campi RFQ" da includere (escludendo BLOB/allegati).
- Suggerito perimetro: colonne di `richieste_offerta`, `dettagli_richiesta`, `richiesta_fornitori`.

2. Whitelist unica
- Aggiungere in `services/dashboard_controller.py` una struttura esplicita dei campi global:
  - campi testata (`id_richiesta`, `riferimento`, `tipo_rdo`, `username`, `note_generali`, `note_formattate`, `numeri_ordine`, date);
  - campi dettaglio/fornitore (`nome_fornitore`, `codice_materiale`, `descrizione_materiale`, `quantita`, `disegno`, `data_consegna_richiesta`, `codice_grezzo`, `disegno_grezzo`, `materiale_conto_lavoro`).

3. Ramo locale
- Estendere solo il blocco OR SQL globale usando whitelist.
- Nessun cambio alla semantica AND dei filtri avanzati.

4. Ramo aggregato
- Conservare short-circuit su campi già in memoria;
- estendere query dettaglio OR ai nuovi campi whitelist non presenti in `all_results`.

5. Verifica coerenza export (ipotesi di hardening)
- Ipotesi: allineare in follow-up anche `mega_export_excel` a `crit['global']` per coerenza output.
- Se fuori scope, documentare esplicitamente il limite noto.

6. Test minimi e rilascio incrementale
- Eseguire test manuali mirati (sezione 12).
- Deploy conservativo, nessuna migrazione schema richiesta.

## 12. Test minimi consigliati
1. Regressione base Global Search
- Verificare che i 6 campi storici continuino a funzionare invariati.

2. Nuovi campi inclusi
- Un test per ciascun nuovo campo whitelist (almeno 1 RFQ match + 1 non match).

3. Combinazione con filtri avanzati
- Verificare `(global OR)` AND (`forn/cod/...`, date, tipo, utente).

4. Modalità locale vs aggregata
- Utente corrente (locale) vs All users/altro utente (aggregata): stessi risultati funzionali a parità dati.

5. Null e tipi misti
- Query su RFQ con campi `None`, campi vuoti, quantità non standard.

6. Tab attive/archiviate
- Verificare isolamento per stato (`attiva` vs `archiviata`).

7. Test di non regressione export (limite noto)
- Confermare e documentare l'attuale differenza tra risultati visualizzati con global e risultati esportati.

## 13. Rollback plan
Rollback conservativo in un commit:

1. Ripristinare whitelist Global Search ai 6 campi attuali.
2. Ripristinare eventuali query dettaglio OR estese al set originario.
3. Verificare con smoke test:
- global su codice materiale;
- global su descrizione;
- combinazione con un filtro avanzato;
- ricerca su tab attive/archiviate.

Nessuna migrazione DB da annullare.

## 14. Conclusione
Dall'analisi codice, la Global Search RFQ non è limitata a soli codice/descrizione: oggi copre 6 campi e si combina in AND con i filtri avanzati.

L'estensione a "tutti i campi RFQ" è fattibile senza refactor massivo con una strategia conservativa basata su whitelist esplicita e allineamento rigoroso dei due percorsi (locale/aggregato), mantenendo UX e semantica attuali.

Ambiguità dichiarata:
- La definizione operativa di "tutti i campi RFQ" va confermata in review (inclusione o meno di campi note e date come testo nella global).
