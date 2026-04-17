# Analisi Global Search RFQ su colonne visibili griglia materiali

## 1. Titolo
Analisi tecnica del comportamento attuale della Global Search RFQ e piano conservativo per estendere la ricerca rapida esclusivamente alle colonne visibili della griglia materiali/fornitori.

## 2. Obiettivo
Obiettivo del documento:
- verificare come funziona oggi la Global Search RFQ;
- mappare i campi reali collegati alle colonne visibili richieste;
- identificare quali di questi campi sono già ricercabili e quali no;
- proporre una strategia minima, reversibile e conservativa per estendere la Global Search solo alle colonne target.

Vincolo di perimetro applicato: esclusione esplicita di campi RFQ non appartenenti alla griglia target (es. note, date, username, tipo, quantità).

## 3. Stato attuale osservato
- Entry point UI Global Search: `MainDashboardToolbar._on_search`, che scrive `search_vars['global']` e invoca `search_requests()`.
  - `ui/components/main_dashboard_toolbar.py:166-220`
- Variabili filtro RFQ definite in dashboard: `search_vars` include `global`, `cod`, `desc`, `cod_grezzo`, `dis_grezzo`, `mat_cl`.
  - `ui/main_dashboard_builder.py:133-155`
- Filtro effettivo RFQ: `DashboardController.search_requests`.
  - `services/dashboard_controller.py:134`
- Global Search RFQ attuale (OR) cerca solo in:
  - `id_richiesta`, `riferimento`, `nome_fornitore`, `codice_materiale`, `descrizione_materiale`, `numeri_ordine`.
  - `services/dashboard_controller.py:212-221`, `359-369`
- Matching:
  - case-insensitive: sì (`LOWER(...) LIKE LOWER(?)` / `.lower()`),
  - substring: sì (`%query%` / `query in ...`),
  - normalizzazione: parziale (trim, blacklist caratteri pericolosi, limite lunghezza input),
  - concatenazione campi: no, OR su clausole separate.
  - `services/dashboard_controller.py:151-193`, `212-224`, `341-377`

## 4. File coinvolti
1. `ui/components/main_dashboard_toolbar.py`
- Handler evento Enter della Global Search (`_on_search`).

2. `ui/main_dashboard_builder.py`
- Definizione `search_vars` e pannello filtri RFQ.

3. `services/dashboard_controller.py`
- Logica di ricerca RFQ (global + filtri avanzati, locale/aggregata).

4. `services/dashboard_search_service.py`
- Supporto a rilevazione stato filtri (`has_active_search_filters`), nessuna logica RFQ di matching campi.

5. `dataflow.py`
- Wrapper `search_requests()` verso controller e routing dashboard.
  - `dataflow.py:2566-2567`

6. `ui/windows/view_request_window.py`
- Costruzione e rendering della griglia materiali/fornitori (`build_grid`) e mapping colonna->campo (`field_map`).

7. `database_manager.py`
- Persistenza e recupero campi materiali (`insert_dettaglio_richiesta`, `get_dettagli_by_richiesta`, `update_dettaglio_field`).

8. `services/rfq_dashboard_service.py`
- Gestisce payload della tabella RFQ dashboard (lista RdO), non la griglia materiali/fornitori.

## 5. Flusso attuale della Global Search
1. L'utente preme Enter nella search bar globale.
- `ui/components/main_dashboard_toolbar.py:166-220`

2. Il testo viene copiato in `search_vars['global']`, poi viene chiamato `search_requests()`.
- `ui/components/main_dashboard_toolbar.py:211-220`

3. `dataflow.MainWindow.search_requests()` delega a `DashboardController.search_requests()`.
- `dataflow.py:2566-2567`

4. Nel ramo RFQ, la query globale è applicata come blocco OR su 6 campi, combinato con AND rispetto ai filtri avanzati.
- `services/dashboard_controller.py:209-243`

5. In modalità aggregata multi-DB, il controllo globale sui campi di dettaglio viene rieseguito sul DB sorgente della singola RdO (query dedicata).
- `services/dashboard_controller.py:348-385`

Nota di contesto: la Global Search opera a livello elenco RdO dashboard; la griglia materiali/fornitori è renderizzata in `ViewRequestWindow` e ne fornisce il mapping campi.

## 6. Colonne target richieste
Colonne target richieste dal task:
- Codice
- Allegato
- Descrizione
- Cod. Grezzo
- Allegato Grezzo
- Mat. C/L

Ambiguità rilevata:
- Nel testo è presente anche la dicitura "7 colonne target", ma l'elenco esplicito contiene 6 colonne. In questo report è stato seguito l'elenco esplicito delle 6 colonne.

## 7. Mapping colonne UI -> campi reali
| Colonna UI | Campo sorgente reale | Dove viene valorizzato | Dove viene renderizzato | Stato in Global Search oggi |
|---|---|---|---|---|
| Codice | `dettagli_richiesta.codice_materiale` | Insert: `database_manager.insert_dettaglio_richiesta` (`database_manager.py:389-400`); update edit cell via `field_map` -> `update_dettaglio_field` (`ui/windows/view_request_window.py:1199-1203`, `1234-1237`) | `build_grid` colonna 0 (`ui/windows/view_request_window.py:993-1013`) | Incluso (`dr.codice_materiale`) |
| Allegato | `dettagli_richiesta.disegno` | Insert: `insert_dettaglio_richiesta` (`database_manager.py:389-400`); update edit cell `field_map` (`ui/windows/view_request_window.py:1201`, `1234-1237`) | `build_grid` colonna 1 (`ui/windows/view_request_window.py:993-1013`) | Escluso |
| Descrizione | `dettagli_richiesta.descrizione_materiale` | Insert: `insert_dettaglio_richiesta` (`database_manager.py:389-400`); update edit cell `field_map` (`ui/windows/view_request_window.py:1202`, `1234-1237`) | `build_grid` colonna 2 (`ui/windows/view_request_window.py:993-1014`) | Incluso (`dr.descrizione_materiale`) |
| Cod. Grezzo | `dettagli_richiesta.codice_grezzo` | Insert: `insert_dettaglio_richiesta` (`database_manager.py:390-400`); import Excel (`database_manager.py:838-845`); update edit cell (`ui/windows/view_request_window.py:1204`, `1234-1237`) | `build_grid` colonna CL (`ui/windows/view_request_window.py:1000-1019`) | Escluso dalla Global; presente in filtro avanzato (`crit['cod_grezzo']`) |
| Allegato Grezzo | `dettagli_richiesta.disegno_grezzo` | Insert: `insert_dettaglio_richiesta` (`database_manager.py:390-400`); import Excel (`database_manager.py:838-845`); update edit cell (`ui/windows/view_request_window.py:1205`, `1234-1237`) | `build_grid` colonna CL (`ui/windows/view_request_window.py:1000-1019`) | Escluso dalla Global; presente in filtro avanzato (`crit['dis_grezzo']`) |
| Mat. C/L | `dettagli_richiesta.materiale_conto_lavoro` | Insert: `insert_dettaglio_richiesta` (`database_manager.py:390-400`); import Excel (`database_manager.py:838-845`); update edit cell (`ui/windows/view_request_window.py:1206`, `1234-1237`) | `build_grid` colonna CL (`ui/windows/view_request_window.py:1000-1019`) | Escluso dalla Global; presente in filtro avanzato (`crit['mat_cl']`) |

Nota tecnica:
- Le ultime 3 colonne sono visualizzate solo se `self.is_conto_lavoro` (tipo RdO conto lavoro).
  - `ui/windows/view_request_window.py:1000-1004`

## 8. Campi attualmente ricercabili
### 8.1 Global Search (oggi)
Campi RFQ globali effettivi:
- `dr.codice_materiale` (colonna target: Codice)
- `dr.descrizione_materiale` (colonna target: Descrizione)
- più 4 campi non target (`id_richiesta`, `riferimento`, `nome_fornitore`, `numeri_ordine`).

Riferimenti:
- `services/dashboard_controller.py:214-221`
- `services/dashboard_controller.py:365-369`

### 8.2 Filtri avanzati (oggi)
Sui campi target sono già presenti filtri avanzati dedicati per:
- `dr.codice_materiale` (`crit['cod']`)
- `dr.descrizione_materiale` (`crit['desc']`)
- `dr.codice_grezzo` (`crit['cod_grezzo']`)
- `dr.disegno_grezzo` (`crit['dis_grezzo']`)
- `dr.materiale_conto_lavoro` (`crit['mat_cl']`)

Riferimento: `services/dashboard_controller.py:236-242`.

Nessun filtro avanzato dedicato al campo `dr.disegno` (colonna Allegato).

## 9. Colonne target oggi escluse dalla ricerca
Esclusione dalla Global Search attuale:
1. Allegato (`dr.disegno`) -> escluso.
2. Cod. Grezzo (`dr.codice_grezzo`) -> escluso dalla global, ricercabile solo via filtro avanzato.
3. Allegato Grezzo (`dr.disegno_grezzo`) -> escluso dalla global, ricercabile solo via filtro avanzato.
4. Mat. C/L (`dr.materiale_conto_lavoro`) -> escluso dalla global, ricercabile solo via filtro avanzato.

Colonne target già incluse in global:
- Codice (`dr.codice_materiale`)
- Descrizione (`dr.descrizione_materiale`)

## 10. Interazione con filtri avanzati
Logica attuale RFQ:
- `(global OR sui campi global)` AND `(filtri avanzati/strutturali/date)`.

Evidenze:
- blocco OR globale aggiunto alle `clauses`, poi clausole aggiuntive in AND.
  - `services/dashboard_controller.py:212-243`
- stesso principio nel ramo aggregato, con `continue` sequenziali.
  - `services/dashboard_controller.py:339-434`

Conseguenza:
- estendere la global alle sole colonne target non cambia l'architettura logica; amplia solo l'OR globale.

## 11. Rischi tecnici e regressivi
1. Divergenza locale vs aggregato
- La global è implementata in due rami distinti (SQL locale e controllo aggregato con query dettaglio). Estensione incompleta in un ramo crea incoerenze.

2. Rischio di sovraccarico nel ramo aggregato
- L'estensione di OR sui campi dettaglio aumenta il lavoro per verifica per-RdO su DB sorgente.

3. Ambiguità semantica "Allegato"
- La colonna target `Allegato` nella griglia materiali mappa a `dettagli_richiesta.disegno` (testo), non alla tabella allegati file `allegati_richiesta`.

4. Coerenza con filtri avanzati
- Tre campi (`codice_grezzo`, `disegno_grezzo`, `materiale_conto_lavoro`) esistono già come filtri avanzati: estenderli anche in global può aumentare overlap funzionale (comportamento comunque coerente con logica OR+AND attuale).

## 12. Strategia conservativa consigliata
Strategia minima, additive, reversibile, entro perimetro 6 colonne target:

1. Whitelist esplicita Global Search limitata ai soli campi target
- `dr.codice_materiale`
- `dr.disegno`
- `dr.descrizione_materiale`
- `dr.codice_grezzo`
- `dr.disegno_grezzo`
- `dr.materiale_conto_lavoro`

2. Nessuna estensione ad altri campi RFQ
- Esclusione esplicita di campi non target (date, note, username, tipo, ecc.).

3. Conservare la semantica attuale
- mantenere OR globale + AND con filtri avanzati, senza cambiare UX/UI.

4. Normalizzazione robusta valori
- query globale: trim/lower già presenti;
- clausole SQL target con `LOWER(COALESCE(campo, '')) LIKE LOWER(?)` per robustezza `None`.

5. Allineamento simmetrico dei due rami
- applicare identico set target sia nel ramo locale sia nel ramo aggregato (query dettaglio).

6. Modifiche piccole e reversibili
- intervento puntuale solo su costruzione `global_clauses` e query OR aggregata; nessun refactor strutturale.

## 13. Piano implementativo proposto a step
1. Definire costante/whitelist locale per i 6 campi target (solo in contesto ricerca RFQ).
2. Aggiornare blocco OR global in `services/dashboard_controller.py` ramo locale con i 6 campi target.
3. Aggiornare query `detail_sql` nel ramo aggregato con gli stessi 6 campi target.
4. Non modificare `search_vars`, layout o pannello filtri.
5. Verificare comportamento su tab `attiva` e `archiviata` in modalità locale/aggregata.
6. Validare rollback rapido ripristinando il blocco OR precedente.

## 14. Test minimi consigliati
1. Global Search su Codice (`codice_materiale`) -> match atteso.
2. Global Search su Allegato (`disegno`) -> match atteso (nuovo).
3. Global Search su Descrizione (`descrizione_materiale`) -> match atteso.
4. Global Search su Cod. Grezzo / Allegato Grezzo / Mat. C/L -> match atteso (nuovi in global).
5. Query che matcha solo campi non target (es. note/username) -> nessun match (comportamento voluto).
6. Combinazione Global + filtro avanzato (`cod_grezzo` o `mat_cl`) -> rispetto logica OR+AND.
7. Verifica parità risultati tra ramo locale e ramo aggregato a parità dati.

## 15. Rollback plan
Rollback in un singolo intervento:
1. Ripristinare OR globale al set precedente.
2. Ripristinare query `detail_sql` aggregata al set precedente.
3. Eseguire smoke test su:
- global `codice_materiale`,
- global `descrizione_materiale`,
- un filtro avanzato (`cod_grezzo`),
- tab attive/archiviate.

Nessuna migrazione DB necessaria.

## 16. Conclusione
La Global Search RFQ attuale copre solo 2 delle 6 colonne target richieste della griglia materiali (`Codice`, `Descrizione`).

L'estensione conservativa è tecnicamente fattibile senza refactor massivo, limitandosi a una whitelist esplicita dei soli 6 campi target e mantenendo invariata la logica corrente (OR globale + AND filtri avanzati).

Ambiguità documentata: il testo cita "7 colonne target" ma l'elenco operativo contiene 6 colonne; il piano è stato costruito sulle 6 colonne esplicitamente elencate.
