# Piano di Allineamento Export VSM alla Vista Filtrata (Saving / Cost Avoidance / Derisking)

## 1. Titolo
Allineamento funzionale `EXPORT = QUELLO CHE VEDO` per i tab VSM della dashboard principale.

## 2. Sintesi esecutiva
Nei tab VSM (`Saving`, `Cost Avoidance`, `Derisking`) la vista usa una pipeline di filtri che include Global Search e filtri contestuali. Gli export Excel, invece, ricalcolano dataset separati e parziali rispetto alla vista. Questo produce divergenza tra griglia e file esportato.

La correzione consigliata, conservativa e a basso rischio, è:
1. forzare all’export il riallineamento della vista tramite la stessa pipeline (`search_requests()` del tab corrente);
2. esportare il dataset **effettivamente visualizzato** (cache per-sheet) invece di ricostruirlo con logica duplicata.

Scope limitato ai soli export VSM dashboard. Nessun impatto su RFQ o KPI.

## 3. Bug confermato
Comportamento attuale nei tab VSM:
- la griglia può essere filtrata (Global Search e/o filtri del tab),
- ma l’export può includere record non presenti nella griglia.

Questo viola la semantica desiderata: `EXPORT = QUELLO CHE VEDO`.

## 4. Ambito esatto coinvolto
Incluso:
- Dashboard tab `Saving`, `Cost Avoidance`, `Derisking`;
- pipeline di caricamento/ricerca VSM che alimenta le griglie;
- rami export VSM/Derisking in `mega_export_excel`.

Escluso:
- RFQ export;
- KPI export;
- layout/UI redesign;
- refactor strutturali estesi.

## 5. Flusso attuale per ciascun tab

### Saving
Vista:
- entrypoint: `DashboardController.search_requests()` dispatch su `_search_vsm_events()` quando tab è `vsm_saving`;
- query globale: `search_vars['global']`;
- pipeline dati vista:
  - `_get_vsm_dataset()`
  - split per tipo evento (`Saving`)
  - `_apply_vsm_filters()` (date/action/repetitive/amount)
  - `filter_vsm_events_by_query()` (Global Search)
  - `_populate_vsm_sheet()`.

Export:
- `_export_vsm_excel()` ricarica DB via `_get_vsm_dataset()` + split tipo + `_apply_vsm_filters()`;
- **non applica Global Search**;
- non usa dataset già visualizzato.

### Cost Avoidance
Vista:
- identica a Saving, con tipo evento `Cost Avoidance`.

Export:
- stesso ramo `_export_vsm_excel()`;
- stessa divergenza (manca Global Search, dataset indipendente dalla griglia).

### Derisking
Vista:
- dispatch `search_requests()` su `_search_derisking_suppliers()` quando tab è `vsm_derisking`;
- carica supplier list con filtro username;
- applica Global Search testuale su campi supplier;
- popola griglia con `_populate_potential_suppliers_sheet()`.

Export:
- `_export_derisking_excel()` usa `load_derisking_suppliers_for_export(username_filter=...)`;
- **non applica Global Search**;
- non usa dataset già visualizzato.

## 6. Causa tecnica della divergenza export vs vista
Causa comune: i tre export VSM usano una pipeline dati separata dalla pipeline vista.

Dettaglio:
- `Saving/Cost Avoidance`: export applica solo subset filtri (username + advanced), ma non Global Search.
- `Derisking`: export applica solo username, ignorando Global Search della griglia.
- in nessun caso export legge il dataset corrente visualizzato nel tab.

## 7. Parti comuni e differenze tra i tre tab
Parti comuni:
- export separato dalla pipeline vista;
- assenza di riuso del dataset visualizzato;
- divergenza funzionale in presenza di filtri globali.

Differenze:
- Saving/CA condividono export `_export_vsm_excel()` con oggetti `VSMEvent`.
- Derisking usa `_export_derisking_excel()` con `PotentialSupplier`.
- Advanced filters sono rilevanti per Saving/CA; Derisking è principalmente supplier + Global Search.

## 8. Strategia di fix consigliata
Strategia unica, minima e reversibile:

1. **Allineamento vista al momento export (tab corrente)**
- In `_export_vsm_excel()` e `_export_derisking_excel()`, chiamare `self.search_requests()` prima di costruire il payload export.
- Questo rende accettabile e coerente il caso “utente ha cambiato filtri ma non ha premuto Search”.

2. **Export da dataset visualizzato (non da ricalcolo DB)**
- Salvare in cache per-sheet il dataset dominio usato per popolare la griglia:
  - in `_populate_vsm_sheet()`: es. `sheet._visible_vsm_events = list(events)`;
  - in `_populate_potential_suppliers_sheet()`: es. `sheet._visible_suppliers = list(suppliers)`.
- In export leggere queste cache dal sheet corrente:
  - Saving/CA: usare `sheet._visible_vsm_events`;
  - Derisking: usare `self.sheet_derisking._visible_suppliers`.

3. **Rimuovere duplicazione di filtro nei rami export VSM**
- Eliminare il ricarico/ri-filtro DB in `_export_vsm_excel()` e `_export_derisking_excel()`.
- Mantenere `services/excel_export_service.py` invariato (accetta già liste prefiltrate).

Questa strategia soddisfa `EXPORT = QUELLO CHE VEDO` con impatto minimo.

## 9. Alternative scartate
1. Duplicare anche Global Search dentro i rami export
- Scartata: aumenta divergenza futura e rischio drift.

2. Leggere direttamente righe `tksheet` per export
- Scartata: i dati in griglia sono formattati/troncati, non adatti come sorgente affidabile per workbook.

3. Refactor esteso in servizi condivisi export/search
- Scartata in questo task: scope troppo ampio rispetto al bug.

## 10. File che probabilmente andranno toccati
Primario:
- `dataflow.py`

Sezione precisa prevista:
- `_populate_vsm_sheet()` (cache dataset visualizzato Saving/CA)
- `_populate_potential_suppliers_sheet()` (cache dataset visualizzato Derisking)
- `_export_vsm_excel()` (riuso cache + riallineamento search)
- `_export_derisking_excel()` (riuso cache + riallineamento search)

Probabile non necessario:
- `services/excel_export_service.py` (già riceve liste oggetti)
- `services/dashboard_controller.py` (dispatch già corretto)

## 11. Rischi reali di regressione
1. Cache dataset non aggiornata in qualche percorso secondario di populate
- Mitigazione: centralizzare assegnazione cache nei due metodi di populate già usati dai flussi principali.

2. Export da tab non corrente/oggetto sheet inatteso
- Mitigazione: mantenere guard su `status` e `sheet` già presenti.

3. Piccolo overhead UI per `search_requests()` invocato all’export
- Accettabile per semantica richiesta e già considerato tollerabile.

## 12. Piano di test manuale post-fix
1. Saving: Global Search attiva
- Applicare query globale che riduce i risultati.
- Export Excel.
- Verificare che file contenga solo righe visibili.

2. Cost Avoidance: Global Search + Advanced Filters
- Impostare filtri avanzati + query globale.
- Export Excel.
- Verificare identità dataset vista/file.

3. Derisking: Global Search supplier
- Applicare query globale su campi supplier.
- Export Excel.
- Verificare che non compaiano supplier esclusi dalla griglia.

4. Tutti e 3 i tab: nessun filtro
- Verificare che export resti equivalente al dataset completo visualizzato.

5. Filtri modificati ma Search non premuto
- Cambiare filtri, lanciare export direttamente.
- Verificare che export rifletta i filtri correnti (vista riallineata in export).

6. Non regressione scope
- RFQ export invariato.
- KPI export invariato.

## 13. Raccomandazione finale
Applicare un fix unico in `dataflow.py` basato su:
- riallineamento esplicito della vista (`search_requests()`) al momento export;
- riuso del dataset visualizzato cache per-sheet;
- eliminazione della ricostruzione dataset nei rami export VSM.

È la soluzione più conservativa, mantiene UX esistente, evita refactor ampio e minimizza il rischio di future divergenze tra vista ed export.
