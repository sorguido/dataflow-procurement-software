# Audit Naming File Output — Export Excel DataFlow Procurement Software

## 1. Titolo
Audit tecnico completo sulle logiche correnti di costruzione del nome file per tutti gli export Excel.

## 2. Sintesi esecutiva
- Il naming degli export Excel non è centralizzato: è distribuito tra più moduli UI/service.
- Esistono pattern con granularità temporale insufficiente (`solo data` o `minuti`), con rischio concreto di sovrascrittura in export ravvicinati.
- Alcuni flussi includono lingua nel nome, altri no; area/tab/sezione non sono codificati in modo uniforme.
- Esiste almeno un caso di naming identico tra flussi diversi: `SQDC_Analysis_RfQ_{id}.xlsx` è usato sia per export utente (EN) sia come nome visualizzato del documento interno SQDC.
- Il fix futuro può restare conservativo: introdurre un builder centralizzato dei filename, mantenendo la logica attuale degli export e cambiando solo la composizione del nome.

## 3. Inventario completo degli export Excel

| # | Export | Area funzionale | Tab/Sezione | Entry point | File/Funzione |
|---|---|---|---|---|---|
| 1 | Export RFQ aggregato (multi-RfQ) | Dashboard principale | RFQ (attive/archiviate) | Pulsante `📥 Export Excel` (global export) | `dataflow.py:3039` -> `services/excel_export_service.py:31` |
| 2 | Export VSM Saving | Dashboard principale | VSM Saving | Pulsante `📥 Export Excel` su tab VSM | `dataflow.py:3124` -> `services/excel_export_service.py:323` |
| 3 | Export VSM Cost Avoidance | Dashboard principale | VSM Cost Avoidance | Pulsante `📥 Export Excel` su tab VSM | `dataflow.py:3124` -> `services/excel_export_service.py:323` |
| 4 | Export Derisking fornitori | Dashboard principale | VSM Derisking | Pulsante `📥 Export Excel` su tab Derisking | `dataflow.py:3153` -> `services/excel_export_service.py:452` |
| 5 | Export riepilogo singola RdO/RfQ | Finestra dettaglio richiesta | Menu export della Request | `📗 Excel` | `ui/windows/view_request_window.py:645` |
| 6 | Export SQDC su file utente | Finestra SQDC Analysis | SQDC | Pulsante `📊 Export Excel` | `ui/windows/sqdc_analysis_window.py:550` |
| 7 | Salvataggio SQDC come Documento Interno (xlsx) | Finestra SQDC Analysis + allegati | SQDC / Documento Interno | Pulsante `💾 Save SQDC` | `ui/windows/sqdc_analysis_window.py:669` + `database_manager.py:765` |
| 8 | Export KPI | Finestra KPI Analysis | RFQ/Saving/Cost Avoidance/Derisking (scope current/all) | Pulsante `📥 Export Excel` | `ui/kpi_window.py:668` |

## 4. Mappa file/funzioni che costruiscono il nome output

| Export | Nome/pattern attuale | Lingua coinvolta nel filename | Timestamp nel filename | Save dialog con default filename |
|---|---|---|---|---|
| RFQ aggregato dashboard | `Export_DataFlow_{YYYYMMDD}.xlsx` (`services/excel_export_service.py:307`) | No (nome uguale IT/EN) | Sì, solo data | Sì (`services/excel_export_service.py:308`) |
| VSM Saving | `Export_VSM_Saving_{YYYYMMDD}.xlsx` (`services/excel_export_service.py:435`) | No marker lingua; `Saving` resta EN | Sì, solo data | Sì (`services/excel_export_service.py:437`) |
| VSM Cost Avoidance | `Export_VSM_Cost_Avoidance_{YYYYMMDD}.xlsx` (`services/excel_export_service.py:435`) | No marker lingua; `Cost_Avoidance` EN | Sì, solo data | Sì (`services/excel_export_service.py:437`) |
| Derisking dashboard | `Export_Derisking_{YYYYMMDD}.xlsx` (`services/excel_export_service.py:518`) | No marker lingua; label EN | Sì, solo data | Sì (`services/excel_export_service.py:520`) |
| Riepilogo singola richiesta IT | `Riepilogo_RdO_{request_id}.xlsx` (`ui/windows/view_request_window.py:688`) | Sì (IT) | No | Sì (`ui/windows/view_request_window.py:773`) |
| Riepilogo singola richiesta EN | `Summary_RfQ_{request_id}.xlsx` (`ui/windows/view_request_window.py:694`) | Sì (EN) | No | Sì (`ui/windows/view_request_window.py:773`) |
| SQDC export IT | `SQDC_Analisi_RdO_{request_id}.xlsx` (`ui/windows/sqdc_analysis_window.py:559`) | Sì (IT) | No | Sì (`ui/windows/sqdc_analysis_window.py:645`) |
| SQDC export EN | `SQDC_Analysis_RfQ_{request_id}.xlsx` (`ui/windows/sqdc_analysis_window.py:562`) | Sì (EN) | No | Sì (`ui/windows/sqdc_analysis_window.py:645`) |
| SQDC Documento Interno (display) | `SQDC_Analysis_RfQ_{request_id}.xlsx` (`ui/windows/sqdc_analysis_window.py:676`) | No (sempre EN) | No | No (save automatico) |
| SQDC Documento Interno (file fisico) | `RfQ{request_id}_SQDC_ID{next_id}.xlsx` (`ui/windows/sqdc_analysis_window.py:775`) | No | Pseudo-univoco via ID DB | No (save automatico) |
| KPI export | `KPI_DataFlow_{YYYYMMDD_HHMM}.xlsx` (`ui/kpi_window.py:729`) | No (uguale IT/EN) | Sì, minuti | Sì (`ui/kpi_window.py:730`) |

Note tecniche SQDC correlate:
- Lookup del documento SQDC su nome fisso EN in `ui/windows/view_request_window.py:389` e `:410`.
- Upsert su allegato SQDC con query `nome_file LIKE 'SQDC_%'` in `database_manager.py:777`.

## 5. Collisioni o ambiguità individuate

1. Collisione forte dashboard RFQ: `Export_DataFlow_{YYYYMMDD}.xlsx`
- Stesso nome per export multipli nello stesso giorno, anche con filtri/tab diversi.

2. Collisione forte VSM (stessa tab stesso giorno)
- `Export_VSM_Saving_{YYYYMMDD}.xlsx` e `Export_VSM_Cost_Avoidance_{YYYYMMDD}.xlsx` cambiano solo per tipo evento, non per ora/esecuzione.

3. Collisione forte Derisking
- `Export_Derisking_{YYYYMMDD}.xlsx` sovrascrivibile su export ripetuti nello stesso giorno.

4. Collisione forte KPI (stesso minuto)
- `KPI_DataFlow_{YYYYMMDD_HHMM}.xlsx` collide se due export avvengono nello stesso minuto (anche da sezioni/scope/lingue diverse).

5. Collisione forte su singola richiesta (RdO summary)
- `Riepilogo_RdO_{id}.xlsx` / `Summary_RfQ_{id}.xlsx` senza timestamp: export ripetuti della stessa richiesta tendono a proporre identico filename.

6. Naming identico tra flussi diversi su SQDC EN
- `SQDC_Analysis_RfQ_{id}.xlsx` usato sia in export utente EN che come `nome_file` del Documento Interno SQDC.
- Non è collisione fisica automatica (percorso interno usa `ID`), ma è ambiguità UX e semantica.

7. Ambiguità lingua
- Diversi export non riportano lingua nel nome (`Export_DataFlow`, `Export_VSM_*`, `Export_Derisking`, `KPI_DataFlow`).

## 6. Valutazione rischio sovrascrittura per ciascun export

| Export | Timestamp | Distintività area/tab/lingua | Rischio collisione | Note UX pratiche |
|---|---|---|---|---|
| RFQ aggregato dashboard | Giorno | Bassa | **Alto** | File identico in export multipli nella stessa giornata |
| VSM Saving | Giorno | Media (tipo evento nel nome) | **Alto** | Ripetizioni nello stesso giorno facilmente sovrascrivibili |
| VSM Cost Avoidance | Giorno | Media | **Alto** | Stesso rischio del Saving |
| Derisking dashboard | Giorno | Media-bassa | **Alto** | Nome corto e poco descrittivo, collisione giornaliera |
| Riepilogo richiesta IT/EN | Nessuno | Media (ID richiesta presente) | **Medio-Alto** | Per la stessa richiesta il nome resta uguale nel tempo |
| SQDC export utente IT/EN | Nessuno | Media (ID presente + lingua IT/EN) | **Medio-Alto** | Re-export stessa richiesta/language sovrascrivibile |
| SQDC Documento Interno (display) | Nessuno | Media | **Medio** | Upsert intenzionale; perdita di versioning percepito |
| SQDC Documento Interno (file fisico) | ID progressivo allegato | Alta | **Basso** | Nome fisico abbastanza unico, ma dipende da `MAX(id)+1` |
| KPI export | Minuti | Bassa (niente tab/scope/lang) | **Alto** | Due export nello stesso minuto possono coincidere |

## 7. Presenza o assenza di centralizzazione
- Centralizzazione **parziale**: i tre export dashboard (RFQ/VSM/Derisking) stanno in `services/excel_export_service.py`.
- Per il resto il naming è **distribuito e duplicato**:
  - `ui/kpi_window.py` per KPI.
  - `ui/windows/view_request_window.py` per export dettaglio richiesta.
  - `ui/windows/sqdc_analysis_window.py` per SQDC export e SQDC internal save.
  - `ui/windows/view_request_window.py` + `database_manager.py` dipendono da stringhe SQDC hardcoded.
- Conclusione: non esiste oggi un builder unico di filename Excel riusato da tutti i flussi.

## 8. Strategia consigliata per uniformare i nomi in futuro
Strategia minima e conservativa (senza redesign):

1. Introdurre un unico builder (es. `utils/export_filename.py`) per i filename Excel.
2. Mantenere i flussi export invariati, sostituendo solo la costruzione del `default_name`.
3. Convenzione proposta:

`DataFlow_[Area]_[TabOrSection]_[ExportType]_[Lang]_[Context]_[YYYY-MM-DD]_[HH-MM-SS-fff].xlsx`

Dove:
- `Area`: `Dashboard`, `Request`, `SQDC`, `KPI`.
- `TabOrSection`: es. `RFQ`, `VSM_Saving`, `VSM_CostAvoidance`, `VSM_Derisking`, `CurrentTab`, `AllTabs`.
- `ExportType`: es. `Summary`, `Analysis`, `Global`, `InternalDoc`.
- `Lang`: `IT` / `EN`.
- `Context`: opzionale ma consigliato, es. `RfQ123`.

Esempi concreti:
- `DataFlow_Dashboard_RFQ_Global_IT_2026-04-18_10-35-22-417.xlsx`
- `DataFlow_Request_RfQ123_Summary_EN_2026-04-18_10-35-25-093.xlsx`
- `DataFlow_KPI_CurrentTab_Derisking_IT_2026-04-18_10-35-29-004.xlsx`

Vincoli cross-platform rispettati:
- solo lettere/numeri/underscore/trattini;
- niente `:` o caratteri non validi Windows;
- estensione `.xlsx` invariata.

## 9. Opzioni di timestamp consigliate

Opzione A — secondi (`YYYY-MM-DD_HH-MM-SS`)
- Pro: semplice, leggibile.
- Contro: non evita collisioni se due export avvengono nello stesso secondo.

Opzione B — millisecondi (`YYYY-MM-DD_HH-MM-SS-fff`) **consigliata**
- Pro: rischio collisione molto basso anche per export ravvicinati.
- Contro: nome leggermente più lungo.

Opzione C — secondi + suffix progressivo se file già esiste (`_01`, `_02`)
- Pro: robusta anche senza millisecondi.
- Contro: richiede controllo filesystem aggiuntivo.

Valutazione esplicita richiesta:
- Per il caso “secondo export 2 secondi dopo” i secondi sono sufficienti.
- Per sicurezza reale (doppio click, export multipli quasi simultanei, più finestre), è preferibile includere anche i millisecondi o un suffix anti-collisione.

## 10. File che probabilmente andrebbero toccati in un futuro fix
- `services/excel_export_service.py` (3 builder di default filename)
- `ui/kpi_window.py` (default filename KPI)
- `ui/windows/view_request_window.py` (default filename summary richiesta; eventuale lookup SQDC se cambia convenzione)
- `ui/windows/sqdc_analysis_window.py` (default filename export SQDC + naming SQDC Documento Interno)
- `database_manager.py` (logica `insert_or_update_allegato_sqdc` e query correlate, se si decide di cambiare naming SQDC interno)
- Nuovo modulo consigliato: `utils/export_filename.py` (helper centralizzato)

## 11. Rischi di regressione di un eventuale intervento
- Rischio funzionale SQDC: il lookup usa oggi un nome hardcoded (`SQDC_Analysis_RfQ_{id}.xlsx`) e query `LIKE 'SQDC_%'`; modifiche non coordinate possono rompere apertura/riconoscimento SQDC esistente.
- Rischio UX/documentazione: utenti e manuali potrebbero aspettarsi nomi storici.
- Rischio test/manual QA: i flussi sono multi-finestra (Dashboard, Request, SQDC, KPI), serve smoke test trasversale per IT/EN.
- Rischio basso sul core business: il cambio è confinato al naming e non richiede alterare calcoli/export content.

## 12. Raccomandazione finale
- Stato attuale: naming eterogeneo, in parte descrittivo ma non uniforme e spesso non univoco nel tempo.
- Priorità minima consigliata: centralizzare la sola generazione filename e introdurre timestamp con millisecondi.
- Approccio conservativo: nessun refactor pesante, nessun cambio logica dati/export; solo sostituzione dei punti di costruzione nome.
- Complessità attesa del futuro fix: **Media** (più file coinvolti, ma impatto tecnico circoscritto).
