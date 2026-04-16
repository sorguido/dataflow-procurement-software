# DataFlow – Execution Report refactoring `dataflow.py`

## 1. Executive summary
Refactor eseguito in una singola run, seguendo l'ordine del piano e applicando micro-step conservativi sulle aree ad alto rischio.

Risultato: `dataflow.py` è alleggerito e più orchestrativo; la logica operativa è stata spostata in servizi dedicati senza introdurre dipendenze nuove e senza modifiche intenzionali a UX/layout/business logic.

Gate eseguiti dopo ogni fase significativa:
- `python3 -m compileall dataflow.py services ui utils`
- `python3 -m unittest -q tests.test_vsm_event_model tests.test_vsm_engine tests.test_vsm_persistence tests.test_supplier_category_persistence`

Nota test: `pytest` non disponibile nell'ambiente; usato `unittest` equivalente sui test richiesti.

## 2. Fasi eseguite
1. Baseline e safety gates: eseguita.
2. Estrazione export Excel (U10): eseguita.
3. Estrazione selection/ownership policy (U1): eseguita.
4. Estrazione actions policy (U2): eseguita.
5. Estrazione sheet factories (U3): eseguita.
6. Estrazione VSM data pipeline + derisking presentation (U4-U5): eseguita in micro-step conservativi.
7. Estrazione VSM command handlers (U6): eseguita in micro-step conservativi.
8. Estrazione RFQ load/render + RFQ command handlers (U7-U8): eseguita in micro-step conservativi.
9. Consolidamento search/filter orchestration (U9): eseguito conservativamente con freeze semantico.
10. Estrazione settings preferences (U11): eseguita.
11. Estrazione backup/migrazione DataFlow folder (U12): eseguita in forma prudente/parziale.
12. Hardening startup/restart e cleanup kernel: eseguito in forma minima conservativa.

## 3. Fasi eseguite parzialmente
- Fase 6 (U4-U5):
  - estratti dataset/filter VSM e presentation Derisking;
  - mantenuto `_populate_vsm_sheet` nel kernel per ridurre rischio su metadata/allineamento tksheet.
- Fase 7 (U6):
  - estratti core command VSM (status mapping, delete, duplicate core);
  - mantenuti nel kernel i wrapper UI/dialog/debounce e parte di orchestrazione.
- Fase 8 (U7-U8):
  - estratti caricamento/filter RFQ e payload render + command core RFQ;
  - mantenuti nel kernel `on_sheet_double_click` e parte open/edit RFQ per preservare lifecycle/read-only flow.
- Fase 11 (U12):
  - estratti core backup manuale/autobackup + helper validazione location;
  - mantenuta nel kernel la regia completa di `select_standard_dataflow_location`.
- Fase 12:
  - estratti helper restart command/spawn;
  - mantenuta nel kernel la regia startup/restart (`main_task`, `__main__`, `_pending_restart`).

## 4. Fasi non eseguite e motivo
Nessuna fase completamente saltata.

Sotto-parti lasciate nel kernel per scelta conservativa (rischio alto, vincoli B1/B4/B5):
- parte render VSM e orchestrazione UI sensibile a metadata;
- parte open/edit RFQ e lifecycle dialog;
- regia principale migrazione DataFlow folder;
- regia principale startup/restart.

## 5. File creati
- `services/excel_export_service.py`
- `services/dashboard_selection_policy.py`
- `services/dashboard_actions_policy.py`
- `ui/sheet_factories.py`
- `services/vsm_dashboard_service.py`
- `services/derisking_dashboard_service.py`
- `services/vsm_command_service.py`
- `services/rfq_command_service.py`
- `services/rfq_dashboard_service.py`
- `services/dashboard_search_service.py`
- `services/settings_preferences_service.py`
- `services/settings_maintenance_service.py`
- `services/dataflow_location_service.py`
- `services/restart_lifecycle_service.py`

## 6. File modificati
- `dataflow.py`
- `services/rfq_command_service.py` (esteso durante la fase RFQ)

## 7. Responsabilità effettivamente uscite da `dataflow.py`
- Export Excel RFQ/VSM/Derisking (core build/save).
- Selection/ownership policy condivisa dashboard.
- Actions capability/menu policy.
- Factory di costruzione sheet RFQ/VSM/Derisking e callback binding standard.
- Pipeline dataset/filter VSM (caricamento e filtri avanzati).
- Presentation suppliers Derisking (rows/metadata/auto-size/populate).
- Core command VSM (mapping status, delete events/suppliers, duplicate core).
- RFQ load/filter per status e payload render rows/metadata.
- Core command RFQ (status update, delete con allegati, duplicate full, create shell).
- Helper search/filter (active-filters check, filtering helpers VSM/Derisking).
- Persistenza preferenze settings (lingua/valuta/autobackup).
- Core manutenzione backup manuale/autobackup.
- Helper validazione location migration (normalize/writable/conflict detection).
- Helper lifecycle restart (resolve path, build command, spawn post-mainloop).

## 8. Responsabilità rimaste nel kernel
- Orchestrazione globale MainWindow/SettingsWindow.
- Wiring UI, dialog lifecycle, debounce e refresh post-dialog.
- Regia azioni ad alto coupling con widget/metadata (`_populate_vsm_sheet`, parti command handler UI).
- Semantica search/filter orchestrata dal controller con dispatch per tab.
- Regia completa migrazione DataFlow folder (progress UI, config rollback, copy orchestration).
- Regia completa startup/restart (`main_task`, `__main__`, `_pending_restart`).

## 9. Invarianti verificate
- `init_i18n()` timing/import order: preservato (nessun cambiamento di ordine sensibile).
- Tkinter lifecycle/dialog behavior: preservati; nessun passaggio a `tkinter.messagebox`.
- Restart lifecycle: preservata la regia nel kernel; estratti solo helper command/spawn.
- Database lifecycle/schema: nessuna modifica schema DB; mantenuto uso `DatabaseManager` con context manager.
- Ownership/permissions semantics: preservate via policy centralizzata + guard nei command wrapper.
- Metadata `tksheet` semantics: preservate (`_sheet_rows_metadata`, `_event_metadata`, `_supplier_metadata`).
- Multi-user/multi-db behavior: preservato nei servizi aggregati e nei filtri username.
- Search/filter semantics freeze:
  - global search OR;
  - advanced filters AND;
  - differenze RFQ/VSM/Derisking mantenute.
- Export behavior: mantenuto output equivalente con estrazione del solo core operativo.
- Compatibilità Linux/Windows: mantenuta (path handling, restart spawn differenziato).
- Nessuna dipendenza nuova: confermato.
- UX/layout invariati: nessuna modifica intenzionale.

## 10. Gate/test eseguiti
Eseguiti ripetutamente dopo ogni fase significativa:
- `python3 -m compileall dataflow.py services ui utils` (verde)
- `python3 -m unittest -q tests.test_vsm_event_model tests.test_vsm_engine tests.test_vsm_persistence tests.test_supplier_category_persistence` (verde)

Risultato ultimo gate:
- test eseguiti: 63
- esito: `OK`
- presenti log di errore attesi dai test negativi VSM (non fallimenti di suite).

## 11. Rischi residui
- Area VSM render/metadata ancora sensibile (`_populate_vsm_sheet` rimasto nel kernel).
- Flusso migrazione DataFlow folder resta complesso e ad alto impatto file/identity/config.
- Startup/restart resta sensibile all'ambiente runtime (script vs packaged executable).
- Coverage automatica non include smoke UI reali (Tkinter lifecycle manuale ancora necessario).

## 12. Punti da verificare manualmente
1. RFQ: load tab attive/archiviate, status change, delete (con allegati), duplicate, open double-click own/other.
2. VSM/Derisking: edit/delete/duplicate, ownership gating, metadata allineati dopo search/filter/refresh.
3. Search/filter matrix:
   - global-only;
   - advanced-only;
   - combinati;
   - scope utente all/current/other;
   - RFQ vs VSM vs Derisking.
4. Settings:
   - save lingua/valuta + restart prompt;
   - save autobackup path;
   - backup manuale con DB aperto.
5. Migrazione DataFlow folder:
   - caso senza conflitto username;
   - caso con conflitto e cambio username;
   - rollback config su errore copia.
6. Restart:
   - restart da settings;
   - rilancio post-mainloop su Linux/Windows.

## 13. Eventuali deviazioni motivate dal piano
- Applicata strategia anti-macro-estrazione (B1): U4/U6/U8/U12 eseguite a micro-step.
- Applicata strategia anti-facciata (B2): nessun passaggio di `self`/`MainWindow`/`SettingsWindow` completo ai servizi.
- Applicato freeze search/filter (B3): nessuna reinterpretazione logica.
- Applicata prudenza startup/restart (B4): regia mantenuta nel kernel, estratti solo helper di supporto.
- Applicato criterio kernel leggibile (B5): mantenuti wrapper orchestratori espliciti nel kernel.

## 14. Conclusione operativa
Refactor completato nel massimo perimetro sicuro della run corrente.

`dataflow.py` risulta sensibilmente alleggerito e più orientato a orchestration/wiring, senza trasformarlo in dispatcher opaco. Le aree ad alto rischio sono state trattate con estrazioni conservative e wrapper kernel, mantenendo le invarianti funzionali critiche.

Tag strategico di rollback considerato: `dataflow.py-refactoring` (non utilizzato operativamente in questa run).
