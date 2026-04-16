# DataFlow - Analisi architetturale preliminare di `dataflow.py`

## 1. Executive summary
`dataflow.py` (4730 righe) non e solo lungo: e un punto di concentrazione di responsabilita eterogenee ad alta criticita (bootstrap, lifecycle Tk, settings operativi, multi-db, VSM/RFQ CRUD, export Excel, restart, backup, orchestration UI).

La criticita principale non e la quantita di codice, ma l'accoppiamento su stato condiviso (`self.*`), ordine di inizializzazione, metadata runtime di `tksheet`, e side effect (config, filesystem, DB, restart).

Esistono segnali positivi di modularizzazione gia avviata (builder/controller/services), ma il main file resta il punto di integrazione di logiche ancora profonde.

Il tag `dataflow.py-refactoring` va considerato come punto di rollback strategico per le fasi successive.

## 2. Stato attuale del file
Macro-struttura reale (con riferimenti di riga):

- Bootstrap e import ordine sensibile: `1-151`
  - DPI awareness Windows, import cross-platform, `init_i18n()` prima degli import UI (`95`), cleanup startup e logging a import-time (`132-136`).
- `SettingsWindow` (`152-1040`)
  - UI settings + logica lingua/valuta/autobackup + backup manuale DB + cambio posizione DataFlow con copia completa e restart.
- `MainWindow` (`1043-4532`)
  - Orchestrazione dashboard, lifecycle app, caricamento RFQ/VSM/Derisking, CRUD, search/filter, ownership, export.
- Esecuzione principale (`4533-4730`)
  - `main_task`, licenza, identita utente, creazione DB, splash, avvio GUI, restart post-mainloop via `_pending_restart`.

## 3. Mappa delle responsabilita presenti
### 3.1 Macro-aree e classificazione (orchestration vs estraibile vs intrecciata)

| Macro-area | Metodi/blocchi principali | Ruolo attuale | Valutazione |
|---|---|---|---|
| Bootstrap + init ordine sensibile | top-level + `init_i18n()` + import UI | bootstrapping app | Orchestration legittima ma temporalmente sensibile |
| Settings UI base | `SettingsWindow.__init__`, `load_settings`, save lingua/valuta | gestione preferenze UI | Estraibile (media) |
| Backup manuale + autobackup config | `backup_database`, `save_autobackup_settings` | operazioni DB/file da settings | Estraibile (media-bassa) |
| Cambio posizione DataFlow | `select_standard_dataflow_location` | workflow lungo: validazioni + copia + update identity/config + restart | Fortemente intrecciata (bassa) |
| Main dashboard orchestration | `MainWindow.__init__`, `build_main_dashboard`, `DashboardController` | avvio vista principale + wiring | Orchestration legittima |
| Lifecycle/restart/timer | `restart_program`, `check_for_autobackup`, `_pending_restart` | lifecycle app e riavvio | Fortemente intrecciata (bassa) |
| Costruzione sheet + bindings | `create_request_treeview`, `_create_vsm_event_sheet`, `_create_supplier_sheet` | setup UI data-grid | Estraibile (media) |
| VSM data loading/filtering/populate | `_get_vsm_dataset`, `_apply_vsm_filters`, `_populate_vsm_sheet` | pipeline dati VSM | Estraibile (media) |
| VSM CRUD + ownership | `_edit_vsm_event`, `_delete_vsm_events`, `_duplicate_vsm_event`, `_delete_supplier` | operazioni utente e sicurezza ownership | Estraibile (media-bassa) |
| RFQ loading/search/update view | `_load_requests_by_status`, `update_treeview`, `search_requests` (via controller) | query + rendering + metadata ownership | Estraibile (media-bassa) |
| RFQ CRUD/open/new | `delete_selected_request`, `duplicate_selected_request`, `open_new_request_window`, `on_sheet_double_click` | operativita core RFQ | Estraibile (media-bassa) |
| Export Excel | `mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel` | reporting/export | Estraibile alta (funzionalmente separabile) |
| Startup applicativo | `main_task` | licensing + identity + DB + splash + bootstrap GUI | Intrecciata (bassa-media) |

### 3.2 Classi, gruppi metodi, helper
- Classe `SettingsWindow`: UI settings + operazioni manutentive (non solo finestra impostazioni).
- Classe `MainWindow`: contiene orchestration legittima e molta logica applicativa (non solo coordinamento).
- Helper interni importanti:
  - Selezione/ownership: `_get_selected_row_indices`, `_check_if_all_*_are_mine`.
  - Pipeline VSM: `_get_vsm_dataset` -> `_apply_vsm_filters` -> `_populate_vsm_sheet`.
  - Dispatch tab: `get_current_tree_and_status`, `_populate_actions_menu`, `update_button_visibility`.

### 3.3 Blocchi gia parzialmente rifattorizzati ma ancora accoppiati
- `build_main_dashboard(app)` (UI estratta) ma dipende da molti metodi/attributi `MainWindow` e ne crea altri a runtime.
- `DashboardController` (search/refresh estratti) ma accede capillarmente a `self.app.*`.
- `services.app_paths` / `services.startup_service` / `database.db_helpers` alleggeriscono il file, ma la regia e il coupling restano in `dataflow.py`.
- Blocchi VSM marcati come estratti da `VSMManagementWindow` ma ancora nel `MainWindow` (Step 4B/4C/4D).

## 4. Accoppiamenti e dipendenze critiche
### 4.1 Accoppiamenti richiesti
- Con `MainWindow`
  - `SettingsWindow` usa `main_app` per `db_manager` e `restart_program`.
  - Dipendenza forte bidirezionale (settings modifica stato/lifecycle della main window).
- Con `SettingsWindow`
  - Non e solo pannello preferenze: contiene workflow operativi ad alto impatto (copy DataFlow, lock DB, config rollback, restart).
- Con `root`
  - Uso massivo (`self.root` e frequenti `wait_window`, `after`, `bind`, `quit`, `destroy`, `mainloop` coupling).
- Con `db_manager`
  - Coesistono `self.db_manager` persistente e molte aperture ad-hoc con `DatabaseManager(get_db_path())`.
  - In alcuni flussi viene chiuso/riaperto manualmente (backup/settings), aumentando sensibilita temporale.
- Con config/restart flow
  - `config.ini` letto/scritto in piu punti (`SettingsWindow`, `main_task`, backup, identity, language, currency).
  - Restart orchestrato via `_pending_restart` e rilancio post-mainloop.
- Con search/filter state
  - Stato distribuito su molte `StringVar`/DateEntry (`search_vars`, `search_tipo`, `date_entries`, `vsm_*_var`).
  - Search RFQ nel controller, search VSM/Derisking in `MainWindow`.
- Con metadata `tksheet`
  - Uso di attributi dinamici su widget: `_sheet_rows_metadata`, `_event_metadata`, `_supplier_metadata`, `_vsm_col_widths`, `_vsm_align_cols`, `_vsm_headers`.
  - Ownership/action logic dipende direttamente da questi metadata runtime.
- Con ownership logic
  - Politiche di permesso nel main file (`_check_if_all_selected_are_mine`, `_check_if_all_vsm_events_are_mine`, `_check_if_all_suppliers_are_mine`).
- Con callback/event binding
  - Bind centrali in builder: notebook change, root click globale, selezione/doppio click per ogni sheet.
  - Effetti a cascata su `update_button_visibility`, clear selection, apertura dialog.
- Con i18n init order
  - Vincolo esplicito: `init_i18n()` prima import UI (`95`), poi richiamato anche in `__main__`.
  - Ordine errato romperebbe `tr(...)` in import-time UI.

### 4.2 Stato condiviso / ordine di esecuzione critico
- Startup: licenza -> identita -> path/config -> `crea_database_v4()` -> splash -> `MainWindow`.
- Main init: build UI -> controller -> preload sheet VSM/Derisking -> refresh RFQ -> timer autobackup.
- Search/action: dipendenza da tab attivo e metadata correnti del sheet.
- Restart: flag `_pending_restart` deve essere impostato prima dell'uscita da mainloop.

### 4.3 Side effect impliciti/non ovvi
- Operazioni filesystem estese in settings (copy intera cartella DataFlow, rename DB, rmtree rollback).
- Chiusura/riapertura connessioni DB durante backup/cambio path.
- Eliminazione file allegati su delete RFQ prima del delete DB.
- `cleanup_temp_on_startup()` e setup logging eseguiti a import-time.

### 4.4 Dipendenze temporali sensibili
- Startup e i18n/import order.
- Lifecycle dialog Tk (`wait_window` diffuso).
- Timer `after` per autobackup e debounce click.
- Restart differito post-mainloop.

## 5. Aree ad alto rischio regressione
Aree piu pericolose (se toccate senza isolamento rigoroso):

- Startup/licensing/identity/DB bootstrap (`main_task`, `__main__`).
- Restart app (`restart_program` + `_pending_restart` + rilancio post-mainloop).
- Workflow cambio posizione DataFlow (`select_standard_dataflow_location`).
- Lifecycle DB durante backup manuale/autobackup/settings.
- Ownership/security gating basato su metadata sheet (RFQ/VSM/Derisking).
- Search/filter multi-db con filtri misti e global search.
- Export Excel RFQ (metodo lungo con molte condizioni e multi-db source).

Rischi per dominio specifico richiesto:
- Startup: molto alto (ordine stretto + side effect config/DB/UI).
- Restart app: molto alto (process lifecycle + Tk teardown).
- Tkinter lifecycle: alto (wait_window/after/bind diffusi).
- Database lifecycle: alto (mix connessioni persistenti e contestuali).
- Multi-user/multi-db: alto (aggregazione + `source_file` + ownership).
- Export Excel: medio-alto (business/reporting + molte trasformazioni).
- RFQ core: alto (CRUD + allegati + ownership + open/edit path).
- VSM core: medio-alto (pipeline dati, dialog, ownership, tab split).
- Search/filter logic: alto (rami locali/aggregati + filtri avanzati + global search).
- Linux/Windows compatibility: medio-alto (DPI, zoom, path, lock file, restart process flags).

## 6. Aree con buona estraibilita
Candidati con estraibilita alta (senza piano esecutivo, solo diagnosi):

1. Export layer (`mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel`)
- Estraibilita: alta
- Motivo: logica lunga ma funzionalmente coesa e relativamente separabile da lifecycle Tk (UI limitata a prompt lingua + save dialog + esito).
- Dipendenze da sciogliere: accesso a filtri attivi, `tr`, `get_currency_excel_number_format`, helper date.
- Rischio regressione atteso: medio (alto impatto utenti, ma superficie confinabile).
- Precondizioni plausibili: contratto dati input esplicito (request_data/event list), adapter dialog/file-save, test snapshot output.

2. Metadata/ownership validators
- Estraibilita: alta-media
- Motivo: helper gia separati concettualmente.
- Dipendenze: formato metadata sheet.
- Rischio: medio (security behavior).
- Precondizioni: standard unico per metadata row-level.

3. Sheet factory utilities
- Estraibilita: alta-media
- Motivo: i builder sheet sono gia blocchi distinti.
- Dipendenze: callback `MainWindow` e convenzioni metadata.
- Rischio: medio.
- Precondizioni: interfaccia callback stabile.

## 7. Aree con estraibilita media o bassa
### 7.1 Estraibilita media
1. VSM data pipeline (`_get_vsm_dataset`, `_apply_vsm_filters`, `_populate_vsm_sheet`)
- Motivo: pipeline chiara ma legata a UI vars, metadata tksheet e formatting.
- Dipendenze: `vsm_*_var`, `sheet._*`, `current_username`, currency.
- Rischio: medio-alto.
- Precondizioni: separare domain filtering da rendering sheet.

2. RFQ load/search/update view (`_load_requests_by_status`, `update_treeview`, controller search)
- Motivo: gia parzialmente separata via controller, ma con coupling su UI state e tuple aggregate.
- Dipendenze: shape tuple aggregate, filter vars, metadata ownership.
- Rischio: medio-alto.
- Precondizioni: formalizzare DTO/record shape e adapter view.

3. VSM CRUD handlers
- Motivo: logica business + dialog + ownership mischiata.
- Dipendenze: status tab, metadata sheet, servizi persistence, wait_window.
- Rischio: medio-alto.
- Precondizioni: separare command handlers da apertura dialog.

### 7.2 Estraibilita bassa
1. `select_standard_dataflow_location`
- Motivo: unico metodo con UX, validazioni, conflitti username, copy ricorsiva, update DB/config, rollback, restart.
- Dipendenze: `main_app.db_manager`, identity, filesystem, config, restart.
- Rischio: molto alto.
- Precondizioni: decomposizione preventiva in step atomici e idempotenti.

2. `restart_program` + `_pending_restart` + blocco post-mainloop
- Motivo: coupling intrinseco con lifecycle processo/Tk.
- Dipendenze: root teardown, cmd/cwd resolution, platform flags.
- Rischio: molto alto.
- Precondizioni: contratto unico restart manager + test su Linux/Windows.

3. `main_task` startup bootstrap
- Motivo: orchestrazione temporale critica con branching abort paths.
- Dipendenze: licenza, identita, config, DB create, splash, `MainWindow`.
- Rischio: molto alto.
- Precondizioni: separare fasi con stato esplicito e rollback deterministici.

## 8. False friends architetturali
Aree apparentemente facili da spostare ma in realta rischiose:

1. `update_button_visibility` sembra UI pura
- In realta incorpora security/ownership e policy azioni cross-tab.

2. `_populate_vsm_sheet` sembra rendering
- In realta integra logica business (variance, currency, metadata ownership/source).

3. `delete_selected_request` sembra CRUD semplice
- In realta include delete fisico allegati + multi-db implications + ownership checks.

4. `perform_autobackup` sembra utility tecnica
- In realta impatta lifecycle DB/file lock cross-platform e retention policy.

5. `search_requests` (controller) sembra gia isolata
- In realta usa molti dettagli `app.*` e policy multi-db/local fallback.

## 9. Considerazioni preliminari per un futuro piano di refactoring
Indicazioni preliminari (non piano esecutivo):

- Mantenere in `dataflow.py` la regia ad alto livello (entrypoint, wiring top-level), ma ridurre logiche operative profonde nel main file.
- Priorita naturale di candidati futuri: export layer -> policy ownership/metadata -> pipeline VSM/RFQ search/load.
- Trattare come blocchi ad alta cautela: startup, restart, cambio path DataFlow, backup lifecycle.
- Preservare rigidamente:
  - ordine i18n/import,
  - compatibilita Linux/Windows,
  - behavior multi-user/multi-db,
  - policy ownership,
  - dialog standardizzati e `tr(...)`.

## 10. Conclusione operativa
Diagnosi: `dataflow.py` e un orchestratore con sovrapposizione di responsabilita applicative critiche. Alcune aree sono maturate verso la modularizzazione, ma molte dipendenze restano basate su stato condiviso e ordine di esecuzione.

Per la fase successiva, il rischio principale non sara "spostare codice", ma preservare invarianti di lifecycle (Tk/process/DB/config) e di sicurezza funzionale (ownership, multi-db scope, filtri).

Riferimento di sicurezza disponibile per le prossime fasi: tag `dataflow.py-refactoring` (solo come ancora di rollback strategico, non usato operativamente in questa sessione).
