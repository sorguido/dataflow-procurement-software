# DataFlow - Piano completo di refactoring di `dataflow.py`

## 1. Executive summary
Questo piano parte dalla diagnosi gia prodotta in `development/report/dataflow_py_refactoring_analysis.md` e la trasforma in una sequenza operativa completa, pensata per una futura esecuzione autonoma end-to-end.

Obiettivo pratico: ridurre il carico operativo nel kernel `dataflow.py` senza svuotarlo artificialmente. Il file deve restare il cervello applicativo (orchestration, wiring, invarianti globali), mentre le logiche operative specialistiche devono migrare in moduli dedicati quando il rischio e gestibile.

Direzione tecnica:
- Refactor per unita reali (non per etichette astratte).
- Ordine dal rischio piu basso al rischio piu alto.
- Presidio esplicito di lifecycle (Tk, DB, restart), ownership, multi-db, search/filter, export, i18n/import order.
- Reversibilita locale per fase e rollback strategico via tag `dataflow.py-refactoring` (solo riferimento, non uso operativo qui).

Assunzione esplicita: nel workspace e disponibile `Linee guida per AI (Vibe Coding).md`; il piano applica integralmente quei principi (conservativita, no regressioni, no nuove dipendenze, UX invariata, i18n `tr(...)`, dialoghi standardizzati).

## 2. Ruolo finale atteso di `dataflow.py`
Ruolo finale: kernel leggibile di coordinamento applicativo, non monolite operativo.

B1. Cosa deve rimanere sicuramente dentro `dataflow.py`:
- Bootstrap top-level e ordine sensibile import/i18n (blocco iniziale, `init_i18n()` prima degli import UI).
- Entry point `__main__`, creazione `root`, avvio `main_task`, `mainloop`, gestione `_pending_restart` post-mainloop.
- Orchestration ad alto livello di `MainWindow.__init__` (sequenza: build UI -> init controller/services -> first load -> first refresh -> timer).
- Routing di alto livello tab-aware (`get_current_tree_and_status`) e dispatch verso servizi.
- Wiring callback/binding globali (restano nel kernel, ma delegano la logica operativa).
- Gestione invarianti globali (ownership semantics, lifecycle rules, multi-db policy) come policy dichiarata e centralizzata.

B2. Cosa deve uscire sicuramente da `dataflow.py`:
- Logica di export Excel (`mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel`).
- Data transformation VSM/RFQ non strettamente di orchestration.
- Blocchi operativi CRUD RFQ/VSM/Derisking che oggi sono metodi lunghi in `MainWindow`.
- Logiche di backup operativo e copia filesystem profonde (manual backup, autobackup, migrazione DataFlow folder).
- Costruzione dettagliata e policy operative dei `tksheet` (factory + metadata handling).

B3. Cosa puo restare temporaneamente dentro `dataflow.py`:
- `select_standard_dataflow_location` come wrapper orchestratore, con estrazione progressiva della logica interna.
- `restart_program` come wrapper orchestratore, con eventuale service di supporto per command resolution e post-quit restart.
- Parti di startup `main_task` in forma di orchestrazione finche i servizi di bootstrap non sono stabili.

B4. Ruolo finale del file a refactor concluso:
- Kernel di coordinamento esplicito, con metodi corti e leggibili, composti da:
  - precondizioni,
  - dispatch,
  - chiamata a servizio,
  - gestione errore/dialogo,
  - aggiornamento UI finale.

B5. Responsabilita compatibili con kernel/cervello:
- Sequenziamento lifecycle.
- Decisioni di routing e policy globali.
- Gestione eventi e binding ad alto livello.
- Contratti tra UI e servizi.

B6. Responsabilita incompatibili col kernel/cervello:
- Algoritmi operativi lunghi di manipolazione dati/file.
- Formattazione/export report complessi.
- Copy/move/rollback filesystem dettagliati.
- Query/filtering multi-step intrecciati con UI rendering in metodi monolitici.

## 3. Responsabilita che devono restare nel kernel
Elenco concreto su codice reale:

- Bootstrap sensibile:
  - blocco DPI awareness iniziale,
  - `init_i18n()` pre-import UI,
  - setup root/Tk nel main.

- Coordinamento `MainWindow`:
  - `__init__` come orchestratore (non come contenitore di logica operativa).
  - apertura finestre di alto livello (`open_help_window`, `open_settings_window`, `on_kpi_click`) come wrapper.

- Routing e controllo contesto UI:
  - `get_current_tree_and_status`.
  - `on_tab_changed` come entrypoint di orchestrazione update.
  - binding globali root/notebook.

- Presidio invarianti:
  - ownership policy dichiarativa (puo chiamare helper/servizi ma resta responsabilita del kernel farla rispettare).
  - guard lifecycle restart/startup.

- Entry point applicativo:
  - blocco `if __name__ == '__main__':` e gestione `_pending_restart` post-mainloop.

## 4. Responsabilita che devono uscire dal kernel
Elenco concreto su metodi reali:

- Export:
  - `mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel`.

- VSM data pipeline:
  - `_get_vsm_dataset`, `_apply_vsm_filters`, `_populate_vsm_sheet`, `_search_vsm_events`, `_search_derisking_suppliers`.

- VSM operations:
  - `_edit_vsm_event`, `_delete_vsm_events`, `_delete_supplier`, `_duplicate_vsm_event`, `_on_vsm_sheet_double_click`, `_on_supplier_sheet_double_click`.

- RFQ operations:
  - `_load_requests_by_status`, `update_treeview`, `delete_selected_request`, `duplicate_selected_request`, `_change_request_status`, `on_sheet_double_click`, `open_new_request_window`.

- Sheet construction and metadata policy:
  - `create_request_treeview`, `create_cell_select_handler`, `create_row_select_handler`, `_create_vsm_event_sheet`, `_create_supplier_sheet`, `_auto_size_supplier_sheet`, `_populate_potential_suppliers_sheet`.

- Settings operative profonde:
  - `backup_database`, `check_for_autobackup`, `perform_autobackup`, `select_standard_dataflow_location` (in modalita progressiva, con wrapper in kernel).

## 5. Unita reali di estrazione
Unita concrete (metodi che devono muoversi insieme), con target suggeriti per futura run.

| Blocco | Metodi coinvolti | Responsabilita concreta | File target suggerito | Motivo accorpamento | Rischio | Benefici |
|---|---|---|---|---|---|---|
| U1 - Selection and Ownership Policy | `_get_selected_row_indices`, `_check_if_all_selected_are_mine`, `_check_if_all_vsm_events_are_mine`, `_check_if_all_suppliers_are_mine` | Semantica selezione e permessi ownership | `services/dashboard_selection_policy.py` | Policy unica, usata da RFQ e VSM | Medio | Riduce duplicazione e bug di permesso |
| U2 - Actions State Policy | `update_button_visibility`, `_populate_actions_menu` (con `get_current_tree_and_status` mantenuto in kernel) | Enable/disable Actions e menu contestuale tab-aware | `services/dashboard_actions_policy.py` | Coupling forte tra stato selezione e menu | Medio | Logica azioni centralizzata e testabile |
| U3 - Sheet Factory RFQ/VSM | `create_request_treeview`, `create_cell_select_handler`, `create_row_select_handler`, `_create_vsm_event_sheet`, `_create_supplier_sheet` | Creazione/configurazione `tksheet` + binding standard | `ui/sheet_factories.py` | Blocchi omogenei di costruzione UI | Medio | Kernel piu corto e coerenza tksheet |
| U4 - VSM Data Pipeline | `_get_vsm_dataset`, `_apply_vsm_filters`, `_populate_vsm_sheet`, `_load_vsm_events`, `_load_potential_suppliers`, `populate_vsm_username_filter`, `_on_vsm_username_filter_changed` | Carico dataset, filtro, render e metadata VSM/Derisking | `services/vsm_dashboard_service.py` | Pipeline coesa ma oggi spezzata su metodi main | Medio-alto | Riduce accoppiamento con `MainWindow` |
| U5 - Derisking Supplier Presentation | `_auto_size_supplier_sheet`, `_populate_potential_suppliers_sheet`, parti supplier di ricerca/carico | Render e metadata supplier-based | `services/derisking_dashboard_service.py` | Semantica diversa da VSM event-based | Medio | Evita regressioni da mixing VSM/supplier |
| U6 - VSM Command Handlers | `_edit_vsm_event`, `_delete_vsm_events`, `_delete_supplier`, `_duplicate_vsm_event`, `_on_vsm_sheet_double_click`, `_on_supplier_sheet_double_click`, branch VSM in `open_new_event` | CRUD e azioni utente VSM/Derisking | `services/vsm_command_service.py` | Metodi da muovere insieme con ownership e dialog | Medio-alto | Chiarezza operativa e minor monolite |
| U7 - RFQ Load and Render | `_load_requests_by_status`, `update_treeview`, `_format_date_for_display` | Carico RFQ e rendering sheet con metadata | `services/rfq_dashboard_service.py` | Query+render+metadata oggi intrecciati | Medio-alto | Isola logica RFQ dal kernel |
| U8 - RFQ Command Handlers | `archive_selected_request`, `reactivate_selected_request`, `_change_request_status`, `delete_selected_request`, `duplicate_selected_request`, `on_sheet_double_click`, `open_new_request_window`, branch RFQ in `open_new_event` | Operazioni RFQ e apertura editor | `services/rfq_command_service.py` | Operativita RFQ coerente in blocco unico | Alto | Riduce complessita del kernel |
| U9 - Search and Filter Orchestration | `_has_active_search_filters`, `search_requests`, `clear_filters`, `refresh_data`, `populate_username_filter`, `_update_filter_panel_for_current_tab`, `_update_advanced_filters_toggle` | Semantica filtri/search multi-tab | Estendere `services/dashboard_controller.py` + helper `services/dashboard_search_service.py` | Gia parzialmente estratto: completare senza frammentare troppo | Alto | Coerenza ricerca e minor split logico |
| U10 - Excel Export Suite | `mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel` | Export RFQ/VSM/Derisking | `services/excel_export_service.py` + `services/excel_export_builders.py` | Alto grado di coesione e basso coupling lifecycle | Medio | Primo alleggerimento consistente e sicuro |
| U11 - Settings Preferences | `load_settings`, `save_language_settings`, `save_currency_settings`, `select_autobackup_path`, `save_autobackup_settings` | Gestione preferenze e persistenza config | `services/settings_preferences_service.py` | Logica config separabile da UI widgeting | Medio | `SettingsWindow` torna orchestratore UI |
| U12 - Backup and DataFlow Migration | `backup_database`, `check_for_autobackup`, `perform_autobackup`, `select_standard_dataflow_location` | Backup/migrazione ad alto impatto su DB e filesystem | `services/settings_maintenance_service.py` + `services/dataflow_location_service.py` | Flussi lunghi, critici, da isolare con gradualita | Molto alto | Riduzione massima del rischio operativo nel kernel |

Metodi che non vanno separati all'interno dei blocchi:
- `_get_vsm_dataset` e `_apply_vsm_filters`: separazione troppo precoce aumenta mismatch su scope utente.
- `update_treeview` e metadata ownership RFQ: devono restare vicini finche il contratto row metadata non e stabilizzato.
- `delete_selected_request` e delete fisico allegati: non separare senza contratto transazionale esplicito.

Adapter temporanei necessari (futura run):
- Wrapper methods in `MainWindow`/`SettingsWindow` che mantengono signature attuale e delegano al servizio.
- `AppContext` minimale (o passaggio esplicito di dipendenze) per evitare moduli che leggono `self` intero.

Dipendenze da sciogliere prima dello spostamento:
- Accesso diretto a `self.search_vars`, `self.date_entries`, `self.*_metadata`.
- Dipendenza implicita da `current tab` tramite `self.notebook`.
- Mix di UI dialog e logica persistence nello stesso metodo.

## 6. Invarianti non negoziabili
Per ogni invariante: cosa non deve cambiare, rischio attuale, protezione nel piano.

| Invariante | Cosa non deve cambiare | Rischio attuale | Protezione nel piano |
|---|---|---|---|
| I1 - `init_i18n` timing e import order | `init_i18n()` deve restare prima degli import UI che usano `tr(...)` | Refactor startup/import puo rompere traduzioni import-time | Fasi startup/restart solo finali; nessuno spostamento di blocchi top-level nelle prime fasi |
| I2 - Tkinter lifecycle | Semantica `wait_window`, `after`, `bind`, `quit/destroy` invariata | Spostare handler senza preservare timing produce race o finestre zombie | Wrapper in kernel + test smoke lifecycle ad ogni fase UI |
| I3 - Restart lifecycle | `_pending_restart` e rilancio post-mainloop invariati | Rifacendo restart troppo presto si rompe auto-restart cross-platform | Restart toccato solo in fase tarda dedicata, con fallback manuale invariato |
| I4 - Database lifecycle | Nessuna corruzione/lock; chiusure e riaperture DB equivalenti | Backup/migrazione e CRUD hanno aperture multiple delicate | Fasi DB ad alto rischio in coda; test lock-sensitive e rollback locale per fase |
| I5 - Ownership semantics | Utente non puo agire su dati altrui (RFQ/VSM/Supplier) | Refactor action/menu/selection puo bypassare controlli | U1/U2 prima di CRUD extraction; check ownership centralizzati e riusati |
| I6 - Metadata `tksheet` semantics | `_sheet_rows_metadata`, `_event_metadata`, `_supplier_metadata` devono restare coerenti con le righe visualizzate | Render/search/refresh possono disallineare metadata e righe | Contratto metadata esplicito introdotto prima di spostare CRUD |
| I7 - Multi-user/multi-db behavior | Aggregazione e `source_file/is_mine` invariati | Search e load attuali alternano percorso locale/aggregato | Search/load consolidati con test su branch locale/aggregato |
| I8 - Search/filter semantics | Global search OR + filtri avanzati AND; scope utente invariato | Alta complessita e duplicazioni tra controller/main | Fase dedicata search con snapshot comportamentali prima/dopo |
| I9 - Export output behavior | Struttura file, formati numerici, selezione dati esportati equivalenti | Export monolitico e fragile a piccole variazioni | Estrarre con golden sample e confronto output su casi reali |
| I10 - Dialog behavior | Uso dialog standard (`SimpleMessageDialog`, `SimpleYesNoDialog`, `show_error`) invariato | Refactor puo introdurre dialog non standard o cambiare flow UX | Vincolo esplicito in ogni fase: nessuna modifica dialog UX |
| I11 - Linux/Windows compatibility | DPI, path handling, restart, file lock semantics invariati | Migrazioni filesystem e restart molto OS-sensitive | Fasi OS-sensitive in coda con smoke matrix Linux/Windows |
| I12 - Nessuna nuova dipendenza | Nessuna libreria nuova | Tentazione di introdurre utility esterne | Policy hard: solo stdlib + moduli gia presenti |
| I13 - UX/layout invariati | Nessun layout shift o cambio visuale funzionale | Estrazione UI factory puo alterare binding/ordini | Sheet/UI extraction solo con parity check visuale tab-by-tab |

## 7. Sequenza completa del refactor
Sequenza globale proposta per futura run autonoma (dall'inizio alla fine):

1. Baseline e safety gates.
2. Estrazione export Excel (U10).
3. Estrazione selection/ownership policy (U1).
4. Estrazione actions policy (U2).
5. Estrazione sheet factories (U3).
6. Estrazione VSM data pipeline + derisking presentation (U4-U5).
7. Estrazione VSM command handlers (U6).
8. Estrazione RFQ load/render + RFQ command handlers (U7-U8).
9. Consolidamento search/filter orchestration (U9).
10. Estrazione settings preferences (U11).
11. Estrazione backup/migrazione DataFlow folder (U12).
12. Hardening startup/restart e final kernel cleanup.

Perche questo ordine:
- Prime fasi su blocchi ad alta estraibilita e basso impatto lifecycle.
- Blocchi cross-cutting (ownership/actions/sheet metadata) prima di CRUD.
- Search e settings dopo stabilizzazione di load/render.
- Startup/restart e migrazione path solo dopo che il resto e stabile.

## 8. Fasi dettagliate con rischi, mitigazioni, test e rollback
Ogni fase include D1-D11.

### Fase 1 - Baseline e safety gates
D1. Obiettivo: fissare baseline comportamentale e tecnica prima di muovere logica.
D2. Area coinvolta: intera app, senza refactor funzionale.
D3. Metodi/classi/helper: nessuno spostamento; inventario metodi `dataflow.py` + contratti metadata.
D4. File target suggerito: nessuno operativo; solo checklist interna run.
D5. Posizione: obbligatoria prima di tutto.
D6. Rischio specifico: basso tecnico, alto se saltata.
D7. Mitigazioni:
- Definire gate minimi automatici e smoke manuali.
- Definire criteri stop-the-line.
D8. Test minimi dopo fase:
- `python -m compileall dataflow.py services ui utils`
- `pytest -q tests/test_vsm_event_model.py tests/test_vsm_engine.py tests/test_vsm_persistence.py tests/test_supplier_category_persistence.py`
- Smoke: avvio app, cambio tab RFQ/VSM/Derisking, apertura/chiusura Settings.
D9. Rollback locale: n/a (fase non invasiva).
D10. Precondizioni: ambiente runtime coerente con branch.
D11. Invarianti: tutte, come baseline.

### Fase 2 - Estrazione export Excel (U10)
D1. Obiettivo: rimuovere dal kernel la logica export piu lunga e coesa.
D2. Area: RFQ/VSM/Derisking export.
D3. Metodi coinvolti: `mega_export_excel`, `_export_vsm_excel`, `_export_derisking_excel` + helper interni locali.
D4. File target: `services/excel_export_service.py`, `services/excel_export_builders.py`.
D5. Posizione: prima estrazione operativa a rischio medio e bassa dipendenza lifecycle.
D6. Rischio: medio (parita output).
D7. Mitigazioni:
- Mantenere wrapper in `MainWindow` con signature invariata.
- Estrarre per copia controllata, poi switch call-site.
- Nessun cambio dialog/save UX.
D8. Test minimi:
- Gate baseline.
- Smoke export RFQ, VSM, Derisking su dataset reale.
- Confronto file output su campioni noti (header, formati, colonne).
D9. Rollback locale:
- Ripristino wrapper con body originale.
- Revert file service export introdotti in fase.
D10. Precondizioni:
- Contratto input export esplicito (`status`, filtri, dataset, sheet metadata).
D11. Invarianti:
- I9, I10, I12, I13.

### Fase 3 - Estrazione selection/ownership policy (U1)
D1. Obiettivo: centralizzare semantica selezione e ownership.
D2. Area: RFQ/VSM/Derisking action safety.
D3. Metodi: `_get_selected_row_indices`, `_check_if_all_selected_are_mine`, `_check_if_all_vsm_events_are_mine`, `_check_if_all_suppliers_are_mine`.
D4. File target: `services/dashboard_selection_policy.py`.
D5. Posizione: prima di azioni CRUD/menu per ridurre rischio autorizzazioni.
D6. Rischio: medio (security regression).
D7. Mitigazioni:
- Funzioni pure con input esplicito (`sheet`, `selected_indices`, `metadata`).
- Default fail-safe invariato (`False` quando metadata non affidabile).
D8. Test minimi:
- Gate baseline.
- Smoke: selezione multipla e blocco azioni su record altrui in RFQ/VSM/Derisking.
D9. Rollback locale:
- Riportare metodi nel kernel e disattivare import service.
D10. Precondizioni:
- Contratto metadata righe formalizzato.
D11. Invarianti:
- I5, I6, I7.

### Fase 4 - Estrazione actions policy (U2)
D1. Obiettivo: separare policy enable/disable Actions da kernel.
D2. Area: menu contestuale e stato pulsante Actions.
D3. Metodi: `update_button_visibility`, `_populate_actions_menu` (con `get_current_tree_and_status` nel kernel).
D4. File target: `services/dashboard_actions_policy.py`.
D5. Posizione: subito dopo Fase 3 per riuso ownership centralizzata.
D6. Rischio: medio.
D7. Mitigazioni:
- Service riceve snapshot stato, non `self` completo.
- Maintained behavior tab-specific invariato.
D8. Test minimi:
- Gate baseline.
- Smoke menu per tutti i tab, casi no selection/selection singola/multipla/altrui.
D9. Rollback locale:
- Ripristino body originale dei due metodi.
D10. Precondizioni:
- Fase 3 stabile.
D11. Invarianti:
- I5, I6, I13.

### Fase 5 - Estrazione sheet factories (U3)
D1. Obiettivo: ridurre peso UI operativo nel kernel mantenendo wiring.
D2. Area: creazione/configurazione `tksheet`.
D3. Metodi: `create_request_treeview`, `create_cell_select_handler`, `create_row_select_handler`, `_create_vsm_event_sheet`, `_create_supplier_sheet`.
D4. File target: `ui/sheet_factories.py`.
D5. Posizione: dopo policy actions/ownership per preservare metadata contract.
D6. Rischio: medio-alto (binding/metadata desync).
D7. Mitigazioni:
- Conservare stessi attributi custom su sheet (`_event_metadata`, `_supplier_metadata`, `_vsm_*`).
- Non cambiare layout, colonne, larghezze, allineamenti.
D8. Test minimi:
- Gate baseline.
- Smoke visuale su tutti i tab: sorting, double click, selezione, redraw.
D9. Rollback locale:
- Revert factory file + restore metodi in `MainWindow`.
D10. Precondizioni:
- Contratto metadata definito e rispettato.
D11. Invarianti:
- I6, I10, I13.

### Fase 6 - Estrazione VSM data pipeline e Derisking presentation (U4-U5)
D1. Obiettivo: estrarre pipeline dataset/filter/populate VSM senza cambiare semantica.
D2. Area: VSM/Derisking load e rendering dati.
D3. Metodi: `populate_vsm_username_filter`, `_on_vsm_username_filter_changed`, `_get_vsm_dataset`, `_apply_vsm_filters`, `_populate_vsm_sheet`, `_load_vsm_events`, `_load_potential_suppliers`, `_auto_size_supplier_sheet`, `_populate_potential_suppliers_sheet`, `_search_vsm_events`, `_search_derisking_suppliers`.
D4. File target: `services/vsm_dashboard_service.py`, `services/derisking_dashboard_service.py`.
D5. Posizione: prima dei command handlers per stabilizzare stato dati e metadata.
D6. Rischio: alto.
D7. Mitigazioni:
- Estrarre in micro-step interni: dataset -> filtri -> populate -> search.
- Preservare identico mapping `status <-> event_type` e scope utente.
- Nessun cambio di formato visuale (`format_currency_display`, percentuali, colonne).
D8. Test minimi:
- Gate baseline.
- Smoke: filtri VSM avanzati (date, azione, ripetitivo, range importi), user scope all/current/other.
- Verifica metadata ownership su risultati filtrati.
D9. Rollback locale:
- Ripristino `_load_vsm_events`/`_populate_vsm_sheet` nel kernel.
D10. Precondizioni:
- Fase 5 stabile.
D11. Invarianti:
- I5, I6, I7, I8, I13.

### Fase 7 - Estrazione VSM command handlers (U6)
D1. Obiettivo: spostare CRUD/action VSM mantenendo UX e permessi.
D2. Area: edit/delete/duplicate/create VSM e supplier.
D3. Metodi: `_edit_vsm_event`, `_delete_vsm_events`, `_delete_supplier`, `_duplicate_vsm_event`, `_on_vsm_sheet_double_click`, `_on_supplier_sheet_double_click`, branch VSM di `open_new_event`.
D4. File target: `services/vsm_command_service.py`.
D5. Posizione: dopo stabilizzazione pipeline dati/metadata.
D6. Rischio: alto.
D7. Mitigazioni:
- Reuse selection/ownership service da Fase 3.
- Conservare debounce e dialog flow attuali.
- Conservare messaggi e policy read_only su eventi non propri.
D8. Test minimi:
- Gate baseline.
- Smoke: edit own vs other, delete multi-select own/other, duplicate, create da tab Saving/CA/Derisking.
D9. Rollback locale:
- Ripristino handlers nel kernel e disattivazione service.
D10. Precondizioni:
- Fase 6 stabile e metadata affidabili.
D11. Invarianti:
- I5, I6, I7, I10, I13.

### Fase 8 - Estrazione RFQ load/render e command handlers (U7-U8)
D1. Obiettivo: isolare operativita RFQ mantenendo semantica multi-db e ownership.
D2. Area: carico RFQ, rendering sheet, CRUD, duplicate, status change, open/edit.
D3. Metodi: `_load_requests_by_status`, `update_treeview`, `_format_date_for_display`, `archive_selected_request`, `reactivate_selected_request`, `_change_request_status`, `delete_selected_request`, `duplicate_selected_request`, `on_sheet_double_click`, `open_new_request_window`, branch RFQ di `open_new_event`.
D4. File target: `services/rfq_dashboard_service.py`, `services/rfq_command_service.py`.
D5. Posizione: dopo VSM extraction per ridurre simultaneita di rischio cross-domain.
D6. Rischio: alto.
D7. Mitigazioni:
- Preservare delete allegati + delete DB nello stesso flusso.
- Preservare metadata `_sheet_rows_metadata` e highlight scadenze.
- Preservare debouncing apertura `ViewRequestWindow`.
D8. Test minimi:
- Gate baseline.
- Smoke: load tab attive/archiviate, ricerca per utente, delete/duplicate/status, apertura dettaglio own/other.
D9. Rollback locale:
- Restore metodi RFQ nel kernel.
D10. Precondizioni:
- Fase 3 (ownership) stabile.
D11. Invarianti:
- I5, I6, I7, I8, I10, I13.

### Fase 9 - Consolidamento search/filter orchestration (U9)
D1. Obiettivo: completare separazione search/filter senza cambiare semantica query.
D2. Area: dashboard controller + hook kernel.
D3. Metodi: `_has_active_search_filters`, `search_requests`, `clear_filters`, `refresh_data`, `populate_username_filter`, `_update_filter_panel_for_current_tab`, `_update_advanced_filters_toggle`, `on_tab_changed`.
D4. File target: estensione `services/dashboard_controller.py` + `services/dashboard_search_service.py`.
D5. Posizione: dopo RFQ/VSM extraction per agganciare servizi gia separati.
D6. Rischio: molto alto (branching locale/aggregato e global OR).
D7. Mitigazioni:
- Freeze semantico esplicito: global OR + advanced AND.
- Non modificare SQL/predicate se non necessario al disaccoppiamento.
- Logging parity durante transizione.
D8. Test minimi:
- Gate baseline.
- Smoke matrix filtri: global-only, advanced-only, combinati, scope utente all/current/other, tab RFQ/VSM/Derisking.
D9. Rollback locale:
- Ripristino metodi nel `MainWindow` o nel vecchio controller.
D10. Precondizioni:
- U7-U8 stabili.
D11. Invarianti:
- I7, I8, I13.

### Fase 10 - Estrazione settings preferences (U11)
D1. Obiettivo: rendere `SettingsWindow` orchestratore UI, non contenitore di logica config.
D2. Area: lingua/valuta/autobackup settings base.
D3. Metodi: `load_settings`, `save_language_settings`, `save_currency_settings`, `select_autobackup_path`, `save_autobackup_settings`.
D4. File target: `services/settings_preferences_service.py`.
D5. Posizione: dopo dashboard core, prima manutenzione ad alto rischio.
D6. Rischio: medio.
D7. Mitigazioni:
- Mantieni identici testi dialog e restart prompt.
- Mantieni encoding UTF-8 e fallback default.
D8. Test minimi:
- Gate baseline.
- Smoke Settings: load defaults, save lingua/valuta, save autobackup path.
D9. Rollback locale:
- Ripristino metodi in `SettingsWindow`.
D10. Precondizioni:
- Contratto config file consolidato.
D11. Invarianti:
- I1, I10, I12, I13.

### Fase 11 - Estrazione backup e migrazione DataFlow folder (U12)
D1. Obiettivo: isolare flussi operativi piu rischiosi (file/DB/restart).
D2. Area: backup manuale/autobackup/migrazione folder.
D3. Metodi: `backup_database`, `check_for_autobackup`, `perform_autobackup`, `select_standard_dataflow_location`.
D4. File target: `services/settings_maintenance_service.py`, `services/dataflow_location_service.py`.
D5. Posizione: in coda perche altamente sensibile.
D6. Rischio: molto alto.
D7. Mitigazioni:
- Estrarre con wrapper conservativo (stessi dialog, stessi progress update).
- Preservare lock handling Windows e retry behavior.
- Preservare rollback config e cleanup cartella parziale.
D8. Test minimi:
- Gate baseline.
- Smoke: backup manuale con DB aperto, autobackup timer, migrazione folder con e senza conflitto username.
- Verifica riapertura DB post-backup e restart flow.
D9. Rollback locale:
- Reintegro dei metodi nel `SettingsWindow/MainWindow` originali.
D10. Precondizioni:
- Fase 10 stabile.
D11. Invarianti:
- I3, I4, I7, I10, I11, I13.

### Fase 12 - Hardening startup/restart e cleanup finale kernel
D1. Obiettivo: chiudere refactor mantenendo kernel chiaro e non opaco.
D2. Area: `restart_program`, `main_task`, blocco `__main__`, `_pending_restart`.
D3. Metodi: `restart_program`, `main_task` (con eventuale estrazione helper), rilancio post-mainloop.
D4. File target: `services/restart_lifecycle_service.py`, `services/startup_orchestrator_service.py` (opzionale, con wrapper nel kernel).
D5. Posizione: ultima fase, massima sensibilita temporale.
D6. Rischio: molto alto.
D7. Mitigazioni:
- Non cambiare ordine funzionale startup.
- Tenere nel kernel la regia finale anche se helper vengono estratti.
- Test manuale Linux e Windows obbligatorio.
D8. Test minimi:
- Gate baseline.
- Smoke completo: first run senza identity/licenza, restart da settings, riapertura app, tab operativi.
D9. Rollback locale:
- Restore blocchi startup/restart originali in `dataflow.py`.
D10. Precondizioni:
- Tutte le fasi precedenti verdi.
D11. Invarianti:
- I1, I2, I3, I4, I11, I13.

## 9. Trappole architetturali del piano
Sezione onesta sui punti dove una futura run autonoma puo deragliare.

- Trappola 1: "estrarre rendering" spezzando metadata.
  - `_populate_vsm_sheet` e `update_treeview` sembrano view-only ma sono anche policy di ownership/source.
- Trappola 2: moduli facciata ancora accoppiati a `self`.
  - Se i servizi accettano `app` intero, il monolite cambia solo posizione.
- Trappola 3: over-splitting.
  - Troppi micro-moduli rendono il kernel dispatcher opaco e difficile da seguire.
- Trappola 4: toccare startup/restart troppo presto.
  - Alto rischio di rompere app launch/relaunch cross-platform.
- Trappola 5: cambiare semantica search mentre si "pulisce" il codice.
  - Rischio concreto su global OR + advanced AND + multi-db scope.
- Trappola 6: separare delete RFQ da delete file allegati senza contratto.
  - Possibili orfani file o cancellazioni parziali incoerenti.
- Trappola 7: migrazione DataFlow folder trattata come semplice utility filesystem.
  - In realta include identita, conflitto username, DB rename/update, config rollback, restart.
- Trappola 8: rifattorizzare dialoghi con "migliorie" UX involontarie.
  - Vietato: questa run deve preservare UX/layout e comportamento dialog.
- Trappola 9: introdurre pattern non necessari.
  - Niente architecture astronauting: service semplici, contratti espliciti, no nuove dipendenze.
- Trappola 10: perdita di leggibilita del kernel.
  - Se il kernel diventa solo una sequenza di call criptiche, si perde il ruolo di cervello.

## 10. Valutazione di fattibilita della futura run autonoma
Verdetto: **Si con condizioni forti**.

Il refactor completo e realisticamente eseguibile in una singola run autonoma solo se valgono tutte queste condizioni:
- Esecuzione strettamente in ordine di fase (no parallel refactor su aree ad alto coupling).
- Gate test obbligatori dopo ogni fase, con stop immediato su regressione.
- Wrapper conservativi nel kernel durante tutta la transizione; cleanup solo alla fine.
- Nessuna deviazione scope (no feature nuove, no redesign UX, no schema DB).
- Presidio esplicito delle invarianti I1-I13.

Aree che possono dover restare nel kernel anche a fine run, per pragmatismo:
- Regia finale startup/restart (`__main__`, `_pending_restart`, sequencing principale).
- Orchestrazione ad alto livello di `MainWindow.__init__`.
- Wrapper orchestratore di `select_standard_dataflow_location` se l'estrazione totale aumentasse il rischio.

Compromessi pragmatici preferibili a decomposizione "perfetta":
- Meglio 1-2 metodi orchestratori medi nel kernel che 5 moduli facciata opachi.
- Meglio mantenere restart/startup principalmente nel kernel con helper piccoli, invece di spingere tutto fuori.
- Meglio stabilita semantica che pulizia teorica estrema.

## 11. Struttura target finale suggerita
Struttura concreta e minimale (no overengineering):

- `dataflow.py` (kernel)
  - bootstrap/import order
  - `SettingsWindow` e `MainWindow` come orchestratori
  - routing tab/eventi
  - entrypoint e lifecycle globale

- `services/excel_export_service.py`
  - orchestrazione export RFQ/VSM/Derisking

- `services/excel_export_builders.py`
  - costruzione workbook e formattazioni comuni

- `services/dashboard_selection_policy.py`
  - selezione righe e ownership checks

- `services/dashboard_actions_policy.py`
  - enable/disable Actions e composizione menu

- `ui/sheet_factories.py`
  - factory RFQ/VSM/Derisking `tksheet`

- `services/vsm_dashboard_service.py`
  - dataset/filter/populate/search VSM

- `services/derisking_dashboard_service.py`
  - supplier pipeline dedicata

- `services/vsm_command_service.py`
  - command handlers VSM/Derisking

- `services/rfq_dashboard_service.py`
  - load/render RFQ

- `services/rfq_command_service.py`
  - command handlers RFQ

- `services/dashboard_controller.py`
  - resta centrale per orchestration search/refresh, con responsabilita meglio delimitate

- `services/settings_preferences_service.py`
  - lingua/valuta/autobackup settings base

- `services/settings_maintenance_service.py`
  - backup manuale/autobackup operativo

- `services/dataflow_location_service.py`
  - logica migrazione cartella DataFlow (con wrapper nel kernel)

- `services/restart_lifecycle_service.py` (opzionale, helper)
  - risoluzione comando restart e supporto lifecycle

- `services/startup_orchestrator_service.py` (opzionale, helper)
  - helper startup con regia finale nel kernel

Regola anti-facciata nella struttura target:
- Ogni service deve avere input espliciti (argomenti concreti), non `self` completo.

## 12. Conclusione operativa
Il piano e completo, eseguibile e orientato a una futura run autonoma one-shot, ma resta prudente: il kernel `dataflow.py` viene alleggerito in profondita senza perdere il ruolo di cervello.

Strategia vincente proposta:
- prima estrazioni a rischio medio e alta coesione,
- poi blocchi cross-cutting (ownership/actions/sheet metadata),
- poi CRUD/search complessi,
- infine le zone piu sensibili (backup/migrazione/startup/restart).

Esito atteso a fine refactor:
- `dataflow.py` leggibile e orchestrativo,
- logiche operative profonde in moduli reali,
- invarianti critiche preservate,
- regressioni minimizzate da gating e rollback per fase.
