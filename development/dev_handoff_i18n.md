# Dev Handoff i18n

## Architettura corrente
- Translation service centralizzato: `utils/i18n_utils.py`.
- API runtime ufficiale nei moduli applicativi: `tr(...)`.
- `.po/.mo` mantenuti (gettext).
- `_` / `builtins._` mantenuti solo come ponte legacy nel service, non nei moduli applicativi migrati.

## File modificati
- `dataflow.py`
- `services/dashboard_controller.py`
- `ui/dialogs/manage_supplier_categories_dialog.py`
- `ui/dialogs/vsm_event_dialog.py`
- `ui/dialogs/potential_supplier_dialog.py`
- `ui/components/main_dashboard_toolbar.py`
- `ui/components/collapsible_filters.py`
- `ui/main_dashboard_builder.py`
- `ui/windows/sqdc_analysis_window.py`
- `ui/windows/notes_window.py`
- `ui/windows/edit_suppliers_window.py`
- `ui/windows/edit_reference_window.py`
- `ui/kpi_chart.py`
- `ui/kpi_window.py`
- `utils/i18n_utils.py`
- `development/dev_tools/compile_translations.py`
- `locale/it/LC_MESSAGES/dataflow.po`
- `locale/en/LC_MESSAGES/dataflow.po`
- `locale/it/LC_MESSAGES/dataflow.mo`
- `locale/en/LC_MESSAGES/dataflow.mo`

## Moduli completati
- Migrazione completa `_()` -> `tr(...)` nei moduli applicativi rilevati dalla scansione globale.
- Rimozione import legacy `_` nei moduli migrati.
- Rimozione fallback locale `builtins._` da `dataflow.py`.
- Allineamento cataloghi su tutti i `msgid` runtime (`tr(...)`) rilevati da audit AST globale.
- Hardcoded UI locali residui eliminati nei moduli migrati (`kpi_chart`, `collapsible_filters`, `common_dialogs` già aggiornato nei giri precedenti).

## Moduli ancora aperti
- Nessun modulo aperto nel perimetro i18n applicativo migrato.

## Verifiche eseguite
- Scansione globale `_()`/`builtins._`/import legacy su `*.py`.
- Audit AST globale `tr(...)` vs cataloghi IT/EN: `files_with_missing = 0`.
- Compilazione Python globale: `python3 -m py_compile $(rg --files -g '*.py')` -> OK.
- Compilazione cataloghi: `python3 development/dev_tools/compile_translations.py` -> OK.

## Problemi aperti
- Nessun problema bloccante emerso nel perimetro i18n migrato.

## Rischi residui
- Duplicati storici nei `.po` possibili (non bloccanti per runtime gettext).
- Eventuali stringhe non-UI (branding/placeholder tecnici) intenzionalmente lasciate non tradotte se non parte del flusso localizzazione utente.

## Rollback per blocchi
- Blocco A (migrazione chiamate): revert dei file applicativi migrati (`dataflow.py`, `services/`, `ui/`).
- Blocco B (cataloghi): revert `locale/*/LC_MESSAGES/dataflow.po` e ricompilare `.mo`.
- Blocco C (tooling): revert `development/dev_tools/compile_translations.py`.
