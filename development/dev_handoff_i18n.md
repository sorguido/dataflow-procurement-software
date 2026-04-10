# Dev Handoff i18n

## Architettura corrente
- Translation service centralizzato in `utils/i18n_utils.py`.
- API runtime ufficiale nei moduli migrati: `tr(...)`.
- `.po/.mo` mantenuti.
- `_`/`builtins._` mantenuti solo per compatibilita' legacy fuori scope.

## File modificati (giro finale)
- `ui/dialogs/common_dialogs.py`
- `locale/it/LC_MESSAGES/dataflow.po`
- `locale/en/LC_MESSAGES/dataflow.po`
- `locale/it/LC_MESSAGES/dataflow.mo`
- `locale/en/LC_MESSAGES/dataflow.mo`

## Modifiche completate
- Residui i18n eliminati nei moduli gia' migrati:
  - `CopyProgressWindow`: rimossi hardcoded IT su titolo default e stato iniziale, ora via `tr(...)`.
  - Allineati cataloghi per `OK`, `Copia in corso...`, `Preparazione...`.
- Coerenza codice-cataloghi validata per i flussi migrati:
  - `ui/kpi_window.py`
  - `ui/windows/view_request_window.py`
  - `ui/windows/purchase_order_window.py`
  - `ui/dialogs/common_dialogs.py`
  - `ui/windows/attachment_window.py`

## Verifiche eseguite
- Audit automatico: nessun `msgid` mancante in IT/EN per tutte le chiamate `tr(...)` nei moduli sopra.
- Ricompilazione cataloghi: `python3 dev_tools/compile_translations.py` (OK).
- Compilazione Python: `python3 -m py_compile ...` sui moduli migrati (OK).

## Problemi aperti
- Nessun residuo i18n aperto nei flussi gia' migrati oggetto del refactor.
- Restano potenziali legacy fuori scope nei moduli non migrati.

## Rischi residui
- Catalogo storico ampio e misto IT/EN in aree fuori scope.

## Rollback per blocchi
- Blocco A (codice locale): revert `ui/dialogs/common_dialogs.py`.
- Blocco B (cataloghi): revert `locale/*/LC_MESSAGES/dataflow.po` e ricompilare `.mo`.
