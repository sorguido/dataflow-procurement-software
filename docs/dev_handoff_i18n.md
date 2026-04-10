# Dev Handoff i18n

## Architettura target corrente
- Translation service centralizzato in `utils/i18n_utils.py`.
- API runtime unica: `tr(...)`.
- `.po/.mo` mantenuti; `_`/`builtins._` solo retrocompatibilita' per moduli legacy non migrati.

## File modificati (questo giro)
- `locale/it/LC_MESSAGES/dataflow.po`
- `locale/en/LC_MESSAGES/dataflow.po`
- `locale/it/LC_MESSAGES/dataflow.mo`
- `locale/en/LC_MESSAGES/dataflow.mo`

## Modifiche completate
- Corretto bug KPI in sessione IT tramite allineamento catalogo: header, periodi, label `Year`, tab, bottone export, label KPI card.
- Corretto bug Purchase Order in sessione IT tramite allineamento catalogo: titolo, testi principali, label e bottoni.
- Completata la parte di step 2 su questi flussi: coerenza `tr(...)` <-> cataloghi aggiornata.

## Verifiche eseguite
- Controllo msgid runtime (`tr(...)`) presenti in `it.po` per:
  - `ui/kpi_window.py`
  - `ui/windows/purchase_order_window.py`
- `python3 dev_tools/compile_translations.py`: OK.
- `python3 -m py_compile ui/kpi_window.py ui/windows/purchase_order_window.py utils/i18n_utils.py dev_tools/compile_translations.py`: OK.

## Problemi aperti
- Duplicati storici in `.po` presenti anche prima (non affrontati in questa iterazione).
- Legacy residuo in moduli fuori scope che non usano ancora il pattern `tr(...)` in modo completo.

## Prossimi step
1. Smoke test manuale UI IT/EN sui due flussi (KPI e Purchase Order).
2. Continuare step 2 su moduli prioritari rimanenti, senza espansione fuori scope.

## Rischi residui
- Catalogo storico misto IT/EN fuori dai flussi appena stabilizzati.
- Possibili traduzioni mancanti in aree legacy non migrate.

## Rollback per blocchi
- Blocco A (cataloghi): revert `locale/*/LC_MESSAGES/dataflow.po` + `.mo`.
- Blocco B (service/tooling): revert `utils/i18n_utils.py` e `dev_tools/compile_translations.py`.
- Blocco C (UI migrata): revert selettivo file UI già migrati.
