# Refactor i18n - Translation Service Centralizzato (2026-04-10)

## A. ARCHITETTURA FINALE ADOTTATA

### Moduli introdotti/modificati
- `utils/i18n_utils.py`: evoluto a servizio centralizzato con singleton `TranslationService`.
- Moduli UI prioritari migrati all'accesso esplicito `tr(...)`.

### API centrale scelta
- API ufficiale runtime: `tr(text)`.
- Accesso al service: `get_translation_service()`.
- Bootstrap: `init_i18n(language_code='en')`.
- Retrocompatibilità: `_` resta alias di `tr` e `builtins._` resta installato dal service solo per moduli legacy non ancora migrati.

### Regole d'uso applicate nel refactor
- Nei moduli migrati: niente dipendenza implicita da `builtins._`.
- Nei moduli migrati: traduzione sempre via import esplicito `from utils.i18n_utils import tr`.
- Niente `tr(...)` / `_()` a import-time nei file migrati.
- Riduzione dei branch `if lingua == ...` per testi UI, sostituiti da `tr(...)`.

## B. FILE TOCCATI

- `utils/i18n_utils.py`
  - Introduzione del `TranslationService` centralizzato.
  - Uniformazione `init_i18n`, `get_current_language`, `tr`, alias `_`.
- `ui/kpi_window.py`
  - Migrazione completa `_()` -> `tr(...)` con import esplicito.
  - Rimozione fallback locale su `builtins`.
  - Eliminato import-time translation su `_PERIOD_OPTIONS` (ora token stabile `"ALL"` + label tradotta runtime).
- `ui/windows/view_request_window.py`
  - Import esplicito `tr`.
  - Migrazione chiamate testo a `tr(...)`.
  - Rimozione branch lingua per testo pulsante PO (`tr("📋 Add PO")`).
- `ui/windows/purchase_order_window.py`
  - Migrazione a `tr(...)`.
  - Eliminati branch manuali `get_current_language() == 'it'` per testi UI (titolo, pulsanti, label, messaggi).
- `ui/dialogs/common_dialogs.py`
  - Import esplicito `tr`.
  - Migrazione completa delle stringhe a `tr(...)`.
  - Uniformati pulsanti `OK` su percorso traduzione centralizzato.
- `ui/windows/attachment_window.py`
  - Import esplicito `tr`.
  - Migrazione completa delle stringhe a `tr(...)`.
  - Correzione caso hardcoded (`"Errore Database"`) su traduzione centralizzata.
- `dev_tools/compile_translations.py`
  - Fix percorso locale: ora risolve `locale/` dalla root progetto (non da `dev_tools/locale`).
- `locale/it/LC_MESSAGES/dataflow.po`
  - Aggiunti msgid necessari introdotti dalla rimozione dei branch manuali in `purchase_order_window.py`.
- `locale/en/LC_MESSAGES/dataflow.po`
  - Allineamento msgid equivalenti lato EN.
- `locale/it/LC_MESSAGES/dataflow.mo`
- `locale/en/LC_MESSAGES/dataflow.mo`
  - Rigenerati con script di compilazione.

## C. MODIFICHE CHIAVE

### Pattern fragili eliminati
- `_()` implicito via `builtins` nei moduli migrati (ora `tr` importato esplicitamente).
- Traduzione a import-time nel modulo KPI (periodo `All` ora risolto runtime).
- Branch manuali lingua per testi UI in `purchase_order_window.py`.
- Branch manuale testo PO in `view_request_window.py`.
- Hardcoded non tradotto in `attachment_window.py` (errore DB).

### Pattern legacy ancora presenti (fuori scope in questa iterazione)
- Rimangono moduli non prioritari che usano ancora `_`/`builtins._`.
- Restano aree codebase con branch lingua storici non toccati perché fuori perimetro.
- Persistono msgid storicamente misti IT/EN nel catalogo (non è stata fatta migrazione globale msgid).

## D. RISCHI

### Rischi residui
- Catalogo `.po` storico eterogeneo IT/EN: rischio traduzione incompleta in percorsi non toccati.
- Moduli non migrati possono ancora dipendere da `builtins._`.
- Alcuni flussi export (scelta lingua file Excel) mantengono logica lingua separata per requisiti funzionali, non per testo UI principale.

### Punti da testare manualmente
- Apertura `KpiWindow`: periodo `ALL`, filtri anno/periodo, export KPI.
- `ViewRequestWindow`: toolbar, popup/errori, pulsante PO.
- `PurchaseOrderWindow`: add/delete/validazioni e dialog di conferma in IT/EN.
- `AttachmentWindow`: add/open/download/delete con error handling.
- `LanguagePrompt`, `SimpleMessageDialog`, `SimpleYesNoDialog`.

### Aree volutamente non migrate ora
- Tutti i moduli UI non inclusi nella lista prioritaria.
- Normalizzazione completa msgid legacy del catalogo.
- Refactor architetturale globale dei testi di business/export.

## E. ROLLBACK (PER BLOCCHI)

### Blocco 1 - Core service
- File: `utils/i18n_utils.py`
- Rollback: ripristinare versione precedente del modulo.
- Impatto rollback: i moduli migrati tornano dipendenti da modello legacy.

### Blocco 2 - Migrazione UI prioritaria
- File: `ui/kpi_window.py`, `ui/windows/view_request_window.py`, `ui/windows/purchase_order_window.py`, `ui/dialogs/common_dialogs.py`, `ui/windows/attachment_window.py`
- Rollback: revert selettivo per file/finestra.
- Impatto rollback: regressione locale della coerenza i18n ma senza impatto DB.

### Blocco 3 - Cataloghi/tooling
- File: `locale/*/LC_MESSAGES/dataflow.po`, `locale/*/LC_MESSAGES/dataflow.mo`, `dev_tools/compile_translations.py`
- Rollback: revert dei cataloghi o solo del tooling.
- Impatto rollback: possibile perdita traduzioni nuove ma nessuna modifica dati business.

## F. MINI STANDARD FUTURO

1. Import standard UI:
   - `from utils.i18n_utils import tr`
2. Traduzione UI:
   - usare `tr(...)` solo a runtime (in `__init__`, metodi UI, callback).
3. Vietato:
   - `_()`/`tr()` a import-time per costanti modulo/classi.
   - testi UI con `if get_current_language() == ...` quando è sufficiente `tr(...)`.
4. Popup/dialog/finestre figlie:
   - titolo, testo, bottoni sempre via `tr(...)`.
5. Logica business:
   - mai dipendere da stringhe tradotte per confronti decisionali.
   - usare valori canonici stabili (enum/token/codici), tradurre solo in presentazione.
6. Compatibilità legacy:
   - `builtins._` resta solo ponte temporaneo per moduli non migrati.
   - ogni nuovo modulo deve essere già nel modello `tr(...)` esplicito.

## Verifiche eseguite
- `python3 -m py_compile utils/i18n_utils.py ui/kpi_window.py ui/windows/view_request_window.py ui/windows/purchase_order_window.py ui/dialogs/common_dialogs.py ui/windows/attachment_window.py dev_tools/compile_translations.py`
- `python3 dev_tools/compile_translations.py`
  - `.mo` compilati con successo per `en` e `it`.
