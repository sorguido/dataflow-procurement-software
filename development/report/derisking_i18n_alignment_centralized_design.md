# Derisking I18N Alignment — Centralized Design

## 1. Problema architetturale
Nel dominio Derisking Status, il mapping IT -> EN oggi è frammentato:

- è presente in più file;
- non è nel layer i18n centralizzato (`utils/i18n_utils.py`);
- non è usato in modo uniforme da tutti i punti UI/export.

Questo crea debito tecnico perché viola il principio "single source of truth" del refactor i18n 2026-04-10 (API unica `tr(...)`, separazione canonico/UI, niente logiche distribuite).

---

## 2. Mapping esistente (raccolta)

- export dashboard Derisking:
- `services/excel_export_service.py:491-496`
- mapping: `Nuovo -> New`, `In valutazione -> Under Evaluation`, `Qualificato -> Qualified`, `Scartato -> Rejected`

- export KPI Derisking:
- `services/kpi_excel_export.py:75-80`
- stesso mapping duplicato (`_STATUS_DERISKING_EN`)

- dialog Derisking:
- `ui/dialogs/potential_supplier_dialog.py:39-48`
- mapping canonico status -> label UI via `_status_label(...)` con `tr("New"|"Under Evaluation"|"Qualified"|"Rejected")`

- altri punti:
- griglia Derisking NON usa mapping dominio: passa `tr` diretto su canonico IT (`dataflow.py:1563-1567`, `services/derisking_dashboard_service.py:15`)
- KPI window card dinamiche mostrano status raw (`ui/kpi_window.py:1061-1065`)

Unione completa dominio status Derisking (da export + dialog + KPI export):
- `Nuovo` <-> `New`
- `In valutazione` <-> `Under Evaluation`
- `Qualificato` <-> `Qualified`
- `Scartato` <-> `Rejected`

Nota importante sulla search:
- non esiste mapping EN->IT esplicito nel codice search (`services/dashboard_search_service.py:22-42`);
- la ricerca globale opera per substring sui campi raw, incluso `notes` e `supplier_status`.

---

## 3. Proposta di centralizzazione

- posizione (file): `utils/i18n_utils.py`
- forma (funzioni / struttura): estensione del pattern RFQ già presente (`normalize_rfq_type(...)`, `translate_rfq_type(...)` in `utils/i18n_utils.py:161-229`)

Obiettivo:
- spostare il mapping Derisking Status in i18n centralizzato;
- mantenere `tr(...)` invariato;
- rendere il mapping riusabile da UI, export e dialog come unica fonte.

---

## 4. API proposta (design, NO codice)

- `normalize_derisking_status(value)`
- converte input status (IT/EN/varianti) nel canonico dominio Derisking (IT, coerente con DB attuale)
- gestisce input sporchi/null/legacy con fallback difensivo

- `translate_derisking_status(value)`
- usa `normalize_derisking_status(...)`
- mappa canonico IT -> msgid EN stabile
- invoca `tr(msgid_en)` per rendering UI
- garantisce fallback EN naturale (se manca chiave locale, gettext restituisce msgid EN)

Proprietà attesa:
- nessun branch lingua manuale
- nessun mapping duplicato nei consumer

---

## 5. Flusso target
DB (IT)
-> normalize
-> msgid EN
-> tr()
-> UI / Export / Dialog

Applicazione pratica:
- stesso flusso per tutti i punti che mostrano status Derisking;
- nessun `tr(canonico_IT)` diretto nei consumer.

---

## 6. Migrazione punti esistenti

- UI Derisking
- allineare il populate griglia a `translate_derisking_status(...)` al posto del passaggio diretto `translate_status=tr`
- punto attuale da riallineare: `dataflow.py:1563-1567`, `services/derisking_dashboard_service.py:15`

- Export
- sostituire i dict locali in:
- `services/excel_export_service.py:491-496`
- `services/kpi_excel_export.py:75-80`
- con riuso API centralizzata Derisking status

- Dialog
- sostituire la logica `_status_label/_status_canonical` locale con chiamate al mapping centralizzato (stessa semantica, unica fonte)
- punto attuale: `ui/dialogs/potential_supplier_dialog.py:39-55`

- altri
- KPI window card dinamiche: usare traduzione status centralizzata sui label (`ui/kpi_window.py:1061-1065`) per coerenza UI

---

## 7. Impatto

- UI:
- coerenza completa dei label status Derisking tra griglia, dialog e KPI card

- export:
- output invariato funzionalmente (già corretto), ma senza duplicazione mapping

- search:
- logica invariata; nessuna dipendenza nuova su stringhe tradotte

- DB:
- nullo (canonico e schema invariati)

---

## 8. Rischi

- rischio principale: regressione su valori status legacy/non previsti se normalizzazione non preserva fallback difensivo
- rischio secondario: rollout incompleto (alcuni consumer restano su mapping locale), con incoerenza residuale
- nessun rischio su persistenza dati, perché non cambia DB

---

## 9. Strategia di test

1. Re-test baseline empirica già eseguita:
- UI EN, search `qualified`, export EN con status corretti

2. Verifica coerenza cross-surface status Derisking:
- griglia, dialog, KPI card, export dashboard, export KPI

3. Verifica IT locale:
- stessi record mostrati correttamente in italiano

4. Verifica fallback difensivo:
- record con status non standard/legacy devono restare visualizzabili senza crash

5. Smoke test non-regressione:
- RFQ/Saving/Cost Avoidance invariati

---

## 10. Coerenza con refactor i18n
Verifica piena: SI

- estende il pattern RFQ esistente (normalizzazione + traduzione)
- mantiene `tr(...)` come API ufficiale
- evita mapping distribuiti e duplicati
- evita branch lingua manuali
- non modifica `tr(...)`, DB, business logic o dipendenze
