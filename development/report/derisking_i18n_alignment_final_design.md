# Derisking I18N Alignment — Final Design

## 1. Nuove evidenze runtime
Dai test reali riportati:

- Search: con UI Derisking che mostra `Qualificato`, la query `qualified` restituisce risultati.
- Export: con UI EN e export EN, la colonna `Status` è corretta in inglese.

Lettura operativa:
- l’ecosistema Derisking ha già una compatibilità runtime EN sui percorsi search/export;
- il disallineamento è concentrato sul rendering della griglia Derisking, non sull’intero sistema.

---

## 2. Architettura reale emersa
Pipeline osservata nel codice:

- DB/canonico: status Derisking in IT (`Nuovo`, `In valutazione`, `Qualificato`, `Scartato`) in `models/potential_supplier.py:13-24`.
- Export EN: mapping IT->EN esplicito già presente (`status_export_en`) in `services/excel_export_service.py:491-502`.
- Dialog Derisking: mapping canonico->msgid EN via `_status_label` + reverse mapping via `_status_canonical` in `ui/dialogs/potential_supplier_dialog.py:39-55`.
- UI griglia Derisking: populate attuale passa `tr` diretto e usa `tr(canonico_IT)` (`dataflow.py:1563-1567`, `services/derisking_dashboard_service.py:15`).

Quindi la pipeline corretta esiste già in alcune aree (export/dialog), ma non è applicata alla griglia.

---

## 3. Punto di incoerenza
Il punto preciso che rompe la pipeline è il populate della griglia Derisking:

- `build_supplier_rows_and_metadata` usa `translate_status(supplier.supplier_status)` (`services/derisking_dashboard_service.py:15`)
- il chiamante passa `translate_status=tr` (`dataflow.py:1566`)

Risultato: `tr(canonico_IT)` invece di `tr(msgid_EN)`.

---

## 4. Mapping esistente
- dove si trova:
- `services/excel_export_service.py:491-496` (`status_export_en`: IT->EN)
- `ui/dialogs/potential_supplier_dialog.py:39-48` (`_status_label`: canonico IT -> `tr("New"|...)`)
- duplicazione aggiuntiva anche in KPI export: `services/kpi_excel_export.py:75-80` (`_STATUS_DERISKING_EN`)

- come viene usato oggi:
- export dashboard Derisking EN converte IT->EN prima dell’output (`services/excel_export_service.py:500-502`)
- dialog mostra label tradotte partendo dal canonico (`ui/dialogs/potential_supplier_dialog.py:39-48`, `:399-403`)

- se è duplicato o centrale:
- oggi è **duplicato**, non centralizzato in un unico punto dominio.

---

## 5. Strategia corretta
Strategia finale (senza nuovo sistema):

- non creare mapping nuovi;
- riusare il mapping Derisking già esistente (semantica export/dialog) come unica fonte dominio;
- applicare quel mapping anche al percorso UI griglia, così la griglia rientra nella stessa pipeline già valida per export;
- lasciare invariati DB, `tr(...)`, business logic e architettura globale.

Principio: la griglia deve consumare lo stesso “status-display resolver” già usato negli altri percorsi Derisking corretti, non chiamare `tr` su canonico IT grezzo.

---

## 6. Punto di intervento minimo
File/funzione minima:

- `dataflow.py` → `_populate_potential_suppliers_sheet(...)` (`dataflow.py:1550-1567`)
- `services/derisking_dashboard_service.py` → `build_supplier_rows_and_metadata(...)` (`services/derisking_dashboard_service.py:6-16`)

Intervento concettuale minimo:
- sostituire nel solo passaggio `translate_status` della griglia il resolver attuale (`tr`) con il resolver dominio già esistente per status Derisking.

---

## 7. Flusso finale corretto
Diagramma logico target:

DB (canonico IT)
-> mapping dominio Derisking esistente (riusato)
-> msgid/label EN compatibile catalogo
-> `tr(...)`
-> UI griglia / Export / (percorsi derivati)

Effetto:
- fallback inglese automatico sul dominio status, perché l’input a `tr` è coerente con msgid EN.

---

## 8. Impatto
- UI:
- positivo diretto su tab Derisking (colonna `Status` coerente con lingua).

- export:
- invariato (già corretto in EN); riallineamento aumenta coerenza UI↔export.

- search:
- invariato per logica; preserva il comportamento runtime già validato.

- DB:
- nullo (nessuna modifica valori canonici/schema).

---

## 9. Rischi reali
- Possibile divergenza residua se restano più mapping status in file diversi e non viene fissata una fonte unica riusata.
- Possibili record legacy con status fuori dominio noto: necessario mantenere fallback difensivo al valore grezzo.
- Nessun rischio strutturale su persistenza o schema, perché il piano opera solo sul rendering.

---

## 10. Strategia di test
1. Ripetere test runtime già fatti: UI EN + search `qualified` + export EN (baseline di regressione zero).
2. Verificare UI Derisking EN: status in colonna sempre EN per tutti i valori canonici noti.
3. Verificare UI Derisking IT: status invariati in italiano.
4. Verificare coerenza UI↔export sugli stessi record (stessa semantica status).
5. Smoke test RFQ/Saving/CA per confermare assenza impatti collaterali.

---

## 11. Coerenza con refactor i18n
Il design è coerente con il refactor 2026-04-10:

- mantiene `tr(...)` come API ufficiale;
- non introduce branch lingua manuali;
- non modifica `tr(...)` globalmente;
- non tocca DB/business;
- applica un fix chirurgico di riallineamento pipeline tramite riuso mapping dominio già esistente.
