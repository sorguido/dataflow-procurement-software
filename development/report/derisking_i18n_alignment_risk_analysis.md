# Derisking i18n Alignment — Analisi e Valutazione Rischi

## 1. Stato attuale Derisking
Il flusso Derisking oggi è composto da una pipeline mista:

- Modello dati con stati canonici italiani: `Nuovo`, `In valutazione`, `Qualificato`, `Scartato` in `models/potential_supplier.py:13-24`.
- Dialog Derisking con mapping canonico -> label UI tradotta (`_status_label`) e reverse mapping in salvataggio (`_status_canonical`) in `ui/dialogs/potential_supplier_dialog.py:39-55` e `:459-464`.
- Popolamento griglia Derisking che passa `translate_status=tr` (`dataflow.py:1563-1567`) e applica `translate_status(supplier.supplier_status)` direttamente (`services/derisking_dashboard_service.py:15`).
- Header griglia già tradotti correttamente via `tr(...)` (`ui/sheet_factories.py:215-218`).

Effetto osservato: in lingua EN, header corretto ma valori `Status` in italiano.

---

## 2. Confronto con pipeline i18n standard
Saving / Cost Avoidance / RFQ seguono una pipeline più coerente tra valore canonico e rendering UI:

- Saving/CA:
- traduzione valori in populate (`tr(event.action)`) in `dataflow.py:1732`
- filtri che confrontano valori tradotti (`tr(event.action)`, `tr("Yes")`) in `services/vsm_dashboard_service.py:102-108`
- RFQ:
- mapping esplicito tipo canonico -> label tradotta (`translate_rfq_type`) in `services/rfq_dashboard_service.py:50-52`, `utils/i18n_utils.py:215-227`

Derisking griglia invece usa `tr(raw_canonico_italiano)` senza mapping intermedio (`services/derisking_dashboard_service.py:15`).

---

## 3. Verifica migrazione i18n

- Stato: parzialmente migrato
- Evidenze tecniche:
- uso di `tr(...)` presente e diffuso in Derisking UI/dialog/service (`ui/dialogs/potential_supplier_dialog.py`, `ui/sheet_factories.py`, `dataflow.py:1566`, `services/derisking_dashboard_service.py:15`)
- assenza di `_()` nei moduli Derisking analizzati (nessuna dipendenza legacy implicita a `builtins._`)
- persistenza di pattern legacy sui valori canonici (stati italiani hardcoded nel model) in `models/potential_supplier.py:13-24`
- mismatch storico schema/model (`DEFAULT 'Prospect'` in DB) in `database_manager.py:272`
- refactor i18n 2026-04-10 non include file Derisking nel perimetro “file toccati” (`development/report/REFACTOR_I18N_TRANSLATION_SERVICE_2026-04-10.md:21-53`) e dichiara aree non prioritarie fuori scope (`:83-87`)

Conclusione della verifica: Derisking non risulta “non migrato”, ma “migrato a metà” (API `tr` adottata, allineamento canonico/rendering non completato).

---

## 4. Possibili motivi di esclusione dal refactor
Ipotesi motivate dal codice e dal report (non dichiarazione esplicita singolo modulo):

- Il refactor 2026-04-10 era scoped su moduli UI prioritari specifici; Derisking non è nella lista (`REFACTOR...md:21-53`).
- Il report indica esplicitamente aree fuori scope e niente refactor globale dei testi business/export (`REFACTOR...md:83-87`).
- Derisking usa una semantica dati separata (`PotentialSupplier`) con dipendenze su DB, KPI, export e search; un riallineamento “profondo” avrebbe rischio più alto rispetto a finestre UI pure.
- Nei flussi export sono ammesse logiche lingua dedicate per requisito funzionale (`REFACTOR...md:74`), e Derisking export infatti usa mapping manuale stato (`services/excel_export_service.py:491-502`).

---

## 5. Root cause reale del disallineamento
Root cause primaria:

- valore canonico `supplier_status` in italiano (`models/potential_supplier.py:13-24`)
- rendering griglia che invoca `tr(...)` direttamente su quel valore (`services/derisking_dashboard_service.py:15`)
- catalogo gettext con chiavi status in inglese (`New`, `Under Evaluation`, `Qualified`, `Rejected`) in `locale/en/LC_MESSAGES/dataflow.po:1864-1874` e `locale/it/LC_MESSAGES/dataflow.po:1878-1888`

Quindi, in EN, `tr("Qualificato")` non mappa e restituisce il testo originale italiano.

---

## 6. Analisi dei rischi di allineamento

### 6.1 Rischi UI
- Rischio regressione basso se l’allineamento resta nel solo layer di presentazione griglia Derisking.
- Rischio incoerenza residua medio perché KPI Derisking espone status raw da DB (`ui/kpi_window.py:1061-1065`), quindi può restare disallineato rispetto alla griglia se non esplicitato come non-obiettivo.

### 6.2 Rischi dati / DB
- Rischio basso se non si toccano valori canonici persistiti.
- Rischio alto se si tenta di cambiare i token canonici DB: impatta insert/update/query (`database_manager.py:2550-2558`, `:2590-2606`, `:2677-2693`).
- Rischio tecnico preesistente da monitorare: presenza legacy `Prospect` nello schema (`database_manager.py:272`) e riferimento a `SUPPLIER_STATUS_PROSPECT` nel model (`models/potential_supplier.py:136`).

### 6.3 Rischi export
- Rischio basso per allineamento solo UI-griglia.
- Rischio medio/alto se si altera il canonico: export Derisking EN dipende da mapping manuale da valori italiani (`services/excel_export_service.py:491-502`).

### 6.4 Rischi search/filter
- Rischio medio funzionale/UX: global search Derisking filtra su `supplier_status` raw (`services/dashboard_search_service.py:27-31`), quindi la ricerca per label EN può restare non intuitiva anche dopo allineamento visuale della colonna.
- Rischio regressione basso se il comportamento search non viene toccato in questa fase.

### 6.5 Rischi architetturali
- Rischio basso se intervento localizzato e senza nuovi layer.
- Rischio medio sistemico se si allarga lo scope: Derisking combina i18n UI moderno con canonico storico language-bound, quindi un “riallineamento totale” richiederebbe coordinamento su più moduli.

---

## 7. Valutazione complessiva

- Allineamento consigliato: SI
- Motivazione: è consigliabile solo un allineamento minimale di presentazione nel flusso griglia Derisking, mantenendo invariati canonico DB e logiche business/export/search. Questo è coerente con il principio `tr(...)` runtime e minimizza il rischio regressione.

---

## 8. Piano di allineamento (NO CODICE)

### Step 1
Definire esplicitamente il perimetro: allineare solo la visualizzazione `Status` della griglia Derisking al mapping già usato nel dialog, senza toccare valori canonici persistiti.

### Step 2
Applicare il mapping nel solo punto di preparazione righe Derisking (service di populate), preservando fallback difensivo su valori status non mappati.

### Step 3 (solo se necessario)
Eseguire verifica di coerenza trasversale post-allineamento su KPI/search/export e documentare eventuali disallineamenti residui come backlog separato, senza estendere il change in questa iterazione.

---

## 9. Strategia di test manuale

1. Lingua EN: aprire tab Derisking e verificare che la colonna `Status` mostri label inglesi per record in tutti gli stati noti.
2. Lingua IT: verificare che gli stessi record restino corretti in italiano.
3. Creazione/modifica supplier da dialog: salvare e verificare coerenza tra valore scelto nel dialog e valore mostrato in griglia.
4. Export Derisking in IT/EN: verificare che gli status esportati restino coerenti con comportamento atteso attuale.
5. Global search Derisking: validare comportamento con query su status (termini IT/EN) e registrare eventuale limite UX senza cambiarne la logica.
6. Smoke test Saving/Cost Avoidance/RFQ su traduzioni colonne/valori per escludere regressioni.

---

## 10. Coerenza con refactor i18n 2026-04-10

- Allineato ai principi: SI
- Note:
- rispetta API ufficiale `tr(...)` e runtime-only translation (`REFACTOR...md:10`, `:110-113`)
- evita dipendenza da `_()` e branch lingua manuali
- non introduce nuovi sistemi i18n
- non altera logica business/canonico DB, in linea con approccio conservativo dichiarato nel refactor
