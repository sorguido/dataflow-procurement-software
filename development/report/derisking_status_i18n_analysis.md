# Derisking Status i18n — Analisi e Piano di Intervento

## 1. Sintesi del problema
Nel tab Derisking, con lingua applicazione in inglese, l'header della colonna `Status` e le altre colonne sono tradotti correttamente, ma i valori riga (`Qualificato`, `In valutazione`, ecc.) restano in italiano.

Il problema non è nella configurazione globale gettext: è circoscritto al flusso di preparazione dati Derisking (valori cella), non al rendering degli header.

---

## 2. Comportamento corretto (baseline)
Nei tab Saving / Cost Avoidance / RFQ la pipeline applica una traduzione coerente tra valore canonico e label UI:

- Saving/Cost Avoidance:
  - gli header sono tradotti in factory (`ui/sheet_factories.py:71-100`)
  - i valori UI vengono tradotti in populate/filter (`dataflow.py:1729-1734`, `services/vsm_dashboard_service.py:102-108`)
- RFQ:
  - la colonna tipo RdO usa mapping esplicito canonico -> label tradotta (`services/rfq_dashboard_service.py:44-52`, `utils/i18n_utils.py:215-227`)

Quindi negli altri tab la traduzione non è un `tr(raw_db_value)` diretto, ma passa da un layer coerente con i valori canonici.

---

## 3. Analisi tecnica

### 3.1 Flusso Derisking
Pipeline attuale:

1. Recupero supplier dal DB -> `PotentialSupplier.supplier_status`
   - sorgente modello/canonico: `models/potential_supplier.py:13-24`
2. Popolamento sheet Derisking
   - `dataflow.py:_populate_potential_suppliers_sheet` passa `translate_status=tr` (`dataflow.py:1550-1567`)
   - `services/derisking_dashboard_service.py:build_supplier_rows_and_metadata` costruisce la riga e fa:
     - `translate_status(supplier.supplier_status)` (`services/derisking_dashboard_service.py:15`)
3. Rendering tksheet
   - header tradotto correttamente in `create_supplier_sheet` (`ui/sheet_factories.py:215-218`)
   - body mostra i valori già preparati dal service (`services/derisking_dashboard_service.py:78-85`)

### 3.2 Punto di rottura
Il punto di rottura è nella preparazione del valore `Status` per il body riga:

- `supplier_status` è canonico italiano nel modello (`Nuovo`, `In valutazione`, `Qualificato`, `Scartato`) in `models/potential_supplier.py:13-24`
- il service Derisking applica `tr(...)` direttamente a quel valore italiano (`services/derisking_dashboard_service.py:15`)
- ma il catalogo gettext è basato su msgid inglesi per questi stati (`New`, `Under Evaluation`, `Qualified`, `Rejected`) e non su msgid italiani (`locale/en/LC_MESSAGES/dataflow.po:1864-1874`, `locale/it/LC_MESSAGES/dataflow.po:1878-1888`)

Con lingua EN, `tr("Qualificato")` non trova una chiave valida e restituisce il testo originale -> resta italiano in UI.

### 3.3 Confronto con altri tab
Differenza concreta:

- RFQ usa mapping dedicato (`translate_rfq_type`) prima del rendering, non `tr()` diretto sul valore canonico (`services/rfq_dashboard_service.py:50-52`, `utils/i18n_utils.py:215-227`)
- Saving/CA usano valori coerenti con chiavi traducibili nei punti di populate/filter (`dataflow.py:1732`, `services/vsm_dashboard_service.py:102-108`)
- Derisking sheet usa `tr(raw_status_canonico_italiano)` senza mapping intermedio (`services/derisking_dashboard_service.py:15`)

Nota importante: nel dialog Derisking il mapping corretto esiste già (`_status_label`) e converte canonico -> chiave traducibile (`ui/dialogs/potential_supplier_dialog.py:39-49`), ma questa logica non viene riusata nel populate della griglia.

---

## 4. Root Cause
La root cause è una divergenza locale nel layer di preparazione dati Derisking:

- i valori status persistiti/modello sono canonici italiani
- il populate Derisking invoca `tr(...)` direttamente su quei valori italiani
- il catalogo gettext traduce gli status tramite chiavi inglesi

Quindi non manca gettext e non manca `tr` negli header: manca (solo in Derisking grid body) il mapping canonico-status -> chiave traducibile, già presente in altre pipeline/tab.

Classificazione della differenza:

- service: SI (punto principale)
- controller: NO
- rendering UI: NO (render mostra correttamente ciò che riceve)

---

## 5. Piano di intervento (NO CODICE)

### Step 1
Intervento minimo e localizzato nel solo flusso Derisking di preparazione righe (`services/derisking_dashboard_service.py` via chiamata da `dataflow.py:_populate_potential_suppliers_sheet`):

- sostituire la traduzione diretta `tr(raw_status)` con la stessa semantica già usata nel dialog Derisking (mapping status canonico -> label traducibile -> `tr`)
- mantenere invariati DB, modello, controller, sheet factory e altri tab

### Step 2
Verifica di non regressione limitata a:

- tab Derisking (solo colonna `Status` body)
- sanity check Saving / Cost Avoidance / RFQ per confermare assenza impatti collaterali

---

## 6. Impatto atteso

- UX: valori `Status` coerenti con lingua applicazione anche nel body Derisking
- coerenza: allineamento Derisking alla pipeline i18n già usata in RFQ/Saving/CA
- rischio regressione: basso, perché il change è localizzato al solo punto di formattazione visuale Derisking

---

## 7. Rischi

Rischi reali residui:

- presenza di eventuali status legacy non mappati esplicitamente: in quel caso resterà attivo il fallback al valore grezzo
- nessun rischio architetturale atteso se il perimetro resta confinato al solo mapping visuale Derisking

---

## 8. Strategia di test manuale

1. Impostare lingua EN, aprire Derisking con record in tutti gli stati; verificare che `Status` mostri `New / Under Evaluation / Qualified / Rejected`.
2. Passare a IT senza riavvio anomalo del flusso; verificare che gli stessi record mostrino `Nuovo / In valutazione / Qualificato / Scartato`.
3. Creare e modificare un supplier da dialog Derisking, salvare, tornare alla griglia e verificare coerenza tra valore selezionato nel dialog e valore mostrato in tab.
4. Verificare che header e valori di Saving/Cost Avoidance/RFQ non cambino comportamento rispetto alla baseline.

---

## 9. Verifica coerenza con linee guida progetto

- Modifica localizzata: OK
- Nessuna regressione: OK (con test manuale mirato)
- Coerenza con i18n esistente: OK
