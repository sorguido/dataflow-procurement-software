# DATAFLOW — REPORT TECNICO DETTAGLIATO POST-MIGRAZIONE I18N EN-FIRST

## Contesto e perimetro analisi
- Baseline usata per confronto: tag `pre-migration-i18n-tentative` (commit `a161de2`), perché nel repository non risulta presente un tag nominato esattamente `pre-migration-i18n`.
- Commit analizzato: `HEAD` (`4345733`).
- Richiesta eseguita in sola analisi: nessuna modifica aggiuntiva al codice applicativo durante questa attività di report.

---

## A. Vista Generale

### A1. Pattern principali di modifica nei `.py`
1. Migrazione meccanica delle sorgenti i18n: sostituzione diffusa di `tr("..." italiano)` con `tr("..." inglese)`.
2. Allineamento dei confronti UI che usavano valori tradotti italiani (es. `Tutte`, `Sì`) a sorgenti EN (`All`, `Yes`).
3. Minimi aggiustamenti di branch legati alla lingua (rimozione di ternarie IT/EN superflue) per mantenere una sola source string EN.
4. Compatibilità legacy conservata dove il dato canonico di business è ancora italiano (RFQ type, action/driver VSM), ma con display source EN.

### A2. Distinzione tipologie di modifica
- **Sostituzioni semplici `tr("IT") -> tr("EN")`**:
  - Dominanti nella quasi totalità dei file UI (`dataflow.py`, dialog/window, dashboard controller, KPI).
- **Disaccoppiamenti lingua/logica**:
  - Confronti filtro da `tr("Tutte")` a `tr("All")`.
  - Confronto booleano ripetitivo da `tr("Sì")` a `tr("Yes")`.
  - Rimozione branch per label variabile in base a `get_current_language()`.
- **Normalizzazioni legacy/minime compatibilità**:
  - `translate_rfq_type()` mantiene canonical IT (`Fornitura piena` / `Conto lavoro`) ma traduce da source EN (`Full Supply` / `Work Order`).
  - Mapping display/internal in `VSMEventDialog` aggiornato a source EN, mantenendo internal value legacy dove necessario.
- **Modifiche collaterali inevitabili**:
  - Uniformazione titoli/label/error message in EN nei punti toccati dalla migrazione i18n.

---

## B. Dettaglio File Python Toccati

| File | Motivo modifica | Tipologia | Impatto funzionale atteso | Rischio regressione | Solo i18n o anche flow/logica |
|---|---|---|---|---|---|
| `dataflow.py` | File UI principale con maggior volume di `tr(...)` | Sostituzioni i18n + pochi confronti/branch normalizzati | Testi EN-first, filtri coerenti con source EN | Medio | Prevalente i18n, con minimi touch su condizioni UI |
| `services/dashboard_controller.py` | Filtri ricerca usavano label tradotta italiana | Sostituzioni + normalizzazione confronti | Ricerca RFQ invariata, confronto su `All` | Basso | i18n + minima logica UI filtro |
| `services/rfq_pdf_export_service.py` | Messaggi warning export PDF e label | Sostituzioni stringhe | Export invariato, solo testo | Basso | i18n puro |
| `ui/dialogs/common_dialogs.py` | Dialog standardizzati via `tr(...)` | Sostituzioni stringhe | Messaggistica EN-first | Basso | i18n puro |
| `ui/dialogs/manage_supplier_categories_dialog.py` | Label/messaggi dialog categorie | Sostituzioni stringhe | UI invariata, testo EN-first | Basso | i18n puro |
| `ui/dialogs/potential_supplier_dialog.py` | Messaggi/suggerimenti duplicati fornitore | Sostituzioni stringhe | UX invariata, testo EN-first | Basso | i18n puro |
| `ui/dialogs/rfq_pdf_export_dialog.py` | Titoli e note export PDF + branch lingua | Sostituzioni + rimozione ternaria lingua | Stessa UX, un’unica source EN | Basso | i18n + piccolo branch UI |
| `ui/dialogs/vsm_event_dialog.py` | Dialog VSM con mapping display/internal | Sostituzioni + compatibilità mapping | Business invariato, display EN-first | Medio | i18n + compatibilità legacy minima |
| `ui/kpi_window.py` | Label tabs/cards/KPI via `tr(...)` | Sostituzioni stringhe | KPI UI EN-first | Basso | i18n puro |
| `ui/main_dashboard_builder.py` | Tab dashboard e filtri avanzati | Sostituzioni + source option EN (`All`) | Rendering tab/filtri invariato | Basso | i18n + minima logica UI |
| `ui/windows/attachment_window.py` | Messaggi attach + guard read-only | Sostituzioni stringhe | Flusso attach invariato | Basso | i18n puro |
| `ui/windows/edit_reference_window.py` | Dialog modifica riferimento | Sostituzioni stringhe | UI invariata | Basso | i18n puro |
| `ui/windows/edit_suppliers_window.py` | Dialog fornitori e warning duplicati | Sostituzioni stringhe | UX invariata | Basso | i18n puro |
| `ui/windows/notes_window.py` | Label/messaggi note | Sostituzioni stringhe | UI invariata | Basso | i18n puro |
| `ui/windows/purchase_order_window.py` | Label/messaggi PO | Sostituzioni stringhe | UI invariata | Basso | i18n puro |
| `ui/windows/sqdc_analysis_window.py` | Label/messaggi SQDC | Sostituzioni stringhe | UI invariata | Basso | i18n puro |
| `ui/windows/view_request_window.py` | Finestra ampia RFQ con molte label + branch export label | Sostituzioni + branch normalizzato | Flussi RFQ invariati, source EN unica | Medio | i18n + piccolo branch UI |
| `utils/i18n_utils.py` | Bridge canonical RFQ type -> traduzione | Compatibilità minima | Preserva canonical legacy, display EN-first | Medio | i18n + compatibilità legacy |

---

## C. Esempi Concreti

### C1. Sostituzione semplice `tr("...")`

1) 
- **file**: `dataflow.py`
- **contesto**: titolo finestra principale
- **prima**: `self.root.title(tr("DataFlow Procurement Software - Cruscotto Principale"))`
- **dopo**: `self.root.title(tr("DataFlow Procurement Software - Main Dashboard"))`
- **motivo**: source string EN-first.

2)
- **file**: `ui/windows/attachment_window.py`
- **contesto**: warning selezione allegato
- **prima**: `tr("Seleziona un allegato da eliminare.")`
- **dopo**: `tr("Select an attachment to delete.")`
- **motivo**: migrazione meccanica i18n.

3)
- **file**: `ui/dialogs/rfq_pdf_export_dialog.py`
- **contesto**: titolo dialog
- **prima**: `self.title(tr("Esporta RFQ PDF"))`
- **dopo**: `self.title(tr("Export RFQ PDF"))`
- **motivo**: sorgente EN stabile per gettext.

### C2. Confronto/branch dipendente lingua reso EN-first

1)
- **file**: `services/dashboard_controller.py`
- **contesto**: filtro tipo in ricerca
- **prima**: `if self.app.search_tipo.get() != tr("Tutte"):`
- **dopo**: `if self.app.search_tipo.get() != tr("All"):`
- **motivo**: confronto indipendente da source IT.

2)
- **file**: `dataflow.py`
- **contesto**: filtro ripetitivo VSM
- **prima**: `want = repetitive_filter == tr("Sì")`
- **dopo**: `want = repetitive_filter == tr("Yes")`
- **motivo**: mantenere confronto coerente con nuove sorgenti EN.

3)
- **file**: `ui/windows/view_request_window.py`
- **contesto**: label export con branch lingua
- **prima**: `export_label = tr("Export") if get_current_language() == "en" else tr("Esporta")`
- **dopo**: `export_label = tr("Export")`
- **motivo**: eliminare branch IT/EN superfluo, single source EN.

### C3. Compatibilità legacy senza cambiare business logic

1)
- **file**: `utils/i18n_utils.py`
- **contesto**: traduzione tipo RFQ canonical legacy
- **prima**:
  - `if canonical == "Fornitura piena": return tr("Fornitura piena")`
  - `elif canonical == "Conto lavoro": return tr("Conto lavoro")`
- **dopo**:
  - `if canonical == "Fornitura piena": return tr("Full Supply")`
  - `elif canonical == "Conto lavoro": return tr("Work Order")`
- **motivo**: mantenere canonical DB legacy ma source EN-first.

2)
- **file**: `ui/dialogs/vsm_event_dialog.py`
- **contesto**: mapping action display -> internal
- **prima**: `if display == tr("Negoziazione"): return "Negoziazione"`
- **dopo**: `if display == tr("Negotiation"): return "Negoziazione"`
- **motivo**: internal legacy invariato, solo display/source aggiornato.

3)
- **file**: `ui/dialogs/vsm_event_dialog.py`
- **contesto**: mapping driver display -> internal
- **prima**: `if display == tr("Prezzo"): return "Prezzo"`
- **dopo**: `if display == tr("Price"): return "Prezzo"`
- **motivo**: compatibilità con dati/modello esistente senza cambiare logica.

---

## D. Modifiche Non Banali (nei `.py`)

1. **Normalizzazione confronti filtro tipo (`All`)**
- **dove**: `services/dashboard_controller.py`, `dataflow.py`, `ui/main_dashboard_builder.py`
- **necessità**: i confronti basati su source IT (`Tutte`) avrebbero mantenuto coupling IT-first.
- **rischio introdotto**: basso.
- **perché low-risk**: stessa semantica, stessa variabile UI, solo chiave sorgente cambiata.

2. **Normalizzazione filtro ripetitivo (`Yes`)**
- **dove**: `dataflow.py`
- **necessità**: evitare dipendenza da literal italiano nel branch.
- **rischio**: basso-medio (coinvolge filtro runtime).
- **perché low-risk**: modifica puntuale, stesso flow e stessa condizione logica.

3. **Rimozione branch lingua esplicito nelle label export/PDF**
- **dove**: `ui/windows/view_request_window.py`, `ui/dialogs/rfq_pdf_export_dialog.py`
- **necessità**: eliminare path distinti IT/EN e usare una sola source EN tradotta da gettext.
- **rischio**: basso.
- **perché low-risk**: branch solo testuale, nessun effetto su business flow.

4. **Bridge compatibilità RFQ type in `translate_rfq_type`**
- **dove**: `utils/i18n_utils.py`
- **necessità**: mantenere canonical legacy (`Fornitura piena`/`Conto lavoro`) ma display EN-first.
- **rischio**: medio.
- **perché low-risk**: funzione dedicata, invariato il valore canonical in ingresso/uscita logica, cambia solo key di traduzione.

5. **Mapping display/internal in VSM dialog**
- **dove**: `ui/dialogs/vsm_event_dialog.py`
- **necessità**: non rompere internal value legacy mentre si passa a source EN.
- **rischio**: medio.
- **perché low-risk**: mapping esplicito, comportamento simmetrico mantenuto.

---

## E. Modifiche Fuori Scope Potenziali

- **Esito**: non emergono modifiche Python chiaramente fuori scope rispetto alla migrazione i18n EN-first.
- Le differenze rilevate sono coerenti con:
  - sostituzione sorgenti testo,
  - normalizzazione confronti lingua-dipendenti,
  - compatibilità minima su mapping legacy.
- **Nota**: non risultano refactor architetturali, né variazioni di algoritmi business, né introduzione dipendenze.

---

## F. Chiusura

1. **Il diff Python è prevalentemente meccanico i18n?**
- Sì. Il diff è prevalentemente meccanico i18n (sostituzioni di stringhe sorgente `tr(...)`), con pochi interventi mirati su confronti/branch lingua e compatibilità legacy.

2. **File che meritano review manuale prioritaria**
- `dataflow.py` (ampiezza diff elevata, include confronti filtro e flussi dashboard/VSM/export)
- `ui/windows/view_request_window.py` (molte stringhe + branch label export)
- `ui/dialogs/vsm_event_dialog.py` (mapping display/internal legacy)
- `utils/i18n_utils.py` (bridge canonical RFQ type)
- `services/dashboard_controller.py` (confronti filtro `All`)

3. **Punti a rischio più alto in test manuale**
- Filtri dashboard RFQ/VSM che confrontano valori localizzati (`All`, `Yes`).
- Dialog VSM su action/driver (display EN vs internal legacy IT).
- Flussi export (Excel/PDF) con label dinamiche e messaggistica condizionale.

