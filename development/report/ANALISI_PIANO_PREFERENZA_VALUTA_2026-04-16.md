# DataFlow — Analisi + Piano minimo (Preferenza Globale Valuta)

Data: 2026-04-16  
Modalità: solo analisi e piano (nessuna modifica codice)

## A. Analisi stato attuale

### A1) Dove vengono salvate oggi le impostazioni utente

- Persistenza principale: file `config.ini` in path cross-platform gestito da `utils/user_utils.py` (`get_app_data_dir()`, `get_config_file()`).
- Sezioni/chiavi oggi rilevate da codice:
- `User`: `first_name`, `last_name`, `username`.
- `Settings`: `language`, `dataflow_base_dir`, `custom_db_path`, `license_accepted`, `vsm_pagamenti_coefficient`, `rfq_pdf_logo_file`.
- `AutoBackup`: `enabled`, `hour`, `path`.
- Finestra impostazioni attuale (`SettingsWindow` in `dataflow.py`) espone oggi:
- posizione DataFlow,
- backup manuale/automatico,
- lingua.

### A2) Dove viene gestita oggi la formattazione importi

- Formatter condiviso in `utils/format_utils.py`:
- `format_currency_display(val, show_symbol=True)` (migliaia + decimali IT, simbolo opzionale).
- Formatter KPI locale in `ui/kpi_window.py`:
- `_fmt_money(v, with_symbol=True)` (attualmente 0 decimali, separatore migliaia, simbolo opzionale).
- Duplicazione reale: esistono due formatter monetari distinti (`format_currency_display` e `_fmt_money`) con regole diverse (2 decimali vs 0 decimali).

### A3) Punti dove compare simbolo valuta (hardcoded o implicito)

#### Tabelle principali (tksheet / treeview)

- Tab Saving / Cost Avoidance (tksheet VSM): popolamento in `dataflow.py` (`_populate_vsm_sheet`) tramite `format_currency_display(..., show_symbol=False)`.
- Stato attuale: celle tabellari VSM senza simbolo (già numeriche formattate).
- Tab Derisking (tksheet supplier): nessuna colonna importo.

#### KPI UI

- `ui/kpi_window.py`:
- `_fmt_money(..., with_symbol=True)` usato nelle KPI cards Saving/CA (mostra `€` se default).
- Tabella `Details` Saving/CA usa `_fmt_money(..., with_symbol=False)` (quindi senza simbolo nelle celle tabellari).
- Label KPI hardcoded con `(€)` in alcune card label (`Recurring Impact (€)`, ecc.).
- Assi grafici hardcoded con `Saving (€)` e `Cost Avoidance (€)`.

#### Export Excel

- Export VSM (`dataflow.py`, `_export_vsm_excel`):
- scrive valori numerici raw (float) con `number_format '#,##0.00'`, nessun simbolo in cella.
- Export Derisking (`_export_derisking_excel`): nessun importo monetario.
- Export KPI (`services/kpi_excel_export.py`):
- valori monetari numerici con `_FMT_MONEY = '#,##0.00'` (nessun simbolo nel formato cella),
- ma nomi KPI includono stringhe con `(€)` in più righe (`_rows_saving`, `_rows_ca`).

#### Altri punti UI non tabellari

- `ui/dialogs/vsm_event_dialog.py`: label campo `Annual Spending (€): *`.

### A4) Gestione attuale allineamento colonne

- tksheet VSM (`_create_vsm_event_sheet` in `dataflow.py`):
- colonna importi (indici 4,5) oggi incluse in `align_cols` con `align='center'`.
- allineamento viene riapplicato anche in `_populate_vsm_sheet` (`align='center'`).
- tksheet Derisking (`_create_supplier_sheet`): non ha colonne importo.
- Treeview KPI Details (`ui/kpi_window.py`):
- colonne importo `theor`/`actual` hanno anchor `'e'` (destra) già configurato.

---

## B. Proposta architettura minima

### B1) Persistenza preferenza valuta (senza rompere nulla)

- Riutilizzare `config.ini`, sezione `Settings` (nessun DB, nessuna migration schema).
- Nuova chiave proposta: `currency_display`.
- Valore default: `NONE` (nessuna valuta).
- Valori ammessi (conservativi): `NONE`, `EUR`, `USD` (eventualmente estendibili).

### B2) Centralizzazione minima formattazione importi

- Definire un unico punto di risoluzione preferenza valuta (read-only) da `Settings`.
- Convergere gradualmente i call-site su un formatter unico “display-only” parametrico:
- input: valore numerico + currency code,
- output: stringa locale coerente,
- senza alterare il dato raw.
- Mantenere compatibilità con formatter esistenti tramite wrapper/parametro, evitando refactor estesi.

### B3) Evitare duplicazioni

- Eliminare divergenza tra `format_currency_display` e `_fmt_money` con strategia incrementale:
- prima standardizzare nei punti obbligatori (tabelle principali, KPI, export KPI/VSM),
- lasciare altri contesti invariati finché non richiesti.

### B4) Garanzia sorting/dati invariati

- Dati persistiti: invariati (nessuna modifica DB/model/calcoli).
- Calcoli KPI/VSM: invariati (solo presentation layer).
- Sorting tabelle: rischio se ordinamento usa stringhe renderizzate.
- Mitigazione minima: verificare sorting su colonne importo con `NONE` e `EUR`; se necessario, usare sort key numerica raw già disponibile nel layer dati (senza cambiare contenuto salvato).
- Nota: comportamento sorting numerico tksheet su stringhe formattate va verificato manualmente (non inferibile con certezza dal solo codice letto).

---

## C. UI — proposta inserimento valuta

### C1) Posizionamento combobox in Settings

- Inserire la selezione valuta dentro `SettingsWindow` esistente (`dataflow.py`), senza nuove finestre.
- Posizione consigliata: nuova riga/sezione compatta subito sotto “Language” (stesso pattern `ttk.Label + ttk.Combobox + Save`).
- Opzioni combobox: `No currency` (default), `EUR (€)`, `USD ($)`.

### C2) Impatto layout

- Impatto basso: la finestra usa container verticali (`pack`) e viene centrata dinamicamente.
- Aumento altezza moderato; non richiede redesign strutturale.

### C3) Strategia UI compatta/coerente

- Riutilizzare stile, spacing e bottoni già presenti in Settings.
- Nessuna nuova action globale: salvataggio nello stesso flusso configurazione (`config.ini`).
- Terminologia coerente con localizzazione EN/IT esistente.

---

## D. Allineamento a destra (importi)

### D1) Come allineare a destra

- tksheet (tab Saving/CA): aggiornare colonne importo in configurazione allineamento da `center` a destra (`'e'`/`right`, secondo API già usata da progetto).
- Treeview KPI Details: già a destra (`anchor='e'`) per teorico/effettivo; mantenere.

### D2) Colonne da modificare

- VSM tksheet Saving/CA: colonne “Theoretical” e “Actual” (indici 4 e 5 nel layout attuale).
- KPI Details Treeview: nessuna modifica necessaria (già destra).
- Derisking: nessuna colonna importo (non applicabile).

### D3) Rischi layout/UX

- Rischio lieve di percezione “salti” visivi su larghezze dinamiche VSM.
- Mitigazione: non cambiare dimensionamenti, solo anchor/allineamento.

---

## E. Export Excel

### E1) Dove viene generato oggi

- Export VSM/Derisking: `dataflow.py` (`_export_vsm_excel`, `_export_derisking_excel`).
- Export KPI: `ui/kpi_window.py` (trigger) + `services/kpi_excel_export.py` (builder workbook).

### E2) Applicare valuta senza rompere numeri Excel

- Regola fondamentale: mantenere valore cella numerico (float/int), agendo solo su `number_format`.
- `NONE`: formato numerico corrente (`#,##0.00`).
- Valuta selezionata: formato numerico con simbolo in `number_format` (no concatenazione stringhe nel valore).
- Così restano ordinabilità, formule e filtri numerici Excel.

### E3) Strategie in caso conflitto

- Strategia preferita: numero puro + formato cella (currency-like).
- Strategia da evitare: serializzare come stringa con simbolo (rompe calcoli/ordinamento Excel).

---

## F. Piano implementazione (step-by-step, piccoli e reversibili)

### A1

- Mappare e confermare enum preferenza (`NONE`, `EUR`, `USD`) e default `NONE`.
- Test: avvio senza chiave in config → fallback coerente.

### A2

- Estendere `SettingsWindow` (solo UI + persistenza chiave `Settings.currency_display`).
- Test: salvataggio/riapertura settings mantiene selezione.

### B1

- Introdurre accessor centralizzato della preferenza valuta (read config, fallback `NONE`).
- Test: lettura robusta con valore mancante/non valido.

### B2

- Allineare formatter display monetario nei punti target (VSM tabelle + KPI cards/details) usando preferenza globale.
- Test: toggle `NONE`↔`EUR` aggiorna resa senza cambiare numeri.

### C1

- Applicare allineamento destro colonne importo in tksheet Saving/CA (creazione + repopulate).
- Test: refresh/search/sort non perde allineamento.

### C2

- Verifica Derisking: nessuna colonna importo, nessuna modifica funzionale.
- Test: regressione zero su tab Derisking.

### D1

- Applicare preferenza valuta al rendering KPI (cards + details table) in modo coerente.
- Test: `NONE` senza simboli, `EUR/USD` con simboli.

### D2

- Applicare preferenza valuta a export VSM e KPI via `number_format` (valore numerico invariato).
- Test: in Excel i valori restano numerici (sum/sort/filter funzionanti).

### E1

- Allineare label testuali KPI con politica valuta (es. `(€)` dinamico o neutro), solo nei punti target.
- Test: lingua EN/IT + valuta NONE/EUR.

### E2

- Smoke test cross-platform (Linux/Windows) su apertura settings, rendering tab, export.
- Test: nessuna dipendenza nuova, nessuna modifica schema DB.

---

## G. Rischi e mitigazioni

- Rischio tecnico: divergenza formatter (2 funzioni monetarie con regole diverse).
- Mitigazione: unificare progressivamente tramite punto centralizzato e parametri espliciti.

- Rischio tecnico: sorting importi su stringhe formattate nelle tabelle tksheet.
- Mitigazione: test manuali su sort asc/desc; se necessario, key numerica raw senza cambiare storage.

- Rischio UX: incoerenza tra celle, card KPI, label assi e label KPI con `(€)`.
- Mitigazione: definire policy chiara:
- `NONE` => UI neutra (niente simboli nei valori),
- valuta selezionata => simbolo solo nei valori e/o anche nelle label secondo decisione prodotto esplicita.

- Rischio regressione export: simbolo inserito come testo (non formato).
- Mitigazione: usare solo `number_format` mantenendo cella numerica.

- Rischio i18n: nuove stringhe non tradotte in Settings.
- Mitigazione: aggiungere solo poche label riusando pattern traduzioni esistente.

---

## H. Checklist test manuale

1. Default senza valuta
- Aprire app con config senza `currency_display`.
- Verificare tabelle Saving/CA, KPI details/cards, export preview: valori senza simbolo.

2. Selezione EUR
- Impostare `EUR` in Settings.
- Verificare simbolo su valori nei punti target e allineamento destro colonne importo in tabelle.

3. Ritorno a Nessuna valuta
- Reimpostare `No currency`.
- Verificare scomparsa simboli e persistenza dopo riavvio.

4. KPI coerenti
- Confrontare stesso dataset con `NONE` e `EUR`: numeri identici, cambia solo rappresentazione.

5. Export Excel coerente
- Export VSM + KPI con `NONE` e `EUR`.
- In Excel: celle importo restano numeriche (SUM, sort, filter ok).

6. Allineamento colonne
- Tab Saving/CA: colonne importo allineate a destra dopo load, refresh, search, sort.
- KPI Details: confermare destra già mantenuta.

---

## Note di chiarezza (punti non completamente verificabili da sola lettura codice)

- Ordinamento numerico tksheet sulle colonne importo formattate come stringa non è garantibile al 100% da analisi statica: richiede test manuale esplicito.
- Ambito “KPI” su label assi/label card con `(€)` va confermato a livello prodotto: il requisito parla di visualizzazione valuta, ma non esplicita se tutte le label testuali debbano diventare dinamiche o restare neutre.
