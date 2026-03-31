# Plan: KPI Derisking — Anagrafica Fornitori Potenziali

## TL;DR
Sostituire i calcoli KPI Derisking basati su `vsm_events` con quelli basati su `potential_suppliers`. Nessuna metrica economica. Card dinamiche per stato. Grafico e tabella per categoria. 4 file coinvolti.

## Steps

### Phase 1 — `services/kpi_engine.py` — replace `get_derisking_kpi()`
1. Sostituire l'intera funzione `get_derisking_kpi()` (righe 465–516) con una nuova che legge da `potential_suppliers`:
   - `total_suppliers`: `SELECT COUNT(*) FROM potential_suppliers`
   - `unique_categories`: `COUNT(DISTINCT TRIM(category))` escludendo null/vuoti
   - `status_counts`: `GROUP BY supplier_status ORDER BY COUNT(*) DESC` → dict[str, int]
   - `category_counts`: `GROUP BY TRIM(category) ORDER BY COUNT(*) DESC` → dict[str, int]
   - Firma invariata (date_from/date_to/year accettati ma ignorati per compatibilità)
   - Struttura risultato: `{total_suppliers, unique_categories, status_counts, category_counts}`

### Phase 2 — `services/kpi_chart_data.py` — replace `get_derisking_chart_data()`
2. Sostituire la funzione `get_derisking_chart_data()` (righe 300–338):
   - Rimuovere dipendenza da `_build_month_buckets`, `vsm_events`, `event_type`
   - Query: `SELECT TRIM(category), COUNT(*) FROM potential_suppliers WHERE ... GROUP BY TRIM(category) ORDER BY COUNT(*) DESC`
   - Restituire `list[{'label': str, 'count': int}]` — stesso formato usato da `draw_bar_chart`
   - Firma invariata (date params accettati ma ignorati)

### Phase 3 — `ui/kpi_window.py` — 5 modifiche coordinate

3a. **`__init__`**: aggiungere attributo `self._derisking_status_frame = None` per il frame dinamico degli stati (riga ~170, dopo `self._derisking_labels: dict = {}`).

3b. **`_build_tab_derisking()`** (riga 346–350): ristrutturare completamente senza usare `_build_section()` (per poter ospitare card dinamiche). Costruire manualmente:
   - `LabelFrame` KPI con 2 card fisse:
     - `total_suppliers` ("Totale Fornitori Potenziali")
     - `unique_categories` ("Categorie Uniche")
     - Salvarle in `self._derisking_labels["total_suppliers"]` e `self._derisking_labels["unique_categories"]`
   - `ttk.Frame` vuoto `self._derisking_status_frame` per card dinamiche (riga subordinata)
   - `LabelFrame` Chart con `tk.Canvas` registrato in `self._chart_canvases['derisking']`
   - `LabelFrame` Details con Treeview registrato in `self._detail_trees['derisking']`

   Strategia: **Opzione A** — `_build_tab_derisking` costruisce tutto manualmente, NON usa `_build_section`. Zero rischi di regressione sugli altri tab.

3c. **`_update_derisking_cards()`** (riga 942): riscrivere:
   - Aggiornare card fisse con `_fmt_int(data.get("total_suppliers", 0))` e `_fmt_int(data.get("unique_categories", 0))`
   - Svuotare `self._derisking_status_frame` (distruggere tutti i widget figli)
   - Per ogni `(stato, count)` in `data.get("status_counts", {}).items()` creare una KPI card via `_build_kpi_card()` e aggiornarne subito il valore

3d. **`_build_detail_table()`** — ramo `elif section_key == 'derisking'` (riga 490–494): cambiare col_specs da `(period, Nuovi Fornitori)` a:
   ```python
   col_specs = [
       ('category', _t_ui(is_ita, 'Categoria', 'Category'),  160, 'w'),
       ('count',    _t_ui(is_ita, 'Fornitori',  'Suppliers'), 100, 'center'),
   ]
   ```
   Il metodo `_populate_table()` usa `d['label']` e `d['count']` — già compatibile col nuovo formato.

3e. **`_render_derisking_chart()`** (riga ~928): aggiornare titolo e label:
   - title: `"Fornitori per categoria"` / `"Suppliers per category"`
   - y_label: `"Fornitori"` / `"Suppliers"`
   - x_label: `"Categoria"` / `"Category"`

### Phase 4 — `services/kpi_excel_export.py` — update `_rows_derisking()`

4. Sostituire `_rows_derisking()` (righe 257–265):
   - Riga fissa: `total_suppliers` → `_t(is_ita, "Totale Fornitori Potenziali", "Total Potential Suppliers")`
   - Riga fissa: `unique_categories` → `_t(is_ita, "Categorie Uniche", "Unique Categories")`
   - Righe dinamiche: per ogni `(stato, count)` in `data.get('status_counts', {}).items()` → una riga con etichetta = `stato` e valore = count
   - `_build_derisking()` non cambia (itera già su `_rows_derisking`)

---

## Relevant Files

- `services/kpi_engine.py` — sostituire `get_derisking_kpi()` (~riga 465)
- `services/kpi_chart_data.py` — sostituire `get_derisking_chart_data()` (~riga 300)
- `ui/kpi_window.py` — 5 punti: `__init__`, `_build_tab_derisking`, `_update_derisking_cards`, `_build_detail_table`, `_render_derisking_chart`
- `services/kpi_excel_export.py` — sostituire `_rows_derisking()` (~riga 257)

NON toccare:
- KPI RFQ, Saving, Cost Avoidance
- `_build_section()` (codice condiviso da tutti i tab)
- Global search, logica VSM, export Excel salvo `_rows_derisking`

---

## Verification

1. `python3 -m unittest discover -s tests -q` → 63+ OK
2. Aprire finestra KPI, tab Derisking:
   - card "Totale Fornitori Potenziali" mostra count corretto
   - card "Categorie Uniche" mostra count corretto
   - card per ogni stato presente (es. "Nuovo — 2", "Qualificato — 1")
   - nessuna card per stati assenti
3. Grafico bar chart mostra fornitori per categoria (desc per count)
4. Tabella Details mostra `Categoria | Fornitori` (non più `Periodo | Nuovi Fornitori`)
5. Nessuna metrica economica visibile

---

## Decisions

- Filtri temporali (date_from/date_to/year) ignorati per `potential_suppliers` (nessuna data KPI rilevante); accettati per firma invariata
- Ordinamento stati: per numerosità decrescente (`ORDER BY COUNT(*) DESC`)
- Ordinamento categorie: per numerosità decrescente (`ORDER BY COUNT(*) DESC`)
- Card dinamiche ricostruite da zero a ogni `_update_derisking_cards()` (semplice, nessun caching)
- `_build_section()` non modificata: usata invariata dagli altri 3 tab (zero rischio regressione)
- `_populate_table('derisking', data)` già compatibile: usa `d['label']` e `d['count']`

## Considerazione critica — costruzione tab Derisking

Il metodo `_build_section()` costruisce TUTTO in sequenza: cards_frame, chart, table.
Per aggiungere card dinamiche serve accesso al `cards_label_frame` interno.

Opzioni:
- **A** (scelta): `_build_tab_derisking` costruisce manualmente cards+chart+table — più verbose ma zero rischio di regressioni su altri tab.
- **B**: modificare `_build_section` per accettare un callback post-cards — più pulito ma tocca codice condiviso da tutti i tab.

**Scelta: Opzione A**.

## Limite residuo

Il filtro Anno/Periodo nella KpiWindow non ha effetto sul tab Derisking (i fornitori potenziali non hanno una `data_creazione` KPI-significativa). Le card mostrano sempre il totale complessivo. Questo è il comportamento corretto per questa versione.
