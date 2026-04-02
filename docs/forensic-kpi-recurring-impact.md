# Report Forensico: KPI DataFlow — Riverbero OPEX Ripetitivo 2027

Data analisi: 2 aprile 2026

---

## A. Risposta Breve

| Domanda | Risposta |
|---|---|
| Il riverbero 24 mesi è implementato? | **Sì**, completamente, nel VSM Engine e nella tabella DB |
| Il 2027 esiste nei dati KPI (vsm_impacts)? | **Sì** — i record con `anno=2027` esistono fisicamente nel DB |
| Il problema è nel calcolo o nella visualizzazione? | **Esclusivamente nella visualizzazione** — 4 difetti distinti, tutti in reporting/UI |

---

## B. Evidenze dal Codice

### B.1 — Dove il riverbero è implementato (correttamente)

| File | Funzione | Responsabilità |
|---|---|---|
| `services/vsm_engine.py` | `_calculate_distribution_months()` | Genera lista di 24 tuple `(year, month)` a partire da `event_date` se `opex_ripetitivo=True` |
| `services/vsm_engine.py` | `_distribute_value()` | Distribuisce il valore totale su quei 24 mesi con pro-rata primo mese |
| `services/vsm_engine.py` | `generate_impacts_for_event()` | Assembla i `VSMImpact` con `year` e `month` del periodo reale (es. 2027-01) |
| `models/vsm_impact.py` | `VSMImpact` (dataclass) | Modello con campi `year: int` e `month: int` — contengono l'anno/mese di impatto effettivo, non quello dell'evento |

**Prova concreta**: un evento `Saving`, `opex_ripetitivo=True`, `event_date=2026-03-15` genera 24 `VSMImpact` con `(year=2026, month=3)` … `(year=2027, month=2)`. I record con `anno=2027` sono nel DB nella tabella `vsm_impacts`.

---

### B.2 — Dove il 2027 viene perso (4 difetti)

#### Difetto 1 — Year selector non include il 2027
**File:** `services/kpi_engine.py` — `get_available_years()` (riga ~583)

```python
def get_available_years(db_path=None) -> list:
    c.execute(
        "SELECT DISTINCT strftime('%Y', event_date) "
        "FROM vsm_events WHERE event_date IS NOT NULL"
    )
```

Interroga `vsm_events.event_date`, non `vsm_impacts.anno`. Se tutti gli eventi sono stati creati nel 2026, il combobox mostra solo `[2026]`. L'anno **2027 non appare mai** nel selettore, anche se `vsm_impacts` contiene righe con `anno=2027`.

---

#### Difetto 2 — Filtro KPI engine ancorato su event_date, non su vi.anno
**File:** `services/kpi_engine.py` — `_sum_impacts()` + `get_saving_kpi()`

```python
d_clauses, d_params = _build_date_filter("ve.event_date", date_from, date_to, year)
# produce: strftime('%Y', ve.event_date) = '2026'
```

La query SQL risultante:

```sql
SELECT SUM(vi.valore_teorico), SUM(vi.valore_effettivo)
FROM vsm_impacts vi
JOIN vsm_events ve ON vi.event_id = ve.event_id
WHERE vi.tipo_valore = 'Saving' AND strftime('%Y', ve.event_date) = '2026'
```

Con `year=2026`: somma **tutti** gli impatti da eventi del 2026, inclusi quelli con `vi.anno=2027`. La KPI card `recurring_impact` mostra quindi il **totale 24 mesi** (2026+2027), non solo la quota 2026 — valore sovrastimato.

Con `year=2027` (se selezionabile): mostrerebbe solo impatti da eventi creati nel 2027 — **i mesi 2027 degli eventi 2026 sarebbero completamente invisibili**.

---

#### Difetto 3 — Chart: bucket capped a TODAY e limitato all'anno selezionato
**File:** `services/kpi_chart_data.py` — `_build_month_buckets()`

```python
today_ym = _date.today().strftime('%Y-%m')   # oggi = '2026-04'
if year is not None:
    buckets = [f'{year:04d}-{m:02d}' for m in range(1, 13)]
    return [b for b in buckets if b <= today_ym]  # filtra i mesi futuri
```

Con `year=2026` e today=`2026-04-02`: i bucket sono `['2026-01', '2026-02', '2026-03', '2026-04']`. I mesi futuri 2026 (Mag-Dic) **e tutti i mesi 2027** sono assenti. Anche se il `lookup` SQL contiene chiavi `'2027-01'` … `'2027-02'`, queste vengono silently discarded dal `lookup.get(b, ...)` che itera solo sui bucket.

---

#### Difetto 4 — Chart data query: stesso filtro su ve.event_date
**File:** `services/kpi_chart_data.py` — `get_saving_chart_data()`

```python
clauses, params = _dclauses('ve.event_date', date_from, date_to, year)
w = _where(["vi.tipo_valore = ?"], clauses)
db.cursor.execute(
    f"""SELECT printf('%04d-%02d', vi.anno, vi.mese),
               SUM(vi.valore_teorico), SUM(vi.valore_effettivo)
        FROM vsm_impacts vi
        JOIN vsm_events ve ON vi.event_id = ve.event_id
        {w}
        GROUP BY vi.anno, vi.mese""",
    ['Saving'] + params,
)
lookup = { b: (float(t), float(a)) for b, t, a in db.cursor.fetchall() }
```

Il `lookup` SQL **contiene effettivamente** le chiavi `'2027-01'`/`'2027-02'` (per un evento 2026-03 ripetitivo) — la query è corretta in questo senso. Ma il loop finale:

```python
return [{'label': _label(b), 'theoretical': lookup.get(b, ...)[0], ...} for b in buckets]
```

itera solo sui `buckets` (max `'2026-04'`), quindi le chiavi 2027 nel `lookup` non vengono mai lette.

---

### B.3 — Excel Export

**File:** `services/kpi_excel_export.py` — `build_kpi_workbook()`

Riceve i dati già calcolati da `get_saving_kpi()` e non fa query proprie. Ha **gli stessi dati distorti** del Difetto 2: `recurring_impact` include il totale 24 mesi, non la quota per anno. Non ha colonne temporali per periodo, quindi il problema del Difetto 3/4 non si manifesta nell'export (ma la card aggregata è essa stessa imprecisa).

---

## C. Flusso Reale End-to-End

```
1. Il buyer inserisce un evento VSM Saving, opex_ripetitivo=True, event_date=2026-03-15
       ↓
2. vsm_engine.generate_impacts_for_event()
   → _calculate_distribution_months(): 24 tuple (2026-03) … (2027-02)
   → _distribute_value(): 24 quote mensili con pro-rata mese 1
   → 24 VSMImpact con year/month reali → salvati in vsm_impacts
       ↓
3. Apertura KpiWindow
   → _populate_year_filter() chiama get_available_years()
   → Query su vsm_events.event_date → anni: [2026]
   → Anno 2027 NON disponibile nel combobox (anche se vsm_impacts.anno ha 2027)
       ↓
4. Utente seleziona Year = 2026 → _load_kpi_data(year=2026)
       ↓
5. get_saving_kpi(year=2026)
   → _build_date_filter("ve.event_date", year=2026)
   → _sum_impacts: WHERE strftime('%Y', ve.event_date) = '2026'
   → Somma TUTTI gli impatti (inclusi mesi 2027) → recurring_impact = valore 24 mesi totale
   → Mostra nella card: valore SOVRASTIMATO per il 2026 (include anche 2027)
       ↓
6. _update_charts() → get_saving_chart_data(year=2026)
   → _build_month_buckets(year=2026):
       buckets = ['2026-01','2026-02','2026-03','2026-04']  ← cap a oggi
   → SQL lookup: contiene (lookup['2027-01'], lookup['2027-02']) ma questi NON vengono mai letti
   → Chart mostra: 4 barre (Gen-Apr 2026), nessun mese 2027
       ↓
7. Utente prova Year = 2027 → NON DISPONIBILE nel combobox → impossibile navigare
```

---

## D. Root Cause Più Probabile

**La causa principale è un'unica scelta di design replicata in tre punti del codice**: i filtri temporali usano `ve.event_date` (l'anno di creazione dell'evento) come proxy per tutti i calcoli KPI, invece di usare `vi.anno`/`vi.mese` (il periodo di competenza economica dell'impatto).

Questo è corretto per eventi one-shot (dove `event_date` e periodo di impatto coincidono), ma **strutturalmente sbagliato per eventi ricorrenti** il cui riverbero attraversa anni di calendario.

Il problema è rinforzato da due scelte ortogonali:

- `get_available_years()` guarda solo le creation dates negli eventi, non i periodi negli impatti.
- `_build_month_buckets()` cap a `today_ym` rende invisibili anche i mesi futuri dell'anno corrente (Mag-Dic 2026), non solo il 2027.

---

## E. Impatto Utente/Business

Il buyer oggi non può:

1. **Vedere i saving 2027 derivanti da contratti stipulati nel 2026** — i 12–24 mesi di carryover OPEX sono calcolati correttamente ma non sono raggiungibili tramite alcun filtro della UI KPI.

2. **Navigare al 2027 nel selettore Year** — l'anno 2027 non esiste nel combobox fino a quando il buyer non inserisce un evento con `event_date` nel 2027.

3. **Leggere correttamente la KPI card `recurring_impact` per Anno 2026** — il valore mostrato è la somma delle 24 quote mensili (2026+2027), **non** solo la quota 2026. Un saving di €12.000 annui su 2 anni compare come `recurring_impact ≈ €23.xxx` invece di ~€10.xxx (quota 2026 pro-rata).

4. **Vedere il grafico mensile completo dell'anno 2026** — il chart è troncato al mese corrente (Aprile), nascondendo i mesi futuri dell'anno in corso (Mag-Dic 2026).

---

## F. Fix Candidate (Solo Enunciazione, No Codice)

### Fix 1 — Minimo: includere vsm_impacts.anno nel Year selector

**Scope file:** `services/kpi_engine.py` — `get_available_years()`

`get_available_years()`: aggiungere una terza query `SELECT DISTINCT anno FROM vsm_impacts` all'unione degli anni.

| | |
|---|---|
| **Rischio regressione** | Basso — aggiunge anni al combobox, non rimuove nulla |
| **Pro** | Il buyer può selezionare 2027 e navigare ai dati di carryover già presenti nel DB |
| **Contro** | Senza il Fix 2, la vista 2027 mostrerebbe solo gli impatti da eventi creati nel 2027 (vuoti), non i carryover da 2026; richiede il Fix 2 per essere significativo |

---

### Fix 2 — Medio: cambiare il filtro temporale da event_date a vi.anno/vi.mese

**Scope file:** `services/kpi_engine.py` e `services/kpi_chart_data.py`

In `_sum_impacts` e nelle query chart, sostituire il filtro `strftime('%Y', ve.event_date) = ?` con `vi.anno = ?` (per il filtro year) e `vi.anno || '-' || printf('%02d', vi.mese)` per i range mensili.

| | |
|---|---|
| **Rischio regressione** | Medio — cambia la semantica dei KPI esistenti; la card `recurring_impact` anno 2026 mostrerebbe solo la quota effettivamente maturata nel 2026 (≠ valore attuale che include 2027). Richiede test. |
| **Pro** | Allineamento tra semantica business ("saving del 2027") e periodo di visualizzazione; risolve sia la card che il chart |
| **Contro** | Breaking change semantico per gli utenti abituati all'aggregazione corrente |

---

### Fix 3 — Complementare: rimuovere il cap a today nei bucket modalità Year

**Scope file:** `services/kpi_chart_data.py` — `_build_month_buckets()`

Quando `year is not None`, rimuovere il `filter b <= today_ym`. I mesi futuri sarebbero mostrati nel chart con valore 0 (comportamento già corretto per mesi senza dati — `lookup.get(b, (0.0, 0.0))`).

| | |
|---|---|
| **Rischio regressione** | Molto basso — impatta solo la presentazione (aggiunge barre a zero) |
| **Pro** | Il buyer vede l'intero anno 2026 nel chart; se combinato con Fix 2, vedrà anche le quote 2027 |
| **Contro** | Il chart mostra mesi con saldo zero (potenzialmente confusivo), ma è corretto e informativo |
