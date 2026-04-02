# Piano Tecnico: Fix KPI Semantics + Carry-over to Next Year

---

## 1. Executive Summary

**Cosa va cambiato:** tre aree distinte — (a) il filtro temporale dei KPI monetari, che deve passare dall'anno di creazione dell'evento (`ve.event_date`) al periodo di competenza dell'impatto (`vi.anno`/`vi.mese`); (b) il popolamento del selettore Year e dei bucket del chart, che devono includere anni futuri già presenti in `vsm_impacts`; (c) l'aggiunta di una nuova KPI `carry_over_to_next_year`.

**Perché:** la semantica "anno 2026 = quota economica maturata nel 2026" è l'unica coerente con la distribuzione temporale già calcolata dal VSM Engine. L'implementazione attuale aggrega per anno di inserimento evento, ignorando la distribuzione temporale pre-calcolata.

**Rischio principale:** i valori monetari nelle card `recurring_impact` / `non_recurring_impact` cambieranno per tutti gli anni già in uso. Questo è il comportamento corretto ma costituisce un **breaking change semantico osservabile dall'utente.** I valori precedenti erano sistematicamente sovrastimati (includevano porzioni future).

---

## 2. Functional Target

| Elemento UI | Comportamento attuale (errato) | Comportamento target (corretto) |
|---|---|---|
| **Year selector** | Contiene solo gli anni presenti in `vsm_events.event_date` | Contiene l'unione degli anni da `vsm_events.event_date`, `richieste_offerta.data_emissione`, `potential_suppliers.created_at` **e** `vsm_impacts.anno` |
| **KPI cards — valori monetari** | Somma tutti gli impatti da eventi creati nell'anno N (include porzioni 2027) | Somma gli impatti con periodo di competenza nell'anno N |
| **KPI card — percentuali** | Statistiche calcolate su eventi creati nell'anno N | Invariate — rimangono su `ve.event_date` (semanticamente corrette per questa KPI) |
| **Carry-over to next year (€)** | Non esiste | Mostra il valore effettivo pianificato per l'anno N+1 da impatti già registrati; "—" se non è selezionato un anno specifico |
| **Chart** | Troncato al mese corrente; solo mesi dell'anno selezionato fino ad oggi | 12 mesi dell'anno selezionato (Gen–Dic), mesi futuri mostrati a 0; valori per competenza (`vi.anno`/`vi.mese`) |
| **Details table** | Popolata dagli stessi bucket del chart | Invariata strutturalmente — segue il chart (stessa source list) |
| **Export Excel** | Valori monetari sovrastimati; carry-over assente | Valori corretti per competenza; carry-over incluso se anno selezionato, omesso altrimenti |

---

## 3. Impacted Files and Responsibilities

| File | Funzioni/Classi | Motivo della modifica | Tipo modifica |
|---|---|---|---|
| `services/kpi_engine.py` | `_sum_impacts()` | Parametri `event_date_clauses`/`event_date_params` → sostituiti con clausole basate su `vi.anno`/`vi.mese` | Query filter semantics |
| `services/kpi_engine.py` | nuovo helper `_build_impact_period_filter()` | Costruisce clausole WHERE per `vi.anno`/`vi.mese` da `year` o `date_from`/`date_to` | Nuova funzione helper |
| `services/kpi_engine.py` | `get_saving_kpi()` | Usa impact-period filter per monetari; aggiunge query carry-over | Query filter + nuovo campo result |
| `services/kpi_engine.py` | `get_cost_avoidance_kpi()` | Idem | Query filter + nuovo campo result |
| `services/kpi_engine.py` | `get_available_years()` | Aggiunge query su `vsm_impacts.anno` all'unione | Year discovery |
| `services/kpi_chart_data.py` | `_build_month_buckets()` | Rimuove cap `b <= today_ym` quando `year is not None` | Bucket generation |
| `services/kpi_chart_data.py` | `get_saving_chart_data()` | Cambia `_dclauses('ve.event_date', ...)` in filtro per `vi.anno`/`vi.mese` | Query filter |
| `services/kpi_chart_data.py` | `get_cost_avoidance_chart_data()` | Idem | Query filter |
| `ui/kpi_window.py` | `_build_tab_saving()` | Aggiunge item carry-over alla lista `items` | UI card wiring |
| `ui/kpi_window.py` | `_build_tab_cost_avoidance()` | Idem | UI card wiring |
| `ui/kpi_window.py` | `_update_saving_cards()` | Aggiunge `carry_over_to_next_year` a `money_keys`; gestisce `None` → "—" | Card update logic |
| `ui/kpi_window.py` | `_update_ca_cards()` | Idem | Card update logic |
| `services/kpi_excel_export.py` | `_rows_saving()` | Aggiunge riga carry-over condizionale | Export mapping |
| `services/kpi_excel_export.py` | `_rows_ca()` | Idem | Export mapping |

**File NON coinvolti:** `services/vsm_engine.py`, `models/vsm_event.py`, `models/vsm_impact.py`, `database_manager.py`, tutti i file fuori scope KPI.

---

## 4. Proposed Implementation Plan

### Step 1 — `kpi_engine.py`: nuovo helper `_build_impact_period_filter()`
**Scopo:** costruire clausole SQL per filtrare su `vi.anno`/`vi.mese` in modo equivalente a quanto fa `_build_date_filter` per le date. Questo helper è interno (`_` prefix), analogo agli altri helper già presenti nel file.

**Logica attesa:**
- `year` selezionato → `vi.anno = ?`, params = `[year]`
- `date_from`/`date_to` → `printf('%04d-%02d', vi.anno, vi.mese) >= ? AND <= ?`, params = `[date_from[:7], date_to[:7]]` (troncato a `YYYY-MM`)
- Nessun filtro → clausole vuote, params vuoti

**File:** `services/kpi_engine.py`
**Dipendenze:** nessuna (step autonomo)
**Rischio regressione:** basso (funzione nuova, non modifica funzionante)
**Rollback:** cancellare il helper

---

### Step 2 — `kpi_engine.py`: modifica `_sum_impacts()`
**Scopo:** i parametri `event_date_clauses` / `event_date_params` diventano `period_clauses` / `period_params` e vengono applicati su `vi.anno`/`vi.mese` invece di `ve.event_date`. Il JOIN con `vsm_events` rimane necessario perché `opex_filter` filtra su `ve.opex_ripetitivo`.

**Attenzione:** questo step da solo NON rompe nulla finché i caller non vengono aggiornati nel passo successivo. Aggiornare il nome del parametro è sufficiente se si aggiornano i call-site contestualmente.

**File:** `services/kpi_engine.py`
**Dipendenze:** Step 1 (usa il nuovo helper nei call-site)
**Rischio regressione:** medio — cambia i valori aggregati
**Rollback:** ripristino parametro `event_date_clauses` e logica originale in una singola funzione

---

### Step 3 — `kpi_engine.py`: modifica `get_saving_kpi()` e `get_cost_avoidance_kpi()`
**Scopo:** per i totali monetari (`theoretical_saving`, `actual_saving`, `recurring_impact`, `non_recurring_impact`, e gli analoghi di CA), costruire le clausole di periodo tramite il nuovo `_build_impact_period_filter()` invece di `_build_date_filter("ve.event_date", ...)`.

**Le percentuali (`d_ev_clauses` / `d_ev_params`) rimangono invariate** — usano ancora `_build_date_filter("event_date", ...)` per filtrare gli eventi nella stessa finestra temporale. Questo è semanticamente corretto: "le % di risparmio sono caratteristiche degli eventi negoziati in quell'anno."

**Nota sull'existing `d_clauses` vs nuovo `impact_clauses`:** nelle due funzioni esistono già due fetch separati: uno con `_sum_impacts` (per monetari) e uno con `event_type` su `vsm_events` (per percentuali). Il refactoring riguarda solo la prima parte.

**File:** `services/kpi_engine.py`
**Dipendenze:** Step 1, Step 2
**Rischio regressione:** medio — valori monetari cambiano, percentuali invariate
**Rollback:** singolo cambio di quale helper costruisce `d_clauses`

---

### Step 4 — `kpi_engine.py`: aggiunta carry-over in `get_saving_kpi()` e `get_cost_avoidance_kpi()`
**Scopo:** aggiungere al `result` dict il campo `carry_over_to_next_year` (float se `year` selezionato, `None` altrimenti).

**Formula (Saving):**
```sql
SELECT COALESCE(SUM(vi.valore_effettivo), 0.0)
FROM vsm_impacts vi
WHERE vi.anno = (year + 1) AND vi.tipo_valore = 'Saving'
```
Nessun JOIN necessario — la query opera solo su `vsm_impacts`. L'indice esistente `idx_vsm_impacts_period` su `(anno, mese)` viene sfruttato automaticamente.

**Stesso pattern per Cost Avoidance** (`tipo_valore = 'Cost Avoidance'`, chiave `"carry_over_to_next_year"` nel dict CA).

**Se `year` è None:** `result["carry_over_to_next_year"] = None` — il valore `None` è il segnale per la UI di mostrare "—" invece di "0 €".

**File:** `services/kpi_engine.py`
**Dipendenze:** Step 3
**Rischio regressione:** basso — aggiunta di campo, nessuna modifica ai campi esistenti
**Rollback:** rimuovere la query e la chiave dal dict result

---

### Step 5 — `kpi_engine.py`: modifica `get_available_years()`
**Scopo:** aggiungere `SELECT DISTINCT anno FROM vsm_impacts WHERE anno IS NOT NULL` all'unione degli anni già interrogata.

**File:** `services/kpi_engine.py`
**Dipendenze:** nessuna (step autonomo)
**Rischio regressione:** basso — operazione additive (aggiunge anni, non rimuove)
**Rollback:** rimuovere la terza query dall'unione

---

### Step 6 — `kpi_chart_data.py`: rimozione cap `today_ym` in `_build_month_buckets()`
**Scopo:** quando `year is not None`, restituire tutti i 12 bucket dell'anno senza filtrare `b <= today_ym`. I mesi futuri senza dati avranno valore 0 nel chart (comportamento già corretto nel lookup dict).

**Attenzione:** il cap rimane in vigore per i preset rolling (`date_from`/`date_to`), dove il tetto al mese corrente rimane semanticamente corretto.

**File:** `services/kpi_chart_data.py`
**Dipendenze:** nessuna (step autonomo)
**Rischio regressione:** basso — impatta solo la presentazione visiva del chart
**Rollback:** ripristino di una riga (`return [b for b in buckets if b <= today_ym]`)

---

### Step 7 — `kpi_chart_data.py`: modifica filter in `get_saving_chart_data()` e `get_cost_avoidance_chart_data()`
**Scopo:** cambiare `_dclauses('ve.event_date', date_from, date_to, year)` in un filtro su `vi.anno`/`vi.mese`. Pattern analogo al Step 1 ma nel modulo `kpi_chart_data.py`.

**Opzioni di implementazione:** (1) duplicare il helper `_build_impact_period_filter` localmente nel modulo, oppure (2) importarlo da `kpi_engine.py`. Si raccomanda (1) per mantenere i due moduli indipendenti, coerentemente con lo stile esistente (entrambi definiscono già un proprio `_dclauses` e `_where`).

**Il GROUP BY di entrambe le funzioni è già corretto** (`GROUP BY vi.anno, vi.mese`) — non va cambiato.

**File:** `services/kpi_chart_data.py`
**Dipendenze:** Step 6
**Rischio regressione:** medio — cambia i valori nel grafico
**Rollback:** ripristino `_dclauses('ve.event_date', ...)`

---

### Step 8 — `kpi_window.py`: aggiunta card carry-over
**Scopo:** aggiungere `carry_over_to_next_year` nelle liste `items` di `_build_tab_saving()` e `_build_tab_cost_avoidance()`. Aggiornare `_update_saving_cards()` e `_update_ca_cards()` per gestire il valore `None` (mostrare "—") e `float` (formattare come moneta).

**Layout:** la griglia è a 4 colonne. Con 9 card (8 esistenti + 1 nuova), la terza riga avrà 1 card nelle prime 2 colonne e 2 vuote. Questo è accettabile e coerente con `_build_kpi_cards()` che usa `col = idx % cols` — non richiede alcuna modifica alla logica della griglia.

**File:** `ui/kpi_window.py`
**Dipendenze:** Step 4
**Rischio regressione:** basso — aggiunta, non modifica strutturale
**Rollback:** rimuovere l'item dalla lista

---

### Step 9 — `kpi_excel_export.py`: aggiunta carry-over nell'export
**Scopo:** aggiungere riga carry-over in `_rows_saving()` e `_rows_ca()`. La riga va inclusa solo se il valore nel dict non è `None` (cioè solo quando era selezionato un anno specifico). Se `None`, la riga viene omessa dall'export.

**File:** `services/kpi_excel_export.py`
**Dipendenze:** Step 4
**Rischio regressione:** basso — aggiunta condizionale, non modifica righe esistenti
**Rollback:** rimuovere la riga condizionale

---

## 5. Carry-over KPI Definition

**Cosa misura esattamente:** il valore effettivo (già pianificato nel DB) che sarà di competenza economica dell'anno N+1, derivante da tutti gli eventi già registrati con impatti futuri. Risponde alla domanda: *"Quanto saving/CA ho già 'in canna' per l'anno prossimo, da contratti già firmati?"*

**Su quali record si basa:**
```sql
SELECT COALESCE(SUM(valore_effettivo), 0.0)
FROM vsm_impacts
WHERE anno = (N + 1)
  AND tipo_valore = 'Saving'   -- o 'Cost Avoidance' rispettivamente
```
Nessun JOIN necessario. Nessun filtro su `opex_ripetitivo`: gli impatti in N+1 possono esistere solo per eventi ripetitivi (per definizione del VSM Engine), quindi il vincolo è implicito nella presenza dei record.

**Come si calcola per anno N:** una query diretta su `vsm_impacts.anno = N+1`. L'indice `idx_vsm_impacts_period` è su `(anno, mese)` → query efficiente.

**Se non esiste N+1 (nessun impatto futuro in DB):** `COALESCE(..., 0.0)` restituisce `0.0`. La UI mostra `"0 €"`, il che è semanticamente corretto (nessun carry-over pianificato).

**Se `year` è None (rolling o all):** la funzione restituisce `None` nel dict. La UI mostra `"—"`. L'export Excel omette la riga. Questo è corretto: "carry-over verso l'anno prossimo" è privo di significato senza un anno di riferimento.

**Se l'evento è one-shot (opex_ripetitivo=False):** il VSM Engine genera un solo impatto nel mese dell'evento. Nessun impatto in N+1 verrà mai generato per un evento one-shot. Il carry-over non è influenzato.

**Punto aperto da decidere prima dell'implementazione:** quando l'utente usa un preset rolling (`date_from`/`date_to`) invece di selezionare un anno specifico, il carry-over deve essere `None` (→ "—") oppure va calcolato come "valore anno N+1 dove N = anno del `date_to`"? Raccomandazione: **None** per semplicità e chiarezza semantica.

---

## 6. Query/Filter Semantics Review

### `_sum_impacts()` — clausole su `ve.event_date`

| | |
|---|---|
| **Comportamento attuale** | `WHERE vi.tipo_valore = ? AND strftime('%Y', ve.event_date) = '2026'` → somma tutti gli impatti da eventi creati nel 2026, inclusi quelli con `vi.anno=2027` |
| **Comportamento target** | `WHERE vi.tipo_valore = ? AND vi.anno = 2026` → somma solo gli impatti con periodo di competenza 2026 |
| **Motivazione** | Allineamento tra semantica business e granularità del DB; rimuove sovrastima sistematica per eventi ricorrenti |

### `get_saving_kpi()` — `d_clauses` per statistiche %

| | |
|---|---|
| **Comportamento attuale** | `SELECT ... FROM vsm_events WHERE event_type='Saving' AND strftime('%Y', event_date)='2026'` |
| **Comportamento target** | **Invariato** — le percentuali di saving sono caratteristiche degli eventi negoziati in quell'anno |
| **Motivazione** | "Average saving % 2026" = media % delle trattative condotte nel 2026, non degli impatti economici di quell'anno |

### `get_cost_avoidance_kpi()` — stesso pattern di `get_saving_kpi()`

Identico al punto precedente. Invariato per le statistiche %.

### `get_saving_chart_data()` — `_dclauses('ve.event_date', ...)`

| | |
|---|---|
| **Comportamento attuale** | `WHERE vi.tipo_valore='Saving' AND strftime('%Y', ve.event_date)='2026'` → include mesi 2027 nei dati ma il `lookup` li scarta perché i bucket si fermano a `today_ym` |
| **Comportamento target** | `WHERE vi.tipo_valore='Saving' AND vi.anno=2026` → solo gli impatti con competenza 2026 compaiono nel chart; bucket = tutti i 12 mesi del 2026 |
| **Motivazione** | Consistenza tra KPI cards e chart; eliminazione del "dato silenziosamente perduto" nel lookup |

### `get_cost_avoidance_chart_data()` — identico

---

## 7. UI Layout Plan

**Stato attuale Saving tab:** 8 card in griglia 4 colonne × 2 righe.
```
[ Theoretical ] [ Actual ] [ Avg % ] [ Best % ]
[ Worst %     ] [ Median ] [ Rec.  ] [ Non-Rec ]
```

**Stato target con carry-over:** 9 card, terza riga parziale.
```
[ Theoretical ] [ Actual ] [ Avg %  ] [ Best %    ]
[ Worst %     ] [ Median ] [ Rec.   ] [ Non-Rec.  ]
[ Carry-over  ] [         ] [        ] [           ]
```

La terza riga ha 1 card e 3 celle vuote. La funzione `_build_kpi_cards()` usa `col = idx % 4` e `row = idx // 4` — il posizionamento avviene automaticamente aggiungendo l'item alla lista. Non è necessaria alcuna modifica alla logica di griglia.

**Rischio overflow:** nessuno. Le card usano `sticky="ew"` con `columnconfigure(weight=1)` — la riga parziale si adatta correttamente.

**Etichetta nella card:**
- IT: `"Carry-over anno successivo (€)"`
- EN: `"Carry-over to next year (€)"`

I 18 caratteri + `(€)` rientrano nel limite senza wraplength issues (`wraplength=200` già impostato nelle card).

**Stessa struttura per Cost Avoidance tab:** identica modifica con stesso posizionamento.

---

## 8. Excel Export Impact

**Valori monetari che cambieranno dopo la correzione (Step 3):**

| Campo dict | Prima (errato) | Dopo (corretto) |
|---|---|---|
| `theoretical_saving` | Totale teorico da eventi del 2026 (include mesi 2027) | Solo quota teorica di competenza 2026 |
| `actual_saving` | Totale effettivo da eventi del 2026 (include mesi 2027) | Solo quota effettiva di competenza 2026 |
| `recurring_impact` | Effettivo ricorrente da eventi del 2026 (24 mesi) | Effettivo ricorrente con anno=2026 (quota anno) |
| `non_recurring_impact` | Effettivo one-shot da eventi del 2026 | Invariato in pratica (one-shot: sempre 1 impatto nello stesso anno) |
| Analoghi CA | Stessa distorsione | Stessa correzione |

**Percentuali (best/worst/avg/median):** invariate nell'export.

**Carry-over nel export:**
- Incluso in `_rows_saving()` e `_rows_ca()` **solo se il valore non è `None`** (`if data.get('carry_over_to_next_year') is not None`).
- Posizione: ultima riga di ciascuna sezione (dopo `non_recurring_impact`).
- Formato: `_FMT_MONEY` (già definito nel modulo).
- Il foglio Summary include la riga tramite `_rows_saving()` — nessuna modifica a `_build_summary` necessaria.

---

## 9. Validation Plan

### VT-1 — Setup dati di test
- Inserire un evento Saving, `opex_ripetitivo=True`, `event_date=2026-03-15`, importo_bdg=120.000 €, importo_negoziato=108.000 €, `percent_realizzo=100%` → valore teorico = 12.000 €
- Il VSM Engine genera 24 impatti: marzo 2026 (pro-rata) … febbraio 2028
- Verificare i valori attesi per anno 2026, 2027, 2028 prima di procedere

### VT-2 — Verifica Year selector
- Aprire KpiWindow quando l'unico evento ha `event_date=2026`
- Verificare che il combobox Year contenga `[2026, 2027, 2028]`
- Prima della fix: combobox contiene solo `[2026]`

### VT-3 — Verifica KPI card 2026 (Saving)
- Selezionare Year = 2026
- Verificare che `Theoretical Saving` ≈ quota di competenza 2026 (non somma 24 mesi)
- Verificare che `Recurring Impact` ≈ quota effettiva 2026

### VT-4 — Verifica KPI card 2027 (Saving)
- Selezionare Year = 2027
- Verificare che `Theoretical Saving` ≈ quota di competenza 2027
- Prima della fix: impossibile navigare al 2027

### VT-5 — Verifica chart anno completo 2026
- Selezionare Year = 2026 con today = 2 aprile 2026
- Verificare che il chart mostri **12 barre** (Gen–Dic), non 4
- Mesi Gen–Apr: valori > 0; Mag–Dic: 0 (nessun impatto futuro nel 2026 per questo evento)

### VT-6 — Verifica carry-over 2026
- Con Year = 2026: `Carry-over to next year (€)` mostra ≈ quota 2027
- Con Year = 2027: `Carry-over to next year (€)` mostra ≈ quota 2028
- Con filtro "12M" (rolling): `Carry-over to next year (€)` mostra "—"
- Con "All": `Carry-over to next year (€)` mostra "—"

### VT-7 — Anti-regressione: evento one-shot
- Inserire un evento Saving, `opex_ripetitivo=False`, `event_date=2026-05-10`, valore 5.000 €
- Con Year = 2026: `Theoretical Saving` include i 5.000 €
- Con Year = 2027: i 5.000 € NON compaiono
- `Carry-over to next year (€)` con Year = 2026: i 5.000 € NON contribuiscono

### VT-8 — Verifica export Excel
- Fare export con Year = 2026: `Saving → Recurring Impact` = valore corretto 2026 (non 24 mesi)
- Verifica presenza riga "Carry-over to next year (€)" nel foglio Saving e Cost Avoidance
- Fare export con filtro "12M": riga carry-over assente nel file Excel generato

### VT-9 — Anti-regressione RFQ e Derisking
- KPI sezione RFQ invariate (filtrano su `data_emissione`)
- KPI sezione Derisking invariate (filtrano su `created_at`)
- I count nel selettore Year per RFQ/Derisking non devono cambiare

---

## 10. Risks and Guardrails

**Rischio 1 — Semantico: breaking change sui valori monetari aggregati**
Il valore `recurring_impact` cambierà (scenderà) per anni già in uso.
*Guardrail:* comunicare esplicitamente nella release note che i valori riflettono ora la competenza economica annuale.

**Rischio 2 — UX: mesi futuri a zero nel chart**
Il chart anno 2026 mostrerà barre vuote per i mesi futuri — potenzialmente confuso.
*Guardrail:* i mesi futuri a zero sono già il comportamento standard per mesi senza dati. Nessuna modifica visiva necessaria.

**Rischio 3 — Export: carry-over omesso silenziosamente nel filtro rolling**
Se l'utente esporta con preset "12M", non troverà il carry-over.
*Guardrail:* la riga è omessa (non mostrata come 0) — coerentemente con la card UI che mostra "—". Comportamento previsto.

**Rischio 4 — Mismatch card/chart/table**
Steps 3 e 7 devono essere applicati insieme per mantenere coerenza tra card e chart.
*Guardrail:* Step 7 da eseguire sullo stesso commit di Step 3; test VT-3+VT-5 vanno verificati insieme.

**Rischio 5 — `date_from`/`date_to` rolling: conversione YYYY-MM**
Il troncamento a `YYYY-MM` per `_build_impact_period_filter` deve essere corretto sui border case.
*Guardrail:* il confronto `printf('%04d-%02d', vi.anno, vi.mese) >= '2025-04'` è lessicograficamente corretto se il formato è standard `YYYY-MM`. Verificare con preset "12M".

**Rischio 6 — `None` vs `0.0` in carry-over**
`"—"` = anno non selezionato; `"0 €"` = nessun carry-over pianificato. Semanticamente distinti.
*Guardrail:* il behavior è non ambiguo per design. Documentare se necessario.

---

## 11. Out of Scope

- `services/vsm_engine.py` — logica di distribuzione 24 mesi corretta, nessuna modifica
- `models/vsm_event.py`, `models/vsm_impact.py` — nessuna modifica
- `database_manager.py` — nessun cambio schema, nessuna migrazione
- `services/vsm_persistence.py` — nessuna modifica
- `services/supplier_persistence.py` — nessuna modifica
- `services/dashboard_controller.py` e tutta la dashboard principale — nessuna modifica
- `ui/kpi_chart.py` (`draw_bar_chart`, `draw_dual_bar_chart`) — ricevono già i dati come lista; nessuna modifica
- `get_derisking_kpi()` e `get_available_years_derisking()` — usano `created_at`, semanticamente corretto
- `get_rfq_kpi()` — usa `data_emissione`, semanticamente corretto
- `tests/test_vsm_engine.py` — la logica del motore non cambia
- Selettore nominativo (Username filter) nella dashboard
- Filtri avanzati VSM nella dashboard (date, azione, ripetitivo, importi)
