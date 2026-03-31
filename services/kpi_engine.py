# -*- coding: utf-8 -*-
"""
KPI Engine — Logica di calcolo KPI per DataFlow.

Fase 2: solo lettura e aggregazione dati, nessuna UI.

Funzioni pubbliche:
    get_rfq_kpi()            → KPI sezione RFQ
    get_saving_kpi()         → KPI sezione Saving
    get_cost_avoidance_kpi() → KPI sezione Cost Avoidance
    get_derisking_kpi()      → KPI sezione Derisking

Ogni funzione:
- accetta db_path, date_from, date_to, year come parametri opzionali
- restituisce un dizionario piatto con valori KPI
- non tocca la UI
- non scrive sul database
- gestisce eccezioni internamente, restituendo valori a zero in caso di errore
"""

import logging
import statistics
from datetime import date as _date
from typing import Optional

from database_manager import DatabaseManager
from services.app_paths import get_db_path
from utils.vsm_config import get_pagamenti_coefficient

logger = logging.getLogger('DataFlow.KPIEngine')

# ---------------------------------------------------------------------------
# Costanti DB
# ---------------------------------------------------------------------------

_TIPO_FULL_SUPPLY = 'Fornitura piena'
_TIPO_WORK_ORDER  = 'Conto lavoro'

# ---------------------------------------------------------------------------
# Helper interni
# ---------------------------------------------------------------------------


def _scalar(cursor, query: str, params: tuple = ()) -> int:
    """Esegue una query scalare e restituisce il primo valore, default 0."""
    cursor.execute(query, params)
    row = cursor.fetchone()
    return row[0] if row and row[0] is not None else 0


def _build_date_filter(
    date_col: str,
    date_from: Optional[str],
    date_to: Optional[str],
    year: Optional[int],
) -> tuple:
    """
    Costruisce clausole WHERE per filtri temporali.

    Il parametro ``year`` ha priorità su ``date_from`` / ``date_to``.
    Le date devono essere in formato ISO (YYYY-MM-DD o prefisso compatibile).

    Returns:
        (clauses: list[str], params: list)
    """
    clauses: list = []
    params: list = []

    if year is not None:
        clauses.append(f"strftime('%Y', {date_col}) = ?")
        params.append(str(year))
    else:
        if date_from:
            clauses.append(f"{date_col} >= ?")
            params.append(date_from)
        if date_to:
            clauses.append(f"{date_col} <= ?")
            params.append(date_to)

    return clauses, params


def _where(*clause_lists) -> str:
    """
    Unisce più liste di clausole con AND e restituisce la stringa WHERE completa.
    Restituisce '' se non ci sono clausole.

    Esempio:
        _where(['stato = ?'], ['data_emissione >= ?'])
        → "WHERE stato = ? AND data_emissione >= ?"
    """
    merged = [c for lst in clause_lists for c in lst]
    if not merged:
        return ""
    return "WHERE " + " AND ".join(merged)


def _safe_pct(numerator: float, denominator: float) -> float:
    """Calcola percentuale con protezione da divisione per zero."""
    if not denominator:
        return 0.0
    return (numerator / denominator) * 100.0


def _pct_stats(values: list) -> dict:
    """
    Calcola average / best / worst / median su una lista di valori numerici.
    Restituisce tutti zero se la lista è vuota.
    """
    if not values:
        return {"average": 0.0, "best": 0.0, "worst": 0.0, "median": 0.0}
    return {
        "average": round(statistics.mean(values), 4),
        "best":    round(max(values), 4),
        "worst":   round(min(values), 4),
        "median":  round(statistics.median(values), 4),
    }


def _sum_impacts(
    cursor,
    tipo_valore: str,
    opex_filter: Optional[int],
    event_date_clauses: list,
    event_date_params: list,
) -> tuple:
    """
    Somma valore_teorico e valore_effettivo da vsm_impacts,
    con JOIN su vsm_events per filtri su tipo evento e data.

    Args:
        tipo_valore:         'Saving' o 'Cost Avoidance'
        opex_filter:         1 → solo ricorrenti, 0 → solo non ricorrenti, None → tutti
        event_date_clauses:  clausole WHERE già qualificate con alias 've.'
        event_date_params:   parametri corrispondenti

    Returns:
        (sum_teorico: float, sum_effettivo: float)
    """
    base_clauses = ["vi.tipo_valore = ?"]
    base_params:  list = [tipo_valore]

    if opex_filter is not None:
        base_clauses.append("ve.opex_ripetitivo = ?")
        base_params.append(opex_filter)

    w = _where(base_clauses, event_date_clauses)
    query = f"""
        SELECT COALESCE(SUM(vi.valore_teorico),   0.0),
               COALESCE(SUM(vi.valore_effettivo), 0.0)
        FROM vsm_impacts vi
        JOIN vsm_events ve ON vi.event_id = ve.event_id
        {w}
    """
    cursor.execute(query, tuple(base_params + event_date_params))
    row = cursor.fetchone()
    if row:
        return float(row[0] or 0.0), float(row[1] or 0.0)
    return 0.0, 0.0


# ---------------------------------------------------------------------------
# API pubblica
# ---------------------------------------------------------------------------


def get_rfq_kpi(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> dict:
    """
    Calcola i KPI della sezione RFQ.

    Returns::

        {
            "rfq_active":       int,   # RFQ con stato 'attiva'
            "rfq_archived":     int,   # RFQ con stato 'archiviata'
            "rfq_total":        int,
            "rfq_not_expired":  int,   # RFQ con data_scadenza >= oggi
            "rfq_expired":      int,   # RFQ con data_scadenza < oggi
            "work_order":       int,   # RFQ di tipo 'Conto lavoro'
            "full_supply":      int,   # RFQ di tipo 'Fornitura piena'
        }

    I filtri ``date_from``, ``date_to``, ``year`` agiscono su ``data_emissione``.
    Non scaduto/scaduto è calcolato rispetto alla data odierna su ``data_scadenza``.
    """
    result = {
        "rfq_active":      0,
        "rfq_archived":    0,
        "rfq_total":       0,
        "rfq_not_expired": 0,
        "rfq_expired":     0,
        "work_order":      0,
        "full_supply":     0,
    }

    try:
        path = db_path or get_db_path()
        today = _date.today().isoformat()   # 'YYYY-MM-DD'

        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            # Clausole data per query su data_emissione
            d_clauses, d_params = _build_date_filter(
                "data_emissione", date_from, date_to, year
            )

            # --- RFQ per stato ---
            for key, stato in [("rfq_active", "attiva"), ("rfq_archived", "archiviata")]:
                result[key] = _scalar(
                    c,
                    f"SELECT COUNT(*) FROM richieste_offerta {_where(['stato = ?'], d_clauses)}",
                    tuple([stato] + d_params),
                )
            result["rfq_total"] = _scalar(
                c,
                f"SELECT COUNT(*) FROM richieste_offerta {_where(d_clauses)}",
                tuple(d_params),
            )

            # --- RFQ per scadenza (rispetto a oggi) ---
            result["rfq_not_expired"] = _scalar(
                c,
                f"SELECT COUNT(*) FROM richieste_offerta"
                f" {_where(['data_scadenza >= ?'], d_clauses)}",
                tuple([today] + d_params),
            )
            result["rfq_expired"] = _scalar(
                c,
                f"SELECT COUNT(*) FROM richieste_offerta"
                f" {_where(['data_scadenza < ?'], d_clauses)}",
                tuple([today] + d_params),
            )

            # --- RFQ per tipo ---
            for key, tipo in [("work_order", _TIPO_WORK_ORDER), ("full_supply", _TIPO_FULL_SUPPLY)]:
                result[key] = _scalar(
                    c,
                    f"SELECT COUNT(*) FROM richieste_offerta {_where(['tipo_rdo = ?'], d_clauses)}",
                    tuple([tipo] + d_params),
                )

    except Exception as e:
        logger.error("[KPIEngine] get_rfq_kpi: %s", e, exc_info=True)

    return result


def get_saving_kpi(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> dict:
    """
    Calcola i KPI della sezione Saving.

    I valori monetari (theoretical/actual/recurring/non-recurring) sono aggregati
    dagli impatti mensili in ``vsm_impacts``.

    Le percentuali (average/best/worst/median) sono calcolate per singolo evento
    come ``(importo_bdg - importo_negoziato) / importo_bdg × 100``,
    considerando solo gli eventi con driver='Prezzo' e importo_bdg > 0.

    Returns::

        {
            "theoretical_saving":   float,
            "actual_saving":        float,
            "average_saving_pct":   float,
            "best_saving_pct":      float,
            "worst_saving_pct":     float,
            "median_saving_pct":    float,
            "recurring_impact":     float,  # effettivo, eventi opex_ripetitivo=True
            "non_recurring_impact": float,  # effettivo, eventi opex_ripetitivo=False
        }
    """
    result = {
        "theoretical_saving":   0.0,
        "actual_saving":        0.0,
        "average_saving_pct":   0.0,
        "best_saving_pct":      0.0,
        "worst_saving_pct":     0.0,
        "median_saving_pct":    0.0,
        "recurring_impact":     0.0,
        "non_recurring_impact": 0.0,
    }

    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            d_clauses, d_params = _build_date_filter(
                "ve.event_date", date_from, date_to, year
            )

            # Totali teorico/effettivo
            teorico, effettivo = _sum_impacts(c, "Saving", None, d_clauses, d_params)
            result["theoretical_saving"] = round(teorico, 2)
            result["actual_saving"]      = round(effettivo, 2)

            # Ricorrente / Non ricorrente (opex_ripetitivo)
            _, rec     = _sum_impacts(c, "Saving", 1, d_clauses, d_params)
            _, non_rec = _sum_impacts(c, "Saving", 0, d_clauses, d_params)
            result["recurring_impact"]     = round(rec, 2)
            result["non_recurring_impact"] = round(non_rec, 2)

            # Percentuali di saving per evento.
            # best / worst / median: statistiche per-evento (invariate semanticamente).
            # average: media PESATA  sum(saving) / sum(base),
            #   coerente con il modello VSM già presente in models/vsm_event.py:
            #   - driver Prezzo:    base = importo_bdg * qty
            #                       saving = (importo_bdg - importo_negoziato) * qty
            #   - driver Pagamenti: base = spending_annuo
            #                       saving = spending * (delta_gg / 30) * coefficiente
            # Se l'evento non ha una base valida (bdg<=0 o spending<=0) viene ignorato
            # senza generare eccezioni né alterare le altre statistiche.
            d_ev_clauses, d_ev_params = _build_date_filter(
                "event_date", date_from, date_to, year
            )
            w_ev = _where(["event_type = ?"], d_ev_clauses)
            c.execute(
                f"""SELECT importo_bdg, importo_negoziato, quantita_annua, driver,
                           spending_annuo, giorni_pagamento_attuali,
                           giorni_pagamento_negoziati, payments_rate
                    FROM vsm_events {w_ev}""",
                tuple(["Saving"] + d_ev_params),
            )

            pcts: list = []
            weighted_num = 0.0   # somma saving di tutti gli eventi validi
            weighted_den = 0.0   # somma basi di tutti gli eventi validi

            for (bdg, neg, qty, drv,
                 spending, gg_att, gg_neg, p_rate) in c.fetchall():

                if drv == "Pagamenti":
                    # Base economica: spending annuo
                    # Saving: spending * (delta_giorni / 30) * coefficiente_opportunità
                    if spending and spending > 0 \
                            and gg_att is not None and gg_neg is not None:
                        coeff = (p_rate / 100.0) if p_rate is not None \
                            else get_pagamenti_coefficient()
                        delta_gg   = (gg_neg or 0) - (gg_att or 0)
                        saving_ev  = spending * (delta_gg / 30.0) * coeff
                        weighted_num += saving_ev
                        weighted_den += spending
                        pcts.append(_safe_pct(saving_ev, spending))
                else:
                    # Driver Prezzo (default): base = importo_bdg × qty
                    if bdg and bdg > 0:
                        q          = qty if qty and qty > 0 else 1.0
                        saving_ev  = ((bdg or 0) - (neg or 0)) * q
                        base_ev    = bdg * q
                        weighted_num += saving_ev
                        weighted_den += base_ev
                        pcts.append(_safe_pct((bdg or 0) - (neg or 0), bdg))

            stats = _pct_stats(pcts)
            result["average_saving_pct"] = round(_safe_pct(weighted_num, weighted_den), 4)
            result["best_saving_pct"]    = stats["best"]
            result["worst_saving_pct"]   = stats["worst"]
            result["median_saving_pct"]  = stats["median"]

    except Exception as e:
        logger.error("[KPIEngine] get_saving_kpi: %s", e, exc_info=True)

    return result


def get_cost_avoidance_kpi(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> dict:
    """
    Calcola i KPI della sezione Cost Avoidance.

    Struttura analoga a :func:`get_saving_kpi`.

    Le percentuali sono calcolate come
    ``(importo_richiesto_iniziale - importo_negoziato) / importo_richiesto_iniziale × 100``
    sugli eventi con importo_richiesto_iniziale > 0 e driver != 'Pagamenti'.

    Returns::

        {
            "theoretical_cost_avoidance": float,
            "actual_cost_avoidance":       float,
            "average_pct":                 float,
            "best_pct":                    float,
            "worst_pct":                   float,
            "median_pct":                  float,
            "recurring":                   float,
            "non_recurring":               float,
        }
    """
    result = {
        "theoretical_cost_avoidance": 0.0,
        "actual_cost_avoidance":       0.0,
        "average_pct":                 0.0,
        "best_pct":                    0.0,
        "worst_pct":                   0.0,
        "median_pct":                  0.0,
        "recurring":                   0.0,
        "non_recurring":               0.0,
    }

    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            d_clauses, d_params = _build_date_filter(
                "ve.event_date", date_from, date_to, year
            )

            teorico, effettivo = _sum_impacts(c, "Cost Avoidance", None, d_clauses, d_params)
            result["theoretical_cost_avoidance"] = round(teorico, 2)
            result["actual_cost_avoidance"]      = round(effettivo, 2)

            _, rec     = _sum_impacts(c, "Cost Avoidance", 1, d_clauses, d_params)
            _, non_rec = _sum_impacts(c, "Cost Avoidance", 0, d_clauses, d_params)
            result["recurring"]     = round(rec, 2)
            result["non_recurring"] = round(non_rec, 2)

            # Percentuale CA per evento.
            # Cost Avoidance usa sempre importo_richiesto_iniziale come base:
            # non esiste un driver 'Pagamenti' per questo tipo.
            d_ev_clauses, d_ev_params = _build_date_filter(
                "event_date", date_from, date_to, year
            )
            w = _where(
                ["event_type = ?", "importo_richiesto_iniziale > 0"],
                d_ev_clauses,
            )
            c.execute(
                f"SELECT importo_richiesto_iniziale, importo_negoziato FROM vsm_events {w}",
                tuple(["Cost Avoidance"] + d_ev_params),
            )
            pcts = [
                _safe_pct(row[0] - row[1], row[0])
                for row in c.fetchall()
                if row[0] and row[0] > 0
            ]
            stats = _pct_stats(pcts)
            result["average_pct"] = stats["average"]
            result["best_pct"]    = stats["best"]
            result["worst_pct"]   = stats["worst"]
            result["median_pct"]  = stats["median"]

    except Exception as e:
        logger.error("[KPIEngine] get_cost_avoidance_kpi: %s", e, exc_info=True)

    return result


def get_derisking_kpi(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> dict:
    """
    KPI Derisking: conta i fornitori potenziali creati nel periodo selezionato.

    Vengono inclusi solo i record con created_at valorizzato (non NULL) e
    dentro il range temporale. I record legacy con created_at=NULL sono sempre
    esclusi — comportamento voluto per questa versione.

    Se non è selezionato alcun filtro temporale (All), vengono inclusi tutti
    i record con created_at non NULL.

    Returns:
        {
            "total_suppliers":   int,
            "unique_categories": int,
            "status_counts":     dict[str, int],  # TRIM(stato) → count, desc
            "category_counts":   dict[str, int],  # TRIM(categoria) → count, desc
        }
    """
    result = {
        "total_suppliers":   0,
        "unique_categories": 0,
        "status_counts":     {},
        "category_counts":   {},
    }
    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            # Filtro temporale su created_at — esclude sempre i record legacy NULL.
            # Uso DATE(created_at) per normalizzare la data: i record storici possono
            # avere created_at in formato 'YYYY-MM-DDTHH:MM:SS' e la comparazione
            # stringa con date_to='YYYY-MM-DD' fallirebbe per lo stesso giorno.
            d_clauses, d_params = _build_date_filter(
                "DATE(created_at)", date_from, date_to, year
            )
            base = ["created_at IS NOT NULL"]

            # A) Totale fornitori nel range
            result["total_suppliers"] = _scalar(
                c,
                f"SELECT COUNT(*) FROM potential_suppliers {_where(base, d_clauses)}",
                tuple(d_params),
            )

            # B) Categorie uniche nel range — escludi NULL / vuoto / solo-spazi
            cat_extra = ["category IS NOT NULL", "TRIM(category) != ''"]
            result["unique_categories"] = _scalar(
                c,
                f"SELECT COUNT(DISTINCT TRIM(category))"
                f" FROM potential_suppliers"
                f" {_where(base + cat_extra, d_clauses)}",
                tuple(d_params),
            )

            # C) Count per stato nel range — TRIM + escludi NULL / vuoto / solo-spazi
            st_extra = ["supplier_status IS NOT NULL", "TRIM(supplier_status) != ''"]
            c.execute(
                f"SELECT TRIM(supplier_status), COUNT(*)"
                f" FROM potential_suppliers"
                f" {_where(base + st_extra, d_clauses)}"
                f" GROUP BY TRIM(supplier_status)"
                f" ORDER BY COUNT(*) DESC",
                tuple(d_params),
            )
            result["status_counts"] = {row[0]: row[1] for row in c.fetchall()}

            # D) Count per categoria nel range — TRIM + escludi NULL / vuoto / solo-spazi
            c.execute(
                f"SELECT TRIM(category), COUNT(*)"
                f" FROM potential_suppliers"
                f" {_where(base + cat_extra, d_clauses)}"
                f" GROUP BY TRIM(category)"
                f" ORDER BY COUNT(*) DESC",
                tuple(d_params),
            )
            result["category_counts"] = {row[0]: row[1] for row in c.fetchall()}

    except Exception as e:
        logger.error("[KPIEngine] get_derisking_kpi: %s", e, exc_info=True)
    return result


def get_available_years_derisking(db_path: Optional[str] = None) -> list:
    """
    Restituisce la lista ordinata degli anni con almeno un fornitore potenziale
    con created_at valorizzato.

    Usata dalla KpiWindow per popolare il filtro Year nel tab Derisking:
    mostra solo anni realmente presenti nell'anagrafica fornitori, senza
    inquinare con anni da vsm_events o richieste_offerta.

    Returns:
        list[int]: anni in ordine crescente; lista vuota in caso di errore.
    """
    years: set = set()
    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor
            c.execute(
                "SELECT DISTINCT strftime('%Y', created_at) "
                "FROM potential_suppliers WHERE created_at IS NOT NULL"
            )
            for (y,) in c.fetchall():
                if y:
                    years.add(int(y))
    except Exception as e:
        logger.error("[KPIEngine] get_available_years_derisking: %s", e, exc_info=True)
    return sorted(years)


def get_available_years(db_path: Optional[str] = None) -> list:
    """
    Restituisce la lista ordinata degli anni presenti nel DB nelle sorgenti
    utilizzate dai KPI (vsm_events.event_date e richieste_offerta.data_emissione).

    Usato dalla KpiWindow per popolare dinamicamente il filtro Year.

    Returns:
        list[int]: anni in ordine crescente; lista vuota in caso di errore.
    """
    years: set = set()
    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor
            c.execute(
                "SELECT DISTINCT strftime('%Y', event_date) "
                "FROM vsm_events WHERE event_date IS NOT NULL"
            )
            for (y,) in c.fetchall():
                if y:
                    years.add(int(y))
            c.execute(
                "SELECT DISTINCT strftime('%Y', data_emissione) "
                "FROM richieste_offerta WHERE data_emissione IS NOT NULL"
            )
            for (y,) in c.fetchall():
                if y:
                    years.add(int(y))
            # Fornitori potenziali (Derisking) — anni da created_at
            c.execute(
                "SELECT DISTINCT strftime('%Y', created_at) "
                "FROM potential_suppliers WHERE created_at IS NOT NULL"
            )
            for (y,) in c.fetchall():
                if y:
                    years.add(int(y))
    except Exception as e:
        logger.error("[KPIEngine] get_available_years: %s", e, exc_info=True)
    return sorted(years)
