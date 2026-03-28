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
from typing import Optional

from database_manager import DatabaseManager
from services.app_paths import get_db_path

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
            "rfq_active":        int,   # RFQ con stato 'attiva'
            "rfq_archived":      int,   # RFQ con stato 'archiviata'
            "rfq_total":         int,
            "offers_active":     int,   # Offerte ricevute per RFQ attive
            "offers_archived":   int,   # Offerte ricevute per RFQ archiviate
            "offers_total":      int,
            "work_order":        int,   # RFQ di tipo 'Conto lavoro'
            "full_supply":       int,   # RFQ di tipo 'Fornitura piena'
            "offers_per_rfq_avg": float,
        }

    I filtri ``date_from``, ``date_to``, ``year`` agiscono su ``data_emissione``.
    """
    result = {
        "rfq_active":         0,
        "rfq_archived":       0,
        "rfq_total":          0,
        "offers_active":      0,
        "offers_archived":    0,
        "offers_total":       0,
        "work_order":         0,
        "full_supply":        0,
        "offers_per_rfq_avg": 0.0,
    }

    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            # Clausole data per query dirette (singola tabella, no alias)
            d_clauses, d_params = _build_date_filter(
                "data_emissione", date_from, date_to, year
            )
            # Clausole data per query con JOIN (richieste_offerta aliasata 'ro')
            d_clauses_ro, _ = _build_date_filter(
                "ro.data_emissione", date_from, date_to, year
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

            # --- RFQ per tipo ---
            for key, tipo in [("work_order", _TIPO_WORK_ORDER), ("full_supply", _TIPO_FULL_SUPPLY)]:
                result[key] = _scalar(
                    c,
                    f"SELECT COUNT(*) FROM richieste_offerta {_where(['tipo_rdo = ?'], d_clauses)}",
                    tuple([tipo] + d_params),
                )

            # --- Offerte ricevute: distinct (id_richiesta, nome_fornitore) ---
            # Un'offerta = una risposta univoca di un fornitore per una RFQ
            _offers_sql = """
                SELECT COUNT(*) FROM (
                    SELECT DISTINCT dr.id_richiesta, orr.nome_fornitore
                    FROM offerte_ricevute orr
                    JOIN dettagli_richiesta dr ON orr.id_dettaglio = dr.id_dettaglio
                    JOIN richieste_offerta ro ON dr.id_richiesta = ro.id_richiesta
                    {where}
                )
            """
            for key, stato in [("offers_active", "attiva"), ("offers_archived", "archiviata")]:
                result[key] = _scalar(
                    c,
                    _offers_sql.format(where=_where(["ro.stato = ?"], d_clauses_ro)),
                    tuple([stato] + d_params),
                )
            result["offers_total"] = _scalar(
                c,
                _offers_sql.format(where=_where(d_clauses_ro)),
                tuple(d_params),
            )

            # --- Offerte medie per RFQ ---
            if result["rfq_total"] > 0:
                result["offers_per_rfq_avg"] = round(
                    result["offers_total"] / result["rfq_total"], 2
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

            # Percentuale di saving per evento.
            # Include tutti i driver; gli eventi senza importo_bdg valido
            # restituiscono 0 da _safe_pct e non alterano la statistica.
            d_ev_clauses, d_ev_params = _build_date_filter(
                "event_date", date_from, date_to, year
            )
            w = _where(
                ["event_type = ?", "importo_bdg > 0"],
                d_ev_clauses,
            )
            c.execute(
                f"SELECT importo_bdg, importo_negoziato FROM vsm_events {w}",
                tuple(["Saving"] + d_ev_params),
            )
            pcts = [
                _safe_pct(row[0] - row[1], row[0])
                for row in c.fetchall()
                if row[0] and row[0] > 0
            ]
            stats = _pct_stats(pcts)
            result["average_saving_pct"] = stats["average"]
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
    Calcola i KPI della sezione Derisking.

    Conta i fornitori unici introdotti come nuovi, ignorando valori vuoti/null.

    Returns::

        {
            "unique_new_suppliers_introduced": int,
        }
    """
    result = {
        "unique_new_suppliers_introduced": 0,
    }

    try:
        path = db_path or get_db_path()
        with DatabaseManager(path, read_only=True) as db:
            c = db.cursor

            d_clauses, d_params = _build_date_filter(
                "event_date", date_from, date_to, year
            )

            w = _where(
                [
                    "event_type = ?",
                    "new_supplier IS NOT NULL",
                    "TRIM(new_supplier) != ''",
                ],
                d_clauses,
            )
            result["unique_new_suppliers_introduced"] = _scalar(
                c,
                f"SELECT COUNT(DISTINCT TRIM(new_supplier)) FROM vsm_events {w}",
                tuple(["Derisking"] + d_params),
            )

    except Exception as e:
        logger.error("[KPIEngine] get_derisking_kpi: %s", e, exc_info=True)

    return result
