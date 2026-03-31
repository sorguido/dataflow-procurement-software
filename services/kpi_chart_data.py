# -*- coding: utf-8 -*-
"""
KPI Chart Data — serie temporali deterministiche per i grafici della finestra KPI.

Il dominio temporale è SEMPRE costruito dai filtri UI, non dai dati.
Ogni bucket è garantito nel range; i bucket senza dati hanno valore 0.

Funzioni pubbliche:
    get_rfq_chart_data(...)          → list[{'label', 'count'}]
    get_saving_chart_data(...)       → list[{'label', 'theoretical', 'actual'}]
    get_cost_avoidance_chart_data(…) → list[{'label', 'theoretical', 'actual'}]
    get_derisking_chart_data(...)    → list[{'label', 'count'}]
"""

import logging
from datetime import date as _date
from typing import Optional

from database_manager import DatabaseManager
from services.app_paths import get_db_path

logger = logging.getLogger('DataFlow.KpiChartData')


# ---------------------------------------------------------------------------
# Bucket generator — cuore della logica deterministica
# ---------------------------------------------------------------------------

def _build_month_buckets(
    date_from: Optional[str],
    date_to:   Optional[str],
    year:      Optional[int],
    db_path:   Optional[str] = None,
) -> list:
    """
    Costruisce la lista ORDINATA di bucket mensili (formato 'YYYY-MM')
    basandosi ESCLUSIVAMENTE sui parametri filtro.

    Regole:
      YEAR selezionato   → 12 mesi fissi  YYYY-01 … YYYY-12
                           capati al mese corrente (no mesi futuri).
      Preset (rolling)   → mesi da (oggi - span) fino al mese corrente.
                           Usa date_from / date_to già calcolati dalla UI.
      Nessun filtro (All)→ mesi dal più vecchio record nel DB fino al mese
                           corrente; fallback lista vuota su errore.

    Returns:
        list[str]: es. ['2025-10', '2025-11', '2025-12', '2026-01']
    """
    today_ym = _date.today().strftime('%Y-%m')   # tetto: nessun mese futuro

    # ------- CASE 1: anno fisso → 12 bucket -------
    if year is not None:
        buckets = [f'{year:04d}-{m:02d}' for m in range(1, 13)]
        return [b for b in buckets if b <= today_ym]

    # ------- CASE 2: rolling preset (date_from + date_to fornite dalla UI) -------
    if date_from and date_to:
        try:
            df = _date.fromisoformat(date_from)
            dt = _date.fromisoformat(date_to)
        except ValueError:
            df = dt = _date.today()

        # Normalizza al primo giorno del mese per entrambi gli estremi
        first_ym = f'{df.year:04d}-{df.month:02d}'
        last_ym  = min(f'{dt.year:04d}-{dt.month:02d}', today_ym)

        return _month_range(first_ym, last_ym)

    # ------- CASE 3: nessun filtro (All) → dal minimo del DB a oggi -------
    try:
        path = db_path or get_db_path()
        first_ym = _db_min_month(path)
    except Exception:
        first_ym = None

    if not first_ym:
        return []
    return _month_range(first_ym, today_ym)


def _month_range(first_ym: str, last_ym: str) -> list:
    """Genera lista inclusiva di 'YYYY-MM' da first_ym a last_ym."""
    if first_ym > last_ym:
        return []
    result = []
    y, m = int(first_ym[:4]), int(first_ym[5:7])
    ey, em = int(last_ym[:4]), int(last_ym[5:7])
    while (y, m) <= (ey, em):
        result.append(f'{y:04d}-{m:02d}')
        m += 1
        if m > 12:
            m = 1
            y += 1
    return result


def _db_min_month(path: str) -> Optional[str]:
    """
    Restituisce il bucket 'YYYY-MM' più antico tra richieste_offerta e vsm_events.
    Ritorna None se nessun record esiste o sulla tabella mancante.
    """
    candidates = []
    try:
        with DatabaseManager(path, read_only=True) as db:
            # richieste_offerta
            try:
                db.cursor.execute(
                    "SELECT MIN(strftime('%Y-%m', data_emissione))"
                    " FROM richieste_offerta WHERE data_emissione IS NOT NULL"
                )
                row = db.cursor.fetchone()
                if row and row[0]:
                    candidates.append(row[0])
            except Exception:
                pass
            # vsm_events
            try:
                db.cursor.execute(
                    "SELECT MIN(strftime('%Y-%m', event_date))"
                    " FROM vsm_events WHERE event_date IS NOT NULL"
                )
                row = db.cursor.fetchone()
                if row and row[0]:
                    candidates.append(row[0])
            except Exception:
                pass
    except Exception:
        pass
    return min(candidates) if candidates else None


def _label(ym: str) -> str:
    """Etichetta compatta: 'YYYY-MM' → 'MM/YY' (es. '25-10' → '10/25')."""
    if not ym or len(ym) < 7:
        return ym
    return f'{ym[5:7]}/{ym[2:4]}'   # MM/YY


def _dclauses(col: str, date_from, date_to, year) -> tuple:
    clauses: list = []
    params:  list = []
    if year is not None:
        clauses.append(f"strftime('%Y', {col}) = ?")
        params.append(str(year))
    else:
        if date_from:
            clauses.append(f"{col} >= ?")
            params.append(date_from)
        if date_to:
            clauses.append(f"{col} <= ?")
            params.append(date_to)
    return clauses, params


def _where(*clause_lists) -> str:
    merged = [c for lst in clause_lists for c in lst]
    return ("WHERE " + " AND ".join(merged)) if merged else ""


# ---------------------------------------------------------------------------
# API pubblica
# ---------------------------------------------------------------------------

def get_rfq_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Andamento RFQ emesse nel tempo, per bucket mensile determinato dal filtro.

    Returns:
        list of {'label': str, 'count': int}
        — tutti i bucket del range, 0 per i mesi senza dati.
    """
    path    = db_path or get_db_path()
    buckets = _build_month_buckets(date_from, date_to, year, path)
    if not buckets:
        return []

    try:
        clauses, params = _dclauses('data_emissione', date_from, date_to, year)
        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                f"SELECT strftime('%Y-%m', data_emissione), COUNT(*)"
                f" FROM richieste_offerta"
                f" {_where(clauses)}"
                f" GROUP BY 1",
                params,
            )
            lookup = {b: int(c or 0) for b, c in db.cursor.fetchall()}
    except Exception as exc:
        logger.error('[KpiChartData] get_rfq_chart_data: %s', exc)
        lookup = {}

    return [{'label': _label(b), 'count': lookup.get(b, 0)} for b in buckets]


def get_saving_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Andamento Theoretical vs Actual Saving per bucket mensile determinato dal filtro.

    Returns:
        list of {'label': str, 'theoretical': float, 'actual': float}
    """
    path    = db_path or get_db_path()
    buckets = _build_month_buckets(date_from, date_to, year, path)
    if not buckets:
        return []

    try:
        clauses, params = _dclauses('ve.event_date', date_from, date_to, year)
        w = _where(["vi.tipo_valore = ?"], clauses)
        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                f"""SELECT printf('%04d-%02d', vi.anno, vi.mese),
                           SUM(vi.valore_teorico),
                           SUM(vi.valore_effettivo)
                    FROM vsm_impacts vi
                    JOIN vsm_events ve ON vi.event_id = ve.event_id
                    {w}
                    GROUP BY vi.anno, vi.mese""",
                ['Saving'] + params,
            )
            lookup = {
                b: (float(t or 0), float(a or 0))
                for b, t, a in db.cursor.fetchall()
            }
    except Exception as exc:
        logger.error('[KpiChartData] get_saving_chart_data: %s', exc)
        lookup = {}

    return [
        {
            'label':       _label(b),
            'theoretical': lookup.get(b, (0.0, 0.0))[0],
            'actual':      lookup.get(b, (0.0, 0.0))[1],
        }
        for b in buckets
    ]


def get_cost_avoidance_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Andamento Theoretical vs Actual Cost Avoidance per bucket mensile.

    Returns:
        list of {'label': str, 'theoretical': float, 'actual': float}
    """
    path    = db_path or get_db_path()
    buckets = _build_month_buckets(date_from, date_to, year, path)
    if not buckets:
        return []

    try:
        clauses, params = _dclauses('ve.event_date', date_from, date_to, year)
        w = _where(["vi.tipo_valore = ?"], clauses)
        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                f"""SELECT printf('%04d-%02d', vi.anno, vi.mese),
                           SUM(vi.valore_teorico),
                           SUM(vi.valore_effettivo)
                    FROM vsm_impacts vi
                    JOIN vsm_events ve ON vi.event_id = ve.event_id
                    {w}
                    GROUP BY vi.anno, vi.mese""",
                ['Cost Avoidance'] + params,
            )
            lookup = {
                b: (float(t or 0), float(a or 0))
                for b, t, a in db.cursor.fetchall()
            }
    except Exception as exc:
        logger.error('[KpiChartData] get_cost_avoidance_chart_data: %s', exc)
        lookup = {}

    return [
        {
            'label':       _label(b),
            'theoretical': lookup.get(b, (0.0, 0.0))[0],
            'actual':      lookup.get(b, (0.0, 0.0))[1],
        }
        for b in buckets
    ]


def get_derisking_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Fornitori per categoria — per il bar chart del tab KPI Derisking.

    date_from / date_to / year: accettati per compatibilità firma, ignorati.

    Returns:
        list of {'label': str, 'count': int}  — desc per count
    """
    path = db_path or get_db_path()
    try:
        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                "SELECT TRIM(category), COUNT(*)"
                " FROM potential_suppliers"
                " WHERE category IS NOT NULL AND TRIM(category) != ''"
                " GROUP BY TRIM(category)"
                " ORDER BY COUNT(*) DESC"
            )
            return [{'label': row[0], 'count': int(row[1] or 0)}
                    for row in db.cursor.fetchall()]
    except Exception as exc:
        logger.error('[KpiChartData] get_derisking_chart_data: %s', exc)
        return []
