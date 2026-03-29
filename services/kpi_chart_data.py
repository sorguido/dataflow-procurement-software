# -*- coding: utf-8 -*-
"""
KPI Chart Data — serie temporali per i grafici della finestra KPI.

Nessuna UI, nessuna logica KPI aggregata, nessun accesso diretto alla UI.
Riceve gli stessi parametri filtro dell'engine e restituisce serie temporali
già bucketizzate, pronte per il rendering su Canvas.

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
# Helpers privati
# ---------------------------------------------------------------------------

def _bucket_pattern(
    date_from: Optional[str],
    date_to:   Optional[str],
    year:      Optional[int],
) -> str:
    """
    Sceglie il pattern strftime per il bucketing in base all'ampiezza del filtro.

    Logica:
      - year     → mensile  ('%Y-%m', max 12 bucket)
      - rolling ≤ 35 gg    → giornaliero  ('%Y-%m-%d')
      - rolling > 35 gg    → mensile      ('%Y-%m')
      - nessun filtro      → mensile      ('%Y-%m', tutti i mesi nel DB)
    """
    if year is not None:
        return '%Y-%m'
    if date_from and date_to:
        try:
            span = (_date.fromisoformat(date_to) - _date.fromisoformat(date_from)).days
            return '%Y-%m-%d' if span <= 35 else '%Y-%m'
        except (ValueError, TypeError):
            pass
    return '%Y-%m'


def _make_label(bucket: str, pattern: str) -> str:
    """
    Etichetta compatta leggibile dal bucket SQLite.

    '%Y-%m'    '2025-10' → '25-10'
    '%Y-%m-%d' '2025-10-15' → '10-15'
    """
    if not bucket:
        return ''
    if pattern == '%Y':
        return bucket          # '2025'
    return bucket[-5:]         # '25-10'  or '10-15'


def _dclauses(col: str, date_from, date_to, year) -> tuple:
    """Costruisce clausole e parametri per il filtro temporale."""
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
    Andamento RFQ nel tempo.

    Returns:
        list of {'label': str, 'count': int}
    """
    try:
        pat  = _bucket_pattern(date_from, date_to, year)
        path = db_path or get_db_path()
        clauses, params = _dclauses('data_emissione', date_from, date_to, year)

        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                f"SELECT strftime('{pat}', data_emissione), COUNT(*)"
                f" FROM richieste_offerta"
                f" {_where(clauses)}"
                f" GROUP BY 1 ORDER BY 1",
                params,
            )
            return [
                {'label': _make_label(b, pat), 'count': int(c or 0)}
                for b, c in db.cursor.fetchall()
                if b
            ]
    except Exception as exc:
        logger.error('[KpiChartData] get_rfq_chart_data: %s', exc)
        return []


def get_saving_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Andamento Theoretical vs Actual Saving nel tempo.

    Bucketing per (anno, mese) da vsm_impacts; filtro data su vsm_events.event_date.

    Returns:
        list of {'label': str, 'theoretical': float, 'actual': float}
    """
    try:
        path = db_path or get_db_path()
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
                    GROUP BY vi.anno, vi.mese
                    ORDER BY vi.anno, vi.mese""",
                ['Saving'] + params,
            )
            return [
                {
                    'label':       _make_label(b, '%Y-%m'),
                    'theoretical': float(t or 0),
                    'actual':      float(a or 0),
                }
                for b, t, a in db.cursor.fetchall()
                if b
            ]
    except Exception as exc:
        logger.error('[KpiChartData] get_saving_chart_data: %s', exc)
        return []


def get_cost_avoidance_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Andamento Theoretical vs Actual Cost Avoidance nel tempo.

    Struttura identica a get_saving_chart_data.

    Returns:
        list of {'label': str, 'theoretical': float, 'actual': float}
    """
    try:
        path = db_path or get_db_path()
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
                    GROUP BY vi.anno, vi.mese
                    ORDER BY vi.anno, vi.mese""",
                ['Cost Avoidance'] + params,
            )
            return [
                {
                    'label':       _make_label(b, '%Y-%m'),
                    'theoretical': float(t or 0),
                    'actual':      float(a or 0),
                }
                for b, t, a in db.cursor.fetchall()
                if b
            ]
    except Exception as exc:
        logger.error('[KpiChartData] get_cost_avoidance_chart_data: %s', exc)
        return []


def get_derisking_chart_data(
    db_path:   Optional[str] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    year:      Optional[int] = None,
) -> list:
    """
    Nuovi fornitori introdotti nel tempo.

    Returns:
        list of {'label': str, 'count': int}
    """
    try:
        pat  = _bucket_pattern(date_from, date_to, year)
        path = db_path or get_db_path()
        clauses, params = _dclauses('event_date', date_from, date_to, year)

        w = _where(
            ["event_type = ?",
             "new_supplier IS NOT NULL",
             "TRIM(new_supplier) != ''"],
            clauses,
        )

        with DatabaseManager(path, read_only=True) as db:
            db.cursor.execute(
                f"SELECT strftime('{pat}', event_date),"
                f"       COUNT(DISTINCT TRIM(new_supplier))"
                f" FROM vsm_events"
                f" {w}"
                f" GROUP BY 1 ORDER BY 1",
                ['Derisking'] + params,
            )
            return [
                {'label': _make_label(b, pat), 'count': int(c or 0)}
                for b, c in db.cursor.fetchall()
                if b
            ]
    except Exception as exc:
        logger.error('[KpiChartData] get_derisking_chart_data: %s', exc)
        return []
