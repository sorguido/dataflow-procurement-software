# -*- coding: utf-8 -*-
"""
KPI Chart — rendering Canvas puro (tkinter, zero dipendenze esterne).

Funzioni pubbliche:
    draw_bar_chart(canvas, data, y_fmt='int')
        data: list of {'label': str, 'value': float|int}

    draw_dual_bar_chart(canvas, data, label1='', label2='')
        data: list of {'label': str, 'theoretical': float, 'actual': float}

Entrambe gestiscono dataset vuoti mostrando "No data available".
Entrambe si adattano alle dimensioni correnti del Canvas.
"""

import tkinter as tk

# ---------------------------------------------------------------------------
# Palette — corporate, muted, coerente con DataFlow
# ---------------------------------------------------------------------------
_C_BAR1  = '#4472C4'   # blu — serie Theoretical / serie singola
_C_BAR2  = '#ED7D31'   # arancio — serie Actual
_C_AXIS  = '#B0B0B0'   # grigio assi
_C_GRID  = '#E8E8E8'   # grigio griglia
_C_TEXT  = '#555555'   # testo etichette
_C_BG    = '#F8F8F8'   # sfondo canvas

# Layout (pixel)
_ML = 58   # margine sinistro  — spazio etichette Y
_MR = 10   # margine destro
_MT = 10   # margine superiore (barra singola) / spazio legenda (dual)
_MB = 44   # margine inferiore — spazio etichette X
_LEGEND_H = 18   # altezza riservata alla legenda nei dual chart


# ---------------------------------------------------------------------------
# API pubblica
# ---------------------------------------------------------------------------

def draw_bar_chart(canvas: tk.Canvas, data: list, y_fmt: str = 'int') -> None:
    """
    Disegna un bar chart a serie singola.

    Args:
        canvas: tk.Canvas sul quale disegnare
        data:   list of {'label': str, 'value': float|int}
        y_fmt:  'int' → etichette Y come interi
                'money' → etichette Y K/M
    """
    _clear(canvas)
    W, H = _canvas_size(canvas)
    if W < 40 or H < 40:
        return

    if not data:
        _draw_no_data(canvas, W, H)
        return

    plot_w = W - _ML - _MR
    plot_h = H - _MT - _MB

    vals  = [float(d.get('value', 0)) for d in data]
    max_v = max(vals) or 1.0

    _draw_axes_and_grid(canvas, _ML, _MT, plot_w, plot_h, max_v, y_fmt)

    n    = len(data)
    slot = plot_w / n
    bw   = max(4, int(slot * 0.65))

    for i, (d, v) in enumerate(zip(data, vals)):
        x_c   = int(_ML + slot * (i + 0.5))
        bar_h = int((v / max_v) * plot_h)
        x0    = x_c - bw // 2
        x1    = x0 + bw
        y1    = _MT + plot_h
        y0    = y1 - bar_h
        if bar_h > 0:
            canvas.create_rectangle(x0, y0, x1, y1,
                                    fill=_C_BAR1, outline='', width=0)
        _draw_x_label(canvas, x_c, y1, d.get('label', ''), n)


def draw_dual_bar_chart(
    canvas: tk.Canvas,
    data:   list,
    label1: str = '',
    label2: str = '',
) -> None:
    """
    Disegna un grouped bar chart a due serie (Theoretical / Actual).

    Args:
        canvas: tk.Canvas
        data:   list of {'label': str, 'theoretical': float, 'actual': float}
        label1: testo legenda serie 1 (blu)
        label2: testo legenda serie 2 (arancio)
    """
    _clear(canvas)
    W, H = _canvas_size(canvas)
    if W < 40 or H < 40:
        return

    if not data:
        _draw_no_data(canvas, W, H)
        return

    mt     = _MT + _LEGEND_H     # top rialzato per la legenda
    plot_w = W - _ML - _MR
    plot_h = H - mt - _MB

    t_vals = [float(d.get('theoretical', 0)) for d in data]
    a_vals = [float(d.get('actual',      0)) for d in data]
    max_v  = max(max(t_vals, default=0), max(a_vals, default=0)) or 1.0

    _draw_axes_and_grid(canvas, _ML, mt, plot_w, plot_h, max_v, 'money')
    _draw_legend(canvas, W, _MT, label1, label2)

    n      = len(data)
    slot   = plot_w / n
    pair_w = max(8, int(slot * 0.72))
    bw     = max(3, pair_w // 2 - 1)

    for i, (d, tv, av) in enumerate(zip(data, t_vals, a_vals)):
        x_left = int(_ML + slot * i + (slot - pair_w) / 2)
        base_y = mt + plot_h

        # Serie 1 — Theoretical (blu)
        h1 = int((tv / max_v) * plot_h)
        if h1 > 0:
            canvas.create_rectangle(
                x_left, base_y - h1, x_left + bw, base_y,
                fill=_C_BAR1, outline='', width=0,
            )
        # Serie 2 — Actual (arancio)
        h2 = int((av / max_v) * plot_h)
        if h2 > 0:
            canvas.create_rectangle(
                x_left + bw + 1, base_y - h2,
                x_left + bw * 2 + 1, base_y,
                fill=_C_BAR2, outline='', width=0,
            )

        x_c = x_left + bw
        _draw_x_label(canvas, x_c, base_y, d.get('label', ''), n)


# ---------------------------------------------------------------------------
# Helper interni
# ---------------------------------------------------------------------------

def _clear(canvas: tk.Canvas) -> None:
    canvas.delete('all')
    canvas.configure(bg=_C_BG)


def _canvas_size(canvas: tk.Canvas) -> tuple:
    canvas.update_idletasks()
    return canvas.winfo_width(), canvas.winfo_height()


def _draw_no_data(canvas: tk.Canvas, W: int, H: int) -> None:
    canvas.create_text(
        W // 2, H // 2,
        text='No data available',
        fill='#AAAAAA',
        font=(None, 9, 'italic'),
    )


def _draw_axes_and_grid(
    canvas:  tk.Canvas,
    ml:      int,
    mt:      int,
    plot_w:  int,
    plot_h:  int,
    max_v:   float,
    y_fmt:   str,
) -> None:
    """Disegna assi, griglie orizzontali e tick/etichette Y."""
    # Asse Y
    canvas.create_line(ml, mt, ml, mt + plot_h, fill=_C_AXIS, width=1)
    # Asse X
    canvas.create_line(ml, mt + plot_h, ml + plot_w, mt + plot_h,
                       fill=_C_AXIS, width=1)
    # Griglie orizzontali e tick Y  (3 livelli: 33%, 66%, 100%)
    for frac in (1/3, 2/3, 1.0):
        y   = mt + plot_h - int(frac * plot_h)
        val = max_v * frac
        # Linea griglia
        canvas.create_line(ml, y, ml + plot_w, y,
                           fill=_C_GRID, width=1)
        # Tick asse Y
        canvas.create_line(ml - 4, y, ml, y, fill=_C_AXIS, width=1)
        # Etichetta Y
        canvas.create_text(
            ml - 6, y,
            text=_fmt_y(val, y_fmt),
            fill=_C_TEXT,
            font=(None, 7),
            anchor='e',
        )


def _draw_legend(
    canvas: tk.Canvas,
    W:      int,
    top_y:  int,
    label1: str,
    label2: str,
) -> None:
    """Legenda compatta in alto a destra del canvas."""
    if not label1 and not label2:
        return

    y_c = top_y + _LEGEND_H // 2  # centro verticale area legenda
    # Calcola posizione: parte da destra e va a sinistra
    rx = W - _MR - 2

    # Serie 2 — Actual (arancio, rightmost)
    if label2:
        lbl = label2[:12]
        canvas.create_rectangle(rx - 8, y_c - 4, rx, y_c + 4,
                                 fill=_C_BAR2, outline='')
        canvas.create_text(rx - 11, y_c, text=lbl,
                           fill=_C_TEXT, font=(None, 7), anchor='e')
        rx = rx - 11 - len(lbl) * 5 - 10

    # Serie 1 — Theoretical (blu)
    if label1:
        lbl = label1[:12]
        canvas.create_rectangle(rx - 8, y_c - 4, rx, y_c + 4,
                                 fill=_C_BAR1, outline='')
        canvas.create_text(rx - 11, y_c, text=lbl,
                           fill=_C_TEXT, font=(None, 7), anchor='e')


def _draw_x_label(
    canvas:  tk.Canvas,
    x_c:     int,
    base_y:  int,
    text:    str,
    n_total: int,
) -> None:
    """Scrive l'etichetta asse X sotto la barra."""
    trunc_to = 6 if n_total > 9 else 8
    font_sz  = 6 if n_total > 9 else 7
    label    = _trunc(text, trunc_to)
    canvas.create_text(
        x_c, base_y + 5,
        text=label,
        fill=_C_TEXT,
        font=(None, font_sz),
        anchor='n',
    )


def _trunc(s: str, n: int) -> str:
    if not s:
        return ''
    return (s[:n] + '\u2026') if len(s) > n else s   # '\u2026' = …


def _fmt_y(v: float, fmt: str) -> str:
    """Formato etichette asse Y."""
    if fmt == 'money':
        av = abs(v)
        if av >= 1_000_000:
            return f'{v / 1_000_000:.1f}M'
        if av >= 1_000:
            return f'{v / 1_000:.0f}K'
        return f'{v:.0f}'
    # int / default
    if v >= 10_000:
        return f'{int(v / 1000)}K'
    return str(int(round(v)))
