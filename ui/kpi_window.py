# -*- coding: utf-8 -*-
"""
KpiWindow - Finestra KPI Analysis per DataFlow Procurement Software.

Fase 1: Shell UI senza logica dati.
- Header con titolo, filtri periodo, selezione anno, Export Excel (placeholder)
- Navigazione a tab: RFQ / Saving / Cost Avoidance / Derisking
- Ogni sezione: area KPI cards + area chart (placeholder) + area tabella (placeholder)
"""

import tkinter as tk
from tkinter import ttk
import builtins

from utils.window_utils import center_window
from utils.resource_utils import set_window_icon

# Compatibilità: _() è installata in builtins da init_i18n().
# Se non disponibile (es. test unitari), usa dummy.
if not hasattr(builtins, '_'):
    builtins._ = lambda x: x


class KpiWindow(tk.Toplevel):
    """Finestra KPI Analysis - Fase 1: struttura UI."""

    _PERIOD_OPTIONS = ["1M", "3M", "12M", "3Y", "5Y", "10Y", _("All")]
    _YEAR_RANGE = list(range(2020, 2031))

    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("KPI Analysis"))
        self.resizable(True, True)

        # Variabili filtro (placeholder — nessuna logica dati)
        self._selected_period = tk.StringVar(value="12M")
        self._selected_year = tk.StringVar(value="")

        self._build_header()
        self._build_navigation()

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        self.minsize(800, 580)

        # Apre massimizzata — stesso pattern della MainWindow (Linux/Windows)
        try:
            self.attributes("-zoomed", True)   # Linux / X11
        except Exception:
            try:
                self.state("zoomed")            # Windows
            except Exception:
                center_window(self)             # Fallback generico

        self.deiconify()

    # ------------------------------------------------------------------
    # HEADER
    # ------------------------------------------------------------------

    def _build_header(self):
        """Costruisce la barra superiore: titolo, filtri, export."""
        header = ttk.Frame(self, padding=(12, 8, 12, 8))
        header.pack(side="top", fill="x")

        # Titolo
        ttk.Label(
            header,
            text=_("KPI Analysis"),
            font=(None, 14, "bold"),
        ).pack(side="left", padx=(0, 20))

        # Separatore verticale visivo
        ttk.Separator(header, orient="vertical").pack(side="left", fill="y", padx=(0, 12))

        # Label "Period:"
        ttk.Label(header, text=_("Period:")).pack(side="left", padx=(0, 4))

        # Pulsanti periodo (radio-style)
        period_frame = ttk.Frame(header)
        period_frame.pack(side="left", padx=(0, 12))
        for option in self._PERIOD_OPTIONS:
            ttk.Radiobutton(
                period_frame,
                text=option,
                variable=self._selected_period,
                value=option,
            ).pack(side="left", padx=2)

        # Separatore verticale
        ttk.Separator(header, orient="vertical").pack(side="left", fill="y", padx=(0, 12))

        # Label "Year:"
        ttk.Label(header, text=_("Year:")).pack(side="left", padx=(0, 4))

        # Combobox anno
        year_values = [""] + [str(y) for y in self._YEAR_RANGE]
        ttk.Combobox(
            header,
            textvariable=self._selected_year,
            values=year_values,
            width=6,
            state="readonly",
        ).pack(side="left", padx=(0, 20))

        # Export Excel (placeholder — destra)
        ttk.Button(
            header,
            text=_("📥 Export Excel"),
            command=self._on_export_excel,
        ).pack(side="right", padx=(8, 0))

        # Separatore orizzontale sotto l'header
        ttk.Separator(self, orient="horizontal").pack(side="top", fill="x")

    # ------------------------------------------------------------------
    # NAVIGAZIONE (Notebook)
    # ------------------------------------------------------------------

    def _build_navigation(self):
        """Costruisce il Notebook con le 4 sezioni."""
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(side="top", fill="both", expand=True, padx=12, pady=(8, 12))

        tab_rfq = ttk.Frame(self._notebook)
        tab_saving = ttk.Frame(self._notebook)
        tab_ca = ttk.Frame(self._notebook)
        tab_derisking = ttk.Frame(self._notebook)

        self._notebook.add(tab_rfq, text=_("  RFQ  "))
        self._notebook.add(tab_saving, text=_("  Saving  "))
        self._notebook.add(tab_ca, text=_("  Cost Avoidance  "))
        self._notebook.add(tab_derisking, text=_("  Derisking  "))

        self._build_tab_rfq(tab_rfq)
        self._build_tab_saving(tab_saving)
        self._build_tab_cost_avoidance(tab_ca)
        self._build_tab_derisking(tab_derisking)

    # ------------------------------------------------------------------
    # TAB: RFQ
    # ------------------------------------------------------------------

    def _build_tab_rfq(self, parent):
        kpi_items = [
            _("RFQ Active"),
            _("RFQ Archived"),
            _("RFQ Total"),
            _("Offers Active"),
            _("Offers Archived"),
            _("Offers Total"),
            _("Work Order"),
            _("Full Supply"),
        ]
        self._build_section(parent, kpi_items)

    # ------------------------------------------------------------------
    # TAB: Saving
    # ------------------------------------------------------------------

    def _build_tab_saving(self, parent):
        kpi_items = [
            _("Theoretical Saving"),
            _("Actual Saving"),
            _("Average Saving %"),
            _("Best Saving %"),
            _("Worst Saving %"),
            _("Median Saving %"),
            _("Recurring Impact"),
            _("Non-Recurring Impact"),
        ]
        self._build_section(parent, kpi_items)

    # ------------------------------------------------------------------
    # TAB: Cost Avoidance
    # ------------------------------------------------------------------

    def _build_tab_cost_avoidance(self, parent):
        kpi_items = [
            _("Theoretical Cost Avoidance"),
            _("Actual Cost Avoidance"),
            _("Average %"),
            _("Best %"),
            _("Worst %"),
            _("Median %"),
            _("Recurring"),
            _("Non-Recurring"),
        ]
        self._build_section(parent, kpi_items)

    # ------------------------------------------------------------------
    # TAB: Derisking
    # ------------------------------------------------------------------

    def _build_tab_derisking(self, parent):
        kpi_items = [
            _("Unique New Suppliers Introduced"),
        ]
        self._build_section(parent, kpi_items)

    # ------------------------------------------------------------------
    # COSTRUZIONE SEZIONE GENERICA
    # ------------------------------------------------------------------

    def _build_section(self, parent, kpi_items):
        """
        Costruisce la struttura comune a ogni sezione:
        1. Area KPI cards
        2. Area chart (placeholder)
        3. Area tabella (placeholder)
        """
        outer = ttk.Frame(parent, padding=(8, 8, 8, 8))
        outer.pack(fill="both", expand=True)

        # --- 1. KPI Cards area ---
        cards_label_frame = ttk.LabelFrame(outer, text=_("KPI"), padding=(10, 6))
        cards_label_frame.pack(side="top", fill="x", pady=(0, 8))

        self._build_kpi_cards(cards_label_frame, kpi_items)

        # --- 2. Chart area (placeholder) ---
        chart_label_frame = ttk.LabelFrame(outer, text=_("Chart"), padding=(10, 6))
        chart_label_frame.pack(side="top", fill="both", expand=True, pady=(0, 8))

        ttk.Label(
            chart_label_frame,
            text=_("[ Chart — coming soon ]"),
            foreground="gray",
            font=(None, 9, "italic"),
        ).pack(expand=True)

        # --- 3. Table area (placeholder) ---
        table_label_frame = ttk.LabelFrame(outer, text=_("Details"), padding=(10, 6))
        table_label_frame.pack(side="top", fill="x")

        ttk.Label(
            table_label_frame,
            text=_("[ Data table — coming soon ]"),
            foreground="gray",
            font=(None, 9, "italic"),
        ).pack(expand=True, pady=6)

    # ------------------------------------------------------------------
    # KPI CARDS
    # ------------------------------------------------------------------

    def _build_kpi_cards(self, parent, kpi_items):
        """
        Dispone le KPI card in una griglia con 4 colonne.
        Ogni card mostra: etichetta + valore placeholder.
        """
        cols = 4
        for idx, label in enumerate(kpi_items):
            col = idx % cols
            row = idx // cols
            self._build_kpi_card(parent, label, row=row, col=col)

        # Distribuisci spazio orizzontale uniformemente
        for c in range(cols):
            parent.columnconfigure(c, weight=1)

    def _build_kpi_card(self, parent, label, row, col):
        """Crea una singola KPI card."""
        card = ttk.Frame(parent, relief="groove", borderwidth=1, padding=(10, 8))
        card.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        ttk.Label(
            card,
            text=label,
            font=(None, 8),
            foreground="#555555",
            wraplength=140,
            justify="center",
        ).pack()

        ttk.Label(
            card,
            text="—",
            font=(None, 16, "bold"),
            foreground="#888888",
        ).pack(pady=(4, 0))

    # ------------------------------------------------------------------
    # PLACEHOLDER HANDLERS
    # ------------------------------------------------------------------

    def _on_export_excel(self):
        """Placeholder: Export Excel (non implementato in Fase 1)."""
        pass
