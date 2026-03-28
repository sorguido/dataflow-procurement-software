# -*- coding: utf-8 -*-
"""
KpiWindow - Finestra KPI Analysis per DataFlow Procurement Software.

Fase 3: UI collegata al KPI engine.
- Header con titolo, filtri periodo, selezione anno, Export Excel (placeholder)
- Navigazione a tab: RFQ / Saving / Cost Avoidance / Derisking
- KPI cards popolate con dati reali da services.kpi_engine
- Chart e Details restano placeholder
"""

import tkinter as tk
from tkinter import ttk
import builtins
import logging

from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from services.kpi_engine import (
    get_rfq_kpi,
    get_saving_kpi,
    get_cost_avoidance_kpi,
    get_derisking_kpi,
)

# Compatibilità: _() è installata in builtins da init_i18n().
# Se non disponibile (es. test unitari), usa dummy.
if not hasattr(builtins, '_'):
    builtins._ = lambda x: x

logger = logging.getLogger('DataFlow.KpiWindow')


# ---------------------------------------------------------------------------
# Helpers di formattazione (solo display, nessuna logica KPI)
# ---------------------------------------------------------------------------

def _fmt_int(v) -> str:
    """Formatta un valore come intero (conteggi RFQ, fornitori…)."""
    try:
        return str(int(v or 0))
    except (TypeError, ValueError):
        return "0"


def _fmt_money(v) -> str:
    """Formatta un valore monetario senza simbolo, separatore migliaia a virgola."""
    try:
        return f"{float(v or 0):,.0f}"
    except (TypeError, ValueError):
        return "0"


def _fmt_pct(v) -> str:
    """Formatta una percentuale con 2 decimali."""
    try:
        return f"{float(v or 0):.2f}%"
    except (TypeError, ValueError):
        return "0.00%"


class KpiWindow(tk.Toplevel):
    """Finestra KPI Analysis — Fase 3: UI collegata all'engine."""

    _PERIOD_OPTIONS = ["1M", "3M", "12M", "3Y", "5Y", "10Y", _("All")]
    _YEAR_RANGE = list(range(2020, 2031))

    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("KPI Analysis"))
        self.resizable(True, True)

        # Variabili filtro (preparate per futura implementazione, non ancora attive)
        self._selected_period = tk.StringVar(value="12M")
        self._selected_year = tk.StringVar(value="")

        # Dizionari engine_key → ttk.Label (valore), popolati da _build_tab_*
        self._rfq_labels: dict = {}
        self._saving_labels: dict = {}
        self._ca_labels: dict = {}
        self._derisking_labels: dict = {}

        self._build_header()
        self._build_navigation()

        # Carica dati reali dall'engine
        self._load_kpi_data()

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
        items = [
            (_("RFQ Active"),      "rfq_active"),
            (_("RFQ Archived"),    "rfq_archived"),
            (_("RFQ Total"),       "rfq_total"),
            (_("Offers Active"),   "offers_active"),
            (_("Offers Archived"), "offers_archived"),
            (_("Offers Total"),    "offers_total"),
            (_("Work Order"),      "work_order"),
            (_("Full Supply"),     "full_supply"),
        ]
        self._rfq_labels = self._build_section(parent, items)

    # ------------------------------------------------------------------
    # TAB: Saving
    # ------------------------------------------------------------------

    def _build_tab_saving(self, parent):
        items = [
            (_("Theoretical Saving"),   "theoretical_saving"),
            (_("Actual Saving"),         "actual_saving"),
            (_("Average Saving %"),      "average_saving_pct"),
            (_("Best Saving %"),         "best_saving_pct"),
            (_("Worst Saving %"),        "worst_saving_pct"),
            (_("Median Saving %"),       "median_saving_pct"),
            (_("Recurring Impact"),      "recurring_impact"),
            (_("Non-Recurring Impact"),  "non_recurring_impact"),
        ]
        self._saving_labels = self._build_section(parent, items)

    # ------------------------------------------------------------------
    # TAB: Cost Avoidance
    # ------------------------------------------------------------------

    def _build_tab_cost_avoidance(self, parent):
        items = [
            (_("Theoretical Cost Avoidance"), "theoretical_cost_avoidance"),
            (_("Actual Cost Avoidance"),       "actual_cost_avoidance"),
            (_("Average %"),                   "average_pct"),
            (_("Best %"),                      "best_pct"),
            (_("Worst %"),                     "worst_pct"),
            (_("Median %"),                    "median_pct"),
            (_("Recurring"),                   "recurring"),
            (_("Non-Recurring"),               "non_recurring"),
        ]
        self._ca_labels = self._build_section(parent, items)

    # ------------------------------------------------------------------
    # TAB: Derisking
    # ------------------------------------------------------------------

    def _build_tab_derisking(self, parent):
        items = [
            (_("Unique New Suppliers Introduced"), "unique_new_suppliers_introduced"),
        ]
        self._derisking_labels = self._build_section(parent, items)

    # ------------------------------------------------------------------
    # COSTRUZIONE SEZIONE GENERICA
    # ------------------------------------------------------------------

    def _build_section(self, parent, items: list) -> dict:
        """
        Costruisce la struttura comune a ogni sezione:
        1. Area KPI cards  2. Area chart (placeholder)  3. Area tabella (placeholder)

        Args:
            parent: widget contenitore
            items:  lista di tuple (label_testo, engine_key)

        Returns:
            dict: engine_key → ttk.Label (il widget del valore, per aggiornamenti)
        """
        outer = ttk.Frame(parent, padding=(8, 8, 8, 8))
        outer.pack(fill="both", expand=True)

        # --- 1. KPI Cards area ---
        cards_label_frame = ttk.LabelFrame(outer, text=_("KPI"), padding=(10, 6))
        cards_label_frame.pack(side="top", fill="x", pady=(0, 8))

        label_refs = self._build_kpi_cards(cards_label_frame, items)

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

        return label_refs

    # ------------------------------------------------------------------
    # KPI CARDS
    # ------------------------------------------------------------------

    def _build_kpi_cards(self, parent, items: list) -> dict:
        """
        Dispone le KPI card in una griglia con 4 colonne.

        Args:
            parent: frame contenitore (LabelFrame)
            items:  lista di tuple (label_testo, engine_key)

        Returns:
            dict: engine_key → ttk.Label del valore
        """
        cols = 4
        label_refs: dict = {}
        for idx, (label_text, key) in enumerate(items):
            col = idx % cols
            row = idx // cols
            value_lbl = self._build_kpi_card(parent, label_text, row=row, col=col)
            label_refs[key] = value_lbl

        for c in range(cols):
            parent.columnconfigure(c, weight=1)

        return label_refs

    def _build_kpi_card(self, parent, label, row, col) -> ttk.Label:
        """
        Crea una singola KPI card.

        Returns:
            ttk.Label: il widget del valore (testo aggiornabile)
        """
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

        value_label = ttk.Label(
            card,
            text="—",
            font=(None, 16, "bold"),
            foreground="#888888",
        )
        value_label.pack(pady=(4, 0))
        return value_label

    # ------------------------------------------------------------------
    # PLACEHOLDER HANDLERS
    # ------------------------------------------------------------------

    def _on_export_excel(self):
        """Placeholder: Export Excel (non implementato)."""
        pass

    # ------------------------------------------------------------------
    # CARICAMENTO DATI (BINDING UI → ENGINE)
    # ------------------------------------------------------------------

    def _load_kpi_data(self):
        """
        Recupera i KPI dall'engine e aggiorna tutte le sezioni.

        Strutturato in modo che in futuro sia sufficiente richiamare
        questo metodo (es. dopo un cambio filtro) per aggiornare la UI.
        """
        try:
            rfq_data = get_rfq_kpi()
        except Exception as e:
            logger.error("[KpiWindow] get_rfq_kpi failed: %s", e)
            rfq_data = {}

        try:
            saving_data = get_saving_kpi()
        except Exception as e:
            logger.error("[KpiWindow] get_saving_kpi failed: %s", e)
            saving_data = {}

        try:
            ca_data = get_cost_avoidance_kpi()
        except Exception as e:
            logger.error("[KpiWindow] get_cost_avoidance_kpi failed: %s", e)
            ca_data = {}

        try:
            derisking_data = get_derisking_kpi()
        except Exception as e:
            logger.error("[KpiWindow] get_derisking_kpi failed: %s", e)
            derisking_data = {}

        self._update_rfq_cards(rfq_data)
        self._update_saving_cards(saving_data)
        self._update_ca_cards(ca_data)
        self._update_derisking_cards(derisking_data)

    # ------------------------------------------------------------------
    # UPDATE CARDS PER SEZIONE
    # ------------------------------------------------------------------

    def _update_rfq_cards(self, data: dict):
        """Aggiorna le card RFQ con i dati restituiti dall'engine."""
        for key, lbl in self._rfq_labels.items():
            v = data.get(key, 0)
            lbl.config(text=_fmt_int(v), foreground="#222222")

    def _update_saving_cards(self, data: dict):
        """Aggiorna le card Saving con i dati restituiti dall'engine."""
        pct_keys   = {"average_saving_pct", "best_saving_pct",
                      "worst_saving_pct", "median_saving_pct"}
        money_keys = {"theoretical_saving", "actual_saving",
                      "recurring_impact", "non_recurring_impact"}
        for key, lbl in self._saving_labels.items():
            v = data.get(key, 0)
            if key in pct_keys:
                lbl.config(text=_fmt_pct(v), foreground="#222222")
            elif key in money_keys:
                lbl.config(text=_fmt_money(v), foreground="#222222")
            else:
                lbl.config(text=str(v or 0), foreground="#222222")

    def _update_ca_cards(self, data: dict):
        """Aggiorna le card Cost Avoidance con i dati restituiti dall'engine."""
        pct_keys   = {"average_pct", "best_pct", "worst_pct", "median_pct"}
        money_keys = {"theoretical_cost_avoidance", "actual_cost_avoidance",
                      "recurring", "non_recurring"}
        for key, lbl in self._ca_labels.items():
            v = data.get(key, 0)
            if key in pct_keys:
                lbl.config(text=_fmt_pct(v), foreground="#222222")
            elif key in money_keys:
                lbl.config(text=_fmt_money(v), foreground="#222222")
            else:
                lbl.config(text=str(v or 0), foreground="#222222")

    def _update_derisking_cards(self, data: dict):
        """Aggiorna le card Derisking con i dati restituiti dall'engine."""
        for key, lbl in self._derisking_labels.items():
            v = data.get(key, 0)
            lbl.config(text=_fmt_int(v), foreground="#222222")
