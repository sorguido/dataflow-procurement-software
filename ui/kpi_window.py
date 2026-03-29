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
from tkinter import ttk, filedialog
import builtins
import logging
from datetime import date, timedelta, datetime as _dt

from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from ui.dialogs.common_dialogs import LanguagePrompt, SimpleMessageDialog
from services.kpi_engine import (
    get_rfq_kpi,
    get_saving_kpi,
    get_cost_avoidance_kpi,
    get_derisking_kpi,
    get_available_years,
)
from services.kpi_excel_export import build_kpi_workbook
from services.kpi_chart_data import (
    get_rfq_chart_data,
    get_saving_chart_data,
    get_cost_avoidance_chart_data,
    get_derisking_chart_data,
)
from ui.kpi_chart import draw_bar_chart, draw_dual_bar_chart

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
    """Formatta un valore monetario: formato italiano (punto migliaia) con simbolo €.

    Esempi:  20000 → "20.000 €"   |   1234567 → "1.234.567 €"
    """
    try:
        return f"{float(v or 0):,.0f}".replace(",", ".") + " €"
    except (TypeError, ValueError):
        return "0 €"


def _fmt_pct(v) -> str:
    """Formatta una percentuale con 2 decimali."""
    try:
        return f"{float(v or 0):.2f}%"
    except (TypeError, ValueError):
        return "0.00%"


def _t_ui(is_ita, ita, eng):
    """Helper bilingua per stringhe UI calcolate a runtime (fuori da _())."""
    return ita if is_ita else eng


# ---------------------------------------------------------------------------
# Dialog: scelta ambito export KPI
# ---------------------------------------------------------------------------

class KpiExportScopeDialog(tk.Toplevel):
    """Dialog per scegliere se esportare la sezione corrente o tutte le sezioni.

    Segue la stessa struttura di NewRdOTypeDialog in common_dialogs.py.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("Esporta KPI"))
        self.scope = None        # 'current' | 'all' | None (annullato)
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=_("Seleziona cosa esportare:"),
            font=(None, 10),
        ).pack(pady=(0, 15))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(0, 10))

        ttk.Button(
            btn_frame,
            text=_("\U0001f4cb Sezione corrente"),
            command=lambda: self._choose("current"),
            width=22,
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text=_("\U0001f4ca Tutte le sezioni"),
            command=lambda: self._choose("all"),
            width=22,
        ).pack(side="left", padx=5)

        ttk.Button(
            frame,
            text=_("\u274c Annulla"),
            command=self.destroy,
        ).pack(pady=(5, 0))

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        center_window(self)

    def _choose(self, scope):
        self.scope = scope
        self.destroy()


class KpiWindow(tk.Toplevel):
    """Finestra KPI Analysis — Fase 3 + Export Excel."""

    _PERIOD_OPTIONS = ["1M", "3M", "12M", "3Y", "5Y", "10Y", _("All")]

    # Giorni rolling per ogni preset periodo
    _ROLLING_DAYS = {
        "1M": 30, "3M": 90, "12M": 365, "3Y": 1095, "5Y": 1825, "10Y": 3650
    }

    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("KPI Analysis"))
        self.resizable(True, True)

        # Variabili filtro
        self._selected_period = tk.StringVar(value="")   # "" = nessun preset attivo
        self._selected_year   = tk.StringVar(value="")

        # Refs ai widget filtro (per binding callbacks)
        self._period_buttons: list = []
        self._year_combo = None

        # Dizionari engine_key → ttk.Label (valore), popolati da _build_tab_*
        self._rfq_labels: dict = {}
        self._saving_labels: dict = {}
        self._ca_labels: dict = {}
        self._derisking_labels: dict = {}

        # Dati correnti: popolati da _load_kpi_data, riusati da _on_export_excel
        self._current_kpi_data: dict = {}

        # Grafici: canvas per sezione e ultimi dati caricati (per resize redraw)
        self._chart_canvases: dict = {}
        self._chart_data:     dict = {}

        # Tabelle Details: treeview per sezione
        self._detail_trees: dict = {}

        self._build_header()
        self._build_navigation()

        # Popola anni disponibili e imposta filtro default (anno corrente)
        self._populate_year_filter()
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

        # Pulsanti periodo (radio-style) — mutuamente esclusivi con filtro anno
        period_frame = ttk.Frame(header)
        period_frame.pack(side="left", padx=(0, 12))
        for option in self._PERIOD_OPTIONS:
            btn = ttk.Radiobutton(
                period_frame,
                text=option,
                variable=self._selected_period,
                value=option,
                command=self._on_period_selected,
            )
            btn.pack(side="left", padx=2)
            self._period_buttons.append(btn)

        # Separatore verticale
        ttk.Separator(header, orient="vertical").pack(side="left", fill="y", padx=(0, 12))

        # Label "Year:"
        ttk.Label(header, text=_("Year:")).pack(side="left", padx=(0, 4))

        # Combobox anno — mutuamente esclusivo con preset periodo
        self._year_combo = ttk.Combobox(
            header,
            textvariable=self._selected_year,
            values=[],
            width=6,
            state="readonly",
        )
        self._year_combo.pack(side="left", padx=(0, 20))
        self._year_combo.bind("<<ComboboxSelected>>", self._on_year_selected)

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
            (_("RFQ Active"),       "rfq_active"),
            (_("RFQ Archived"),     "rfq_archived"),
            (_("RFQ Total"),        "rfq_total"),
            (_("RFQ Not Expired"),  "rfq_not_expired"),
            (_("RFQ Expired"),      "rfq_expired"),
            (_("Work Order"),       "work_order"),
            (_("Full Supply"),      "full_supply"),
        ]
        self._rfq_labels = self._build_section(parent, items, section_key='rfq')

    # ------------------------------------------------------------------
    # TAB: Saving
    # ------------------------------------------------------------------

    def _build_tab_saving(self, parent):
        items = [
            (_("Theoretical Saving"),   "theoretical_saving"),
            (_("Actual Saving"),         "actual_saving"),
            (_("Average Theoretical Saving %"), "average_saving_pct"),
            (_("Best Saving %"),         "best_saving_pct"),
            (_("Worst Saving %"),        "worst_saving_pct"),
            (_("Median Saving %"),       "median_saving_pct"),
            (_("Recurring Impact (€)"),     "recurring_impact"),
            (_("Non-Recurring Impact (€)"),  "non_recurring_impact"),
        ]
        self._saving_labels = self._build_section(parent, items, section_key='saving')

    # ------------------------------------------------------------------
    # TAB: Cost Avoidance
    # ------------------------------------------------------------------

    def _build_tab_cost_avoidance(self, parent):
        items = [
            (_("Theoretical Cost Avoidance"), "theoretical_cost_avoidance"),
            (_("Actual Cost Avoidance"),       "actual_cost_avoidance"),
            (_("Average Theoretical CA %"),    "average_pct"),
            (_("Best %"),                      "best_pct"),
            (_("Worst %"),                     "worst_pct"),
            (_("Median %"),                    "median_pct"),
            (_("Recurring (\u20ac)"),     "recurring"),
            (_("Non-Recurring (\u20ac)"), "non_recurring"),
        ]
        self._ca_labels = self._build_section(parent, items, section_key='ca')

    # ------------------------------------------------------------------
    # TAB: Derisking
    # ------------------------------------------------------------------

    def _build_tab_derisking(self, parent):
        items = [
            (_("Unique New Suppliers Introduced"), "unique_new_suppliers_introduced"),
        ]
        self._derisking_labels = self._build_section(parent, items, section_key='derisking')

    # ------------------------------------------------------------------
    # COSTRUZIONE SEZIONE GENERICA
    # ------------------------------------------------------------------

    def _build_section(self, parent, items: list, section_key: str = '') -> dict:
        """
        Costruisce la struttura comune a ogni sezione:
        1. Area KPI cards  2. Area chart (Canvas)  3. Area tabella (placeholder)

        Args:
            parent:      widget contenitore
            items:       lista di tuple (label_testo, engine_key)
            section_key: chiave per i dizionari _chart_canvases / _chart_data

        Returns:
            dict: engine_key → ttk.Label (il widget del valore, per aggiornamenti)
        """
        outer = ttk.Frame(parent, padding=(8, 8, 8, 8))
        outer.pack(fill="both", expand=True)

        # --- 1. KPI Cards area ---
        cards_label_frame = ttk.LabelFrame(outer, text=_("KPI"), padding=(10, 6))
        cards_label_frame.pack(side="top", fill="x", pady=(0, 8))

        label_refs = self._build_kpi_cards(cards_label_frame, items)

        # --- 2. Chart area ---
        chart_label_frame = ttk.LabelFrame(outer, text=_("Chart"), padding=(4, 4))
        chart_label_frame.pack(side="top", fill="both", expand=True, pady=(0, 8))

        canvas = tk.Canvas(
            chart_label_frame,
            height=190,
            bg='#F8F8F8',
            highlightthickness=0,
        )
        canvas.pack(fill="both", expand=True)

        if section_key:
            self._chart_canvases[section_key] = canvas
            canvas.bind(
                '<Configure>',
                lambda e, k=section_key: self._on_chart_resize(k),
            )

        # --- 3. Table area ---
        table_label_frame = ttk.LabelFrame(outer, text=_("Details"), padding=(4, 4))
        table_label_frame.pack(side="top", fill="x")

        self._build_detail_table(table_label_frame, section_key)

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
            wraplength=200,
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
    # DETAIL TABLE
    # ------------------------------------------------------------------

    def _build_detail_table(self, parent, section_key: str) -> None:
        """Costruisce il Treeview read-only nella sezione Details."""
        if not section_key:
            return

        is_ita = _("All") != "All"

        if section_key == 'rfq':
            col_specs = [
                ('period', _t_ui(is_ita, 'Periodo',    'Period'),              70, 'center'),
                ('count',  _t_ui(is_ita, 'RFQ Emesse', 'RFQ Issued'),         100, 'center'),
            ]
        elif section_key == 'saving':
            col_specs = [
                ('period', _t_ui(is_ita, 'Periodo',           'Period'),               70, 'center'),
                ('theor',  _t_ui(is_ita, 'Saving Teorico',    'Theoretical Saving'),  150, 'e'),
                ('actual', _t_ui(is_ita, 'Saving Effettivo',  'Actual Saving'),       150, 'e'),
            ]
        elif section_key == 'ca':
            col_specs = [
                ('period', _t_ui(is_ita, 'Periodo',                   'Period'),                        70, 'center'),
                ('theor',  _t_ui(is_ita, 'Cost Avoidance Teorico',    'Theoretical Cost Avoidance'),   180, 'e'),
                ('actual', _t_ui(is_ita, 'Cost Avoidance Effettivo',  'Actual Cost Avoidance'),        180, 'e'),
            ]
        elif section_key == 'derisking':
            col_specs = [
                ('period', _t_ui(is_ita, 'Periodo',        'Period'),                     70, 'center'),
                ('count',  _t_ui(is_ita, 'Nuovi Fornitori','New Suppliers Introduced'),  160, 'center'),
            ]
        else:
            return

        col_ids = [c[0] for c in col_specs]

        frame = ttk.Frame(parent)
        frame.pack(fill='x')

        tree = ttk.Treeview(
            frame,
            columns=col_ids,
            show='headings',
            height=6,
            selectmode='browse',
        )

        for col_id, heading, width, anchor in col_specs:
            tree.heading(col_id, text=heading)
            tree.column(col_id, width=width, minwidth=40, anchor=anchor, stretch=True)

        vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        tree.pack(side='left', fill='x', expand=True)

        self._detail_trees[section_key] = tree

    def _populate_table(self, key: str, data: list) -> None:
        """Popola il Treeview Details con i dati del grafico (stessi bucket)."""
        tree = self._detail_trees.get(key)
        if tree is None:
            return
        for row in tree.get_children():
            tree.delete(row)
        for d in data:
            if key in ('rfq', 'derisking'):
                tree.insert('', 'end', values=(d['label'], _fmt_int(d['count'])))
            else:  # saving, ca
                tree.insert('', 'end', values=(
                    d['label'],
                    _fmt_money(d['theoretical']),
                    _fmt_money(d['actual']),
                ))

    # ------------------------------------------------------------------
    # FILTRI TEMPORALI
    # ------------------------------------------------------------------

    def _period_to_dates(self, period: str) -> tuple:
        """
        Traduce un preset periodo in (date_from, date_to) come stringhe ISO.

        Semantica rolling (tutti i valori calcolati a ritroso da oggi):
            1M  = ultimi 30 giorni
            3M  = ultimi 90 giorni
            12M = ultimi 365 giorni
            3Y  = ultimi 3 anni  (1 095 giorni)
            5Y  = ultimi 5 anni  (1 825 giorni)
            10Y = ultimi 10 anni (3 650 giorni)
            All = nessun filtro  → (None, None)
        """
        today = date.today()
        days  = self._ROLLING_DAYS.get(period)
        if days is not None:
            return (today - timedelta(days=days)).isoformat(), today.isoformat()
        return None, None  # All o preset non riconosciuto

    def _populate_year_filter(self):
        """
        Popola il combobox Year con gli anni realmente presenti nel DB.

        Default: anno corrente se disponibile, altrimenti il più recente.
        Stato iniziale: anno attivo, nessun preset periodo selezionato.
        """
        years = get_available_years()
        year_values = [""] + [str(y) for y in years]
        if self._year_combo is not None:
            self._year_combo["values"] = year_values

        current_year = str(date.today().year)
        if current_year in year_values:
            self._selected_year.set(current_year)
        elif years:
            self._selected_year.set(str(years[-1]))
        else:
            self._selected_year.set("")

        # Anno attivo come default → nessun preset periodo
        self._selected_period.set("")

    def _on_period_selected(self):
        """
        Callback: l'utente ha cliccato un preset periodo.
        Mutua esclusione: azzera l'anno e ricarica i KPI.
        """
        self._selected_year.set("")
        self._load_kpi_data()

    def _on_year_selected(self, event=None):
        """
        Callback: l'utente ha selezionato un anno dal combobox.
        Mutua esclusione: azzera il preset periodo e ricarica i KPI.
        """
        self._selected_period.set("")
        self._load_kpi_data()

    # ------------------------------------------------------------------
    # PLACEHOLDER HANDLERS
    # ------------------------------------------------------------------

    def _on_export_excel(self):
        """Export Excel KPI: dialog scope → dialog lingua → build workbook → salva."""

        # 1. Scelta ambito (sezione corrente / tutte le sezioni)
        scope_dlg = KpiExportScopeDialog(self)
        self.wait_window(scope_dlg)
        if not scope_dlg.scope:
            return

        # 2. Scelta lingua (riusa LanguagePrompt già presente nel progetto)
        lang_prompt = LanguagePrompt(self)
        self.wait_window(lang_prompt)
        if not lang_prompt.choice:
            return

        lang    = lang_prompt.choice
        is_ita  = (lang == 'ita')

        # 3. Sezione attiva nel notebook
        try:
            tab_idx = self._notebook.index(self._notebook.select())
            _tab_map = ['RFQ', 'Saving', 'Cost Avoidance', 'Derisking']
            current_section = _tab_map[tab_idx] if tab_idx < len(_tab_map) else 'RFQ'
        except Exception:
            current_section = 'RFQ'

        # 4. Etichetta filtro (dipende dalla lingua scelta)
        year_str = self._selected_year.get()
        period   = self._selected_period.get()
        if year_str:
            filter_label = (_t_ui(is_ita, "Anno: ", "Year: ")) + year_str
        elif period:
            filter_label = (_t_ui(is_ita, "Periodo: ", "Period: ")) + period
        else:
            filter_label = _t_ui(is_ita, "Tutti i dati", "All data")

        # 5. Costruisci workbook (riusa i dati già caricati)
        kpi = self._current_kpi_data
        try:
            wb = build_kpi_workbook(
                rfq_data=kpi.get('rfq', {}),
                saving_data=kpi.get('saving', {}),
                ca_data=kpi.get('ca', {}),
                derisking_data=kpi.get('derisking', {}),
                filter_label=filter_label,
                scope=scope_dlg.scope,
                current_section=current_section,
                lang=lang,
            )
        except Exception as e:
            logger.error("[KpiWindow] build_kpi_workbook failed: %s", e, exc_info=True)
            SimpleMessageDialog(
                self, _("Errore Esportazione"),
                _("Errore durante l'esportazione: {}").format(e), "error"
            )
            return

        # 6. Dialog salvataggio
        ts           = _dt.now().strftime('%Y%m%d_%H%M')
        default_name = f"KPI_DataFlow_{ts}.xlsx"
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title=_("Salva Export KPI"),
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not save_path:
            try:
                wb.close()
            except Exception:
                pass
            return

        # 7. Salva e notifica
        try:
            wb.save(save_path)
            SimpleMessageDialog(
                self, _("Successo"),
                _("Export KPI completato:\n{}").format(save_path), "info"
            )
            logger.info("[KpiWindow] Export KPI salvato: %s", save_path)
        except Exception as e:
            logger.error("[KpiWindow] wb.save failed: %s", e, exc_info=True)
            SimpleMessageDialog(
                self, _("Errore Esportazione"),
                _("Errore durante l'esportazione: {}").format(e), "error"
            )
        finally:
            try:
                wb.close()
            except Exception:
                pass


    # ------------------------------------------------------------------
    # CARICAMENTO DATI (BINDING UI → ENGINE)
    # ------------------------------------------------------------------

    def _load_kpi_data(self):
        """
        Recupera i KPI dall'engine con i filtri attivi e aggiorna tutte le sezioni.

        Logica filtri (mutuamente esclusivi, see _on_period_selected / _on_year_selected):
        - Anno selezionato    → year=<int> all'engine, date_from/date_to=None
        - Preset periodo      → date_from/date_to rolling, year=None
        - Nessun filtro (All) → tutti i dati (nessun parametro passato)
        """
        year_str = self._selected_year.get()
        period   = self._selected_period.get()

        year      = None
        date_from = None
        date_to   = None

        if year_str:
            try:
                year = int(year_str)
            except ValueError:
                pass
        elif period:
            date_from, date_to = self._period_to_dates(period)

        kw = dict(date_from=date_from, date_to=date_to, year=year)

        try:
            rfq_data = get_rfq_kpi(**kw)
        except Exception as e:
            logger.error("[KpiWindow] get_rfq_kpi failed: %s", e)
            rfq_data = {}

        try:
            saving_data = get_saving_kpi(**kw)
        except Exception as e:
            logger.error("[KpiWindow] get_saving_kpi failed: %s", e)
            saving_data = {}

        try:
            ca_data = get_cost_avoidance_kpi(**kw)
        except Exception as e:
            logger.error("[KpiWindow] get_cost_avoidance_kpi failed: %s", e)
            ca_data = {}

        try:
            derisking_data = get_derisking_kpi(**kw)
        except Exception as e:
            logger.error("[KpiWindow] get_derisking_kpi failed: %s", e)
            derisking_data = {}

        self._current_kpi_data = {
            'rfq':       rfq_data,
            'saving':    saving_data,
            'ca':        ca_data,
            'derisking': derisking_data,
        }

        self._update_rfq_cards(rfq_data)
        self._update_saving_cards(saving_data)
        self._update_ca_cards(ca_data)
        self._update_derisking_cards(derisking_data)

        # Aggiorna i grafici subito dopo le card (schedule per permettere
        # al layout di stabilizzarsi prima del rendering Canvas)
        self.after(80, self._update_charts)

    # ------------------------------------------------------------------
    # GRAFICI
    # ------------------------------------------------------------------

    def _update_charts(self) -> None:
        """Recupera le serie temporali e ridisegna tutti i chart Canvas attivi."""
        year_str = self._selected_year.get()
        period   = self._selected_period.get()
        year      = None
        date_from = None
        date_to   = None

        if year_str:
            try:
                year = int(year_str)
            except ValueError:
                pass
        elif period:
            date_from, date_to = self._period_to_dates(period)

        kw = dict(date_from=date_from, date_to=date_to, year=year)

        # RFQ
        canvas = self._chart_canvases.get('rfq')
        if canvas:
            data = get_rfq_chart_data(**kw)
            self._chart_data['rfq'] = data
            self._render_rfq_chart(canvas, data)
            self._populate_table('rfq', data)

        # Saving
        canvas = self._chart_canvases.get('saving')
        if canvas:
            data = get_saving_chart_data(**kw)
            self._chart_data['saving'] = data
            self._render_saving_chart(canvas, data)
            self._populate_table('saving', data)

        # Cost Avoidance
        canvas = self._chart_canvases.get('ca')
        if canvas:
            data = get_cost_avoidance_chart_data(**kw)
            self._chart_data['ca'] = data
            self._render_ca_chart(canvas, data)
            self._populate_table('ca', data)

        # Derisking
        canvas = self._chart_canvases.get('derisking')
        if canvas:
            data = get_derisking_chart_data(**kw)
            self._chart_data['derisking'] = data
            self._render_derisking_chart(canvas, data)
            self._populate_table('derisking', data)

    def _on_chart_resize(self, key: str) -> None:
        """Ridisegna il chart `key` al resize del Canvas (debounced 40 ms)."""
        attr = f'_resize_job_{key}'
        existing = getattr(self, attr, None)
        if existing:
            try:
                self.after_cancel(existing)
            except Exception:
                pass
        job = self.after(40, lambda k=key: self._redraw_stored_chart(k))
        setattr(self, attr, job)

    def _redraw_stored_chart(self, key: str) -> None:
        """Ridisegna usando i dati già in cache (nessuna query DB)."""
        canvas = self._chart_canvases.get(key)
        data   = self._chart_data.get(key)
        if canvas is None or data is None:
            return
        if key == 'rfq':
            self._render_rfq_chart(canvas, data)
        elif key == 'saving':
            self._render_saving_chart(canvas, data)
        elif key == 'ca':
            self._render_ca_chart(canvas, data)
        elif key == 'derisking':
            self._render_derisking_chart(canvas, data)

    def _render_rfq_chart(self, canvas, data: list) -> None:
        is_ita = _("All") != "All"
        draw_bar_chart(
            canvas,
            [{'label': d['label'], 'value': d['count']} for d in data],
            y_fmt='int',
            title=_t_ui(is_ita, "RFQ emesse per periodo", "RFQ issued per period"),
            y_label=_t_ui(is_ita, "Numero RFQ", "No. RFQ"),
            x_label=_t_ui(is_ita, "Periodo", "Period"),
        )

    def _render_saving_chart(self, canvas, data: list) -> None:
        is_ita = _("All") != "All"
        draw_dual_bar_chart(
            canvas,
            data,
            label1=_t_ui(is_ita, "Teorico", "Theoretical"),
            label2=_t_ui(is_ita, "Effettivo", "Actual"),
            title=_t_ui(is_ita, "Saving teorico vs effettivo per periodo",
                        "Theoretical vs Actual Saving per period"),
            y_label="Saving (€)",
            x_label=_t_ui(is_ita, "Periodo", "Period"),
        )

    def _render_ca_chart(self, canvas, data: list) -> None:
        is_ita = _("All") != "All"
        draw_dual_bar_chart(
            canvas,
            data,
            label1=_t_ui(is_ita, "Teorico", "Theoretical"),
            label2=_t_ui(is_ita, "Effettivo", "Actual"),
            title=_t_ui(is_ita, "Cost avoidance teorico vs effettivo per periodo",
                        "Theoretical vs Actual CA per period"),
            y_label="Cost Avoidance (€)",
            x_label=_t_ui(is_ita, "Periodo", "Period"),
        )

    def _render_derisking_chart(self, canvas, data: list) -> None:
        is_ita = _("All") != "All"
        draw_bar_chart(
            canvas,
            [{'label': d['label'], 'value': d['count']} for d in data],
            y_fmt='int',
            title=_t_ui(is_ita, "Nuovi fornitori introdotti per periodo",
                        "New suppliers introduced per period"),
            y_label=_t_ui(is_ita, "Nuovi fornitori", "New suppliers"),
            x_label=_t_ui(is_ita, "Periodo", "Period"),
        )

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
