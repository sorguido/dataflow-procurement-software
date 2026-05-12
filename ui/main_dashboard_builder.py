"""
ui/main_dashboard_builder.py

Pure UI builder for MainWindow's dashboard.
Extracted conservatively from MainWindow.__init__ in dataflow.py
as part of release 2.1.0 refactoring.

RESPONSIBILITIES
- Create all dashboard widgets
- Assign them to app.<attribute>
- Set grid layout
- Register event bindings

NOT RESPONSIBLE FOR
- Loading data into sheets (_load_vsm_events stays in __init__)
- Populating filter combo values (populate_vsm_username_filter stays in __init__)
- Triggering refresh_data / update_button_visibility / check_for_autobackup
"""

import os
import tkinter as tk
from tkinter import ttk
import webbrowser

from PIL import Image, ImageTk
from tkcalendar import DateEntry

from utils.resource_utils import resource_path
from utils.i18n_utils import tr, get_current_language
from ui.components.main_dashboard_toolbar import MainDashboardToolbar
from ui.components.collapsible_filters import CollapsibleFilters


def build_main_dashboard(app):
    """
    Build all UI widgets for the main dashboard and attach them to *app*
    (a MainWindow instance).  This function is a pure widget builder:
    it creates, configures, and lays out every widget, but does not
    load data or trigger any application logic.

    Parameters
    ----------
    app : MainWindow
        The MainWindow instance.  All widgets are assigned as app.<attr>.
    """

    # -----------------------------------------------------------------------
    # Frame top: toolbar row (button bar at the top of the window)
    # -----------------------------------------------------------------------
    frame_top = ttk.Frame(app.root)
    try:
        logo_path = resource_path(os.path.join("add_data", "logo_dataflow.png"))
        if os.path.exists(logo_path):
            img = Image.open(logo_path)
            # BUG #51 FIX: Check dimensioni valide prima di divisione per evitare ZeroDivisionError
            if img.width > 0 and img.height > 0:
                img.thumbnail((int(40 * (img.width / img.height)), 40), Image.Resampling.LANCZOS)
                app.logo_photo = ImageTk.PhotoImage(img)
                ttk.Label(frame_top, image=app.logo_photo).pack(side="left", padx=(0, 20), anchor="w")
    except Exception as e:
        print(f"Errore caricamento logo: {e}")

    # --- Pulsanti Operativi (Riga Superiore) ---
    # 1. New Event (dinamico: RdO o VSM in base al tab attivo - Step 4D.6)
    app.btn_new_rdo = ttk.Button(frame_top, text=tr("➕ New Event"), command=app.open_new_event)
    app.btn_new_rdo.pack(side="left", padx=(0, 10))

    # 2. Actions dropdown (sostituisce Delete/Duplicate/Archive/Reactivate)
    # Pattern desktop classico: Menubutton con menu contestuale
    app.btn_actions = ttk.Menubutton(
        frame_top,
        text=tr("⚡ Actions"),
        state="disabled"
    )
    app.btn_actions.pack(side="left", padx=(0, 10))

    # Menu popup per azioni contestuali (popolato dinamicamente)
    app.actions_menu = tk.Menu(app.btn_actions, tearoff=0)
    app.btn_actions.config(menu=app.actions_menu)

    # 3. Export Excel (Export Globale)
    app.btn_mega_export = ttk.Button(frame_top, text=tr("📥 Export Excel"), command=app.mega_export_excel)
    app.btn_mega_export.pack(side="left", padx=(0, 20))

    # 4. KPI Dashboard
    app.btn_kpi = ttk.Button(frame_top, text=tr("≋ KPI"), command=app.on_kpi_click)
    app.btn_kpi.pack(side="left", padx=(0, 20))

    # --- MODIFICA: Aggiunto pulsante Licenza e riordinato ---
    app.btn_guida = ttk.Button(frame_top, text=tr("❓ Help"), command=app.open_help_window)
    app.btn_guida.pack(side="right")
    app.btn_license = ttk.Button(frame_top, text=tr("≡ License"), command=app.open_license_window)
    app.btn_license.pack(side="right", padx=(0, 10))
    app.btn_settings = ttk.Button(frame_top, text=tr("⚙️ Settings"), command=app.open_settings_window)
    app.btn_settings.pack(side="right", padx=(0, 10))
    # --- FINE MODIFICA ---

    # -----------------------------------------------------------------------
    # CONVERSIONE LAYOUT PRINCIPALE A GRID()
    # Grid è più robusto di pack() per gestire widget con show/hide dinamici
    # -----------------------------------------------------------------------
    app.root.grid_rowconfigure(3, weight=1)  # Row 3 (notebook) si espande verticalmente
    app.root.grid_columnconfigure(0, weight=1)  # Colonna 0 si espande orizzontalmente

    # Row 0: Toolbar pulsanti
    frame_top.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

    # Row 1: Global Search Toolbar (Step 2-4)
    app.main_dashboard_toolbar = MainDashboardToolbar(app.root, app)
    app.main_dashboard_toolbar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

    # Row 2: Filtri collassabili (Step 5)
    app.collapsible_filters = CollapsibleFilters(app.root, label_text=tr("Search Filters"))
    # Salva configurazione grid per expand()/collapse()
    app.collapsible_filters.set_grid_config(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))
    # Default collapsed: grid_remove() lo nasconde senza lasciare gap
    app.collapsible_filters.collapse()

    # search_frame è il filters_frame interno di CollapsibleFilters.
    # Ospita due sub-frame: rfq_filter_subframe e vsm_filter_subframe,
    # visibili alternativamente in base al tab attivo (context-aware).
    # Col=0 si espande (contenuto context), col=1 fisso (pulsanti condivisi).
    search_frame = app.collapsible_filters.filters_frame
    search_frame.columnconfigure(0, weight=1)

    # -----------------------------------------------------------------------
    # Sub-frame RFQ (visibile di default: il primo tab è RFQ)
    # -----------------------------------------------------------------------
    app.rfq_filter_subframe = ttk.Frame(search_frame)
    app.rfq_filter_subframe.grid(row=0, column=0, sticky="nsew")
    _rf = app.rfq_filter_subframe  # alias locale per brevità

    app.search_vars = {name: tk.StringVar() for name in ['global', 'num', 'ref', 'forn', 'cod', 'desc', 'ord', 'cod_grezzo', 'dis_grezzo', 'mat_cl']}
    app.search_tipo = tk.StringVar(value=tr("All"))

    ttk.Label(_rf, text=tr("RfQ Number:")).grid(row=0, column=0, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['num']).grid(row=0, column=1, sticky="ew")
    ttk.Label(_rf, text=tr("RfQ Type:")).grid(row=0, column=2, sticky="w")
    ttk.Combobox(_rf, textvariable=app.search_tipo, values=[tr("All"), tr("Full Supply"), tr("Work Order")], state="readonly").grid(row=0, column=3, sticky="ew")
    ttk.Label(_rf, text=tr("Reference:")).grid(row=1, column=0, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['ref']).grid(row=1, column=1, sticky="ew")
    ttk.Label(_rf, text=tr("Supplier:")).grid(row=1, column=2, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['forn']).grid(row=1, column=3, sticky="ew")
    ttk.Label(_rf, text=tr("Material Code:")).grid(row=2, column=0, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['cod']).grid(row=2, column=1, sticky="ew")
    ttk.Label(_rf, text=tr("Material Description:")).grid(row=2, column=2, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['desc']).grid(row=2, column=3, sticky="ew")
    ttk.Label(_rf, text=tr("Order Number:")).grid(row=0, column=4, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['ord']).grid(row=0, column=5, sticky="ew")
    ttk.Label(_rf, text=tr("Raw Code:")).grid(row=3, column=0, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['cod_grezzo']).grid(row=3, column=1, sticky="ew")
    ttk.Label(_rf, text=tr("Raw Attachment:")).grid(row=3, column=2, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['dis_grezzo']).grid(row=3, column=3, sticky="ew")
    ttk.Label(_rf, text=tr("Material for Processing:")).grid(row=3, column=4, sticky="w")
    ttk.Entry(_rf, textvariable=app.search_vars['mat_cl']).grid(row=3, column=5, sticky="ew")

    ttk.Label(_rf, text=tr("User:")).grid(row=1, column=4, sticky="w")
    default_user_value = app.current_username if getattr(app, 'current_username', '') else app.all_users_placeholder
    app.username_filter_var = tk.StringVar(value=default_user_value)
    app.user_filter_combo = ttk.Combobox(
        _rf,
        textvariable=app.username_filter_var,
        state="readonly",
        values=[default_user_value]
    )
    app.user_filter_combo.grid(row=1, column=5, sticky="ew")
    app.user_filter_combo.bind("<<ComboboxSelected>>", lambda _e: app.refresh_data())

    app.date_entries = {}
    for i, (lbl, key) in enumerate([(tr("From:"), "emm_da"), (tr("To:"), "emm_a"), (tr("From:"), "scad_da"), (tr("To:"), "scad_a")]):
        row, col_lbl, col_entry = (4 + i // 2, (i % 2) * 2, (i % 2) * 2 + 1)
        prefix = tr("Issue Date ") if i < 2 else tr("Expiry Date ")
        ttk.Label(_rf, text=prefix + lbl).grid(row=row, column=col_lbl, sticky="w")
        de = DateEntry(_rf, date_pattern='dd/mm/yyyy', locale=('it_IT' if get_current_language() == 'it' else 'en_US'))
        de.grid(row=row, column=col_entry, sticky="ew")
        de.delete(0, 'end')
        app.date_entries[key] = de
    for i in range(1, 6, 2):
        _rf.grid_columnconfigure(i, weight=1)

    # -----------------------------------------------------------------------
    # Sub-frame VSM (nascosto di default: il primo tab è RFQ)
    # -----------------------------------------------------------------------
    app.vsm_filter_subframe = ttk.Frame(search_frame)
    app.vsm_filter_subframe.grid(row=0, column=0, sticky="nsew")
    app.vsm_filter_subframe.grid_remove()  # nascosto finché non si attiva un tab VSM
    _vsf = app.vsm_filter_subframe

    # --- Variabili filtro VSM ---
    default_vsm_user = app.current_username if getattr(app, 'current_username', '') else app.all_users_placeholder
    app.vsm_username_filter_var = tk.StringVar(value=default_vsm_user)
    app.vsm_action_var = tk.StringVar()
    app.vsm_repetitive_var = tk.StringVar()
    app.vsm_theoretical_from_var = tk.StringVar()
    app.vsm_theoretical_to_var = tk.StringVar()
    app.vsm_actual_from_var = tk.StringVar()
    app.vsm_actual_to_var = tk.StringVar()

    # --- Riga 0: filtri comuni a tutti i tab VSM (Utente, Dal, Al) ---
    app.vsm_user_filter_label = ttk.Label(_vsf, text=tr("User:"))
    app.vsm_user_filter_label.grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
    _vsm_cb = ttk.Combobox(
        _vsf,
        textvariable=app.vsm_username_filter_var,
        state="readonly",
        width=20,
    )
    _vsm_cb.grid(row=0, column=1, sticky="w", pady=5)
    _vsm_cb.bind("<<ComboboxSelected>>", lambda _e: app._on_vsm_username_filter_changed())
    app.vsm_user_filter_combos = [_vsm_cb]

    app.vsm_date_from_label = ttk.Label(_vsf, text=tr("From:"))
    app.vsm_date_from_label.grid(row=0, column=2, sticky="w", padx=(10, 5), pady=5)
    app.vsm_date_from_entry = DateEntry(
        _vsf, date_pattern='dd/mm/yyyy',
        locale=('it_IT' if get_current_language() == 'it' else 'en_US'),
    )
    app.vsm_date_from_entry.grid(row=0, column=3, sticky="w", pady=5)
    app.vsm_date_from_entry.delete(0, 'end')

    app.vsm_date_to_label = ttk.Label(_vsf, text=tr("To:"))
    app.vsm_date_to_label.grid(row=0, column=4, sticky="w", padx=(10, 5), pady=5)
    app.vsm_date_to_entry = DateEntry(
        _vsf, date_pattern='dd/mm/yyyy',
        locale=('it_IT' if get_current_language() == 'it' else 'en_US'),
    )
    app.vsm_date_to_entry.grid(row=0, column=5, sticky="w", pady=5)
    app.vsm_date_to_entry.delete(0, 'end')
    app._vsm_common_date_widgets = [
        app.vsm_date_from_label,
        app.vsm_date_from_entry,
        app.vsm_date_to_label,
        app.vsm_date_to_entry,
    ]

    # --- Spec frame Saving / Cost Avoidance (riga 1, condiviso) ---
    _vsm_sc_spec = ttk.Frame(_vsf)
    _vsm_sc_spec.grid(row=1, column=0, columnspan=6, sticky="ew")
    ttk.Label(_vsm_sc_spec, text=tr("Action:")).grid(row=0, column=0, sticky="w", padx=(0, 5), pady=3)
    ttk.Combobox(
        _vsm_sc_spec,
        textvariable=app.vsm_action_var,
        state="readonly",
        values=["", tr("Negotiation"), "Derisking", tr("Other")],
        width=18,
    ).grid(row=0, column=1, sticky="w", pady=3)
    ttk.Label(_vsm_sc_spec, text=tr("Repetitive:")).grid(row=0, column=2, sticky="w", padx=(10, 5), pady=3)
    ttk.Combobox(
        _vsm_sc_spec,
        textvariable=app.vsm_repetitive_var,
        state="readonly",
        values=["", tr("Yes"), tr("No")],
        width=8,
    ).grid(row=0, column=3, sticky="w", pady=3)
    ttk.Label(_vsm_sc_spec, text=tr("Theoretical From:")).grid(row=1, column=0, sticky="w", padx=(0, 5), pady=3)
    ttk.Entry(_vsm_sc_spec, textvariable=app.vsm_theoretical_from_var, width=12).grid(row=1, column=1, sticky="w", pady=3)
    ttk.Label(_vsm_sc_spec, text=tr("To:")).grid(row=1, column=2, sticky="w", padx=(10, 5), pady=3)
    ttk.Entry(_vsm_sc_spec, textvariable=app.vsm_theoretical_to_var, width=12).grid(row=1, column=3, sticky="w", pady=3)
    ttk.Label(_vsm_sc_spec, text=tr("Actual From:")).grid(row=2, column=0, sticky="w", padx=(0, 5), pady=3)
    ttk.Entry(_vsm_sc_spec, textvariable=app.vsm_actual_from_var, width=12).grid(row=2, column=1, sticky="w", pady=3)
    ttk.Label(_vsm_sc_spec, text=tr("To:")).grid(row=2, column=2, sticky="w", padx=(10, 5), pady=3)
    ttk.Entry(_vsm_sc_spec, textvariable=app.vsm_actual_to_var, width=12).grid(row=2, column=3, sticky="w", pady=3)

    # --- Spec frame Derisking (riga 1, vuoto — i fornitori potenziali non hanno Ripetitivo) ---
    _vsm_dr_spec = ttk.Frame(_vsf)
    _vsm_dr_spec.grid(row=1, column=0, columnspan=6, sticky="ew")
    _vsm_dr_spec.grid_remove()  # nascosto di default; mostrato solo nel tab Derisking

    # Mappa tab status → spec frame (per _update_filter_panel_for_current_tab)
    app._vsm_spec_frames = {
        'vsm_saving': _vsm_sc_spec,
        'vsm_cost_avoidance': _vsm_sc_spec,
        'vsm_derisking': _vsm_dr_spec,
    }

    # -----------------------------------------------------------------------
    # Pulsanti condivisi (sempre visibili accanto al sub-frame attivo)
    # -----------------------------------------------------------------------
    btn_search_frame = ttk.Frame(search_frame)
    btn_search_frame.grid(row=0, column=1, sticky="ns", padx=20)
    ttk.Button(btn_search_frame, text=tr("🔍 Search"), command=app.search_requests).pack(fill="x", expand=True, pady=2)
    ttk.Button(btn_search_frame, text=tr("🧹 Clear Filters"), command=app.clear_filters).pack(fill="x", expand=True, pady=2)

    # -----------------------------------------------------------------------
    # Row 3: Notebook (dopo i filtri collassabili)
    # -----------------------------------------------------------------------
    app.notebook = ttk.Notebook(app.root)
    app.notebook.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)

    app.tab_attive = ttk.Frame(app.notebook)
    app.tab_archiviate = ttk.Frame(app.notebook)
    app.notebook.add(app.tab_attive, text=tr("Active RfQs"))
    app.notebook.add(app.tab_archiviate, text=tr("Archived RfQs"))

    # Tab VSM (Step 4A: Direct Tab Integration)
    app.tab_saving = ttk.Frame(app.notebook)
    app.tab_cost_avoidance = ttk.Frame(app.notebook)
    app.tab_derisking = ttk.Frame(app.notebook)
    app.notebook.add(app.tab_saving, text=tr("Saving"))
    app.notebook.add(app.tab_cost_avoidance, text=tr("Cost Avoidance"))
    app.notebook.add(app.tab_derisking, text=tr("Derisking"))

    # Step 4B: Crea sheet VSM per ogni tab.
    app.sheet_saving = app._create_vsm_event_sheet(app.tab_saving, event_type="Saving")
    app.sheet_cost_avoidance = app._create_vsm_event_sheet(app.tab_cost_avoidance, event_type="Cost Avoidance")
    app.sheet_derisking = app._create_supplier_sheet(app.tab_derisking)

    # -----------------------------------------------------------------------
    # Row 4: Footer
    # -----------------------------------------------------------------------
    footer_frame = ttk.Frame(app.root)
    footer_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
    ttk.Label(footer_frame, text=tr("v.2.3.0 - Developed by ")).pack(side="left")
    name_label = ttk.Label(footer_frame, text="Guido Sorarù", foreground="blue", cursor="hand2")
    name_label.pack(side="left")
    name_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/guido-soraru-buyer/"))
    ttk.Label(footer_frame, text=tr(" © 2025–2026 - Released under GNU GPLv3 license")).pack(side="left")

    # -----------------------------------------------------------------------
    # RFQ treeviews
    # -----------------------------------------------------------------------
    app.tree_attive = app.create_request_treeview(app.tab_attive)
    app.tree_archiviate = app.create_request_treeview(app.tab_archiviate)

    # -----------------------------------------------------------------------
    # Bindings
    # -----------------------------------------------------------------------
    app.notebook.bind("<<NotebookTabChanged>>", app.on_tab_changed)
    app.root.bind("<Button-1>", app._on_root_click, add="+")
