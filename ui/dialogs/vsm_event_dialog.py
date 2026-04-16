"""
VSM Event Dialog - Dialog per creazione/modifica eventi VSM.
Form dinamico che mostra solo campi pertinenti al tipo evento selezionato.
"""

import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import logging
from datetime import datetime

# Import da database e services
from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from services.vsm_persistence import (
    save_event_with_impacts,
    update_event_with_impacts,
    get_event_with_impacts,
    VSMError
)

# Import da utils
from utils.i18n_utils import tr, get_current_language
from utils.resource_utils import set_window_icon

# Import models
from models.vsm_event import VSMEvent

# Import custom dialog per messaggi uniformi
from ui.dialogs.common_dialogs import SimpleMessageDialog, SimpleYesNoDialog

# Import per validazione importi con formato virgola
from utils.format_utils import parse_float_from_comma_string, format_quantity_display, format_amount_display
from utils.vsm_config import get_pagamenti_coefficient, set_pagamenti_coefficient

logger = logging.getLogger(__name__)


class VSMEventDialog(tk.Toplevel):
    """
    Dialog per creazione/modifica eventi VSM.
    Form dinamico con campi condizionali basati su event_type.
    """
    
    def __init__(self, parent, current_username, event_type="Saving", event_id=None, read_only=False):
        """
        Inizializza il dialog.
        
        Args:
            parent: Widget parent
            current_username: Username utente corrente
            event_type: Tipo evento di default ("Saving", "Cost Avoidance", "Derisking")
            event_id: Se None, modalità create; se int, modalità edit
            read_only: Se True, apre in sola lettura (eventi di altri utenti)
        """
        super().__init__(parent)
        
        # Parametri
        self.current_username = current_username
        self.event_id = event_id
        self.is_edit_mode = event_id is not None
        self.read_only = read_only
        self.result = None  # Usato per indicare successo salvataggio
        
        # Setup finestra
        self.withdraw()
        set_window_icon(self)
        if read_only:
            self.title(tr("View VSM Event"))
        else:
            self.title(tr("Edit VSM Event") if self.is_edit_mode else tr("New VSM Event"))
        self.transient(parent)
        self.resizable(False, False)
        
        # Widget refs
        self.event_type_var = tk.StringVar(value=event_type)
        self.action_var = tk.StringVar()
        self.opex_ripetitivo_var = tk.BooleanVar(value=False)
        self.driver_var = tk.StringVar()
        
        # Build UI
        self._build_ui()
        
        # Imposta valori iniziali tradotti
        self._set_action_display("Negoziazione")
        self._set_driver_display("Prezzo")
        
        # Popola campo User con username corrente (sempre, sia in CREATE che in EDIT)
        self.entry_buyer.configure(state="normal")
        self.entry_buyer.insert(0, self.current_username)
        self.entry_buyer.configure(state="disabled")
        
        # Se modalità edit, carica dati
        if self.is_edit_mode:
            self._load_event_data()
            if self.read_only:
                self._apply_read_only()
        else:
            # Modalità create: applica dinamismo iniziale
            self._on_event_type_changed()
        
        # Finalize window geometry once (while still hidden) to avoid visible resize/flicker
        self.update_idletasks()
        base_width = self.winfo_reqwidth()
        base_height = self.winfo_reqheight()
        target_width = base_width * 2
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        target_x = max(0, (screen_w - target_width) // 2)
        target_y = max(0, (screen_h - base_height) // 2)
        self.geometry(f"{target_width}x{base_height}+{target_x}+{target_y}")
        
        self.grab_set()
        self.deiconify()
    
    # === HELPER METHODS PER CONVERSIONE VALORI TRADOTTI/INTERNI ===
    
    def _get_action_internal(self):
        """Converte il valore display di action al valore interno (italiano)."""
        display = self.action_var.get()
        # Reverse mapping: cerca la corrispondenza italiana
        if display == tr("Negotiation"):
            return "Negoziazione"
        elif display == "Derisking":
            return "Derisking"
        elif display == tr("Other"):
            return "Altro"
        return display  # Fallback per valori già in italiano
    
    def _set_action_display(self, internal_value):
        """Converte il valore interno (italiano) al valore display (tradotto)."""
        if internal_value == "Negoziazione":
            self.action_var.set(tr("Negotiation"))
        elif internal_value == "Derisking":
            self.action_var.set("Derisking")
        elif internal_value == "Altro":
            self.action_var.set(tr("Other"))
        else:
            self.action_var.set(internal_value)
    
    def _get_driver_internal(self):
        """Converte il valore display di driver al valore interno (italiano)."""
        display = self.driver_var.get()
        # Reverse mapping
        if display == tr("Price"):
            return "Prezzo"
        elif display == tr("Payments"):
            return "Pagamenti"
        return display  # Fallback
    
    def _set_driver_display(self, internal_value):
        """Converte il valore interno (italiano) al valore display (tradotto)."""
        if internal_value == "Prezzo":
            self.driver_var.set(tr("Price"))
        elif internal_value == "Pagamenti":
            self.driver_var.set(tr("Payments"))
        else:
            self.driver_var.set(internal_value)
    
    # === UI BUILDING ===
    
    def _build_ui(self):
        """Costruisce l'interfaccia utente."""
        # Main container
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        
        # === SEZIONE GENERALE ===
        general_frame = ttk.LabelFrame(main_frame, text=tr("General Information"), padding="10")
        general_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        general_frame.columnconfigure(1, weight=1)
        
        # Data Evento
        ttk.Label(general_frame, text=tr("Event Date: *")).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_date = DateEntry(
            general_frame,
            width=15,
            date_pattern='dd/mm/yyyy',
            locale='it_IT' if get_current_language() == 'it' else 'en_US'
        )
        self.entry_date.grid(row=0, column=1, sticky="w", pady=5)
        
        # Tipo Evento (pre-condizionato dal tab attivo, non modificabile)
        ttk.Label(general_frame, text=tr("Event Type:")).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_event_type = ttk.Entry(
            general_frame,
            textvariable=self.event_type_var,
            state="disabled",
            width=22
        )
        self.entry_event_type.grid(row=1, column=1, sticky="ew", pady=5)
        
        # Azione
        ttk.Label(general_frame, text=tr("Action:")).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        self.combo_action = ttk.Combobox(
            general_frame,
            textvariable=self.action_var,
            values=[tr("Negotiation"), "Derisking", tr("Other")],
            state="readonly",
            width=20
        )
        self.combo_action.grid(row=2, column=1, sticky="ew", pady=5)
        # Entry per Azione (stesso sistema di Tipo Evento, usato in modalità Derisking)
        self.entry_action = ttk.Entry(
            general_frame,
            textvariable=self.action_var,
            state="disabled",
            width=22
        )
        
        # User (auto-valorizzato con username corrente, non modificabile)
        ttk.Label(general_frame, text=tr("User:")).grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_buyer = ttk.Entry(general_frame, width=30, state="disabled")
        self.entry_buyer.grid(row=3, column=1, sticky="ew", pady=5)
        
        # === SEZIONE NEW SUPPLIER INTRODUCED (solo Derisking) ===
        self.new_supplier_frame = ttk.LabelFrame(main_frame, text=tr("New Supplier Introduced"), padding="10")
        self.new_supplier_frame.columnconfigure(0, weight=1)
        self.entry_new_supplier = ttk.Entry(self.new_supplier_frame, width=40)
        self.entry_new_supplier.grid(row=0, column=0, sticky="ew", pady=5)
        self.entry_new_supplier.bind("<Return>", self._on_new_supplier_enter)
        # (not gridded here — shown dynamically by _on_event_type_changed)

        # === SEZIONE DESCRIZIONE ===
        desc_frame = ttk.LabelFrame(main_frame, text=tr("Description"), padding="10")
        desc_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        desc_frame.columnconfigure(0, weight=1)
        desc_frame.rowconfigure(0, weight=1)
        
        # Campo descrizione (full width, espandibile)
        self.text_description = tk.Text(desc_frame, height=3, width=40, wrap="word")
        self.text_description.grid(row=0, column=0, sticky="nsew", pady=5)
        
        # === SEZIONE ECONOMICA (dinamica) ===
        self.economic_frame = ttk.LabelFrame(main_frame, text=tr("Economic Data"), padding="10")
        self.economic_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self.economic_frame.columnconfigure(0, weight=1)
        self.economic_frame.columnconfigure(1, weight=1)
        
        # --- SUB-FRAME PREZZO (driver Prezzo) ---
        self.price_fields_frame = ttk.Frame(self.economic_frame)
        self.price_fields_frame.columnconfigure(1, weight=1)
        
        # Importo Budget (Saving only)
        self.lbl_importo_bdg = ttk.Label(self.price_fields_frame, text=tr("Budget Amount: *"))
        self.entry_importo_bdg = ttk.Entry(self.price_fields_frame, width=20)
        
        # Importo Richiesto Iniziale (Cost Avoidance only)
        self.lbl_importo_richiesto = ttk.Label(self.price_fields_frame, text=tr("Initial Requested Amount: *"))
        self.entry_importo_richiesto = ttk.Entry(self.price_fields_frame, width=20)
        
        # Importo Negoziato (Saving + Cost Avoidance)
        self.lbl_importo_negoziato = ttk.Label(self.price_fields_frame, text=tr("Negotiated Amount: *"))
        self.entry_importo_negoziato = ttk.Entry(self.price_fields_frame, width=20)
        
        # Annual Q.ty
        self.lbl_quantita_annua = ttk.Label(self.price_fields_frame, text=tr("Annual Q.ty:"))
        self.entry_quantita_annua = ttk.Entry(self.price_fields_frame, width=20)
        self.entry_quantita_annua.insert(0, "1")  # Default 1
        
        # Percentuale Realizzo
        self.lbl_percent_realizzo = ttk.Label(self.price_fields_frame, text=tr("% Realization:"))
        self.entry_percent_realizzo = ttk.Entry(self.price_fields_frame, width=10)
        self.entry_percent_realizzo.insert(0, "100")  # Default 100%
        
        # --- SUB-FRAME PAGAMENTI (driver Pagamenti) ---
        self.payment_fields_frame = ttk.Frame(self.economic_frame)
        self.payment_fields_frame.columnconfigure(1, weight=1)
        
        # Spending Annuo
        self.lbl_spending_annuo = ttk.Label(self.payment_fields_frame, text=tr("Annual Spending (€): *"))
        self.entry_spending_annuo = ttk.Entry(self.payment_fields_frame, width=10)
        
        # Termini Pagamento Attuali
        self.lbl_giorni_attuali = ttk.Label(self.payment_fields_frame, text=tr("Current Payment Terms (days): *"))
        self.entry_giorni_attuali = ttk.Entry(self.payment_fields_frame, width=8)
        
        # Termini Pagamento Negoziati
        self.lbl_giorni_negoziati = ttk.Label(self.payment_fields_frame, text=tr("Negotiated Payment Terms (days): *"))
        self.entry_giorni_negoziati = ttk.Entry(self.payment_fields_frame, width=8)
        
        # Financial Impact rate (% per 30 days)
        self.lbl_payments_rate = ttk.Label(self.payment_fields_frame, text=tr("Financial Impact (% per 30 days):"))
        self.entry_payments_rate = ttk.Entry(self.payment_fields_frame, width=8)
        # Valore iniziale: ultima preferenza utente salvata (default 0.50)
        self.entry_payments_rate.insert(0, format_amount_display(get_pagamenti_coefficient() * 100))
        
        # --- DRIVER (always visible, outside sub-frames) ---
        self.lbl_driver = ttk.Label(self.economic_frame, text=tr("Driver:"))
        self.combo_driver = ttk.Combobox(
            self.economic_frame,
            textvariable=self.driver_var,
            values=[tr("Price"), tr("Payments")],
            state="readonly",
            width=18
        )
        
        # --- DERISKING INFO LABEL (created once, shown conditionally) ---
        self.lbl_derisking_info = ttk.Label(
            self.economic_frame,
            text=tr("Derisking events do not generate economic impacts.\nFill only descriptive sections."),
            foreground="blue",
            font=("Calibri", 9, "italic")
        )
        
        # Binding su combobox driver per show/hide dinamico
        self.combo_driver.bind("<<ComboboxSelected>>", self._on_driver_changed)
        
        # === SEZIONE DISTRIBUZIONE ===
        dist_frame = ttk.LabelFrame(main_frame, text=tr("Value Distribution"), padding="10")
        dist_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        dist_frame.columnconfigure(1, weight=1)
        
        # OPEX Ripetitivo
        self.check_opex_ripetitivo = ttk.Checkbutton(
            dist_frame,
            text=tr("Repetitive OPEX (multi-month distribution)"),
            variable=self.opex_ripetitivo_var,
            command=self._on_opex_changed
        )
        self.check_opex_ripetitivo.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        
        # Info pro-rata (label informativo)
        info_label = ttk.Label(
            dist_frame,
            text=tr("• Repetitive events: distribution up to 24 months with first month pro-rata\n• One-shot events: single impact in event month"),
            foreground="gray",
            font=("Calibri", 9, "italic")
        )
        info_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 5))
        
        # === PULSANTI ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        
        self.btn_cancel = ttk.Button(
            btn_frame,
            text=tr("❌ Cancel"),
            command=self.destroy
        )
        self.btn_cancel.pack(side="right", padx=(5, 0))
        
        self.btn_save = ttk.Button(
            btn_frame,
            text=tr("💾 Save"),
            command=self._validate_and_save
        )
        self.btn_save.pack(side="right")

    # === MODALITÀ SOLA LETTURA ===

    def _apply_read_only(self):
        """Attiva la modalità sola lettura: disabilita tutti i campi, nasconde Salva."""
        self.btn_save.pack_forget()
        self._disable_widgets_recursive(self)
        # Riabilita il bottone Chiudi (disabilitato dal loop ricorsivo)
        self.btn_cancel.configure(state="normal", text=tr("✖ Close"))

    def _disable_widgets_recursive(self, widget):
        """Disabilita ricorsivamente tutti i widget interattivi nel dialog."""
        for child in widget.winfo_children():
            try:
                child.configure(state="disabled")
            except tk.TclError:
                pass
            self._disable_widgets_recursive(child)

    def _on_event_type_changed(self):
        """
        Handler per cambio event_type.
        Mostra/nasconde campi economici in base al tipo.
        """
        event_type = self.event_type_var.get()
        
        # Hide all economic sub-frames and widgets explicitly
        self.price_fields_frame.grid_remove()
        self.payment_fields_frame.grid_remove()
        self.lbl_driver.grid_remove()
        self.combo_driver.grid_remove()
        self.lbl_derisking_info.grid_remove()
        self.new_supplier_frame.grid_remove()
        
        if event_type == "Saving":
            # Saving: show driver combo (row 0) and delegate field layout to _on_driver_changed
            self.lbl_driver.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.combo_driver.grid(row=0, column=1, sticky="ew", pady=5)
            # Call driver handler to position appropriate sub-frame at row 1
            self._on_driver_changed()
            # Mostra Combobox Azione, nascondi Entry
            self.combo_action.grid()
            self.entry_action.grid_remove()
        
        elif event_type == "Cost Avoidance":
            # Cost Avoidance: show driver combo (row 0) and delegate field layout to _on_driver_changed
            self.lbl_driver.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.combo_driver.grid(row=0, column=1, sticky="ew", pady=5)
            # Call driver handler to position appropriate sub-frame at row 1
            self._on_driver_changed()
            # Mostra Combobox Azione, nascondi Entry
            self.combo_action.grid()
            self.entry_action.grid_remove()
        
        elif event_type == "Derisking":
            # Derisking: show info label only
            self.lbl_derisking_info.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            # Mostra Entry Azione (stesso sistema di Tipo Evento), nascondi Combobox
            self.action_var.set("Derisking")
            self.combo_action.grid_remove()
            self.entry_action.grid(row=2, column=1, sticky="ew", pady=5)
            # Mostra sezione New Supplier Introduced (solo per Derisking)
            self.new_supplier_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
    
    def _on_opex_changed(self):
        """Handler per cambio checkbox OPEX ripetitivo (future enhancement)."""
        # Per ora non serve logica specifica
        # In futuro potremmo abilitare campo num_mesi_ripetizione
        pass
    
    def _on_driver_changed(self, event=None):
        """
        Handler per cambio driver.
        Mostra/nasconde SUB-FRAME in base al driver selezionato.
        """
        driver_internal = self._get_driver_internal()
        event_type = self.event_type_var.get()
        
        # Solo per event_type con campi economici
        if event_type not in ["Saving", "Cost Avoidance"]:
            return
        
        # Hide both sub-frames first
        self.price_fields_frame.grid_remove()
        self.payment_fields_frame.grid_remove()
        
        if driver_internal == "Prezzo":
            # Show price fields frame at row 1
            self.price_fields_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            
            # Layout widgets inside price_fields_frame based on event_type
            if event_type == "Saving":
                self.lbl_importo_bdg.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_importo_bdg.grid(row=0, column=1, sticky="ew", pady=5)
                
                self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_importo_negoziato.grid(row=1, column=1, sticky="ew", pady=5)
                
                self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_quantita_annua.grid(row=2, column=1, sticky="ew", pady=5)
                
                self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_percent_realizzo.grid(row=3, column=1, sticky="ew", pady=5)
                
                # Hide Cost Avoidance specific field
                self.lbl_importo_richiesto.grid_remove()
                self.entry_importo_richiesto.grid_remove()
                
            elif event_type == "Cost Avoidance":
                self.lbl_importo_richiesto.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_importo_richiesto.grid(row=0, column=1, sticky="ew", pady=5)
                
                self.lbl_importo_negoziato.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_importo_negoziato.grid(row=1, column=1, sticky="ew", pady=5)
                
                self.lbl_quantita_annua.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_quantita_annua.grid(row=2, column=1, sticky="ew", pady=5)
                
                self.lbl_percent_realizzo.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
                self.entry_percent_realizzo.grid(row=3, column=1, sticky="ew", pady=5)
                
                # Hide Saving specific field
                self.lbl_importo_bdg.grid_remove()
                self.entry_importo_bdg.grid_remove()
            
        elif driver_internal == "Pagamenti":
            # Show payment fields frame at row 1 (SAME position as price frame)
            self.payment_fields_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            
            # Layout widgets inside payment_fields_frame
            self.lbl_spending_annuo.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_spending_annuo.grid(row=0, column=1, sticky="ew", pady=5)
            
            self.lbl_giorni_attuali.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_giorni_attuali.grid(row=1, column=1, sticky="ew", pady=5)
            
            self.lbl_giorni_negoziati.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_giorni_negoziati.grid(row=2, column=1, sticky="ew", pady=5)
            
            self.lbl_payments_rate.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_payments_rate.grid(row=3, column=1, sticky="ew", pady=5)
    
    def _on_new_supplier_enter(self, event=None):
        """Salvataggio immediato del campo new_supplier su Enter (solo edit mode con evento già persistito)."""
        if not self.is_edit_mode or self.event_id is None:
            return
        value = self.entry_new_supplier.get().strip()
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                db_manager.update_vsm_event_new_supplier(self.event_id, value)
        except Exception as e:
            logger.error(f"Errore aggiornamento new_supplier: {e}", exc_info=True)

    def _load_event_data(self):
        """Carica dati evento esistente (modalità edit)."""
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                event, _impacts = get_event_with_impacts(db_manager, self.event_id)
            
            # Salva evento caricato per preservare campi non mostrati in UI
            self._loaded_event = event
            
            # Popola form
            if event.event_date:
                self.entry_date.set_date(event.event_date)
            
            self.event_type_var.set(event.event_type)
            
            # Gestione campo Azione: forza coerenza per eventi Derisking
            if event.event_type == "Derisking":
                # Se evento Derisking, forza action a "Derisking" indipendentemente dal valore DB
                self.action_var.set("Derisking")
            else:
                self._set_action_display(event.action)
            
            # Campo User già popolato con current_username in __init__, non sovrascrivere
            
            self.text_description.insert("1.0", event.description)
            
            # Campi economici
            # Usa format_amount_display per mostrare sempre 2 decimali con virgola
            if event.importo_bdg:
                self.entry_importo_bdg.insert(0, format_amount_display(event.importo_bdg))
            if event.importo_richiesto_iniziale:
                self.entry_importo_richiesto.insert(0, format_amount_display(event.importo_richiesto_iniziale))
            if event.importo_negoziato:
                self.entry_importo_negoziato.insert(0, format_amount_display(event.importo_negoziato))
            
            # Annual Q.ty: default to 1 if 0 or None for backward compatibility
            quantita_display = event.quantita_annua if event.quantita_annua and event.quantita_annua > 0 else 1.0
            self.entry_quantita_annua.delete(0, tk.END)
            self.entry_quantita_annua.insert(0, format_amount_display(quantita_display))
            
            self.entry_percent_realizzo.delete(0, tk.END)
            self.entry_percent_realizzo.insert(0, format_amount_display(event.percent_realizzo))
            
            if event.driver:
                # Handle legacy drivers: convert Volume/Altro to Prezzo for safety
                if event.driver in ["Volume", "Altro"]:
                    logger.warning(f"Legacy driver '{event.driver}' found for event {self.event_id}, defaulting to 'Prezzo'")
                    self._set_driver_display("Prezzo")
                else:
                    self._set_driver_display(event.driver)
            
            # Campi Pagamenti
            if event.spending_annuo:
                self.entry_spending_annuo.insert(0, format_amount_display(event.spending_annuo))
            if event.giorni_pagamento_attuali is not None:
                self.entry_giorni_attuali.insert(0, str(event.giorni_pagamento_attuali))
            if event.giorni_pagamento_negoziati is not None:
                self.entry_giorni_negoziati.insert(0, str(event.giorni_pagamento_negoziati))
            
            # payments_rate: usa valore storico dell'evento; se assente usa preferenza config
            self.entry_payments_rate.delete(0, tk.END)
            if event.payments_rate is not None:
                self.entry_payments_rate.insert(0, format_amount_display(event.payments_rate))
            else:
                self.entry_payments_rate.insert(0, format_amount_display(get_pagamenti_coefficient() * 100))
            
            self.opex_ripetitivo_var.set(event.opex_ripetitivo)
            
            # Campo new_supplier (Derisking)
            self.entry_new_supplier.delete(0, tk.END)
            self.entry_new_supplier.insert(0, event.new_supplier or "")
            
            # Applica dinamismo form (event_type e driver)
            self._on_event_type_changed()
            self._on_driver_changed()
        
        except Exception as e:
            logger.error(f"Errore caricamento evento {self.event_id}: {e}", exc_info=True)
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Unable to load event:\n{}").format(e),
                "error"
            )
            self.destroy()
    
    def _validate_and_save(self):
        """Valida input e salva evento VSM."""
        try:
            # === VALIDAZIONE CAMPI OBBLIGATORI ===
            event_date = self.entry_date.get_date()
            if not event_date:
                raise ValueError(tr("Event date required."))
            
            event_type = self.event_type_var.get()
            if not event_type:
                raise ValueError(tr("Event type required."))
            
            action = self._get_action_internal()
            if not action:
                raise ValueError(tr("Action required."))
            
            buyer = self.entry_buyer.get().strip()
            if not buyer:
                raise ValueError(tr("Buyer required."))
            
            # Campi testuali
            description = self.text_description.get("1.0", tk.END).strip()
            
            # Reference: preserva valore storico in EDIT, vuoto in CREATE
            reference = self._loaded_event.reference if (self.is_edit_mode and hasattr(self, '_loaded_event')) else ""
            
            # Campi economici (condizionali in base a event_type e driver)
            importo_bdg = None
            importo_negoziato = None
            importo_richiesto_iniziale = None
            percent_realizzo = 100.0
            spending_annuo = None
            giorni_pagamento_attuali = None
            giorni_pagamento_negoziati = None
            payments_rate = None
            driver = self._get_driver_internal()
            
            if event_type == "Saving":
                # Validazioni specifiche per driver
                if driver == "Prezzo":
                    # Driver Prezzo: valida importi (solo numeri con virgola, es. 1,23)
                    try:
                        importo_bdg = parse_float_from_comma_string(self.entry_importo_bdg.get().strip())
                    except ValueError:
                        raise ValueError(tr("Budget Amount not valid. Use only numbers with comma (e.g. 1,23)."))
                    
                    try:
                        importo_negoziato = parse_float_from_comma_string(self.entry_importo_negoziato.get().strip())
                    except ValueError:
                        raise ValueError(tr("Negotiated Amount not valid. Use only numbers with comma (e.g. 1,23)."))
                    
                    try:
                        quantita_annua = parse_float_from_comma_string(self.entry_quantita_annua.get().strip())
                    except ValueError:
                        raise ValueError(tr("Annual Q.ty not valid. Use only numbers with comma (e.g. 1,5)."))
                    if quantita_annua <= 0:
                        raise ValueError(tr("Annual Q.ty must be greater than zero."))
                    
                    try:
                        percent_realizzo = parse_float_from_comma_string(self.entry_percent_realizzo.get().strip())
                    except ValueError:
                        raise ValueError(tr("% Realization not valid. Use only numbers with comma (e.g. 99,5)."))
                    if not (0 <= percent_realizzo <= 100):
                        raise ValueError(tr("% Realization must be between 0 and 100."))
                    
                    # Campi Pagamenti a NULL
                    spending_annuo = None
                    giorni_pagamento_attuali = None
                    giorni_pagamento_negoziati = None
                    
                elif driver == "Pagamenti":
                    # Driver Pagamenti: valida spending e giorni
                    try:
                        spending_annuo = parse_float_from_comma_string(self.entry_spending_annuo.get().strip())
                    except ValueError:
                        raise ValueError(tr("Annual Spending not valid. Use only numbers with comma (e.g. 100000,00)."))
                    if spending_annuo <= 0:
                        raise ValueError(tr("Annual Spending must be positive."))
                    
                    try:
                        giorni_pagamento_attuali = int(self.entry_giorni_attuali.get().strip())
                        if giorni_pagamento_attuali < 0:
                            raise ValueError(tr("Current Payment Terms cannot be negative."))
                    except ValueError as e:
                        if "invalid literal" in str(e):
                            raise ValueError(tr("Current Payment Terms must be a valid integer."))
                        raise
                    
                    try:
                        giorni_pagamento_negoziati = int(self.entry_giorni_negoziati.get().strip())
                        if giorni_pagamento_negoziati < 0:
                            raise ValueError(tr("Negotiated Payment Terms cannot be negative."))
                    except ValueError as e:
                        if "invalid literal" in str(e):
                            raise ValueError(tr("Negotiated Payment Terms must be a valid integer."))
                        raise
                    
                    # Warning opzionale se delta negativo (peggioramento)
                    if giorni_pagamento_negoziati < giorni_pagamento_attuali:
                        risposta = SimpleYesNoDialog(
                            self,
                            tr("Warning"),
                            tr("Negotiated terms ({}) lower than current terms ({}).\nThis generates a negative impact (worsening of payment terms).\n\nDo you confirm you want to proceed?").format(
                                  giorni_pagamento_negoziati,
                                  giorni_pagamento_attuali
                              )
                        ).result
                        if not risposta:
                            return
                    
                    # payments_rate: valida e parse (stessa policy degli altri campi numerici)
                    try:
                        payments_rate = parse_float_from_comma_string(self.entry_payments_rate.get().strip())
                    except ValueError:
                        raise ValueError(tr("Financial Impact not valid. Use only numbers with comma (e.g. 0,50)."))
                    if payments_rate <= 0:
                        raise ValueError(tr("Financial Impact must be greater than zero."))
                    
                    # Campi Prezzo a NULL, percent_realizzo fisso a 100 (tecnico, ignorato)
                    importo_bdg = None
                    importo_negoziato = None
                    percent_realizzo = 100.0
                
                else:
                    # Driver non supportato: non dovrebbe mai accadere con combobox readonly
                    raise ValueError(tr(f"Driver '{driver}' non supportato per eventi Saving."))
            
            elif event_type == "Cost Avoidance":
                # Cost Avoidance richiede importo_richiesto_iniziale e importo_negoziato
                try:
                    importo_richiesto_iniziale = parse_float_from_comma_string(self.entry_importo_richiesto.get().strip())
                except ValueError:
                    raise ValueError(tr("Initial Requested Amount not valid. Use only numbers with comma (e.g. 1,23)."))
                
                try:
                    importo_negoziato = parse_float_from_comma_string(self.entry_importo_negoziato.get().strip())
                except ValueError:
                    raise ValueError(tr("Negotiated Amount not valid. Use only numbers with comma (e.g. 1,23)."))
                
                # Annual Q.ty for driver Prezzo
                if driver == "Prezzo":
                    try:
                        quantita_annua = parse_float_from_comma_string(self.entry_quantita_annua.get().strip())
                    except ValueError:
                        raise ValueError(tr("Annual Q.ty not valid. Use only numbers with comma (e.g. 1,5)."))
                    if quantita_annua <= 0:
                        raise ValueError(tr("Annual Q.ty must be greater than zero."))
                else:
                    quantita_annua = 0.0  # Not used for non-Prezzo drivers
                
                # Cost Avoidance sempre con percent_realizzo
                try:
                    percent_realizzo = parse_float_from_comma_string(self.entry_percent_realizzo.get().strip())
                except ValueError:
                    raise ValueError(tr("% Realization not valid. Use only numbers with comma (e.g. 99,5)."))
                if not (0 <= percent_realizzo <= 100):
                    raise ValueError(tr("% Realization must be between 0 and 100."))
            
            # Flags
            opex_ripetitivo = self.opex_ripetitivo_var.get()
            
            # === CREA/AGGIORNA EVENTO ===
            event = VSMEvent(
                id=self.event_id,  # None per create, int per edit
                event_date=datetime.combine(event_date, datetime.min.time()),
                username=self.current_username,
                buyer=buyer,
                event_type=event_type,
                action=action,
                description=description,
                reference=reference,
                importo_bdg=importo_bdg if importo_bdg is not None else 0.0,
                importo_negoziato=importo_negoziato if importo_negoziato is not None else 0.0,
                importo_richiesto_iniziale=importo_richiesto_iniziale,
                percent_realizzo=percent_realizzo,
                driver=driver,
                spending_annuo=spending_annuo if spending_annuo is not None else 0.0,
                giorni_pagamento_attuali=giorni_pagamento_attuali,
                giorni_pagamento_negoziati=giorni_pagamento_negoziati,
                quantita_annua=quantita_annua if 'quantita_annua' in locals() else 0.0,
                payments_rate=payments_rate,
                opex_ripetitivo=opex_ripetitivo,
                new_supplier=self.entry_new_supplier.get().strip() if event_type == "Derisking" else ""
            )
            
            # Salva tramite persistence layer
            with DatabaseManager(get_db_path()) as db_manager:
                if self.is_edit_mode:
                    update_event_with_impacts(db_manager, event)
                    msg = tr("VSM Event updated successfully.")
                else:
                    event_id = save_event_with_impacts(db_manager, event)
                    msg = tr("VSM Event created successfully.")
                    logger.info(f"Evento VSM creato con ID: {event_id}")
            
            # Aggiorna preferenza utente se il driver è Pagamenti
            if driver == "Pagamenti" and payments_rate is not None:
                set_pagamenti_coefficient(payments_rate / 100.0)
            
            # Usa dialog custom con font uniforme
            SimpleMessageDialog(self, tr("Success"), msg, "info")
            
            # Imposta result per indicare successo
            self.result = True
            self.destroy()
        
        except ValueError as e:
            SimpleMessageDialog(self, tr("Validation"), str(e), "warning")
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore salvataggio evento VSM: {e}", exc_info=True)
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Unable to save event:\n{}").format(e),
                "error"
            )
        except Exception as e:
            logger.error(f"Errore inatteso salvataggio evento: {e}", exc_info=True)
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Unexpected error:\n{}").format(e),
                "error"
            )
