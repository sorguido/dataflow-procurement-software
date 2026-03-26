"""
VSM Event Dialog - Dialog per creazione/modifica eventi VSM.
Form dinamico che mostra solo campi pertinenti al tipo evento selezionato.
"""

import tkinter as tk
from tkinter import ttk, messagebox
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
from utils.i18n_utils import _, get_current_language
from utils.resource_utils import set_window_icon
from utils.window_utils import center_window

# Import models
from models.vsm_event import VSMEvent

logger = logging.getLogger(__name__)


class VSMEventDialog(tk.Toplevel):
    """
    Dialog per creazione/modifica eventi VSM.
    Form dinamico con campi condizionali basati su event_type.
    """
    
    def __init__(self, parent, current_username, event_type="Saving", event_id=None):
        """
        Inizializza il dialog.
        
        Args:
            parent: Widget parent
            current_username: Username utente corrente
            event_type: Tipo evento di default ("Saving", "Cost Avoidance", "Derisking")
            event_id: Se None, modalità create; se int, modalità edit
        """
        super().__init__(parent)
        
        # Parametri
        self.current_username = current_username
        self.event_id = event_id
        self.is_edit_mode = event_id is not None
        self.result = None  # Usato per indicare successo salvataggio
        
        # Setup finestra
        self.withdraw()
        set_window_icon(self)
        self.title(_("Modifica Evento VSM") if self.is_edit_mode else _("Nuovo Evento VSM"))
        self.transient(parent)
        self.resizable(False, False)
        
        # Widget refs
        self.event_type_var = tk.StringVar(value=event_type)
        self.action_var = tk.StringVar(value="Negoziazione")
        self.opex_ripetitivo_var = tk.BooleanVar(value=False)
        self.driver_var = tk.StringVar(value="Prezzo")
        
        # Build UI
        self._build_ui()
        
        # Popola campo User con username corrente (sempre, sia in CREATE che in EDIT)
        self.entry_buyer.configure(state="normal")
        self.entry_buyer.insert(0, self.current_username)
        self.entry_buyer.configure(state="disabled")
        
        # Se modalità edit, carica dati
        if self.is_edit_mode:
            self._load_event_data()
        else:
            # Modalità create: applica dinamismo iniziale
            self._on_event_type_changed()
        
        # Finalize window
        center_window(self)
        self.wait_visibility()
        self.grab_set()
        self.deiconify()
    
    def _build_ui(self):
        """Costruisce l'interfaccia utente."""
        # Main container
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)
        
        # === SEZIONE GENERALE ===
        general_frame = ttk.LabelFrame(main_frame, text=_("Informazioni Generali"), padding="10")
        general_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        general_frame.columnconfigure(1, weight=1)
        
        # Data Evento
        ttk.Label(general_frame, text=_("Data Evento: *")).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_date = DateEntry(
            general_frame,
            width=15,
            date_pattern='dd/mm/yyyy',
            locale='it_IT' if get_current_language() == 'it' else 'en_US'
        )
        self.entry_date.grid(row=0, column=1, sticky="w", pady=5)
        
        # Tipo Evento (pre-condizionato dal tab attivo, non modificabile)
        ttk.Label(general_frame, text=_("Tipo Evento:")).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_event_type = ttk.Entry(
            general_frame,
            textvariable=self.event_type_var,
            state="disabled",
            width=22
        )
        self.entry_event_type.grid(row=1, column=1, sticky="w", pady=5)
        
        # Azione
        ttk.Label(general_frame, text=_("Azione: *")).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        self.combo_action = ttk.Combobox(
            general_frame,
            textvariable=self.action_var,
            values=["Negoziazione", "Derisking", "Altro"],
            state="readonly",
            width=20
        )
        self.combo_action.grid(row=2, column=1, sticky="w", pady=5)
        
        # User (auto-valorizzato con username corrente, non modificabile)
        ttk.Label(general_frame, text=_("User:")).grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
        self.entry_buyer = ttk.Entry(general_frame, width=30, state="disabled")
        self.entry_buyer.grid(row=3, column=1, sticky="ew", pady=5)
        
        # === SEZIONE DESCRIZIONE ===
        desc_frame = ttk.LabelFrame(main_frame, text=_("Descrizione"), padding="10")
        desc_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        desc_frame.columnconfigure(0, weight=1)
        desc_frame.rowconfigure(0, weight=1)
        
        # Campo descrizione (full width, espandibile)
        self.text_description = tk.Text(desc_frame, height=3, width=40, wrap="word")
        self.text_description.grid(row=0, column=0, sticky="nsew", pady=5)
        
        # === SEZIONE ECONOMICA (dinamica) ===
        self.economic_frame = ttk.LabelFrame(main_frame, text=_("Dati Economici"), padding="10")
        self.economic_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self.economic_frame.columnconfigure(1, weight=1)
        
        # Importo Budget (Saving only)
        self.lbl_importo_bdg = ttk.Label(self.economic_frame, text=_("Importo a Budget: *"))
        self.entry_importo_bdg = ttk.Entry(self.economic_frame, width=20)
        
        # Importo Richiesto Iniziale (Cost Avoidance only)
        self.lbl_importo_richiesto = ttk.Label(self.economic_frame, text=_("Importo Richiesto Iniziale: *"))
        self.entry_importo_richiesto = ttk.Entry(self.economic_frame, width=20)
        
        # Importo Negoziato (Saving + Cost Avoidance)
        self.lbl_importo_negoziato = ttk.Label(self.economic_frame, text=_("Importo Negoziato: *"))
        self.entry_importo_negoziato = ttk.Entry(self.economic_frame, width=20)
        
        # Percentuale Realizzo
        self.lbl_percent_realizzo = ttk.Label(self.economic_frame, text=_("% Realizzo:"))
        self.entry_percent_realizzo = ttk.Entry(self.economic_frame, width=10)
        self.entry_percent_realizzo.insert(0, "100")  # Default 100%
        
        # Driver (opzionale)
        self.lbl_driver = ttk.Label(self.economic_frame, text=_("Driver:"))
        self.combo_driver = ttk.Combobox(
            self.economic_frame,
            textvariable=self.driver_var,
            values=["Prezzo", "Pagamenti", "Volume", "Altro"],
            state="readonly",
            width=18
        )
        
        # === SEZIONE DISTRIBUZIONE ===
        dist_frame = ttk.LabelFrame(main_frame, text=_("Distribuzione Valore"), padding="10")
        dist_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        dist_frame.columnconfigure(1, weight=1)
        
        # OPEX Ripetitivo
        self.check_opex_ripetitivo = ttk.Checkbutton(
            dist_frame,
            text=_("OPEX Ripetitivo (distribuzione multi-mese)"),
            variable=self.opex_ripetitivo_var,
            command=self._on_opex_changed
        )
        self.check_opex_ripetitivo.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        
        # Info pro-rata (label informativo)
        info_label = ttk.Label(
            dist_frame,
            text=_("• Eventi ripetitivi: distribuzione fino a 24 mesi con pro-rata primo mese\n"
                   "• Eventi one-shot: impatto singolo nel mese evento"),
            foreground="gray",
            font=("Calibri", 9, "italic")
        )
        info_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 5))
        
        # === PULSANTI ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        
        ttk.Button(
            btn_frame,
            text=_("❌ Annulla"),
            command=self.destroy
        ).pack(side="right", padx=(5, 0))
        
        ttk.Button(
            btn_frame,
            text=_("💾 Salva"),
            command=self._validate_and_save
        ).pack(side="right")
    
    def _on_event_type_changed(self):
        """
        Handler per cambio event_type.
        Mostra/nasconde campi economici in base al tipo.
        """
        event_type = self.event_type_var.get()
        
        # Rimuovi tutti i campi economici
        for widget in self.economic_frame.winfo_children():
            widget.grid_forget()
        
        row = 0
        
        if event_type == "Saving":
            # Saving: importo_bdg + importo_negoziato
            self.lbl_importo_bdg.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_bdg.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_importo_negoziato.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_negoziato.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_percent_realizzo.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_percent_realizzo.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_driver.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.combo_driver.grid(row=row, column=1, sticky="w", pady=5)
        
        elif event_type == "Cost Avoidance":
            # Cost Avoidance: importo_richiesto_iniziale + importo_negoziato
            self.lbl_importo_richiesto.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_richiesto.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_importo_negoziato.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_importo_negoziato.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_percent_realizzo.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.entry_percent_realizzo.grid(row=row, column=1, sticky="w", pady=5)
            row += 1
            
            self.lbl_driver.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            self.combo_driver.grid(row=row, column=1, sticky="w", pady=5)
        
        elif event_type == "Derisking":
            # Derisking: nessun campo economico obbligatorio
            # Solo label informativo
            info_label = ttk.Label(
                self.economic_frame,
                text=_("Gli eventi Derisking non generano impatti economici.\n"
                       "Compilare solo sezioni descrittive."),
                foreground="blue",
                font=("Calibri", 9, "italic")
            )
            info_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
    
    def _on_opex_changed(self):
        """Handler per cambio checkbox OPEX ripetitivo (future enhancement)."""
        # Per ora non serve logica specifica
        # In futuro potremmo abilitare campo num_mesi_ripetizione
        pass
    
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
            self.action_var.set(event.action)
            # Campo User già popolato con current_username in __init__, non sovrascrivere
            
            self.text_description.insert("1.0", event.description)
            
            # Campi economici
            if event.importo_bdg:
                self.entry_importo_bdg.insert(0, str(event.importo_bdg))
            if event.importo_richiesto_iniziale:
                self.entry_importo_richiesto.insert(0, str(event.importo_richiesto_iniziale))
            if event.importo_negoziato:
                self.entry_importo_negoziato.insert(0, str(event.importo_negoziato))
            
            self.entry_percent_realizzo.delete(0, tk.END)
            self.entry_percent_realizzo.insert(0, str(event.percent_realizzo))
            
            if event.driver:
                self.driver_var.set(event.driver)
            
            self.opex_ripetitivo_var.set(event.opex_ripetitivo)
            
            # Applica dinamismo form
            self._on_event_type_changed()
        
        except Exception as e:
            logger.error(f"Errore caricamento evento {self.event_id}: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile caricare l'evento:\n{}").format(e),
                parent=self
            )
            self.destroy()
    
    def _validate_and_save(self):
        """Valida input e salva evento VSM."""
        try:
            # === VALIDAZIONE CAMPI OBBLIGATORI ===
            event_date = self.entry_date.get_date()
            if not event_date:
                raise ValueError(_("Data evento obbligatoria."))
            
            event_type = self.event_type_var.get()
            if not event_type:
                raise ValueError(_("Tipo evento obbligatorio."))
            
            action = self.action_var.get()
            if not action:
                raise ValueError(_("Azione obbligatoria."))
            
            buyer = self.entry_buyer.get().strip()
            if not buyer:
                raise ValueError(_("Buyer obbligatorio."))
            
            # Campi testuali
            description = self.text_description.get("1.0", tk.END).strip()
            
            # Reference: preserva valore storico in EDIT, vuoto in CREATE
            reference = self._loaded_event.reference if (self.is_edit_mode and hasattr(self, '_loaded_event')) else ""
            
            # Campi economici (condizionali)
            importo_bdg = 0.0
            importo_negoziato = 0.0
            importo_richiesto_iniziale = None
            percent_realizzo = 100.0
            driver = self.driver_var.get()
            
            if event_type == "Saving":
                # Saving richiede importo_bdg e importo_negoziato
                try:
                    importo_bdg = float(self.entry_importo_bdg.get().strip())
                except ValueError:
                    raise ValueError(_("Importo a Budget deve essere un numero valido."))
                
                try:
                    importo_negoziato = float(self.entry_importo_negoziato.get().strip())
                except ValueError:
                    raise ValueError(_("Importo Negoziato deve essere un numero valido."))
            
            elif event_type == "Cost Avoidance":
                # Cost Avoidance richiede importo_richiesto_iniziale e importo_negoziato
                try:
                    importo_richiesto_iniziale = float(self.entry_importo_richiesto.get().strip())
                except ValueError:
                    raise ValueError(_("Importo Richiesto Iniziale deve essere un numero valido."))
                
                try:
                    importo_negoziato = float(self.entry_importo_negoziato.get().strip())
                except ValueError:
                    raise ValueError(_("Importo Negoziato deve essere un numero valido."))
            
            # Percent realizzo (opzionale, default 100)
            try:
                percent_realizzo = float(self.entry_percent_realizzo.get().strip())
                if not (0 <= percent_realizzo <= 100):
                    raise ValueError(_("% Realizzo deve essere tra 0 e 100."))
            except ValueError as e:
                if "could not convert" in str(e):
                    raise ValueError(_("% Realizzo debe essere un numero valido."))
                raise
            
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
                importo_bdg=importo_bdg,
                importo_negoziato=importo_negoziato,
                importo_richiesto_iniziale=importo_richiesto_iniziale,
                percent_realizzo=percent_realizzo,
                driver=driver,
                opex_ripetitivo=opex_ripetitivo
            )
            
            # Salva tramite persistence layer
            with DatabaseManager(get_db_path()) as db_manager:
                if self.is_edit_mode:
                    update_event_with_impacts(db_manager, event)
                    msg = _("Evento VSM aggiornato con successo.")
                else:
                    event_id = save_event_with_impacts(db_manager, event)
                    msg = _("Evento VSM creato con successo (ID: {}).").format(event_id)
            
            messagebox.showinfo(_("Successo"), msg, parent=self)
            
            # Imposta result per indicare successo
            self.result = True
            self.destroy()
        
        except ValueError as e:
            messagebox.showwarning(_("Validazione"), str(e), parent=self)
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore salvataggio evento VSM: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile salvare l'evento:\n{}").format(e),
                parent=self
            )
        except Exception as e:
            logger.error(f"Errore inatteso salvataggio evento: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Errore inatteso:\n{}").format(e),
                parent=self
            )
