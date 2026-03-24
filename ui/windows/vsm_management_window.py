"""
VSM Management Window - Finestra di gestione eventi VSM.
Contiene nested notebook con le 3 tab operative: Saving, Cost Avoidance, Derisking.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from tksheet import Sheet
import logging
from datetime import datetime

# Import da database e services
from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from services.vsm_persistence import (
    delete_event_and_impacts,
    VSMError
)

# Import da utils
from utils.i18n_utils import _

# Import models
from models.vsm_event import VSMEvent

logger = logging.getLogger(__name__)


class VSMManagementWindow(ttk.Frame):
    """
    Contenitore principale per il modulo VSM.
    Embedded come tab nella Main Dashboard.
    """
    
    def __init__(self, parent, current_username, refresh_callback=None):
        """
        Inizializza la finestra di gestione VSM.
        
        Args:
            parent: Widget parent (notebook principale)
            current_username: Username dell'utente corrente per ownership validation
            refresh_callback: Callback opzionale per refresh dashboard principale
        """
        super().__init__(parent)
        
        self.current_username = current_username
        self.refresh_callback = refresh_callback
        self.current_event_type = "Saving"  # Default event type per tab attiva
        
        # Riferimenti ai sheet per ogni tab
        self.sheets = {}
        
        # Build UI
        self._build_ui()
        
        # Initial load
        self.refresh_events()
    
    def _build_ui(self):
        """Costruisce l'interfaccia utente."""
        # Toolbar in alto
        toolbar = ttk.Frame(self)
        toolbar.pack(side="top", fill="x", padx=10, pady=5)
        
        ttk.Button(
            toolbar,
            text=_("➕ Nuovo Evento"),
            command=self.on_new_event
        ).pack(side="left", padx=5)
        
        ttk.Button(
            toolbar,
            text=_("✏️ Modifica"),
            command=self.on_edit_event
        ).pack(side="left", padx=5)
        
        ttk.Button(
            toolbar,
            text=_("🗑 Elimina"),
            command=self.on_delete_event
        ).pack(side="left", padx=5)
        
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        
        ttk.Button(
            toolbar,
            text=_("📊 KPI"),
            command=self.on_kpi_click
        ).pack(side="left", padx=5)
        
        ttk.Button(
            toolbar,
            text=_("🔄 Aggiorna"),
            command=self.refresh_events
        ).pack(side="left", padx=5)
        
        # Notebook con 3 sub-tab per tipo evento
        self.vsm_notebook = ttk.Notebook(self)
        self.vsm_notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tab Saving
        self.tab_saving = ttk.Frame(self.vsm_notebook)
        self.vsm_notebook.add(self.tab_saving, text=_("Saving"))
        self.sheets["Saving"] = self._create_event_sheet(self.tab_saving)
        
        # Tab Cost Avoidance
        self.tab_cost_avoidance = ttk.Frame(self.vsm_notebook)
        self.vsm_notebook.add(self.tab_cost_avoidance, text=_("Cost Avoidance"))
        self.sheets["Cost Avoidance"] = self._create_event_sheet(self.tab_cost_avoidance)
        
        # Tab Derisking
        self.tab_derisking = ttk.Frame(self.vsm_notebook)
        self.vsm_notebook.add(self.tab_derisking, text=_("Derisking"))
        self.sheets["Derisking"] = self._create_event_sheet(self.tab_derisking)
        
        # Bind tab change per refresh dinamico
        self.vsm_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
    
    def _create_event_sheet(self, parent):
        """
        Crea un tksheet per visualizzare eventi VSM.
        Pattern identico a create_request_treeview da dataflow.py.
        
        Args:
            parent: Widget parent
        
        Returns:
            Sheet: Widget tksheet configurato
        """
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        
        # Crea widget tksheet
        sheet = Sheet(
            frame,
            theme="light blue",
            header_font=("Calibri", 11, "bold"),
            font=("Calibri", 11, "normal"),
            headers=[
                _("Data"),
                _("Tipo"),
                _("Azione"),
                _("Descrizione"),
                _("Valore Teorico"),
                _("Realizzo %"),
                _("Ripetitivo"),
                _("Utente")
            ],
            show_header=True,
            show_row_index=False
        )
        
        # Configura larghezze colonne
        sheet.set_column_widths([100, 120, 120, 300, 120, 90, 90, 140])
        
        # Centra colonne numeriche e date
        sheet.align_columns(columns=[0, 1, 2, 4, 5, 6, 7], align="center")
        
        # Abilita bindings
        sheet.enable_bindings()
        
        # Rendi readonly
        for col_idx in range(8):
            sheet.readonly_columns(columns=[col_idx], readonly=True)
        
        # Binding per doppio click (apertura edit)
        sheet.bind("<Double-Button-1>", lambda event: self._on_sheet_double_click(sheet, event))
        
        # Selection handlers
        sheet.extra_bindings("cell_select", lambda event_data: self._update_buttons_state())
        sheet.extra_bindings("row_select", lambda event_data: self._update_buttons_state())
        
        sheet.pack(fill="both", expand=True)
        
        # Metadata storage
        sheet._event_metadata = []  # Lista di dict con event_id, username, is_mine
        
        return sheet
    
    def _on_tab_changed(self, event=None):
        """Handler per cambio tab VSM."""
        # Ottieni nome tab corrente
        current_tab_idx = self.vsm_notebook.index(self.vsm_notebook.select())
        tab_names = ["Saving", "Cost Avoidance", "Derisking"]
        self.current_event_type = tab_names[current_tab_idx]
        
        # Refresh eventi per tab corrente
        self.refresh_events()
    
    def _on_sheet_double_click(self, sheet, event):
        """Handler per doppio click su sheet (apre edit)."""
        self.on_edit_event()
    
    def _update_buttons_state(self):
        """Aggiorna stato pulsanti in base alla selezione (future enhancement)."""
        # Per ora non implementato, ma lascia hook per future
        pass
    
    def refresh_events(self):
        """
        Ricarica gli eventi VSM per la tab corrente.
        Filtra per event_type corrispondente alla tab attiva.
        """
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                # Carica tutti eventi per utente corrente
                all_events = db_manager.get_all_vsm_events(username=self.current_username)
            
            # Filtra per event_type della tab corrente
            filtered_events = [e for e in all_events if e.event_type == self.current_event_type]
            
            # Ottieni sheet corrente
            sheet = self.sheets[self.current_event_type]
            
            # Popola sheet
            self._populate_sheet(sheet, filtered_events)
            
            logger.debug(f"Caricati {len(filtered_events)} eventi {self.current_event_type}")
            
        except DatabaseError as e:
            logger.error(f"Errore caricamento eventi VSM: {e}")
            messagebox.showerror(
                _("Errore Database"),
                _("Impossibile caricare gli eventi VSM: {}\n").format(e),
                parent=self
            )
    
    def _populate_sheet(self, sheet, events):
        """
        Popola un tksheet con lista di eventi VSM.
        
        Args:
            sheet: Widget tksheet da popolare
            events: Lista di VSMEvent
        """
        data_rows = []
        metadata = []
        
        for event in events:
            # Calcola valore teorico
            valore_teorico = event.calculate_theoretical_value()
            
            # Formatta row
            row = [
                event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
                event.event_type,
                event.action,
                (event.description or event.reference or "")[:50],  # Truncate
                f"€ {valore_teorico:,.2f}",
                f"{event.percent_realizzo:.0f}%",
                "✓" if event.opex_ripetitivo else "",
                event.username
            ]
            data_rows.append(row)
            
            # Metadata per ownership e event_id
            metadata.append({
                'event_id': event.id,
                'username': event.username,
                'is_mine': event.username == self.current_username
            })
        
        # Aggiorna sheet
        sheet.set_sheet_data(data_rows)
        sheet._event_metadata = metadata
    
    def on_new_event(self):
        """Handler per creazione nuovo evento VSM."""
        # Import lazy per evitare circular dependencies
        from ui.dialogs.vsm_event_dialog import VSMEventDialog
        
        try:
            # Apri dialog con event_type preimpostato dalla tab corrente
            dialog = VSMEventDialog(
                self,
                current_username=self.current_username,
                event_type=self.current_event_type,
                event_id=None  # None = modalità create
            )
            self.wait_window(dialog)
            
            # Se dialog ha salvato con successo, refresh
            if hasattr(dialog, 'result') and dialog.result:
                self.refresh_events()
                if self.refresh_callback:
                    self.refresh_callback()
        
        except Exception as e:
            logger.error(f"Errore apertura dialog nuovo evento: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile aprire il form: {}").format(e),
                parent=self
            )
    
    def on_edit_event(self):
        """Handler per modifica evento VSM selezionato."""
        # Ottieni sheet corrente
        sheet = self.sheets[self.current_event_type]
        
        # Ottieni selezione
        selected_rows = list(sheet.get_selected_rows())
        
        if not selected_rows:
            messagebox.showwarning(
                _("Nessuna Selezione"),
                _("Seleziona un evento da modificare."),
                parent=self
            )
            return
        
        if len(selected_rows) > 1:
            messagebox.showwarning(
                _("Selezione Multipla"),
                _("Seleziona un solo evento per la modifica."),
                parent=self
            )
            return
        
        # Ottieni event_id e ownership
        row_idx = selected_rows[0]
        if row_idx >= len(sheet._event_metadata):
            return
        
        metadata = sheet._event_metadata[row_idx]
        event_id = metadata['event_id']
        is_mine = metadata['is_mine']
        
        # Valida ownership
        if not is_mine:
            messagebox.showerror(
                _("Operazione Non Consentita"),
                _("Puoi modificare solo i tuoi eventi VSM."),
                parent=self
            )
            return
        
        # Apri dialog edit
        from ui.dialogs.vsm_event_dialog import VSMEventDialog
        
        try:
            dialog = VSMEventDialog(
                self,
                current_username=self.current_username,
                event_type=self.current_event_type,
                event_id=event_id  # event_id not None = modalità edit
            )
            self.wait_window(dialog)
            
            # Refresh se salvato
            if hasattr(dialog, 'result') and dialog.result:
                self.refresh_events()
                if self.refresh_callback:
                    self.refresh_callback()
        
        except Exception as e:
            logger.error(f"Errore apertura dialog modifica evento: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile aprire il form: {}").format(e),
                parent=self
            )
    
    def on_delete_event(self):
        """Handler per eliminazione evento(i) VSM selezionato(i)."""
        # Ottieni sheet corrente
        sheet = self.sheets[self.current_event_type]
        
        # Ottieni selezione
        selected_rows = list(sheet.get_selected_rows())
        
        if not selected_rows:
            messagebox.showwarning(
                _("Nessuna Selezione"),
                _("Seleziona uno o più eventi da eliminare."),
                parent=self
            )
            return
        
        # Raccolta event_id e validazione ownership
        events_to_delete = []
        for row_idx in selected_rows:
            if row_idx >= len(sheet._event_metadata):
                continue
            
            metadata = sheet._event_metadata[row_idx]
            
            # Valida ownership
            if not metadata['is_mine']:
                messagebox.showerror(
                    _("Operazione Non Consentita"),
                    _("Puoi eliminare solo i tuoi eventi VSM.\nAlcuni eventi selezionati appartengono ad altri utenti."),
                    parent=self
                )
                return
            
            events_to_delete.append(metadata['event_id'])
        
        # Conferma eliminazione
        count = len(events_to_delete)
        if not messagebox.askyesno(
            _("Conferma Eliminazione"),
            _("Sei sicuro di voler eliminare {} evento(i) VSM?\nQuesta operazione non può essere annullata.").format(count),
            parent=self
        ):
            return
        
        # Elimina eventi
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                for event_id in events_to_delete:
                    delete_event_and_impacts(db_manager, event_id)
            
            messagebox.showinfo(
                _("Successo"),
                _("{} evento(i) VSM eliminato(i) con successo.").format(count),
                parent=self
            )
            
            # Refresh
            self.refresh_events()
            if self.refresh_callback:
                self.refresh_callback()
        
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore eliminazione eventi VSM: {e}")
            messagebox.showerror(
                _("Errore Eliminazione"),
                _("Impossibile eliminare gli eventi:\n{}").format(e),
                parent=self
            )
    
    def on_kpi_click(self):
        """Handler per pulsante KPI (placeholder per step futuro)."""
        messagebox.showinfo(
            _("KPI Analysis"),
            _("Funzionalità KPI Analysis in arrivo nel prossimo step.\n\n"
              "Prossime funzionalità:\n"
              "• Aggregazioni mensili/trimestrali/annuali\n"
              "• Confronto Saving vs Cost Avoidance\n"
              "• Grafici Valore Teorico vs Effettivo\n"
              "• Export Excel/CSV"),
            parent=self
        )
