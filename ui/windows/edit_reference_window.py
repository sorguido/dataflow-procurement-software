"""
EditReferenceWindow - Finestra per la modifica del riferimento di una RdO.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk
import logging

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from utils.i18n_utils import tr
from ui.dialogs.common_dialogs import show_error, show_info

logger = logging.getLogger(__name__)


class EditReferenceWindow(tk.Toplevel):
    def __init__(self, parent, request_id):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.request_id = request_id
        self.title(tr("Edit Reference"))
        self.db_path = get_db_path()
        self.transient(parent)
        
        # Frame pulsanti (sempre in fondo)
        btn_frame = ttk.Frame(self)
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text=tr("💾 Save"), command=self.save_changes).pack(side="right")
        ttk.Button(btn_frame, text=tr("❌ Cancel"), command=self.destroy).pack(side="right", padx=10)
        
        # Frame contenuto (espandibile)
        frame = ttk.Frame(self, padding="10")
        frame.pack(side="top", fill="both", expand=True)
        
        ttk.Label(frame, text=tr("Edit Reference:")).pack(anchor="w")
        self.entry_riferimento = ttk.Entry(frame, width=70)
        self.entry_riferimento.pack(fill="x", expand=True, pady=5)
        
        self.load_current_reference()
        center_window(self)
        self.wait_visibility()
        self.grab_set()
    
    def load_current_reference(self):
        try:
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                result = db_manager.get_riferimento(self.request_id)
            if result and result[0]:
                self.entry_riferimento.insert(0, result[0])
        except DatabaseError as e:
            logger.error(f"Errore database in load_current_reference: {e}", exc_info=True)
            show_error(self, tr("Error"), tr("Unable to load reference: {}").format(e))
    
    def save_changes(self):
        try:
            # BUG #46 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                db_manager.update_riferimento(self.request_id, self.entry_riferimento.get().strip())
            show_info(self.master, tr("Success"), tr("Reference updated."))
            self.destroy()
        except DatabaseError as e:
            logger.error(f"Errore database in save_changes (EditReferenceWindow): {e}", exc_info=True)
            show_error(self, tr("Error"), tr("Unable to save: {}").format(e))
