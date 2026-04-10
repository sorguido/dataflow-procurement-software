"""
EditSuppliersWindow - Finestra per la modifica dei fornitori di una RdO.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from utils.i18n_utils import tr

logger = logging.getLogger(__name__)


class EditSuppliersWindow(tk.Toplevel):
    def __init__(self, parent, request_id):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.request_id = request_id
        self.title(tr("Modifica Fornitori - RdO N°{}").format(request_id))
        self.db_path = get_db_path()
        self.transient(parent)
        self.grab_set()
        
        # Frame pulsanti (sempre in fondo)
        btn_frame = ttk.Frame(self)
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text=tr("💾 Salva"), command=self.save_changes).pack(side="right")
        ttk.Button(btn_frame, text=tr("❌ Annulla"), command=self.destroy).pack(side="right", padx=10)
        
        # Frame contenuto (espandibile)
        frame = ttk.Frame(self, padding="10")
        frame.pack(side="top", fill="both", expand=True)
        
        ttk.Label(frame, text=tr("Modifica elenco fornitori (nomi separati da virgola):")).pack(anchor="w")
        self.entry_suppliers = ttk.Entry(frame, width=70)
        self.entry_suppliers.pack(fill="x", expand=True, pady=5)
        
        self.load_current_suppliers()
        center_window(self)
    
    def load_current_suppliers(self):
        try:
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                rows = db_manager.get_fornitori_by_richiesta(self.request_id)
            self.entry_suppliers.insert(0, ", ".join([r[0] for r in rows]))
        except DatabaseError as e:
            logger.error(f"Errore database in load_current_suppliers: {e}", exc_info=True)
            messagebox.showerror(tr("Errore"), tr("Impossibile caricare i fornitori: {}").format(e), parent=self)
    
    def save_changes(self):
        # Blocca se in modalità read-only
        if getattr(self, 'read_only', False):
            messagebox.showwarning(
                tr("Operazione Non Consentita"),
                tr("Non puoi modificare i fornitori di RdO di altri utenti."),
                parent=self
            )
            return
        
        new_suppliers = [n.strip() for n in self.entry_suppliers.get().split(',') if n.strip()]
        
        # Validazione fornitori duplicati (solo se ci sono fornitori)
        if new_suppliers:
            fornitori_lower = [f.lower() for f in new_suppliers]
            duplicati = [f for f in new_suppliers if fornitori_lower.count(f.lower()) > 1]
            duplicati_unici = list(set(duplicati))
            
            if duplicati_unici:
                messagebox.showwarning(
                    tr("Fornitori Duplicati"),
                    tr("Hai inserito lo stesso fornitore più volte:\n\n{}\n\nOgni fornitore deve essere inserito una sola volta.").format(', '.join(sorted(set(duplicati_unici)))),
                    parent=self
                )
                return
        
        try:
            # Recupera i fornitori PRIMA di eliminarli e gli id_dettaglio
            # BUG FIX: Usa context manager per garantire chiusura DB automatica
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                old_suppliers_rows = db_manager.get_fornitori_by_richiesta(self.request_id)
                old_suppliers = [row[0] for row in old_suppliers_rows]
                detail_ids_rows = db_manager.get_dettaglio_ids_by_richiesta(self.request_id)
                detail_ids = [row[0] for row in detail_ids_rows]
                
                # Usa db_manager per salvare con transazione
                print(f"[EditSuppliersWindow] Inizio save_suppliers_with_transaction...")
                db_manager.save_suppliers_with_transaction(self.request_id, new_suppliers, old_suppliers, detail_ids)
                print(f"[EditSuppliersWindow] Fine save_suppliers_with_transaction")
                
                # Verifica immediata che i dati siano stati salvati
                verify_rows = db_manager.get_fornitori_by_richiesta(self.request_id)
                verify_count = len(verify_rows)
                print(f"[EditSuppliersWindow] VERIFICA POST-SALVATAGGIO: {verify_count} fornitori trovati nel DB (attesi: {len(new_suppliers)})")
                print(f"[EditSuppliersWindow] Fornitori salvati: {[r[0] for r in verify_rows]}")
            
            # Context manager ha già chiuso il DB qui
            print(f"[EditSuppliersWindow] DB chiuso dal context manager")
            
            # Messaggio di successo personalizzato
            if new_suppliers:
                messagebox.showinfo(tr("Successo"), tr("Elenco fornitori aggiornato."), parent=self.master)
            else:
                messagebox.showinfo(tr("Successo"), tr("Tutti i fornitori sono stati rimossi."), parent=self.master)
            
            self.destroy()
            
        except DatabaseError as e:
            logger.error(f"Errore database in save_changes (EditSuppliersWindow): {e}", exc_info=True)
            messagebox.showerror(tr("Errore"), tr("Impossibile salvare: {}").format(e), parent=self)
