"""
EditSuppliersWindow - Finestra per la modifica dei fornitori di una RdO.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk
import logging

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from services.supplier_name_suggestion_service import SupplierNameSuggestionService
from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from utils.i18n_utils import tr
from ui.components.supplier_name_suggest import SupplierNameSuggestController
from ui.dialogs.common_dialogs import show_error, show_info, show_warning

logger = logging.getLogger(__name__)


class EditSuppliersWindow(tk.Toplevel):
    def __init__(self, parent, request_id):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.request_id = request_id
        self.title(tr("Modify Suppliers - RfQ N°{}").format(request_id))
        self.db_path = get_db_path()
        self.transient(parent)
        self.grab_set()
        
        # Frame pulsanti (sempre in fondo)
        btn_frame = ttk.Frame(self)
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text=tr("💾 Save"), command=self.save_changes).pack(side="right")
        ttk.Button(btn_frame, text=tr("❌ Cancel"), command=self.destroy).pack(side="right", padx=10)
        
        # Frame contenuto (espandibile)
        frame = ttk.Frame(self, padding="10")
        frame.pack(side="top", fill="both", expand=True)
        
        ttk.Label(frame, text=tr("Modify supplier list (names separated by comma):")).pack(anchor="w")
        self.entry_suppliers = ttk.Entry(frame, width=70)
        self.entry_suppliers.pack(fill="x", expand=True, pady=5)

        self._supplier_index = None
        self._suggest_controller = None
        self._init_supplier_suggestions()
        
        self.load_current_suppliers()
        center_window(self)

    def destroy(self):
        if self._suggest_controller is not None:
            self._suggest_controller.destroy()
            self._suggest_controller = None
        super().destroy()

    def _init_supplier_suggestions(self):
        try:
            self._supplier_index = SupplierNameSuggestionService.build_index(
                getattr(self, "db_path", get_db_path())
            )
            self._suggest_controller = SupplierNameSuggestController(
                self,
                self.entry_suppliers,
                self._get_supplier_suggestions,
                get_query_text=self._get_current_supplier_token,
                apply_suggestion=self._apply_supplier_suggestion,
                min_chars=2,
                max_items=8,
            )
        except Exception as e:
            logger.warning("Suggerimenti fornitori non disponibili: %s", e)
            self._supplier_index = None
            self._suggest_controller = None

    def _get_supplier_suggestions(self, query: str) -> list:
        if not self._supplier_index:
            return []
        return self._supplier_index.suggest(query, limit=8)

    def _get_current_supplier_token(self) -> str:
        text = self.entry_suppliers.get()
        cursor = self.entry_suppliers.index(tk.INSERT)
        left = text[:cursor]
        return left.split(",")[-1].strip()

    def _apply_supplier_suggestion(self, suggestion: str):
        text = self.entry_suppliers.get()
        cursor = self.entry_suppliers.index(tk.INSERT)
        start = text.rfind(",", 0, cursor) + 1
        end = text.find(",", cursor)
        if end < 0:
            end = len(text)

        current_token = text[start:end]
        leading_ws_len = len(current_token) - len(current_token.lstrip())
        trailing_ws_len = len(current_token) - len(current_token.rstrip())
        replacement = (" " * leading_ws_len) + suggestion + (" " * trailing_ws_len)

        updated = text[:start] + replacement + text[end:]
        self.entry_suppliers.delete(0, tk.END)
        self.entry_suppliers.insert(0, updated)
        self.entry_suppliers.icursor(start + leading_ws_len + len(suggestion))

    def _show_soft_duplicate_warning_if_needed(self, supplier_names: list):
        if not self._supplier_index:
            return

        warnings = []
        for name in supplier_names:
            candidates = self._supplier_index.get_soft_duplicate_candidates(name, limit=3)
            if not candidates:
                continue
            warnings.append(f"{name} -> {', '.join(candidates)}")

        if not warnings:
            return

        show_warning(
            self,
            tr("Possible Supplier Duplicate"),
            tr(
                "Some names may refer to existing suppliers.\n"
                "Please verify before saving:\n\n{}"
            ).format("\n".join(warnings)),
        )
    
    def load_current_suppliers(self):
        try:
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                rows = db_manager.get_fornitori_by_richiesta(self.request_id)
            self.entry_suppliers.insert(0, ", ".join([r[0] for r in rows]))
        except DatabaseError as e:
            logger.error(f"Errore database in load_current_suppliers: {e}", exc_info=True)
            show_error(self, tr("Error"), tr("Unable to load suppliers: {}").format(e))
    
    def save_changes(self):
        # Blocca se in modalità read-only
        if getattr(self, 'read_only', False):
            show_warning(self, tr("Operation Not Allowed"), tr("You cannot edit suppliers for other users' RfQs."))
            return
        
        new_suppliers = [n.strip() for n in self.entry_suppliers.get().split(',') if n.strip()]
        
        # Validazione fornitori duplicati (solo se ci sono fornitori)
        if new_suppliers:
            fornitori_lower = [f.lower() for f in new_suppliers]
            duplicati = [f for f in new_suppliers if fornitori_lower.count(f.lower()) > 1]
            duplicati_unici = list(set(duplicati))
            
            if duplicati_unici:
                show_warning(
                    self,
                    tr("Duplicate Suppliers"),
                    tr("You have entered the same supplier multiple times:\n\n{}\n\nEach supplier must be entered only once.").format(', '.join(sorted(set(duplicati_unici))))
                )
                return

        # Warning soft non bloccante per varianti simili.
        self._show_soft_duplicate_warning_if_needed(new_suppliers)
        
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

            # Refresh indice suggerimenti dopo salvataggio per mantenere risultati coerenti.
            if self._supplier_index is not None:
                try:
                    self._supplier_index = SupplierNameSuggestionService.build_index(
                        getattr(self, "db_path", get_db_path())
                    )
                    if self._suggest_controller is not None:
                        self._suggest_controller.refresh()
                except Exception as idx_err:
                    logger.warning("Impossibile aggiornare indice suggerimenti: %s", idx_err)
            
            # Context manager ha già chiuso il DB qui
            print(f"[EditSuppliersWindow] DB chiuso dal context manager")
            
            # Messaggio di successo personalizzato
            if new_suppliers:
                show_info(self.master, tr("Success"), tr("Supplier list updated."))
            else:
                show_info(self.master, tr("Success"), tr("All suppliers have been removed."))
            
            self.destroy()
            
        except DatabaseError as e:
            logger.error(f"Errore database in save_changes (EditSuppliersWindow): {e}", exc_info=True)
            show_error(self, tr("Error"), tr("Unable to save: {}").format(e))
