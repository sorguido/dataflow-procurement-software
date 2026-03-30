"""
PurchaseOrderWindow - Finestra per la gestione dei numeri ordine di acquisto.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
import json
import re
from tksheet import Sheet

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.window_utils import center_window
from utils.i18n_utils import _, get_current_language
from ui.dialogs.common_dialogs import SimpleMessageDialog, SimpleYesNoDialog

logger = logging.getLogger(__name__)


class PurchaseOrderWindow(tk.Toplevel):
    def __init__(self, parent, request_id):
        super().__init__(parent)
        self.request_id = request_id
        self.parent = parent
        
        # Imposta titolo in base alla lingua
        if get_current_language() == 'it':
            self.title("Gestione Numeri Ordine")
        else:
            self.title("Purchase Order Management")
        
        # MODIFICA 2: Finestra 20% più corta (da 500 a 400)
        self.geometry("700x400")
        self.resizable(True, True)
        
        # MODIFICA 4: Mantieni la finestra sempre in primo piano
        self.transient(parent)
        self.grab_set()
        
        # IMPORTANTE: Frame pulsanti PRIMA (side=bottom) - così rimangono sempre visibili
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        # Pulsante Aggiungi con testo tradotto (a sinistra)
        if get_current_language() == 'it':
            add_text = "➕ Aggiungi"
        else:
            add_text = "➕ Add"
        
        add_btn = ttk.Button(
            bottom_frame, 
            text=add_text,
            command=lambda: self.add_po_entry_safe()
        )
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Pulsante Elimina con testo tradotto (a sinistra)
        if get_current_language() == 'it':
            delete_text = "❌ Elimina"
        else:
            delete_text = "❌ Delete"
        
        delete_btn = ttk.Button(
            bottom_frame, 
            text=delete_text,
            command=lambda: self.delete_selected_po_safe()
        )
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        # Pulsante Chiudi con testo tradotto (a destra)
        if get_current_language() == 'it':
            close_text = "Chiudi"
        else:
            close_text = "Close"
        
        close_btn = ttk.Button(
            bottom_frame, 
            text=close_text,
            command=self.on_closing
        )
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # Ora creiamo i frame contenuto (dall'alto verso il basso)
        # Frame superiore con istruzioni
        info_frame = ttk.Frame(self)
        info_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        # Testo istruzioni tradotto
        if get_current_language() == 'it':
            info_text = "Associa numeri di ordine ai fornitori della RdO"
        else:
            info_text = "Associate purchase order numbers with RfQ suppliers"
        
        info_label = ttk.Label(
            info_frame, 
            text=info_text,
            font=('Segoe UI', 10)
        )
        info_label.pack(anchor='w')
        
        # Frame per i controlli di inserimento
        controls_frame = ttk.Frame(self)
        controls_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        # Campo numero ordine con label tradotta
        if get_current_language() == 'it':
            po_label_text = "Numero Ordine:"
        else:
            po_label_text = "PO Number:"
        
        po_label = ttk.Label(controls_frame, text=po_label_text)
        po_label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
        
        self.po_entry = ttk.Entry(controls_frame, width=30)
        self.po_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        
        # ComboBox fornitore con label tradotta
        if get_current_language() == 'it':
            supplier_label_text = "Fornitore:"
        else:
            supplier_label_text = "Supplier:"
        
        supplier_label = ttk.Label(controls_frame, text=supplier_label_text)
        supplier_label.grid(row=0, column=2, sticky='w', padx=5, pady=5)
        
        self.supplier_combo = ttk.Combobox(controls_frame, state='readonly', width=25)
        self.supplier_combo.grid(row=0, column=3, sticky='ew', padx=5, pady=5)
        
        controls_frame.columnconfigure(1, weight=1)
        controls_frame.columnconfigure(3, weight=1)
        
        # Frame per la griglia (espandibile)
        grid_frame = ttk.Frame(self)
        grid_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Headers tradotti per la griglia
        if get_current_language() == 'it':
            headers = ["Numero Ordine", "Fornitore"]
        else:
            headers = ["PO Number", "Supplier"]
        
        # Crea il foglio tksheet
        self.sheet = Sheet(
            grid_frame,
            headers=headers,
            header_font=("Segoe UI", 10, "bold"),
            font=("Segoe UI", 10, "normal"),
            show_row_index=False,
            show_top_left=False,
            empty_horizontal=0,
            empty_vertical=0
        )
        self.sheet.enable_bindings(
            "single_select",
            "row_select",
            "drag_select",
            "column_width_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "copy",
            "cut",
            "paste",
            "delete",
            "edit_cell"
        )
        self.sheet.pack(fill=tk.BOTH, expand=True)
        
        # Salva gli headers per uso futuro
        self.headers = headers
        
        # Gestione chiusura finestra - PRIMA di caricare i dati
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Centra la finestra
        center_window(self)
        
        # Carica i fornitori e i PO esistenti - ALLA FINE
        self.load_suppliers_for_request()
        self.load_po_entries()
    
    def on_closing(self):
        """Salva i dati prima di chiudere."""
        try:
            self.save_po_entries()
        except Exception as e:
            logger.error(f"Errore nel salvataggio PO durante chiusura: {e}")
        
        # Notifica il parent per aggiornare eventualmente l'interfaccia
        if hasattr(self.parent, 'load_po_numbers'):
            try:
                self.parent.load_po_numbers()
            except Exception as e:
                logger.error(f"Errore nell'aggiornamento parent: {e}")
        
        self.destroy()
    
    def load_suppliers_for_request(self):
        """Carica i fornitori associati a questa richiesta."""
        try:
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                suppliers_rows = db_manager.get_fornitori_ordered_for_request(self.request_id)
            suppliers = [row[0] for row in suppliers_rows]
            
            # MODIFICA 3: Non impostare un valore predefinito
            # Forza l'utente a selezionare esplicitamente il fornitore
            self.supplier_combo['values'] = suppliers
            # NON impostare: self.supplier_combo.current(0)
        except DatabaseError as e:
            logger.error(f"Errore nel caricamento fornitori per PO: {e}")
            # Non mostrare messagebox che potrebbe distruggere la finestra
            # Lascia semplicemente il combo vuoto
    
    def load_po_entries(self):
        """Carica i numeri ordine esistenti dal database."""
        try:
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                row = db_manager.get_numeri_ordine(self.request_id)
            
            if row and row[0]:
                po_data = row[0]
                po_list = []
                
                # Tenta di parsare come JSON (nuovo formato)
                try:
                    po_list = json.loads(po_data)
                    if not isinstance(po_list, list):
                        po_list = []
                except (json.JSONDecodeError, TypeError):
                    # Formato vecchio: stringa con virgole
                    # Converti in nuovo formato (senza fornitore associato)
                    old_numbers = [n.strip() for n in po_data.split(',') if n.strip()]
                    po_list = [{"po_number": num, "supplier": ""} for num in old_numbers]
                
                # Popola il foglio
                data = [[entry.get("po_number", ""), entry.get("supplier", "")] for entry in po_list]
                if data:
                    self.sheet.set_sheet_data(data)
                    # Auto-ridimensiona le colonne dopo aver caricato i dati
                    self.auto_resize_columns()
        except DatabaseError as e:
            logger.error(f"Errore nel caricamento numeri ordine: {e}")
    
    def auto_resize_columns(self):
        """Auto-ridimensiona le colonne in base al contenuto."""
        try:
            import tkinter.font as tkfont
            
            # Font per intestazioni e contenuto
            header_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
            content_font = tkfont.Font(family="Segoe UI", size=10, weight="normal")
            
            PADDING_PX = 30  # Padding per evitare troncamenti
            MIN_WIDTH = 150  # Larghezza minima
            
            # Ottieni i dati correnti
            data_rows = self.sheet.get_sheet_data()
            
            # Per ogni colonna
            for col_idx in range(len(self.headers)):
                header_text = self.headers[col_idx]
                max_width = header_font.measure(header_text)
                
                # Controlla il contenuto di tutte le righe
                for row in data_rows:
                    if col_idx < len(row):
                        cell_value = str(row[col_idx])
                        cell_width = content_font.measure(cell_value)
                        max_width = max(max_width, cell_width)
                
                # Calcola larghezza finale con padding e minimo
                column_width = max(int(max_width + PADDING_PX), MIN_WIDTH)
                self.sheet.column_width(column=col_idx, width=column_width)
                
        except Exception as e:
            logger.warning(f"Errore auto-ridimensionamento colonne PO: {e}")
            # Fallback a larghezze fisse
            self.sheet.column_width(column=0, width=200)
            self.sheet.column_width(column=1, width=250)
    
    def add_po_entry_safe(self):
        """Wrapper sicuro per add_po_entry."""
        if hasattr(self, 'sheet'):
            self.add_po_entry()
    
    def delete_selected_po_safe(self):
        """Wrapper sicuro per delete_selected_po."""
        if hasattr(self, 'sheet'):
            self.delete_selected_po()
    
    def add_po_entry(self):
        """Aggiunge un nuovo numero ordine con fornitore associato."""
        po_number = self.po_entry.get().strip()
        supplier = self.supplier_combo.get().strip()
        
        # BUG #35 FIX: Validazione esplicita campo vuoto con feedback utente
        if not po_number:
            SimpleMessageDialog(self, _("Attenzione"), _("Inserisci un numero ordine valido."), "warning")
            return
        
        # BUG #25 FIX: Previeni SQL injection e caratteri pericolosi
        FORBIDDEN_CHARS = re.compile(r"[';\"\\`<>]")
        if FORBIDDEN_CHARS.search(po_number):
            logger.warning(f"Caratteri pericolosi rimossi da PO number: '{po_number}'")
            po_number = FORBIDDEN_CHARS.sub('', po_number)
            self.po_entry.delete(0, tk.END)
            self.po_entry.insert(0, po_number)
        
        if FORBIDDEN_CHARS.search(supplier):
            logger.warning(f"Caratteri pericolosi rimossi da supplier: '{supplier}'")
            supplier = FORBIDDEN_CHARS.sub('', supplier)
        
        if not po_number:
            if get_current_language() == 'it':
                SimpleMessageDialog(self, "Campo obbligatorio", "Inserire il numero ordine.", "warning")
            else:
                SimpleMessageDialog(self, "Required Field", "Please enter the PO number.", "warning")
            return
        
        if not supplier:
            if get_current_language() == 'it':
                SimpleMessageDialog(self, "Campo obbligatorio", "Selezionare un fornitore.", "warning")
            else:
                SimpleMessageDialog(self, "Required Field", "Please select a supplier.", "warning")
            return
        
        # Aggiungi alla griglia
        current_data = self.sheet.get_sheet_data()
        current_data.append([po_number, supplier])
        self.sheet.set_sheet_data(current_data)
        
        # Auto-ridimensiona le colonne dopo l'aggiunta
        self.auto_resize_columns()
        
        # Pulisci i campi
        self.po_entry.delete(0, tk.END)
        
        # Salva automaticamente
        self.save_po_entries()
    
    def delete_selected_po(self):
        """Elimina il numero ordine selezionato."""
        selected = self.sheet.get_currently_selected()
        
        if not selected:
            if get_current_language() == 'it':
                SimpleMessageDialog(self, "Nessuna selezione", "Selezionare una riga da eliminare.", "warning")
            else:
                SimpleMessageDialog(self, "No Selection", "Please select a row to delete.", "warning")
            return
        
        # BUG #18 FIX: Validazione robusta dell'indice riga prima dell'accesso
        row_idx = selected.row if hasattr(selected, 'row') else None
        if row_idx is None:
            logger.warning("delete_selected_po: selected non ha attributo 'row'")
            return
        
        # BUG #18 FIX: Ottieni i dati PRIMA di validare l'indice
        current_data = self.sheet.get_sheet_data()
        
        # BUG #18 FIX: Verifica che l'indice sia valido PRIMA di mostrare il dialog
        if not (0 <= row_idx < len(current_data)):
            logger.error(f"delete_selected_po: Indice {row_idx} fuori range (0-{len(current_data)-1})")
            if get_current_language() == 'it':
                SimpleMessageDialog(self, "Errore", f"Impossibile eliminare: indice riga non valido ({row_idx}).", "error")
            else:
                SimpleMessageDialog(self, "Error", f"Cannot delete: invalid row index ({row_idx}).", "error")
            return
        
        # Conferma eliminazione
        if get_current_language() == 'it':
            confirm = SimpleYesNoDialog(self, "Conferma eliminazione", "Eliminare il numero ordine selezionato?").result
        else:
            confirm = SimpleYesNoDialog(self, "Confirm Deletion", "Delete the selected PO number?").result
        
        if confirm:
            del current_data[row_idx]
            self.sheet.set_sheet_data(current_data)
            self.save_po_entries()
    
    def save_po_entries(self):
        """Salva i numeri ordine nel database in formato JSON."""
        try:
            # Ottieni i dati dal foglio
            data = self.sheet.get_sheet_data()
            
            # Converti in lista di dizionari
            po_list = []
            for row in data:
                if len(row) >= 2 and row[0].strip():
                    po_list.append({
                        "po_number": row[0].strip(),
                        "supplier": row[1].strip() if len(row) > 1 else ""
                    })
            
            # Salva come JSON
            json_data = json.dumps(po_list, ensure_ascii=False)
            
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                db_manager.update_numeri_ordine(self.request_id, json_data)
            
            logger.info(f"Numeri ordine salvati per RdO {self.request_id}: {len(po_list)} entries")
        except DatabaseError as e:
            logger.error(f"Errore nel salvataggio numeri ordine: {e}")
            SimpleMessageDialog(self, _("Errore"), _("Errore nel salvataggio dei numeri ordine."), "error")
