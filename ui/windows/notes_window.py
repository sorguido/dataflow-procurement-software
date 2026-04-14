"""
NotesWindow - Finestra per la gestione delle note di una RdO con formattazione.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk
from ui.dialogs.common_dialogs import SimpleMessageDialog
import logging
import json
import ast

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from utils.i18n_utils import tr

logger = logging.getLogger(__name__)


class NotesWindow(tk.Toplevel):
    def __init__(self, parent, request_id):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.request_id = request_id
        self.db_path = get_db_path()
        self.title(tr("Notes - RfQ N° {}").format(self.request_id))
        self.resizable(True, True)
        self.parent_window = parent
        # BUG #37 FIX: Verifica esistenza finestra prima di modificare attributes
        if hasattr(self, 'parent_window') and self.parent_window:
            try:
                if self.parent_window.winfo_exists():
                    self.parent_window.attributes('-disabled', True)
            except Exception as e:
                logger.debug(f"Impossibile disabilitare parent_window in NotesWindow: {e}")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Pulsanti di salvataggio e chiusura (sempre in fondo)
        button_frame = ttk.Frame(self)
        button_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        ttk.Button(button_frame, text=tr("💾 Save Note"), command=self.save_note).pack(side="right")
        ttk.Button(button_frame, text=tr("❌ Cancel"), command=self.on_close).pack(side="right", padx=10)

        # Frame principale (espandibile)
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(side="top", fill="both", expand=True)

        # Toolbar di formattazione
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill="x", pady=(0, 5))

        ttk.Button(toolbar, text=tr("B Bold").replace("B", "𝐁", 1), command=lambda: self.apply_tag("bold")).pack(side="left")
        ttk.Button(toolbar, text=tr("I Italic").replace("I", "𝑰", 1), command=lambda: self.apply_tag("italic")).pack(side="left", padx=5)
        ttk.Button(toolbar, text=tr("U Underline").replace("U", "U\u0332", 1), command=lambda: self.apply_tag("underline")).pack(side="left")
        # Spacer per allineare i pulsanti a sinistra e lasciare libera la caption bar
        ttk.Label(toolbar, text="").pack(side="left", expand=True, fill="x")

        # Editor di testo con scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True)
        self.text_editor = tk.Text(text_frame, wrap="word", undo=True, font=("Calibri", 11))
        scrollbar = ttk.Scrollbar(text_frame, command=self.text_editor.yview)
        self.text_editor.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.text_editor.pack(side="left", fill="both", expand=True)

        # Configurazione dei tag per la formattazione
        self.text_editor.tag_configure("bold", font=("Calibri", 11, "bold"))
        self.text_editor.tag_configure("italic", font=("Calibri", 11, "italic"))
        self.text_editor.tag_configure("underline", font=("Calibri", 11, "underline"))

        self.load_note()
        center_window(self)

    def apply_tag(self, tag_name):
        try:
            # Controlla se il tag è già applicato alla selezione
            current_tags = self.text_editor.tag_names("sel.first")
            if tag_name in current_tags:
                self.text_editor.tag_remove(tag_name, "sel.first", "sel.last")
            else:
                self.text_editor.tag_add(tag_name, "sel.first", "sel.last")
        except tk.TclError:
            # Nessun testo selezionato, non fare nulla
            pass

    def load_note(self):
        """
        Carica la nota dal database e ricostruisce il testo con la formattazione corretta.
        """
        try:
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                result = db_manager.get_note_formattate(self.request_id)
            
            if result and result[0]:
                # ✅ SICURO: Validazione prima del parsing
                note_data = result[0]
                
                # 1. Limita lunghezza massima
                if len(note_data) > 1000000:  # 1MB max
                    raise ValueError("Nota troppo grande")
                
                # 2. Usa json.loads invece di ast.literal_eval (più sicuro e veloce)
                try:
                    content_dump = json.loads(note_data)
                except json.JSONDecodeError:
                    # Fallback per compatibilità con vecchi dati
                    content_dump = ast.literal_eval(note_data)
                
                # 3. Valida struttura dati
                if not isinstance(content_dump, list):
                    raise ValueError("Formato nota non valido")
                
                # 4. Limita numero di elementi
                if len(content_dump) > 10000:
                    raise ValueError("Nota troppo complessa")
                
                self.text_editor.delete("1.0", tk.END)
                
                # Mantiene un set dei tag di formattazione attivi
                active_tags = set()
                
                # Itera su ogni elemento salvato (testo, inizio tag, fine tag)
                for item in content_dump:
                    key = item[0]
                    value = item[1]

                    if key == "text":
                        # Inserisce il testo applicando i tag attualmente attivi
                        self.text_editor.insert(tk.END, value, tuple(active_tags))
                    elif key == "tagon":
                        # Aggiunge un tag al set di quelli attivi per il testo successivo
                        active_tags.add(value)
                    elif key == "tagoff":
                        # Rimuove un tag dal set di quelli attivi
                        active_tags.discard(value)

        except (DatabaseError, SyntaxError, ValueError) as e:
            logger.error(f"Errore in load_note: {e}", exc_info=True)
            # Se la nota salvata è in un formato vecchio o corrotto, prova a caricarla come testo semplice
            if result and result[0]:
                self.text_editor.delete("1.0", tk.END)
                self.text_editor.insert("1.0", tr("Unable to restore formatting. Note loaded as plain text:\n\n{}").format(result[0]))
            SimpleMessageDialog(self, tr("Note Loading Error"), tr("It was not possible to restore the note formatting. It may have been saved with a previous version.\n\nDetails: {}").format(e), "warning")

    def on_close(self):
        try:
            if hasattr(self, 'parent_window') and self.parent_window and self.parent_window.winfo_exists():
                try:
                    self.parent_window.attributes('-disabled', False)
                except Exception:
                    pass
                try:
                    self.parent_window.focus_set()
                except Exception:
                    pass
        finally:
            self.destroy()

    def save_note(self):
        # Salva il contenuto del widget Text, inclusa la formattazione
        content_dump = self.text_editor.dump("1.0", tk.END, text=True, tag=True)
        
        # Estrae solo il testo effettivo (senza formattazione) per verificare se è vuoto
        text_content = self.text_editor.get("1.0", tk.END).strip()
        
        # Se il contenuto è vuoto, salva NULL invece di una stringa vuota
        if not text_content:
            content_to_save = None
        else:
            # Usiamo repr per salvare una rappresentazione sicura della stringa
            content_to_save = repr(content_dump)

        try:
            # BUG #46 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(getattr(self, 'db_path', get_db_path())) as db_manager:
                db_manager.update_note_formattate(self.request_id, content_to_save)
            # Chiama il metodo del genitore per aggiornare il pulsante
            if hasattr(self.master, 'check_note_status_and_update_button'):
                self.master.check_note_status_and_update_button()
            SimpleMessageDialog(self, tr("Success"), tr("Note saved successfully."), "info")
            self.on_close()
        except DatabaseError as e:
            logger.error(f"Errore database in save_note: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Database Error"), tr("Unable to save note: {}").format(e), "error")
