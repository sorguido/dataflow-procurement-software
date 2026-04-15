"""AttachmentWindow - Finestra per la gestione degli allegati di una RdO.
Estratta da dataflow.py per compatibilità con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk, filedialog
import logging
import os
import time
import shutil
import tempfile
import webbrowser
import threading
from urllib.parse import unquote, urlparse
from tksheet import Sheet
from datetime import datetime

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path, get_fixed_attachments_dir
from utils.window_utils import center_window
from utils.resource_utils import set_window_icon
from utils.i18n_utils import tr
from utils.validation_utils import sanitize_filename
from ui.dialogs.common_dialogs import (
    SimpleMessageDialog,
    SimpleYesNoDialog,
    show_error,
    show_warning,
)

try:
    from tkinterdnd2 import TkinterDnD
except Exception:
    TkinterDnD = None

logger = logging.getLogger(__name__)


class AttachmentWindow(tk.Toplevel):
    def __init__(self, parent, request_id, attachment_type, read_only=False, source_db_path=None):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.transient(parent)
        self.grab_set()
        self.request_id = request_id
        self.attachment_type = attachment_type
        self.read_only = read_only
        
        # Determina quale database usare (locale o remoto) e la cartella Attachments
        if source_db_path and os.path.exists(source_db_path):
            self.db_path = source_db_path
            logger.info(f"[AttachmentWindow] Usando DB remoto: {source_db_path}")
            
            # Calcola path Attachments con fallback robusto
            try:
                db_parent = os.path.dirname(self.db_path)
                dataflow_root = os.path.dirname(db_parent)
                self.attachments_base = os.path.join(dataflow_root, 'Attachments')
                
                if not os.path.isdir(self.attachments_base):
                    logger.warning(f"Cartella Attachments non trovata: {self.attachments_base}")
                    self.attachments_base = None
                else:
                    logger.info(f"[AttachmentWindow] Path Attachments remoto: {self.attachments_base}")
            except Exception as e:
                logger.error(f"Errore calcolo path Attachments remoto: {e}")
                self.attachments_base = None
        else:
            self.db_path = get_db_path()
            logger.info(f"[AttachmentWindow] Usando DB locale: {self.db_path}")
            self.attachments_base = get_fixed_attachments_dir()
        
        # Titolo con suffisso SOLA LETTURA se applicabile
        if self.attachment_type == "Offerta Fornitore":
            title_base = tr("Manage Supplier Offer")
        else:
            title_base = tr("Manage Internal Document")
        
        if self.read_only:
            title_base += tr(" [READ-ONLY]")
        
        self.title(title_base)
        
        # Lista per tracciare file temporanei creati
        self.temp_files = []
        
        # Lista per memorizzare gli ID degli allegati (non visibili nella tabella)
        self.attachment_ids = []
        
        # Handler per cleanup alla chiusura
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Frame per avviso (solo per Offerta Fornitore)
        if self.attachment_type == "Offerta Fornitore":
            frame_warning = ttk.Frame(self)
            frame_warning.pack(side="bottom", fill="x", padx=10, pady=(0, 5))
            warning_label = tk.Label(frame_warning, 
                                    text=tr("⚠️ Please select a supplier from the list below before adding an attachment"), 
                                    fg="red", 
                                    font=(None, 10, "bold"))
            warning_label.pack()
        
        # Frame per i pulsanti (SEMPRE IN FONDO, non espandibile)
        frame_buttons = ttk.Frame(self)
        frame_buttons.pack(side="bottom", fill="x", padx=10, pady=10)
        
        # Pulsanti con gestione read-only
        self.btn_delete = ttk.Button(frame_buttons, text=tr("❌ Delete Selected"), command=self.delete_attachment)
        self.btn_delete.pack(side="right")

        self.btn_add = ttk.Button(frame_buttons, text=tr("➕ Add..."), command=self.add_attachment)
        self.btn_add.pack(side="left")
        
        ttk.Button(frame_buttons, text=tr("📂 Open Selected"), command=self.open_attachment).pack(side="left", padx=10)
        ttk.Button(frame_buttons, text=tr("⬇️ Download..."), command=self.download_attachment).pack(side="left")

        if self.attachment_type == "Offerta Fornitore":
            self.combo_suppliers = ttk.Combobox(frame_buttons, state="readonly")
            self.combo_suppliers.pack(side="left", padx=10)
            self.load_suppliers_for_request()
            
        # Disabilita pulsanti di modifica se in modalità read-only
        if self.read_only:
            self.btn_add.config(state='disabled')
            self.btn_delete.config(state='disabled')
            if hasattr(self, 'combo_suppliers'):
                self.combo_suppliers.config(state='disabled')
        
        # Frame contenuto principale (ESPANDIBILE, sopra i pulsanti)
        frame_main = ttk.Frame(self)
        frame_main.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 0))
        
        # Creiamo un frame per contenere il foglio
        sheet_frame = ttk.Frame(frame_main)
        sheet_frame.pack(fill="both", expand=True)
        
        # Creiamo il widget tksheet
        self.sheet_attachments = Sheet(sheet_frame,
                                       theme="light blue",
                                       header_font=("Calibri", 11, "bold"),
                                       font=("Calibri", 11, "normal"))
        
        # Abilita solo i binding necessari, ESCLUDE edit_cell per impedire modifiche
        self.sheet_attachments.enable_bindings(
            "single_select",
            "row_select",
            "column_width_resize",
            "double_click_column_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "copy"
        )
        
        self.sheet_attachments.pack(fill="both", expand=True)

        # Hint non invasivo tra area lista allegati e pulsanti inferiori.
        # Mantiene il layout esistente senza overlay.
        hint_frame = ttk.Frame(self)
        hint_frame.pack(side="bottom", fill="x", padx=10, pady=(0, 2))
        self.lbl_drop_hint = ttk.Label(
            hint_frame,
            text=tr("Trascina qui i file oppure usa '+ Aggiungi...'"),
            foreground="black",
            anchor="center"
        )
        self.lbl_drop_hint.pack(fill="x")

        self.load_attachments()
        self._init_drag_and_drop()
        
        # Imposta dimensione minima per mostrare tutte le colonne
        self.geometry("850x450")
        self.minsize(800, 400)
        
        center_window(self)
        self.deiconify()

    def _init_drag_and_drop(self):
        """Abilita DnD solo sulla lista allegati, in modalità opzionale/reversibile."""
        self._dnd_available = False
        self._dnd_drop_cmd = None
        self._dnd_backend = "none"
        self._dnd_target = None

        # In read-only lasciamo DnD disattivo per evitare UX incoerente.
        if self.read_only:
            return

        target = self._get_dnd_target_widget()
        if target is None:
            logger.warning("DnD non inizializzato: target widget non disponibile.")
            return

        # Tentativo 1: tkdnd già disponibile nel runtime Tk.
        if self._try_enable_tkdnd(target, backend_name="runtime-tkdnd"):
            return

        # Tentativo 2: bootstrap tkdnd tramite tkinterdnd2 (opzionale).
        if self._try_enable_tkinterdnd2(target):
            return

        logger.info("Drag-and-drop allegati disattivato: backend non disponibile. Fallback picker attivo.")

    def _get_dnd_target_widget(self):
        """Target DnD: solo area bianca centrale lista (MT), non header/hint/pulsanti."""
        if hasattr(self, "sheet_attachments") and hasattr(self.sheet_attachments, "MT"):
            return self.sheet_attachments.MT
        return getattr(self, "sheet_attachments", None)

    def _try_enable_tkdnd(self, target_widget, backend_name):
        """Prova registrazione DnD usando tkdnd già caricato nel runtime Tcl/Tk."""
        try:
            self.tk.call('package', 'require', 'tkdnd')
            self.tk.call('tkdnd::drop_target', 'register', target_widget._w, 'DND_Files')
            self._dnd_drop_cmd = self.register(self._on_tkdnd_drop)

            # Bind su entrambe le varianti evento per compatibilità runtime differenti.
            self.tk.call('bind', target_widget._w, '<<Drop>>', f'{self._dnd_drop_cmd} %D')
            self.tk.call('bind', target_widget._w, '<<Drop:DND_Files>>', f'{self._dnd_drop_cmd} %D')

            self._dnd_available = True
            self._dnd_backend = backend_name
            self._dnd_target = target_widget
            logger.info("Drag-and-drop allegati attivato su area lista (backend=%s).", backend_name)
            return True
        except Exception as dnd_error:
            logger.info("Backend DnD '%s' non disponibile: %s", backend_name, dnd_error)
            return False

    def _try_enable_tkinterdnd2(self, target_widget):
        """Prova bootstrap tkdnd tramite tkinterdnd2 (se installato)."""
        if TkinterDnD is None:
            logger.info("tkinterdnd2 non installato nel runtime Python.")
            return False
        try:
            TkinterDnD._require(self)
        except Exception as e:
            logger.info("tkinterdnd2 presente ma _require fallito: %s", e)
            return False
        return self._try_enable_tkdnd(target_widget, backend_name="tkinterdnd2")

    def _on_tkdnd_drop(self, drop_data):
        """Callback drop tkdnd: normalizza i path e instrada il flusso di upload."""
        try:
            paths = self._parse_drop_paths(drop_data)
            self._handle_dropped_paths(paths)
        except Exception as e:
            logger.error("Errore gestione drop allegati: %s", e, exc_info=True)
            show_error(self, tr("Error"), tr("Unable to add attachment: {}").format(e))
        return "break"

    def _parse_drop_paths(self, raw_data):
        """Parsa i path file dal payload DnD (Windows/Linux, spazi e caratteri speciali)."""
        if not raw_data:
            return []

        try:
            parts = self.tk.splitlist(raw_data)
        except tk.TclError:
            parts = [raw_data]

        parsed = []
        for part in parts:
            item = (part or "").strip()
            if not item:
                continue
            if item.startswith("{") and item.endswith("}"):
                item = item[1:-1]
            if item.startswith('"') and item.endswith('"'):
                item = item[1:-1]

            if item.lower().startswith("file://"):
                url = urlparse(item)
                item = unquote(url.path or "")
                if os.name == "nt" and item.startswith("/") and len(item) > 2 and item[2] == ":":
                    item = item[1:]

            item = os.path.normpath(item)
            if item:
                parsed.append(item)
        return parsed

    def _handle_dropped_paths(self, paths):
        """Applica policy multi-file e instrada su _attach_from_path()."""
        if self.read_only:
            show_warning(self, tr("Operation Not Allowed"), tr("You cannot add attachments to other users' RfQs."))
            return

        if self.attachment_type == "Offerta Fornitore" and len(paths) > 1:
            show_warning(self, tr("Warning"), tr("For supplier offers, drop only one file at a time."))
            return

        valid_files = [p for p in paths if p and os.path.isfile(p)]
        if not valid_files:
            show_warning(self, tr("Warning"), tr("No valid files were dropped."))
            return

        for path in valid_files:
            self._attach_from_path(path)

    def _attach_from_path(self, filepath: str):
        """
        Punto unico di persistenza allegati:
        validazioni -> naming -> copy -> insert DB -> refresh UI.
        """
        if self.read_only:
            show_warning(self, tr("Operation Not Allowed"), tr("You cannot add attachments to other users' RfQs."))
            return

        supplier = self.combo_suppliers.get() if self.attachment_type == "Offerta Fornitore" else "Interno"
        if not supplier and self.attachment_type == "Offerta Fornitore":
            show_warning(self, tr("Warning"), tr("Select a supplier."))
            return

        archive_path = self.attachments_base
        if not archive_path:
            show_error(self, tr("Error"), tr("Attachment path not available."))
            return

        try:
            file_ext = os.path.splitext(filepath)[1]
            sanitized_supplier = sanitize_filename(supplier)
            db_manager_temp = DatabaseManager(self.db_path, read_only=self.read_only)
            try:
                next_id = db_manager_temp.get_max_allegato_id() + 1
            finally:
                try:
                    db_manager_temp.close()
                except Exception:
                    pass

            if self.attachment_type == "Documento Interno":
                new_filename = f"RfQ{self.request_id}_ID{next_id}{file_ext}"
            else:
                new_filename = f"RfQ{self.request_id}_{sanitized_supplier}_ID{next_id}{file_ext}"

            dest_path = os.path.join(archive_path, new_filename)
            shutil.copy(filepath, dest_path)

            with DatabaseManager(self.db_path, read_only=self.read_only) as db_manager:
                db_manager.insert_allegato_richiesta_link(
                    self.request_id,
                    os.path.basename(filepath),
                    self.attachment_type,
                    supplier,
                    new_filename
                )
        except Exception as e:
            show_error(self, tr("Error"), tr("Unable to add attachment: {}").format(e))
        finally:
            self.load_attachments()

    def on_closing(self):
        """Pulisce i file temporanei prima di chiudere la finestra con gestione sicura."""
        # Disabilita i pulsanti per evitare nuove operazioni durante la chiusura
        try:
            for widget in self.winfo_children():
                if isinstance(widget, (ttk.Button, tk.Button)):
                    widget.config(state='disabled')
        except:
            pass
        
        # Garbage collection UNA SOLA VOLTA all'inizio
        import sys
        if sys.platform == 'win32':
            try:
                import gc
                gc.collect()
            except:
                pass
        
        # Attendi eventuali operazioni DB in corso con gestione robusta
        max_wait = 30
        wait_count = 0
        window_destroyed = False
        
        while wait_count < max_wait:
            active_db_threads = [t for t in threading.enumerate() 
                                if 'database' in t.name.lower()]
            if not active_db_threads:
                break
            
            try:
                self.update()
            except Exception as update_error:
                logger.debug(f"Errore update() durante chiusura: {update_error}")
                window_destroyed = True
                break
            
            time.sleep(0.1)
            wait_count += 1
        
        # Pulisci i file temporanei
        for temp_path in self.temp_files:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    logger.info(f"File temporaneo eliminato: {temp_path}")
            except PermissionError:
                logger.debug(f"File temporaneo in uso, verrà eliminato dal cleanup automatico: {temp_path}")
            except Exception as e:
                logger.warning(f"Impossibile eliminare file temporaneo {temp_path}: {e}")
        
        if not window_destroyed:
            try:
                self.destroy()
            except Exception as destroy_error:
                logger.debug(f"Errore destroy() durante chiusura: {destroy_error}")

    def delete_attachment(self):
        if self.read_only:
            SimpleMessageDialog(self, tr("Operation Not Allowed"), tr("You cannot delete attachments for other users' RfQs."), "warning")
            return
        
        selected = self.sheet_attachments.get_currently_selected()
        if not selected or selected.row is None:
            SimpleMessageDialog(self, tr("Warning"), tr("Select an attachment to delete."), "warning")
            return
            
        if SimpleYesNoDialog(self, tr("Delete Confirmation"), tr("Are you sure you want to delete this attachment?")).result:
            row_idx = selected.row
            
            if row_idx < 0 or row_idx >= len(self.attachment_ids):
                logger.error(f"Indice riga non valido in delete: {row_idx}, totale: {len(self.attachment_ids)}")
                SimpleMessageDialog(self, tr("Error"), tr("Unable to identify the selected attachment."), "error")
                return
            
            attachment_id = self.attachment_ids[row_idx]
            try:
                file_to_delete = None
                db_manager = DatabaseManager(self.db_path, read_only=self.read_only)
                try:
                    try:
                        result = db_manager.get_allegato_file_data(attachment_id)
                        if result:
                            nome_file, dati_file, percorso_esterno = result
                            if percorso_esterno:
                                base_path = self.attachments_base
                                if base_path:
                                    file_to_delete = os.path.join(base_path, percorso_esterno)
                    except DatabaseError as fetch_error:
                        logger.warning(f"Impossibile recuperare informazioni allegato da eliminare: {fetch_error}")
                    db_manager.delete_allegato(attachment_id)
                finally:
                    try:
                        db_manager.close()
                    except Exception as close_error:
                        logger.warning(f"Errore chiusura DB in delete_attachment: {close_error}")

                if file_to_delete and os.path.exists(file_to_delete):
                    try:
                        os.remove(file_to_delete)
                        logger.info(f"Allegato eliminato dal disco: {file_to_delete}")
                    except Exception as disk_error:
                        logger.warning(f"Impossibile eliminare il file allegato {file_to_delete}: {disk_error}")
                
                SimpleMessageDialog(self, tr("Deletion"), tr("Attachment deleted."), "info")
                self.load_attachments()
                
                # Se è stato eliminato un documento SQDC, aggiorna il pulsante nella finestra parent
                if self.attachment_type == "Documento Interno":
                    if hasattr(self.master, 'check_sqdc_status_and_update_button'):
                        try:
                            self.master.check_sqdc_status_and_update_button()
                        except Exception as e:
                            logger.warning(f"Impossibile aggiornare pulsante SQDC nel parent: {e}")
                
            except DatabaseError as e: 
                SimpleMessageDialog(self, tr("Database Error"), tr("Unable to delete attachment: {}").format(e), "error")

    def load_attachments(self):
        try:
            with DatabaseManager(self.db_path, read_only=self.read_only) as db_manager:
                has_date_column = db_manager.check_table_has_column('allegati_richiesta', 'data_inserimento')
                rows = db_manager.get_allegati_by_richiesta(self.request_id, self.attachment_type, has_date_column)
            
            if has_date_column:
                self.attachment_ids = [id_allegato for id_allegato, nf, nfile, di in rows]
                
                data_rows = []
                for id_all, nome_fornitore, nome_file, data_inserimento in rows:
                    data_formattata = ""
                    if data_inserimento:
                        date_formats = [
                            '%Y-%m-%d %H:%M:%S',
                            '%Y-%m-%d',
                            '%d/%m/%Y'
                        ]
                        
                        for fmt in date_formats:
                            try:
                                dt = datetime.strptime(str(data_inserimento).strip(), fmt)
                                data_formattata = dt.strftime('%d/%m/%Y')
                                break
                            except (ValueError, TypeError):
                                continue
                        
                        if not data_formattata:
                            logger.warning(f"Formato data non riconosciuto per allegato {id_all}: '{data_inserimento}'")
                            data_formattata = str(data_inserimento) if data_inserimento else ""
                    
                    data_rows.append([str(nome_fornitore), str(nome_file), data_formattata])
                
                headers = [tr("Supplier"), tr("File Name"), tr("Insert Date")]
                self.sheet_attachments.headers(headers)
                self.sheet_attachments.set_sheet_data(data_rows)
                
                self.sheet_attachments.column_width(column=0, width=200)
                self.sheet_attachments.column_width(column=1, width=350)
                self.sheet_attachments.column_width(column=2, width=150)
            else:
                self.attachment_ids = [id_allegato for id_allegato, nf, nfile in rows]
                data_rows = [[str(nome_fornitore), str(nome_file)] for id_all, nome_fornitore, nome_file in rows]
                
                headers = [tr("Supplier"), tr("File Name")]
                self.sheet_attachments.headers(headers)
                self.sheet_attachments.set_sheet_data(data_rows)
                
                self.sheet_attachments.column_width(column=0, width=200)
                self.sheet_attachments.column_width(column=1, width=400)
            
        except DatabaseError as e:
            logger.error(f"Errore database in load_attachments: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Database Error"), tr("Unable to load attachments: {}").format(e), "error")

    def load_suppliers_for_request(self):
        try:
            with DatabaseManager(self.db_path, read_only=self.read_only) as db_manager:
                rows = db_manager.get_fornitori_by_richiesta(self.request_id)
            self.combo_suppliers['values'] = [row[0] for row in rows]
        except DatabaseError as e:
            logger.error(f"Errore database in load_suppliers_for_request: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Database Error"), tr("Unable to load suppliers: {}").format(e), "error")

    def add_attachment(self):
        if self.read_only:
            show_warning(self, tr("Operation Not Allowed"), tr("You cannot add attachments to other users' RfQs."))
            return
        
        self.grab_release()
        try:
            filepath = filedialog.askopenfilename(
                title=tr("Select file to attach"),
                parent=self
            )
        finally:
            self.grab_set()
        
        if not filepath:
            return
        self._attach_from_path(filepath)
    
    def open_attachment(self):
        selected = self.sheet_attachments.get_currently_selected()
        if not selected or selected.row is None:
            SimpleMessageDialog(self, tr("Warning"), tr("Select an attachment to open."), "warning")
            return

        row_idx = selected.row
        
        if row_idx < 0 or row_idx >= len(self.attachment_ids):
            logger.error(f"Indice riga non valido: {row_idx}, totale attachment_ids: {len(self.attachment_ids)}")
            SimpleMessageDialog(self, tr("Error"), tr("Unable to identify the selected attachment. Try reloading the window."), "error")
            return
        
        attachment_id = self.attachment_ids[row_idx]
        
        try:
            with DatabaseManager(self.db_path, read_only=self.read_only) as db_manager:
                result = db_manager.get_allegato_file_data(attachment_id)
            
            if not result:
                SimpleMessageDialog(self, tr("Error"), tr("Attachment not found."), "error")
                return
                
            nome_file, dati_file, percorso_esterno = result
        except DatabaseError as e:
            logger.error(f"Errore database in open_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Database Error"), tr("Unable to retrieve attachment: {}").format(e), "error")
            return

        try:
            if percorso_esterno:
                logger.info(f"Apertura allegato esterno: {nome_file}")
                base_path = self.attachments_base
                if not base_path:
                    logger.error("Percorso archivio non configurato")
                    SimpleMessageDialog(self, tr("Error"), tr("Archive path not configured."), "error")
                    return
                
                full_path = os.path.join(base_path, percorso_esterno)
                
                real_base = os.path.realpath(base_path)
                real_full = os.path.realpath(full_path)
                
                if not real_full.startswith(real_base + os.sep) and real_full != real_base:
                    logger.error(f"Tentativo di accesso non autorizzato a: {real_full}")
                    SimpleMessageDialog(self, tr("Security Error"), tr("Invalid file path. Possible unauthorized access attempt."), "error")
                    return
                
                if not os.path.exists(real_full):
                    logger.error(f"File esterno non trovato: {real_full}")
                    SimpleMessageDialog(self, tr("Error"), tr("Source file not found:\n{}").format(real_full), "error")
                    return
                
                webbrowser.open(f'file:///{real_full}')

            elif dati_file:
                logger.info(f"Apertura allegato interno: {nome_file}")
                file_ext = os.path.splitext(nome_file)[1]
                
                with tempfile.NamedTemporaryFile(mode='wb', suffix=file_ext, delete=False) as temp_file:
                    temp_file.write(dati_file)
                    temp_path = temp_file.name
                
                self.temp_files.append(temp_path)
                
                def delayed_cleanup(path, delay=60):
                    """Elimina il file temporaneo dopo delay secondi con retry se locked."""
                    try:
                        time.sleep(delay)
                        
                        if not os.path.exists(path):
                            return
                        
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                os.remove(path)
                                logger.info(f"File temporaneo pulito automaticamente: {path}")
                                break
                            except (PermissionError, OSError) as e:
                                if attempt < max_retries - 1:
                                    logger.debug(f"File temporaneo ancora in uso, retry {attempt+1}/{max_retries}: {e}")
                                    time.sleep(5)
                                else:
                                    logger.warning(f"File temporaneo non eliminabile dopo {max_retries} tentativi (in uso?): {path}")
                    except Exception as e:
                        logger.warning(f"Impossibile pulire file temporaneo {path}: {e}")
                
                cleanup_thread = threading.Thread(
                    target=delayed_cleanup,
                    args=(temp_path,),
                    name=f"TempFileCleanup-{os.path.basename(temp_path)}",
                    daemon=True
                )
                cleanup_thread.start()
                
                webbrowser.open(f'file:///{temp_path}')
            else:
                logger.error("Allegato senza dati né percorso esterno (open)")
                SimpleMessageDialog(self, tr("Error"), tr("Attachment data not available (neither internal nor external)."), "error")

        except FileNotFoundError as e:
            logger.error(f"File non trovato in open_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Opening Error"), tr("File not found: {}").format(e), "error")
        except PermissionError as e:
            logger.error(f"Permessi insufficienti in open_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Opening Error"), tr("Insufficient permissions to open file: {}").format(e), "error")
        except OSError as e:
            logger.error(f"Errore sistema operativo in open_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Opening Error"), tr("Operating system error: {}").format(e), "error")
        except Exception as e:
            logger.error(f"Errore imprevisto in open_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Opening Error"), tr("Unable to open file: {}").format(e), "error")
    
    def download_attachment(self):
        selected = self.sheet_attachments.get_currently_selected()
        if not selected or selected.row is None:
            SimpleMessageDialog(self, tr("Warning"), tr("Select an attachment to download."), "warning")
            return

        row_idx = selected.row
        
        if row_idx < 0 or row_idx >= len(self.attachment_ids):
            logger.error(f"Indice riga non valido in download: {row_idx}, totale: {len(self.attachment_ids)}")
            SimpleMessageDialog(self, tr("Error"), tr("Unable to identify the selected attachment."), "error")
            return
        
        attachment_id = self.attachment_ids[row_idx]
        
        try:
            with DatabaseManager(self.db_path) as db_manager:
                result = db_manager.get_allegato_file_data(attachment_id)
            
            if not result:
                SimpleMessageDialog(self, tr("Error"), tr("Attachment not found."), "error")
                return
                
            nome_file, dati_file, percorso_esterno = result
        except DatabaseError as e:
            logger.error(f"Errore database in download_attachment: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Database Error"), tr("Unable to retrieve attachment: {}").format(e), "error")
            return

        self.grab_release()
        try:
            save_path = filedialog.asksaveasfilename(
                title=tr("Save attachment as..."),
                initialfile=nome_file,
                parent=self
            )
        finally:
            self.grab_set()
        
        if not save_path:
            return

        try:
            if percorso_esterno:
                base_path = self.attachments_base
                if not base_path:
                    SimpleMessageDialog(self, tr("Error"), tr("Archive path not configured."), "error")
                    return
                
                full_path = os.path.join(base_path, percorso_esterno)
                
                real_base = os.path.realpath(base_path)
                real_full = os.path.realpath(full_path)
                
                if not real_full.startswith(real_base + os.sep) and real_full != real_base:
                    SimpleMessageDialog(self, tr("Security Error"), tr("Invalid file path. Possible unauthorized access attempt."), "error")
                    return
                
                if not os.path.exists(real_full):
                    SimpleMessageDialog(self, tr("Error"), tr("Source file not found:\n{}").format(real_full), "error")
                    return
                
                shutil.copy(real_full, save_path)

            elif dati_file:
                with open(save_path, 'wb') as f:
                    f.write(dati_file)
            else:
                SimpleMessageDialog(self, tr("Error"), tr("Attachment data not available (neither internal nor external)."), "error")
                return

            SimpleMessageDialog(self, tr("Success"), tr("File downloaded successfully to:\n{}").format(save_path), "info")

        except Exception as e:
            SimpleMessageDialog(self, tr("Download Error"), tr("Unable to save file: {}").format(e), "error")
