# RIGHE 1-2 (Importazioni necessarie per il DPI)
import sys
if sys.platform == 'win32':
    from ctypes import windll
else:
    windll = None

# ---------------------------------------------
# RIGHE 3-20: BLOCCO DPI AWARENESS (DEVE ESSERE QUI)
# ---------------------------------------------
if sys.platform == 'win32':
    try:
        # Importiamo windll direttamente se non è già stata importata
        
        # Imposta PerMonitorV2
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        
        if hasattr(windll.shcore, 'SetProcessDpiAwarenessContext'):
            windll.shcore.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        elif hasattr(windll.user32, 'SetProcessDPIAware'):
            windll.user32.SetProcessDPIAware()

    except Exception as e:
        # Ignora errori se le librerie non sono presenti o la funzione non è supportata
        # BUG #21 FIX: Log warning invece di pass silenzioso per diagnostica
        import logging
        logging.getLogger(__name__).debug(f"DPI awareness non disponibile: {e}")

import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
from tksheet import Sheet, natural_sort_key
import os
from database_manager import DatabaseManager, DatabaseError
import tempfile
from tkcalendar import DateEntry
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from copy import copy
import shutil
import configparser
import re
from PIL import Image, ImageTk
import time
import math
import glob
import ast # Aggiunto per la gestione sicura delle note formattate
import json # Aggiunto per parsing sicuro delle note
import logging
from logging.handlers import RotatingFileHandler
import atexit
import gettext
import subprocess
import threading
import unicodedata

# Importa costanti UI/layout
from constants import (
    TASKBAR_BUFFER,
    BASE_ARTICLE_WIDTH,
    CONTO_LAVORO_WIDTH,
    SUPPLIER_COLUMN_WIDTH,
    PADDING,
    BUTTONS_MIN_WIDTH,
    MIN_WINDOW_WIDTH,
    SCREEN_WIDTH_PERCENTAGE,
    SCREEN_HEIGHT_PERCENTAGE
)

# Importa utility stringhe e formattazione
from utils.string_utils import generate_username
from utils.format_utils import (
    parse_float_from_comma_string,
    format_quantity_display,
    format_currency_display,
    get_currency_code,
    get_currency_excel_number_format,
)
from utils.window_utils import calculate_center_position, calculate_optimal_window_size, center_window
from utils.user_utils import get_app_data_dir, get_config_file, load_user_identity, save_user_identity
from utils.resource_utils import resource_path, set_window_icon
from utils.i18n_utils import (
    tr,
    init_i18n,
    get_current_language,
    get_pos_column_text,
    get_qty_column_text,
    normalize_rfq_type,
    translate_rfq_type
)
from utils.validation_utils import sanitize_filename, format_date_for_db, format_price_display

# !!!!! IMPORTANTE: Inizializza le traduzioni PRIMA di importare moduli UI !!!!!
# I moduli UI usano tr() durante l'import, quindi init_i18n() DEVE essere chiamato prima
init_i18n()

# Importa UI components (DOPO init_i18n per avere tr() disponibile)
from ui.kpi_window import KpiWindow
from ui.window_launchers import open_help_window, on_kpi_click, open_license_window
from ui.windows.view_request_window import ViewRequestWindow
from ui.components.main_dashboard_toolbar import MainDashboardToolbar
from ui.components.collapsible_filters import CollapsibleFilters
from ui.main_dashboard_builder import build_main_dashboard
from ui.sheet_factories import (
    create_request_sheet,
    create_vsm_event_sheet,
    create_supplier_sheet,
    create_cell_select_handler as factory_create_cell_select_handler,
    create_row_select_handler as factory_create_row_select_handler,
)
from services.dashboard_controller import DashboardController

# REFACTORING: Import moduli estratti
from services.app_paths import (
    get_user_documents_dataflow_dir,
    get_fixed_db_dir,
    get_fixed_attachments_dir,
    get_db_path,
    reset_db_cache
)
from services.startup_service import (
    cleanup_temp_on_startup,
    setup_logging,
    initialize_dataflow_directory_structure
)
from services.excel_export_service import (
    export_rfq_requests_excel,
    export_vsm_events_excel,
    export_derisking_suppliers_excel,
    load_derisking_suppliers_for_export,
)
from services.dashboard_selection_policy import (
    get_selected_row_indices as policy_get_selected_row_indices,
    check_all_selected_are_mine as policy_check_all_selected_are_mine,
)
from services.dashboard_actions_policy import (
    compute_actions_capabilities,
    build_actions_menu_spec,
)
from services.vsm_dashboard_service import (
    get_vsm_dataset as service_get_vsm_dataset,
    apply_vsm_filters as service_apply_vsm_filters,
)
from services.derisking_dashboard_service import (
    build_supplier_rows_and_metadata,
    auto_size_supplier_sheet as service_auto_size_supplier_sheet,
    populate_supplier_sheet as service_populate_supplier_sheet,
)
from services.vsm_command_service import (
    status_to_event_type,
    delete_vsm_events_by_ids,
    delete_suppliers_by_ids,
    duplicate_vsm_event_by_id,
)
from services.rfq_dashboard_service import (
    load_requests_by_status as service_load_requests_by_status,
    build_rfq_sheet_payload,
)
from services.rfq_command_service import (
    update_request_status,
    delete_requests_with_attachments,
    duplicate_request_full,
    create_request_shell,
)
from services.dashboard_search_service import (
    has_active_search_filters,
    filter_derisking_suppliers_by_query,
    split_vsm_events_by_type,
    filter_vsm_events_by_query,
)
from services.settings_preferences_service import (
    ALLOWED_CURRENCIES,
    load_settings_snapshot,
    save_language_preference,
    save_currency_preference,
    save_autobackup_preferences,
)
from services.settings_maintenance_service import (
    read_autobackup_config,
    read_last_autobackup_date,
    save_last_autobackup_date,
    copy_manual_backup_bundle,
    perform_autobackup_copy,
)
from services.dataflow_location_service import (
    normalize_parent_directory,
    ensure_parent_directory_writable,
    detect_username_conflict,
)
from services.restart_lifecycle_service import (
    resolve_restart_script_path,
    build_restart_command,
    launch_post_mainloop_restart,
)
from database.db_helpers import crea_database_v4
from ui.dialogs.common_dialogs import (
    LanguagePrompt,
    NewRdOTypeDialog,
    UserIdentityDialog,
    CopyProgressWindow,
    SplashScreen,
    SimpleYesNoDialog,
    SimpleMessageDialog,
    LicenseAcceptanceDialog,
    show_error
)

# Esegui pulizia all'avvio
cleanup_temp_on_startup()

# REFACTORING: Setup logging estratto in services.startup_service
logger = setup_logging()

# REFACTORING: Funzioni path management estratte in services.app_paths
# - get_user_documents_dataflow_dir()
# - get_fixed_db_dir()
# - get_fixed_attachments_dir()
# - initialize_dataflow_directory_structure()
# - get_db_path()
# - reset_db_cache()

# REFACTORING: Database helpers estratti in database.db_helpers
# - crea_database_v4()


# FINESTRA IMPOSTAZIONI
# ------------------------------------------------------------------------------------
class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, main_app):
        try:
            super().__init__(parent)
            self.withdraw()
            set_window_icon(self)
            
            self.main_app = main_app
            try:
                self.title(tr("Settings and Maintenance"))
            except Exception as e:
                logger.error(f"Errore nel settare il titolo: {e}")
                self.title(tr("Settings and Maintenance"))
            self.transient(parent)
            self.grab_set()
            
            self.autobackup_enabled = tk.BooleanVar()
            self.autobackup_hour = tk.StringVar()
            self.autobackup_path = tk.StringVar()
            self.language_var = tk.StringVar()
            self.currency_var = tk.StringVar(value=tr("None"))
            # Imposta un valore di default per la lingua (verrà aggiornato da load_settings)
            self.language_var.set("English")

            # Le impostazioni di visualizzazione sono ora gestite automaticamente da Windows DPI

            main_frame = ttk.Frame(self, padding="20")
            main_frame.pack(fill="both", expand=True)

            # --- Sezione Posizione DataFlow Standard ---
            dataflow_frame = ttk.LabelFrame(main_frame, text=tr("Standard DataFlow Location"), padding=10)
            dataflow_frame.pack(fill="x", pady=(0, 15), padx=5)
            
            dataflow_label = ttk.Label(
                dataflow_frame, 
                text=tr("Choose where to save the DataFlow folder (requires restart)."),
                font=(None, 10),
                wraplength=480,
                justify="left"
            )
            dataflow_label.pack(anchor="w", pady=(0, 10))
            
            ttk.Button(
                dataflow_frame, 
                text=tr("📁 Change DataFlow Location..."), 
                command=self.select_standard_dataflow_location
            ).pack()
            
            try:
                current_dataflow = get_user_documents_dataflow_dir()
                ttk.Label(
                    dataflow_frame,
                    text=tr("Current DataFlow folder: {}").format(current_dataflow),
                    font=(None, 9),
                    foreground="gray",
                    wraplength=480,
                    justify="left"
                ).pack(anchor="w", pady=(10, 0))
            except Exception as e:
                logger.error(f"Errore visualizzazione posizione DataFlow corrente: {e}")

            # --- Sezione Backup Manuale ---
            backup_frame = ttk.LabelFrame(main_frame, text=tr("Manual Backup"), padding="10")
            backup_frame.pack(fill="x", pady=(0, 15), padx=5)
            ttk.Label(backup_frame, text=tr("Create an immediate backup of the database."), font=(None, 10), wraplength=500).pack(anchor="w", pady=(0, 10))
            ttk.Button(backup_frame, text=tr("💾 Manual Backup..."), command=self.backup_database).pack()

            # --- Sezione Backup Automatico ---
            autobackup_frame = ttk.LabelFrame(main_frame, text=tr("Daily Automatic Backup"), padding="10")
            autobackup_frame.pack(fill="x", pady=(0, 15), padx=5)

            ttk.Checkbutton(autobackup_frame, text=tr("Enable daily automatic backup (max 3 copies)"), variable=self.autobackup_enabled).pack(anchor="w", pady=(0, 10))
            
            hour_frame = ttk.Frame(autobackup_frame)
            hour_frame.pack(fill="x", pady=5)
            ttk.Label(hour_frame, text=tr("Time:")).pack(side="left", padx=(0, 5))
            ttk.Combobox(hour_frame, textvariable=self.autobackup_hour, values=[f"{h:02}" for h in range(24)], width=5, state="readonly").pack(side="left")

            path_frame = ttk.Frame(autobackup_frame)
            path_frame.pack(fill="x", pady=5)
            ttk.Label(path_frame, text=tr("Save to:")).pack(anchor="w")
            
            path_entry_frame = ttk.Frame(autobackup_frame)
            path_entry_frame.pack(fill="x")
            ttk.Entry(path_entry_frame, textvariable=self.autobackup_path, state="readonly", width=50).pack(side="left", fill="x", expand=True, pady=(0, 5))
            ttk.Button(path_entry_frame, text=tr("📁 Choose..."), command=self.select_autobackup_path).pack(side="left", padx=(5,0), pady=(0,5))

            ttk.Button(autobackup_frame, text=tr("💾 Save Backup Settings"), command=self.save_autobackup_settings).pack(pady=(10,0))

            # --- Sezione Lingua e Valuta ---
            language_frame = ttk.LabelFrame(main_frame, text=tr("Lingua e Valuta"), padding="10")
            language_frame.pack(fill="x", pady=(0, 15), padx=5)
            
            ttk.Label(language_frame, text=tr("Select the interface language. The change requires restarting the application."), font=(None, 10), wraplength=500).pack(anchor="w", pady=(0, 15))
            
            # Riga per il controllo della lingua
            lang_row = ttk.Frame(language_frame)
            lang_row.pack(fill="x", pady=(0, 5))
            
            ttk.Label(lang_row, text=tr("Language:")).pack(side="left", padx=(0, 10))
            language_combo = ttk.Combobox(lang_row, textvariable=self.language_var, values=["English", "Italiano"], state="readonly", width=20)
            language_combo.pack(side="left", padx=(0, 10))
            self.language_combo = language_combo  # Salva riferimento per aggiornamento successivo

            # Riga per preferenza valuta globale
            currency_row = ttk.Frame(language_frame)
            currency_row.pack(fill="x", pady=(8, 0))
            ttk.Label(currency_row, text=tr("Currency")).pack(side="left", padx=(0, 10))
            self.currency_combo = ttk.Combobox(
                currency_row,
                textvariable=self.currency_var,
                values=[tr("None"), "EUR", "USD", "GBP", "CHF"],
                state="readonly",
                width=20,
            )
            self.currency_combo.pack(side="left", padx=(0, 10))
            ttk.Button(
                language_frame,
                text=tr("💾 Save Settings"),
                command=self.save_language_currency_settings,
            ).pack(pady=(12, 0))
            
            # Assicura che il valore nel combobox corrisponda al codice lingua
            def on_language_change(event):
                selected = self.language_var.get()
                # Il valore viene già impostato correttamente dal combobox
                pass
            language_combo.bind("<<ComboboxSelected>>", on_language_change)
            

            try:
                self.load_settings()
                # Aggiorna il combobox dopo aver caricato le impostazioni
                if hasattr(self, 'language_combo'):
                    current_val = self.language_var.get()
                    if current_val == "English":
                        self.language_combo.current(0)
                    elif current_val == "Italiano":
                        self.language_combo.current(1)
            except Exception as e:
                logger.error(f"Errore nel caricare impostazioni all'avvio di SettingsWindow: {e}", exc_info=True)
                # Continua comunque con valori di default
            
            try:
                center_window(self)
            except Exception as e:
                logger.error(f"Errore nel centrare la finestra SettingsWindow: {e}", exc_info=True)
                # Mostra comunque la finestra anche se il centraggio fallisce
                self.deiconify()
                self.geometry("800x600")
        except Exception as e:
            logger.error(f"Errore critico nell'inizializzazione di SettingsWindow: {e}", exc_info=True)
            # Mostra la finestra anche in caso di errore critico
            try:
                self.deiconify()
                self.geometry("800x600")
            except:
                pass

    def load_settings(self):
        """Carica le impostazioni dal file config.ini."""
        try:
            snapshot = load_settings_snapshot(get_config_file())
            self.autobackup_enabled.set(snapshot["autobackup_enabled"])
            self.autobackup_hour.set(snapshot["autobackup_hour"])
            self.autobackup_path.set(snapshot["autobackup_path"])
            self.language_var.set("English" if snapshot["language_code"] == "en" else "Italiano")
            self.currency_var.set(tr("None") if snapshot["currency_code"] == "NONE" else snapshot["currency_code"])
        except Exception as e:
            logger.error(f"Errore critico nel caricare impostazioni: {e}", exc_info=True)
            # Imposta valori di default in caso di errore
            self.autobackup_enabled.set(False)
            self.autobackup_hour.set("12")
            self.autobackup_path.set("")
            self.language_var.set("English")
            self.currency_var.set(tr("None"))

    # La funzione save_display_settings() è stata rimossa perché le impostazioni
    # di visualizzazione sono ora gestite automaticamente da Windows DPI

    def save_language_settings(self):
        """Salva la lingua selezionata nel config.ini."""
        try:
            selected_lang = self.language_var.get()
            if not selected_lang:
                SimpleMessageDialog(self, tr("Warning"), tr("Select a language."), "warning")
                return

            save_language_preference(get_config_file(), selected_lang)
            
            dialog = SimpleYesNoDialog(
                self,
                tr("Success"),
                tr("Language setting saved.\nRestart the application now to apply the changes?")
            )
            if dialog.result:
                # Riavvia l'applicazione
                self.main_app.restart_program()
        except Exception as e:
            logger.error(f"Errore nel salvare la lingua: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Error"), tr("Unable to save language setting: {}").format(e), "error")

    def save_currency_settings(self):
        """Salva la preferenza valuta globale nel config.ini."""
        try:
            save_currency_preference(
                get_config_file(),
                self.currency_var.get(),
                tr("None"),
            )

            dialog = SimpleYesNoDialog(
                self,
                tr("Success"),
                tr("Currency setting saved.")
                + "\n"
                + tr("The change requires restarting the application.")
                + "\n"
                + tr("Restart the application now to apply the changes?")
            )
            if dialog.result:
                self.main_app.restart_program()
        except Exception as e:
            logger.error(f"Errore nel salvare valuta: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Error"), tr("Unable to save currency setting: {}").format(e), "error")

    def save_language_currency_settings(self):
        """Salva lingua/valuta con un solo feedback finale e un solo prompt di restart."""
        try:
            selected_lang = self.language_var.get()
            if not selected_lang:
                SimpleMessageDialog(self, tr("Warning"), tr("Select a language."), "warning")
                return

            snapshot = load_settings_snapshot(get_config_file())
            selected_lang_code = "en" if selected_lang == "English" else "it"

            selected_currency_ui = (self.currency_var.get() or "").strip()
            selected_currency_code = (
                "NONE"
                if selected_currency_ui in {tr("None"), "NONE"}
                else selected_currency_ui.upper()
            )
            if selected_currency_code not in ALLOWED_CURRENCIES:
                selected_currency_code = "NONE"

            language_changed = selected_lang_code != snapshot["language_code"]
            currency_changed = selected_currency_code != snapshot["currency_code"]

            if not language_changed and not currency_changed:
                SimpleMessageDialog(self, tr("Info"), tr("No changes to save."), "info")
                return

            if language_changed:
                save_language_preference(get_config_file(), selected_lang)

            if currency_changed:
                save_currency_preference(
                    get_config_file(),
                    self.currency_var.get(),
                    tr("None"),
                )

            dialog = SimpleYesNoDialog(
                self,
                tr("Success"),
                tr("The change requires restarting the application.")
                + "\n"
                + tr("Restart the application now to apply the changes?"),
            )
            if dialog.result:
                self.main_app.restart_program()
        except Exception as e:
            logger.error(f"Errore nel salvare lingua/valuta: {e}", exc_info=True)
            SimpleMessageDialog(self, tr("Error"), tr("Unable to save: {}").format(e), "error")

    def select_autobackup_path(self):
        path = filedialog.askdirectory(title=tr("Select folder for automatic backups"), parent=self)
        if path: self.autobackup_path.set(path)

    def save_autobackup_settings(self):
        try:
            save_autobackup_preferences(
                get_config_file(),
                enabled=self.autobackup_enabled.get(),
                hour=self.autobackup_hour.get(),
                path=self.autobackup_path.get(),
            )
            SimpleMessageDialog(self, tr("Success"), tr("Backup settings saved."), "info")
        except ValueError:
            SimpleMessageDialog(self, tr("Warning"), tr("To enable automatic backup, specify a path."), "warning")
        except Exception as e:
            SimpleMessageDialog(self, tr("Error"), tr("Unable to save: {}").format(e), "error")

    def backup_database(self):
        """Crea backup manuale copiando i file del database (db, wal, shm)."""
        db_file = get_db_path()
        if not os.path.exists(db_file):
            SimpleMessageDialog(self, tr("Error"), tr("Database file '{}' not found!").format(db_file), "error")
            return
        
        dest = filedialog.asksaveasfilename(
            title=tr("Save backup as..."), 
            initialfile=f"backup_manuale_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db", 
            defaultextension=".db", 
            filetypes=[(tr("Database SQLite"), "*.db"), (tr("All files"), "*.*")], 
            parent=self
        )
        
        if not dest:
            return  # Utente ha annullato
        
        # Normalizza l'estensione del file di destinazione
        if not dest.endswith('.db'):
            dest = dest.rsplit('.', 1)[0] + '.db'
        
        # Chiudi temporaneamente la connessione della MainWindow per permettere il backup
        main_window_was_open = False
        try:
            if hasattr(self.main_app, 'db_manager') and self.main_app.db_manager:
                logger.info("Chiusura connessione MainWindow per backup...")
                self.main_app.db_manager.close()
                main_window_was_open = True
                # Piccolo delay per assicurarsi che la connessione sia completamente chiusa
                import time
                time.sleep(0.2)
        except Exception as e:
            logger.warning(f"Impossibile chiudere connessione MainWindow: {e}")
        
        try:
            copy_result = copy_manual_backup_bundle(
                db_file=db_file,
                dest=dest,
                logger=logger,
            )
            original_size = copy_result["original_size"]
            backup_size = copy_result["backup_size"]
            copied_paths = copy_result["copied_files"]
            
            if backup_size < original_size * 0.5:
                logger.warning(f"Backup manuale potenzialmente incompleto: {backup_size} vs {original_size} bytes")
                dialog = SimpleYesNoDialog(
                    self,
                    tr("Size Warning"), 
                    tr("The created backup is significantly smaller than the original database.\n\nOriginal: {:.2f} MB\nBackup: {:.2f} MB\n\nDo you want to keep it anyway?").format(original_size / (1024*1024), backup_size / (1024*1024))
                )
                if not dialog.result:
                    try:
                        for file_path in copied_paths:
                            if os.path.exists(file_path):
                                os.remove(file_path)
                    except:
                        pass
                    return
            
            # Messaggio di successo con info sui file copiati
            files_copied = [os.path.basename(path) for path in copied_paths if os.path.exists(path)]
            
            SimpleMessageDialog(
                self,
                tr("Success"), 
                tr("Backup created successfully:\n\nFiles copied:\n{}\n\nTotal size: {:.2f} MB").format(
                    '\n'.join(f'  • {f}' for f in files_copied),
                    copy_result["total_size"] / (1024 * 1024)
                ),
                "info"
            )
            logger.info(f"Backup manuale completato: {len(files_copied)} file copiati")
            
        except Exception as e:
            logger.error(f"Errore backup manuale: {e}", exc_info=True)
            SimpleMessageDialog(
                self,
                tr("Error"), 
                tr("Unable to create backup:\n{}").format(e),
                "error"
            )
            # Rimuovi backup parziale/corrotto
            if os.path.exists(dest):
                try:
                    for file_path in [dest, dest.replace('.db', '.db-wal'), dest.replace('.db', '.db-shm')]:
                        if os.path.exists(file_path):
                            os.remove(file_path)
                except:
                    pass
        finally:
            # Riapri la connessione della MainWindow se era aperta
            if main_window_was_open:
                try:
                    logger.info("Riapertura connessione MainWindow dopo backup...")
                    self.main_app.db_manager = DatabaseManager(get_db_path())
                    logger.info("Connessione MainWindow riaperta con successo")
                except Exception as e:
                    logger.error(f"Errore nella riapertura connessione MainWindow: {e}")
                    SimpleMessageDialog(
                        self,
                        tr("Warning"),
                        tr("The backup has been completed, but it was not possible to reopen the main connection.\nIt is recommended to restart the application."),
                        "warning"
                    )

    def select_standard_dataflow_location(self):
        """
        Permette all'utente di scegliere una nuova posizione per la cartella DataFlow.
        
        Passaggi:
        1. Avviso esplicativo con conferma
        2. Selezione cartella
        3. Validazioni (permessi, rete, lunghezza path, unità)
        4. Salvataggio config
        5. Istruzioni per spostare manualmente la cartella
        6. Riavvio applicazione
        """
        logger.info("Avvio procedura cambio posizione cartella DataFlow")
        current_dataflow_dir = get_user_documents_dataflow_dir()
        
        warning_text = tr(
            "⚠️ WARNING: you are about to change the DataFlow folder location.\n\nThe current folder will be automatically copied to the new selected location, including the database and attachments. This operation may take a moment.\n\nThe application will restart when it is complete.\n\nCurrent location:\n{}\n\nDo you want to proceed?"
        ).format(current_dataflow_dir)
        
        dialog = SimpleYesNoDialog(
            self,
            tr("Confirm Location Change"), 
            warning_text,
            icon='warning'
        )
        if not dialog.result:
            logger.info("Utente ha annullato il cambio posizione DataFlow")
            return
        
        if sys.platform == 'win32':
            initial_dir = os.path.dirname(current_dataflow_dir) or os.path.join(os.path.expanduser('~'), 'Documents')
        else:
            initial_dir = os.path.dirname(current_dataflow_dir) or os.path.expanduser('~')
        
        try:
            selected_dir = filedialog.askdirectory(
                title=tr("Select the new DataFlow folder location"),
                initialdir=initial_dir,
                parent=self
            )
        except Exception as e:
            logger.error(f"Errore apertura dialog selezione cartella: {e}")
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Error selecting folder: {}").format(e),
                "error"
            )
            return
        
        if not selected_dir:
            logger.info("Utente ha annullato la selezione della nuova posizione")
            return
        
        normalized_dir = normalize_parent_directory(selected_dir)
        if not normalized_dir:
            SimpleMessageDialog(self, tr("Error"), tr("Invalid path."), "error")
            return
        
        # ✅ CORREZIONE: NON aggiungere "DataFlow" - useremo DataFlow_{username}
        # Il percorso selezionato dall'utente è la directory PARENT dove verrà creata DataFlow_{username}
        logger.info(f"Cartella parent selezionata per DataFlow: {normalized_dir}")
        
        try:
            ensure_parent_directory_writable(normalized_dir)
            logger.info(f"Permessi verifica OK per {normalized_dir}")
        except (OSError, PermissionError) as e:
            logger.error(f"Test permessi fallito per {normalized_dir}: {e}")
            SimpleMessageDialog(
                self,
                tr("Permission Error"),
                tr("Cannot write to the selected folder:\n{}\n\nDetails: {}").format(normalized_dir, e),
                "error"
            )
            return
        
        # Controllo lunghezza
        if len(normalized_dir) > 240:
            logger.warning(f"Percorso DataFlow troppo lungo ({len(normalized_dir)} caratteri)")
            length_warning = tr(
                "The selected path is very long ({} characters).\nWindows may have issues accessing files.\nDo you want to continue anyway?"
            ).format(len(normalized_dir))
            dialog = SimpleYesNoDialog(
                self,
                tr("Path Too Long"),
                length_warning
            )
            if not dialog.result:
                logger.info("Utente ha annullato dopo avviso percorso lungo")
                return
        
        # Controllo unità rimovibile
        try:
            drive_letter = os.path.splitdrive(normalized_dir)[0]
            if drive_letter and drive_letter.upper() not in ['C:', 'D:', 'E:']:
                logger.warning(f"Unità potenzialmente rimovibile: {drive_letter}")
                removable_warning = tr(
                    "⚠️ The selected drive ({}) might be removable.\nIf disconnected, DataFlow will not be able to access the data."
                ).format(drive_letter)
                SimpleMessageDialog(self, tr("Removable Drive?"), removable_warning, "warning")
        except Exception as e:
            logger.error(f"Errore durante controllo unità rimovibile: {e}")
        
        # === INIZIO LOGICA CONTROLLO CONFLITTO USERNAME ===
        # Carica identità utente corrente
        identity = load_user_identity()
        current_username = identity.get('username', '').strip().lower()
        
        if not current_username:
            logger.error("Username corrente non trovato nel config")
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Unable to determine the current user. Restart DataFlow."),
                "error"
            )
            return
        
        # Variabili per gestione cambio username
        final_username = current_username
        username_changed = False
        
        # Loop controllo conflitto username
        while True:
            # Controlla se esiste già un database con questo username nella destinazione
            conflict_info = detect_username_conflict(
                parent_dir=normalized_dir,
                username=final_username,
                logger=logger,
            )
            folder_exists = conflict_info["folder_exists"]
            db_exists = conflict_info["db_exists"]
            logger.info(f"Controllo conflitto per username '{final_username}': folder={folder_exists}, db={db_exists}")
            
            # ✅ CORREZIONE LOGICA: Se ESISTE cartella O database, è un CONFLITTO
            if folder_exists or db_exists:
                # Conflitto rilevato: chiedi se vuole cambiare username
                conflict_message = tr(
                    "⚠️ USER CONFLICT DETECTED\n\nA database associated with user '{}' already exists \nin the selected destination folder.\n\nTo avoid conflicts and data loss, you need to change \nyour username before proceeding.\n\nDo you want to proceed with the username change?"
                ).format(final_username)
                
                dialog = SimpleYesNoDialog(
                    self,
                    tr("Username Conflict"),
                    conflict_message,
                    icon='warning'
                )
                if not dialog.result:
                    # Utente ha rifiutato, annulla tutto
                    logger.info("Utente ha rifiutato il cambio username, operazione annullata")
                    return
                
                # Mostra dialogo cambio identità
                self.withdraw()  # Nascondi finestra settings temporaneamente
                new_identity_dialog = UserIdentityDialog(self)
                self.wait_window(new_identity_dialog)
                self.deiconify()  # Mostra di nuovo
                
                new_identity = getattr(new_identity_dialog, 'result', None)
                if not new_identity:
                    # Utente ha annullato il dialogo identità
                    logger.info("Utente ha annullato il dialogo identità, operazione annullata")
                    return
                
                # Aggiorna username e continua il loop per ricontrollare
                final_username = new_identity['username']
                username_changed = True
                logger.info(f"Nuovo username proposto: {final_username}, rientro nel loop controllo")
            else:
                # ✅ NESSUN CONFLITTO: Username libero, prosegui
                logger.info(f"Username '{final_username}' disponibile nella destinazione (nessun conflitto rilevato)")
                break
        
        # === FINE LOGICA CONTROLLO CONFLITTO USERNAME ===
        
        # A questo punto final_username è libero, procedi con la copia
        source_folder = current_dataflow_dir
        dest_parent = normalized_dir  # Directory parent dove creare DataFlow_{username}
        dest_folder = os.path.join(dest_parent, f"DataFlow_{final_username}")  # Percorso completo destinazione
        
        # Verifica che la cartella sorgente esista
        if not os.path.exists(source_folder):
            logger.error(f"Cartella sorgente non esiste: {source_folder}")
            SimpleMessageDialog(
                self,
                tr("Error"),
                tr("Source DataFlow folder not found:\n{}").format(source_folder),
                "error"
            )
            return
        
        # ✅ CHIUDI DATABASE PRIMA DELLA COPIA (evita WinError 32)
        logger.info("Chiusura database prima della copia...")
        try:
            # Chiudi il DatabaseManager globale se esiste
            if hasattr(self.main_app, 'db_manager') and self.main_app.db_manager:
                self.main_app.db_manager.close()
                logger.info("DatabaseManager principale chiuso")
        except Exception as e:
            logger.warning(f"Errore chiusura DatabaseManager: {e}")
        
        # Mostra finestra progresso copia
        progress_win = CopyProgressWindow(self, title=tr("Copying DataFlow..."))
        progress_win.update_progress(0, tr("Preparing copy..."))
        
        # Backup config originale (per rollback)
        config_backup = None
        try:
            config_file = get_config_file()
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_backup = f.read()
        except Exception as e:
            logger.error(f"Impossibile fare backup config: {e}")
        
        try:
            # === COPIA FISICA COMPLETA CON PROGRESSIONE ===
            logger.info(f"Inizio copia da '{source_folder}' a '{dest_folder}'")
            
            # Conta file totali per barra progresso
            progress_win.update_progress(5, tr("Analyzing files to copy..."))
            total_files = 0
            for root, dirs, files in os.walk(source_folder):
                total_files += len(files)
            
            logger.info(f"File totali da copiare: {total_files}")
            
            if total_files == 0:
                raise Exception(tr("No files to copy in source folder"))
            
            # Copia ricorsiva con aggiornamento progressione
            files_copied = 0
            
            def copy_with_progress(src, dst):
                nonlocal files_copied
                os.makedirs(dst, exist_ok=True)
                
                for item in os.listdir(src):
                    s = os.path.join(src, item)
                    d = os.path.join(dst, item)
                    
                    if os.path.isdir(s):
                        copy_with_progress(s, d)
                    else:
                        # Copia file
                        shutil.copy2(s, d)
                        files_copied += 1
                        
                        # Aggiorna progressione (da 10% a 80%)
                        progress_pct = 10 + int((files_copied / total_files) * 70)
                        file_name = os.path.basename(s)
                        progress_win.update_progress(
                            progress_pct,
                            tr("Copying file {}/{}: {}").format(files_copied, total_files, file_name[:40])
                        )
            
            copy_with_progress(source_folder, dest_folder)
            
            logger.info(f"Copia file completata: {files_copied} file copiati")
            progress_win.update_progress(85, tr("Copy completed, updating configuration..."))
            
            # === AGGIORNA USERNAME NEL DATABASE (SOLO SE CAMBIATO) ===
            if username_changed:
                logger.info(f"Username cambiato da '{current_username}' a '{final_username}', aggiorno database")
                progress_win.update_progress(90, tr("Updating username in database..."))
                
                # Percorso nuovo database
                new_db_path = os.path.join(dest_folder, 'Database', f'dataflow_db_{final_username}.db')
                
                # Rinomina anche il file database se necessario
                old_db_name = f'dataflow_db_{current_username}.db'
                old_db_path = os.path.join(dest_folder, 'Database', old_db_name)
                
                if os.path.exists(old_db_path) and old_db_path != new_db_path:
                    logger.info(f"Rinomino database da '{old_db_name}' a 'dataflow_db_{final_username}.db'")
                    shutil.move(old_db_path, new_db_path)
                
                # Aggiorna username in tutte le RdO
                try:
                    # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
                    with DatabaseManager(new_db_path) as db_manager:
                        rows_updated = db_manager.update_all_usernames(final_username)
                    logger.info(f"Username aggiornato in {rows_updated} RdO")
                except Exception as db_error:
                    logger.error(f"Errore aggiornamento username in DB: {db_error}", exc_info=True)
                    raise
            
            # === AGGIORNA CONFIG.INI ===
            progress_win.update_progress(95, tr("Saving configuration..."))
            
            config = configparser.ConfigParser(interpolation=None)
            config_file = get_config_file()
            
            if os.path.exists(config_file):
                config.read(config_file)
            
            if 'Settings' not in config:
                config['Settings'] = {}
            if 'User' not in config:
                config['User'] = {}
            
            # Salva nuovo percorso base
            config['Settings']['dataflow_base_dir'] = dest_parent
            
            # Rimuovi legacy custom_db_path se presente
            if config.has_option('Settings', 'custom_db_path'):
                config.remove_option('Settings', 'custom_db_path')
            
            # Se username è cambiato, aggiorna anche sezione User
            if username_changed:
                config['User']['first_name'] = new_identity['first_name']
                config['User']['last_name'] = new_identity['last_name']
                config['User']['username'] = final_username
            
            with open(config_file, 'w', encoding='utf-8') as f:
                config.write(f)
            
            logger.info(f"Config aggiornato con nuovo percorso: {dest_parent}")
            
            progress_win.update_progress(100, tr("Operation completed!"))
            time.sleep(0.5)
            progress_win.destroy()
            
            # === MESSAGGIO SUCCESSO ===
            username_info = ""
            if username_changed:
                username_info = tr("\n\n✓ Username updated from '{}' to '{}'").format(current_username, final_username)
            
            success_msg = tr(
                "✓ OPERATION COMPLETED SUCCESSFULLY\n\nThe DataFlow folder has been successfully copied to:\n{dest}\n\nFiles copied: {count}{username_change}\n\n⚠️ IMPORTANT:\n- The ORIGINAL folder in '{src}' has NOT been deleted.\n- Before deleting it manually, TEST the correct operation \n  of the copied database.\n- DataFlow will restart automatically."
            ).format(
                dest=dest_folder,
                count=files_copied,
                username_change=username_info,
                src=source_folder
            )
            
            SimpleMessageDialog(self, tr("Operation Completed"), success_msg, "info")
            
            # ✅ SALVA ESPLICITAMENTE LA NUOVA IDENTITÀ (se cambiata)
            if username_changed:
                save_user_identity(new_identity['first_name'], new_identity['last_name'], final_username)
                logger.info(f"Identità salvata nel config: {final_username}")
            
            # Invalida cache e riavvia
            reset_db_cache()
            logger.info("Cache DB invalidata, riavvio applicazione")
            self.destroy()
            self.main_app.restart_program()
            
        except Exception as e:
            # === GESTIONE ERRORE CON ROLLBACK ===
            logger.error(f"Errore durante copia DataFlow: {e}", exc_info=True)
            
            try:
                progress_win.destroy()
            except:
                pass
            
            # Ripristina config backup se disponibile
            if config_backup:
                try:
                    with open(get_config_file(), 'w', encoding='utf-8') as f:
                        f.write(config_backup)
                    logger.info("Config.ini ripristinato da backup")
                except Exception as restore_err:
                    logger.error(f"Impossibile ripristinare config: {restore_err}")
            
            # Tenta di eliminare cartella parziale (se creata)
            if os.path.exists(dest_folder):
                try:
                    shutil.rmtree(dest_folder, ignore_errors=True)
                    logger.info(f"Cartella parziale eliminata: {dest_folder}")
                except Exception as cleanup_err:
                    logger.error(f"Impossibile eliminare cartella parziale: {cleanup_err}")
            
            error_msg = tr(
                "❌ OPERATION FAILED\n\nUnable to complete moving the DataFlow folder.\n\nError details:\n{error}\n\nOriginal settings have been restored.\nSee the log file for more details."
            ).format(error=str(e))
            
            SimpleMessageDialog(self, tr("Move Error"), error_msg, "error")

# ------------------------------------------------------------------------------------
# FINESTRA PRINCIPALE
# ------------------------------------------------------------------------------------
class MainWindow:
    def __init__(self, root):
        self.root = root;
        set_window_icon(self.root)
        self.root.title(tr("DataFlow Procurement Software - Main Dashboard"))
        
        # Windows: la normal size di default è troppo piccola; imposta base minima
        # esplicita prima del maximize per evitare restore/normal size minuscola.
        if sys.platform == 'win32':
            self.root.geometry("1200x768")
            self.root.minsize(1000, 700)

        # Avvia finestra massimizzata (compatibilità Linux/Windows)
        try:
            # Linux/X11: usa attributes -zoomed
            self.root.attributes("-zoomed", True)
        except:
            try:
                # Windows: usa state zoomed
                self.root.state("zoomed")
            except:
                # Fallback: se nessuno dei due funziona, ignora
                pass
        
        self.all_users_placeholder = tr("All users")
        self.username_filter_var = None
        self.user_filter_combo = None
        self.vsm_username_filter_var = None
        self.vsm_user_filter_combos = []
        self._load_identity_from_config()
        self.last_backup_date = None; self.db_path_standard = self.get_standard_db_path()
        
        # BUG #45 FIX: Inizializza ID del timer autobackup per permettere cancellazione
        self._autobackup_timer_id = None
        
        # BUG #48 FIX: Inizializza ID del timer SQL warning per permettere cancellazione
        self._sql_warning_after_id = None
        
        # BUG #50 FIX: Inizializza flag debounce per doppio click apertura RdO
        self._opening_request = False
        
        # Inizializza il database manager con il percorso completo del database
        self.db_manager = DatabaseManager(get_db_path())
        
        # --- INIZIO MODIFICA: Rilevamento DB Temporaneo (escludendo DB personalizzato) ---
        self.active_db_path = get_db_path()
        
        # Determina il percorso "di default" considerando eventuali impostazioni legacy
        # Il warning deve apparire solo se il database è DIVERSO da quello di default
        # RIMOSSO: logica e visualizzazione warning DB provvisorio
        
        # --- Costruzione UI dashboard ---
        build_main_dashboard(self)

        # --- Controller di orchestrazione dashboard ---
        self.dashboard_controller = DashboardController(self)

        # Step 4C: Caricamento iniziale dati VSM (runtime data init, non UI construction)
        # Popola ogni sheet con i dati correnti dell'utente
        for event_type, sheet in [
            ("Saving", self.sheet_saving),
            ("Cost Avoidance", self.sheet_cost_avoidance),
        ]:
            self._load_vsm_events(event_type, sheet)
        # Derisking: usa il nuovo backend PotentialSupplier (separato da VSMEvent)
        self._load_potential_suppliers(self.sheet_derisking)
        self.populate_vsm_username_filter()

        self.refresh_data(); self.update_button_visibility(); self.check_for_autobackup()

    # --- INIZIO NUOVI METODI LICENZA ---
    def open_license_window(self):
        open_license_window(self)

    # --- FINE NUOVI METODI LICENZA ---

    def _load_identity_from_config(self):
        identity = load_user_identity()
        self.current_first_name = identity.get('first_name', '')
        self.current_last_name = identity.get('last_name', '')
        self.current_username = identity.get('username', '')
        self.current_full_name = identity.get('full_name', '').strip()

    def ensure_user_identity(self, force_prompt=False):
        """Garantisce che nome, cognome e username siano impostati."""
        self._load_identity_from_config()
        needs_prompt = force_prompt or not self.current_first_name or not self.current_last_name or not self.current_username
        if not needs_prompt:
            return True
        
        identity_kwargs = {
            'first_name': self.current_first_name,
            'last_name': self.current_last_name
        }
        
        while True:
            dialog = UserIdentityDialog(self.root, **identity_kwargs)
            self.root.wait_window(dialog)
            result = getattr(dialog, 'result', None)
            if not result:
                if not self.root.winfo_exists():
                    return False
                SimpleMessageDialog(self.root, tr("Missing Data"), tr("To use DataFlow, you must enter your first and last name."), "warning")
                continue
            try:
                save_user_identity(result['first_name'], result['last_name'], result['username'])
                self._load_identity_from_config()
                self.apply_user_identity_to_ui()
                return True
            except Exception as e:
                logger.error(f"Errore salvataggio identità utente: {e}", exc_info=True)
                SimpleMessageDialog(self.root, tr("Error"), tr("Unable to save user data: {}").format(e), "error")
                identity_kwargs = result

    def apply_user_identity_to_ui(self):
        """Applica lo username corrente all'interfaccia (filtri e nuovi inserimenti)."""
        if not self.username_filter_var:
            return
        value = self.current_username if self.current_username else self.all_users_placeholder
        self.username_filter_var.set(value)
        self.refresh_data()

    def populate_username_filter(self):
        self.dashboard_controller.populate_username_filter()

    def _get_active_username_filter(self, var=None):
        """Ritorna lo username filtro attivo oppure None (= tutti gli utenti).

        Accetta un parametro 'var' opzionale per riuso diretto da moduli separati
        (es. VSM) senza duplicare la logica del placeholder e del lower-case.
        Se 'var' non è fornito, usa self.username_filter_var (comportamento RFQ originale).
        """
        target_var = var if var is not None else self.username_filter_var
        if not target_var:
            return None
        value = target_var.get().strip()
        if not value or value == self.all_users_placeholder:
            return None
        return value.lower()

    def populate_vsm_username_filter(self):
        """Aggiorna la lista degli username disponibili nel filtro utente VSM.

        Nota: duplicazione consapevole di populate_username_filter().
        La fonte dati è diversa (vsm_events vs richieste_offerta) e il
        percorso di aggregazione è separato. Un helper comune non è
        giustificato a questo stadio.
        """
        if not self.vsm_user_filter_combos or not self.vsm_username_filter_var:
            return

        usernames = []

        try:
            with DatabaseManager(get_db_path()) as db_manager:
                all_data = db_manager.get_all_vsm_events_aggregated(get_db_path())
            usernames = list({
                ev.username.strip().lower()
                for ev, _im, _src in all_data
                if ev.username and ev.username.strip()
            })
        except Exception as e:
            logger.warning(f"[populate_vsm_username_filter] Fallback al DB locale: {e}")
            try:
                with DatabaseManager(get_db_path()) as db_manager:
                    local_events = db_manager.get_all_vsm_events()
                usernames = list({
                    ev.username.strip().lower()
                    for ev in local_events
                    if ev.username and ev.username.strip()
                })
            except Exception as e2:
                logger.error(f"[populate_vsm_username_filter] Errore nel fallback: {e2}")

        # Assicura che l'utente corrente sia sempre nella lista
        if self.current_username and self.current_username.lower() not in usernames:
            usernames.append(self.current_username.lower())

        clean_usernames = sorted({u for u in usernames if u})
        values = [self.all_users_placeholder] + clean_usernames
        current_value = self.vsm_username_filter_var.get()

        for combo in self.vsm_user_filter_combos:
            try:
                combo.config(values=values)
            except Exception:
                pass  # Widget potrebbe essere già distrutto

        if current_value not in values:
            self.vsm_username_filter_var.set(self.current_username or self.all_users_placeholder)

    def _on_vsm_username_filter_changed(self):
        """Handler per cambio filtro utente VSM. Ricarica tutti i tab VSM."""
        for event_type, sheet in [
            ("Saving", self.sheet_saving),
            ("Cost Avoidance", self.sheet_cost_avoidance),
        ]:
            self._load_vsm_events(event_type, sheet)
        # Derisking ha un backend separato (PotentialSupplier)
        self._load_potential_suppliers(self.sheet_derisking)

    def _has_active_search_filters(self):
        """Verifica se ci sono filtri di ricerca attivi (escludendo username e stato)"""
        return has_active_search_filters(
            search_values={k: v.get() for k, v in self.search_vars.items()},
            search_tipo_value=self.search_tipo.get(),
            all_label=tr("All"),
            date_values={k: v.get() for k, v in self.date_entries.items()},
        )

    def _assign_request_to_current_user(self, request_id):
        """Associa una RdO all'utente corrente."""
        if not self.current_username:
            return
        try:
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(get_db_path()) as db_manager:
                db_manager.update_request_username(request_id, self.current_username)
        except DatabaseError as e:
            logger.error(f"Impossibile assegnare la RdO {request_id} all'utente {self.current_username}: {e}", exc_info=True)

    def get_standard_db_path(self):
        """Restituisce il percorso standard del database (estensione .db)"""
        return os.path.join(get_fixed_db_dir(), 'gestione_offerte.db')

    def restart_program(self):
        """Riavvia l'applicazione con le nuove impostazioni."""
        python = sys.executable
        is_frozen = hasattr(sys, "_MEIPASS")
        script_path = resolve_restart_script_path(
            file_value=globals().get("__file__"),
            argv0=sys.argv[0] if sys.argv else "",
            cwd=os.getcwd(),
            executable=sys.executable,
            is_frozen=is_frozen,
        )
        
        # Riavvia l'applicazione usando subprocess invece di os.execl
        # Questo gestisce correttamente i percorsi con spazi
        try:
            cmd = build_restart_command(
                python_executable=python,
                script_path=script_path,
                is_frozen=is_frozen,
            )
            
            # Imposta la working directory
            if os.path.dirname(script_path):
                cwd = os.path.dirname(script_path)
            else:
                cwd = os.getcwd()
            
            # Funzione per chiudere tutto: imposta il flag di riavvio e chiude la GUI.
            # Il nuovo processo viene lanciato DOPO che mainloop() è tornato.
            def do_restart():
                global _pending_restart
                try:
                    # Invalida la cache del DB prima del riavvio
                    reset_db_cache()
                    
                    # Memorizza il comando da lanciare dopo la chiusura della GUI
                    _pending_restart = (cmd, cwd)
                    
                    # Chiudi tutte le finestre Tkinter in modo ordinato
                    if hasattr(self, 'root') and self.root:
                        try:
                            # Distruggi tutte le finestre Toplevel
                            for widget in self.root.winfo_children():
                                if isinstance(widget, tk.Toplevel):
                                    try:
                                        widget.destroy()
                                    except:
                                        pass
                            # Esci dal mainloop (root.mainloop() tornerà)
                            self.root.quit()
                        except:
                            pass
                    
                except Exception as e:
                    logger.error(f"Errore nel riavvio dell'applicazione: {e}")
                    try:
                        show_error(
                            tr("Error"),
                            tr("Unable to restart the application automatically.\n\nPlease close and reopen the application manually to apply the changes.\n\nAttempted path: {}").format(script_path),
                            parent=None
                        )
                    except:
                        pass
            
            # Chiudi la finestra di dialogo corrente se esiste
            if hasattr(self, 'master') and self.master:
                try:
                    self.master.destroy()
                except:
                    pass
            
            # Esegui il riavvio dopo un breve delay per permettere la chiusura della finestra corrente
            if hasattr(self, 'root') and self.root:
                self.root.after(100, do_restart)
            else:
                # Se root non è disponibile, esegui immediatamente
                do_restart()
            
        except Exception as e:
            # Se il riavvio fallisce, mostra un messaggio all'utente
            logger.error(f"Errore nel riavvio dell'applicazione: {e}")
            show_error(
                tr("Error"),
                tr("Unable to restart the application automatically.\n\nPlease close and reopen the application manually to apply the changes.\n\nAttempted path: {}").format(script_path),
                parent=self.root if hasattr(self, 'root') else None
            )

    def check_for_autobackup(self):
        config_file = get_config_file()
        enabled, path, hour = read_autobackup_config(config_file)
        if enabled and path and hour:
            try:
                now = datetime.now()
                persisted_last_run_date = read_last_autobackup_date(config_file)
                self.last_backup_date = persisted_last_run_date
                if now.hour == int(hour) and now.date() != persisted_last_run_date:
                    if self.perform_autobackup(path):
                        self.last_backup_date = now.date()
                        save_last_autobackup_date(config_file, now.date())
            except Exception as e:
                print(f"ERRORE AUTOBACKUP: {e}")
        
        # BUG #45 FIX: Cancella timer precedente prima di ri-registrarlo (previene memory leak)
        if self._autobackup_timer_id is not None:
            try:
                self.root.after_cancel(self._autobackup_timer_id)
            except Exception as e:
                logger.warning(f"Impossibile cancellare timer autobackup precedente: {e}")
        
        # Ri-registra timer e salva ID per cancellazione futura
        self._autobackup_timer_id = self.root.after(60000, self.check_for_autobackup)

    def perform_autobackup(self, dest_folder):
        """Esegue backup automatico copiando direttamente il database.
        
        BUG #8 FIX: Aggiunta sincronizzazione e retry logic per evitare race condition.
        """
        logger.info(f"Avvio backup automatico in: {dest_folder}")
        db_file = get_db_path()
        
        # Verifica che non ci sia un backup già in corso
        if hasattr(self, '_backup_in_progress') and self._backup_in_progress:
            logger.warning("Backup già in corso, saltato")
            return False
        
        self._backup_in_progress = True
        
        try:
            result = perform_autobackup_copy(
                db_file=db_file,
                dest_folder=dest_folder,
                logger=logger,
            )
            if not result.get("copied"):
                return False

            files_copied = len(result["copied_files"])
            total_size = result["total_size"]
            original_size = result["original_size"]
            logger.info(
                "Backup automatico completato: %d file copiati, %d bytes totali (%.1f%% dimensione originale)",
                files_copied,
                total_size,
                (total_size / original_size * 100) if original_size else 0.0,
            )
            return True
            
        except Exception as e:
            logger.error(f"Errore backup automatico: {e}", exc_info=True)
            print(f"ERRORE AUTOBACKUP: {e}")
            return False
        finally:
            self._backup_in_progress = False

    def open_help_window(self): open_help_window(self)
    def open_settings_window(self): self.root.wait_window(SettingsWindow(self.root, self))

    def on_kpi_click(self): on_kpi_click(self)
    
    def create_request_treeview(self, parent):
        return create_request_sheet(
            parent=parent,
            on_cell_select=self.create_cell_select_handler(None),
            on_row_select=self.create_row_select_handler(None),
            on_double_click_cb=self.on_sheet_double_click,
        )
    
    def create_cell_select_handler(self, sheet):
        """Crea un handler cell_select che aggiorna lo stato azioni."""
        return factory_create_cell_select_handler(self.update_button_visibility)
    
    def create_row_select_handler(self, sheet):
        """Crea un handler row_select che aggiorna lo stato azioni."""
        return factory_create_row_select_handler(self.update_button_visibility)

    def _create_vsm_event_sheet(self, parent, event_type=None):
        return create_vsm_event_sheet(
            parent=parent,
            event_type=event_type,
            on_cell_select=self.create_cell_select_handler(None),
            on_row_select=self.create_row_select_handler(None),
            on_double_click_cb=self._on_vsm_sheet_double_click,
        )

    # NOTA: I metodi sort_treeview_column e update_sort_indicators sono stati rimossi
    # perché tksheet ha funzionalità di ordinamento integrate che si abilitano automaticamente
    # con enable_bindings(). L'utente può cliccare sugli header delle colonne per ordinare.

    def _create_supplier_sheet(self, parent):
        return create_supplier_sheet(
            parent=parent,
            on_cell_select=self.create_cell_select_handler(None),
            on_row_select=self.create_row_select_handler(None),
            on_double_click_cb=self._on_supplier_sheet_double_click,
        )

    def _on_supplier_sheet_double_click(self, sheet, event=None):
        """
        Handler doppio click sul tab Derisking (fornitori potenziali).

        Apre PotentialSupplierDialog in modalità edit.
        Se il fornitore non appartiene all'utente corrente, apre in read_only.
        Pattern identico a _on_vsm_sheet_double_click (debounce incluso).
        """
        # Debounce: evita aperture multiple rapide
        if hasattr(self, '_opening_supplier_edit') and self._opening_supplier_edit:
            return

        selected_rows = self._get_selected_row_indices(sheet)
        if not selected_rows:
            return  # Silent return — nessuna riga selezionata

        row_idx = selected_rows[0]

        if not hasattr(sheet, '_supplier_metadata') or row_idx >= len(sheet._supplier_metadata):
            return  # Metadata non disponibile o indice fuori range

        metadata = sheet._supplier_metadata[row_idx]
        supplier_id = metadata.get('supplier_id')
        is_mine = metadata.get('is_mine', False)

        if not supplier_id:
            return

        self._opening_supplier_edit = True
        try:
            from ui.dialogs.potential_supplier_dialog import PotentialSupplierDialog
            dlg = PotentialSupplierDialog(
                self.root,
                self.current_username,
                supplier_id=supplier_id,
                read_only=not is_mine,
                refresh_derisking_cb=lambda: self._load_potential_suppliers(self.sheet_derisking),
            )
            self.root.wait_window(dlg)
            if dlg.result:
                self._load_potential_suppliers(sheet)
        except Exception as e:
            logger.error("Errore apertura dialog fornitore: %s", e, exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to open the form: {}").format(e), "error")
        finally:
            self.root.after(300, lambda: setattr(self, '_opening_supplier_edit', False))

    def _load_vsm_events(self, event_type, sheet):
        """
        Carica eventi VSM per un tipo specifico.
        
        ESTRATTO da VSMManagementWindow.refresh_events() (Step 4C).
        
        Args:
            event_type: Tipo evento ("Saving"|"Cost Avoidance"|"Derisking")
            sheet: Widget tksheet da popolare
        """
        try:
            vsm_username_filter = self._get_active_username_filter(self.vsm_username_filter_var)
            all_events, extra_meta = self._get_vsm_dataset(vsm_username_filter)

            # Filtra per event_type preservando corrispondenza indice con extra_meta
            if extra_meta is not None:
                pairs = [(ev, m) for ev, m in zip(all_events, extra_meta) if ev.event_type == event_type]
                filtered_events = [p[0] for p in pairs]
                filtered_meta = [p[1] for p in pairs]
            else:
                filtered_events = [e for e in all_events if e.event_type == event_type]
                filtered_meta = None

            # Applica filtri VSM avanzati (data, azione, ripetitivo, importi)
            filtered_events, filtered_meta = self._apply_vsm_filters(
                filtered_events, event_type, extra_meta=filtered_meta
            )

            # Popola sheet
            self._populate_vsm_sheet(sheet, filtered_events, event_type=event_type, extra_metadata=filtered_meta)
            
            logger.debug(f"Caricati {len(filtered_events)} eventi VSM {event_type}")
            
        except DatabaseError as e:
            logger.error(f"Errore caricamento eventi VSM {event_type}: {e}")
            SimpleMessageDialog(self.root, tr("Database Error"), tr("Unable to load VSM events: {}\n").format(e), "error")

    def _load_potential_suppliers(self, sheet):
        """
        Carica i fornitori potenziali dal DB e popola il tab Derisking.

        Separato da _load_vsm_events perché usa PotentialSupplier, non VSMEvent.
        Rispetta il filtro utente condiviso vsm_username_filter_var.
        """
        try:
            from services.supplier_persistence import get_all_suppliers
            username_filter = self._get_active_username_filter(self.vsm_username_filter_var)
            with DatabaseManager(get_db_path()) as db_manager:
                suppliers = get_all_suppliers(db_manager, username=username_filter)
            self._populate_potential_suppliers_sheet(sheet, suppliers)
            logger.debug(f"Caricati {len(suppliers)} fornitori potenziali")
        except Exception as e:
            logger.error(f"Errore caricamento fornitori potenziali: {e}")
            SimpleMessageDialog(
                self.root,
                tr("Database Error"),
                tr("Unable to load potential suppliers: {}\n").format(e),
                "error",
            )

    def _auto_size_supplier_sheet(self, sheet, data_rows):
        """Calcola larghezze colonne supplier delegando al servizio dedicato."""
        service_auto_size_supplier_sheet(
            sheet=sheet,
            data_rows=data_rows,
            notes_header_text=tr("Notes"),
        )

    def _populate_potential_suppliers_sheet(self, sheet, suppliers, *, resize_columns=True):
        """
        Popola il tksheet Derisking con una lista di PotentialSupplier.

        Imposta sheet._supplier_metadata (lista di dict con supplier_id, username, is_mine)
        e azzera sheet._event_metadata per compatibilità con action-button logic.

        Args:
            sheet:          Widget tksheet del tab Derisking
            suppliers:      list[PotentialSupplier]
            resize_columns: se False, salta il ricalcolo larghezze colonne (usato dal
                            filtro Global Search per evitare micro-spostamenti visivi)
        """
        data_rows, metadata = build_supplier_rows_and_metadata(
            suppliers=suppliers,
            current_username=self.current_username,
            translate_status=tr,
        )
        service_populate_supplier_sheet(
            sheet=sheet,
            data_rows=data_rows,
            metadata=metadata,
            resize_columns=resize_columns,
            notes_header_text=tr("Notes"),
        )

    def _get_vsm_dataset(self, vsm_username_filter):
        """Carica il dataset VSM grezzo in base allo scope utente del filtro UI.

        Unica fonte di verità per lo scope utente nei metodi VSM.
        Usato da _load_vsm_events e _search_vsm_events per evitare duplicazione.

        Args:
            vsm_username_filter: valore da _get_active_username_filter(vsm_username_filter_var).
                None  → tutti gli utenti (aggregazione completa)
                str   → utente specifico; se coincide con current_username usa path locale

        Returns:
            tuple(all_events: list[VSMEvent], extra_meta: list[dict] | None)
        """
        return service_get_vsm_dataset(
            vsm_username_filter=vsm_username_filter,
            current_username=self.current_username,
        )

    def _apply_vsm_filters(self, events, event_type, extra_meta=None):
        """Applica i filtri VSM avanzati (data, azione, ripetitivo, importi) a una lista di eventi.

        Chiamato dopo il filtro per event_type in _load_vsm_events e _search_vsm_events.
        I filtri vuoti vengono ignorati. Restituisce una tupla (filtered_events, filtered_meta)
        con extra_meta allineato agli eventi filtrati (None se non era fornito).
        """
        de_from = getattr(self, "vsm_date_from_entry", None)
        de_to = getattr(self, "vsm_date_to_entry", None)
        av = getattr(self, "vsm_action_var", None)
        rv = getattr(self, "vsm_repetitive_var", None)
        tfv = getattr(self, "vsm_theoretical_from_var", None)
        ttv = getattr(self, "vsm_theoretical_to_var", None)
        afv = getattr(self, "vsm_actual_from_var", None)
        atv = getattr(self, "vsm_actual_to_var", None)

        filters = {
            "date_from": de_from.get().strip() if de_from else "",
            "date_to": de_to.get().strip() if de_to else "",
            "action": av.get().strip() if av else "",
            "repetitive": rv.get().strip() if rv else "",
            "theoretical_from": tfv.get().strip() if tfv else "",
            "theoretical_to": ttv.get().strip() if ttv else "",
            "actual_from": afv.get().strip() if afv else "",
            "actual_to": atv.get().strip() if atv else "",
        }
        return service_apply_vsm_filters(
            events=events,
            event_type=event_type,
            extra_meta=extra_meta,
            filters=filters,
        )

    def _populate_vsm_sheet(self, sheet, events, event_type=None, extra_metadata=None):
        """
        Popola un tksheet con lista di eventi VSM.
        
        ESTRATTO da VSMManagementWindow._populate_sheet() (Step 4C).
        
        Args:
            sheet: Widget tksheet da popolare
            events: Lista di VSMEvent
            event_type: Tipo evento ("Saving"|"Cost Avoidance"|None)
            extra_metadata: Lista opzionale di dict {is_mine, source_file} per riga.
                            Se fornita, sovrascrive il calcolo locale di is_mine.
        """
        data_rows = []
        metadata = []
        currency_code = get_currency_code()
        
        # Usa event_type dalla sheet se non passato direttamente
        if event_type is None:
            event_type = getattr(sheet, '_vsm_event_type', None)
        
        use_dual_value = event_type in ("Saving", "Cost Avoidance")
        
        for i, event in enumerate(events):
            # Calcola valori
            valore_teorico = event.calculate_theoretical_value()
            
            if use_dual_value:
                valore_effettivo = event.calculate_effective_value()
                # Variance %: logica speciale per driver Pagamenti
                if event.driver == "Pagamenti" and event.giorni_pagamento_attuali is not None and event.giorni_pagamento_negoziati is not None:
                    _delta = event.giorni_pagamento_negoziati - event.giorni_pagamento_attuali
                    _variance_pct = (_delta / 30.0) * event.effective_payments_rate_pct
                    _variance_str = f"{_variance_pct:.2f}".replace('.', ',') + "%"
                else:
                    # Variance % = (baseline - negoziato) / baseline * 100
                    # Saving usa importo_bdg; Cost Avoidance usa importo_richiesto_iniziale
                    if event_type == "Cost Avoidance":
                        _baseline = event.importo_richiesto_iniziale or 0.0
                    else:
                        _baseline = event.importo_bdg or 0.0
                    if _baseline != 0.0:
                        _variance_pct = (_baseline - (event.importo_negoziato or 0.0)) / _baseline * 100
                        _variance_str = f"{_variance_pct:.2f}".replace('.', ',') + "%"
                    else:
                        _variance_str = "0%"
                row = [
                    event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
                    event.event_type,
                    tr(event.action),
                    (event.description or event.reference or "")[:50],
                    format_currency_display(valore_teorico, currency_code=currency_code),
                    format_currency_display(valore_effettivo, currency_code=currency_code),
                    f"{event.percent_realizzo:.2f}".replace('.', ',') + "%",
                    _variance_str,
                    "✓" if event.opex_ripetitivo else "",
                    event.username
                ]
            else:
                # LEGACY DEAD CODE — Derisking VSM (precedente modello event-based).
                # Questo branch (non use_dual_value) non viene più raggiunto per il tab Derisking
                # grazie al guard in dashboard_controller.search_requests() (B2 fix).
                # Il tab Derisking ora usa _populate_supplier_sheet() con struttura PotentialSupplier.
                # Da rimuovere nello step successivo insieme al branch C7 di _create_vsm_event_sheet.
                row = [
                    event.event_date.strftime("%d/%m/%Y") if event.event_date else "",
                    event.new_supplier or "",
                    (event.description or event.reference or "")[:50],
                    "✓" if event.opex_ripetitivo else "",
                    event.username
                ]
            data_rows.append(row)
            
            # Metadata per ownership e event_id
            # Se extra_metadata è fornito (aggregazione multi-DB), usa is_mine/source_file
            # dall'aggregatore; altrimenti calcola localmente.
            if extra_metadata is not None and i < len(extra_metadata):
                is_mine = extra_metadata[i].get('is_mine', event.username == self.current_username)
                source_file = extra_metadata[i].get('source_file', 'local')
            else:
                is_mine = event.username == self.current_username
                source_file = 'local'
            metadata.append({
                'event_id': event.id,
                'username': event.username,
                'is_mine': is_mine,
                'source_file': source_file,
            })
        
        # Aggiorna sheet.
        # reset_col_positions=False: lo schema colonne VSM è fisso per tutta la vita dello sheet;
        # evita che MT.reset_col_positions() azzeri transientemente le larghezze a default_column_width,
        # eliminando il micro-tremolio visivo durante search/refresh.
        sheet.set_sheet_data(data_rows, reset_col_positions=False)
        sheet._event_metadata = metadata

        # Larghezze colonne: parte dal template salvato in _create_vsm_event_sheet, poi aggiusta
        # la colonna "Nuovo Fornitore" (Derisking, indice 1) in base al contenuto effettivo.
        col_widths = list(getattr(sheet, '_vsm_col_widths', None) or
                          ([400 if i == 3 else 120 for i in range(10)] if use_dual_value
                           else [400 if i == 2 else 120 for i in range(5)]))

        if not use_dual_value and data_rows:
            # Derisking: calcola larghezza dinamica per "Nuovo Fornitore" (col 1)
            _NEW_SUPPLIER_COL = 1
            _CELL_PADDING = 20
            try:
                import tkinter.font as tkfont
                _cfont = tkfont.Font(family="Calibri", size=11)
                _hfont = tkfont.Font(family="Calibri", size=11, weight="bold")
                _header_w = _hfont.measure(tr("New Supplier")) + 30
                _longest = max((row[_NEW_SUPPLIER_COL] for row in data_rows if row[_NEW_SUPPLIER_COL]), key=len, default="")
                _content_w = _cfont.measure(_longest) + _CELL_PADDING if _longest else 0
                col_widths[_NEW_SUPPLIER_COL] = max(_header_w, _content_w)
            except Exception:
                pass  # Mantiene larghezza esistente se tkfont non disponibile

        sheet.set_column_widths(col_widths)

        # Allineamento centrato: col_options NON viene toccato da reset_col_positions,
        # quindi è già preservato dal populate precedente. La chiamata è mantenuta
        # per garantire la corretta inizializzazione al primo populate.
        align_cols = getattr(sheet, '_vsm_align_cols', None)
        if align_cols:
            sheet.align_columns(columns=align_cols, align="center", redraw=False)
        amount_cols = getattr(sheet, '_vsm_amount_cols', None)
        if amount_cols:
            sheet.align_columns(columns=amount_cols, align="right", redraw=False)
        # Redraw esplicito unico: sostituisce il timer deferred di set_sheet_data e
        # align_columns, garantendo un solo ridisegno con stato già completamente coerente.
        sheet.redraw(redraw_header=True, redraw_row_index=True)

    # ===========================
    # Step 4D.3: VSM CRUD Handlers (implementazione completa)
    # ===========================
    
    def _edit_vsm_event(self):
        """Handler per modifica evento VSM.
        
        Step 4D.3: Implementazione completa con VSMEventDialog.
        Pattern estratto da VSMManagementWindow.on_edit_event().
        """
        sheet, status = self.get_current_tree_and_status()
        if not status.startswith('vsm_'):
            return
        
        # Ottieni selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            SimpleMessageDialog(self.root, tr("No Selection"), tr("Select an event to edit."), "warning")
            return
        
        if len(selected_rows) > 1:
            SimpleMessageDialog(self.root, tr("Multiple Selection"), tr("Select only one event for editing."), "warning")
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
            # Apre in sola lettura invece di bloccare con un errore
            from ui.dialogs.vsm_event_dialog import VSMEventDialog
            event_type = status_to_event_type(status)
            if not event_type:
                return
            dialog = VSMEventDialog(
                self.root,
                current_username=self.current_username,
                event_type=event_type,
                event_id=event_id,
                read_only=True,
            )
            self.root.wait_window(dialog)
            return
        
        # Determina event_type da status
        event_type = status_to_event_type(status)
        if not event_type:
            return  # Fail-safe
        
        # Apri dialog edit
        from ui.dialogs.vsm_event_dialog import VSMEventDialog
        
        try:
            dialog = VSMEventDialog(
                self.root,
                current_username=self.current_username,
                event_type=event_type,
                event_id=event_id
            )
            self.root.wait_window(dialog)
            
            # Refresh se salvato
            if hasattr(dialog, 'result') and dialog.result:
                self._load_vsm_events(event_type, sheet)
                logger.info(f"Evento VSM {event_id} modificato con successo")
        
        except Exception as e:
            logger.error(f"Errore apertura dialog modifica evento VSM: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to open the form: {}").format(e), "error")
    
    def _delete_vsm_events(self):
        """Handler per eliminazione eventi VSM o fornitori potenziali.

        Step 4D.3: Implementazione completa con delete_event_and_impacts.
        Pattern estratto da VSMManagementWindow.on_delete_event().
        Derisking usa il backend fornitori separato (_delete_supplier).
        """
        sheet, status = self.get_current_tree_and_status()
        if not status.startswith('vsm_'):
            return

        # Derisking: routing al backend fornitori potenziali
        if status == 'vsm_derisking':
            self._delete_supplier()
            return

        # Ottieni selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            SimpleMessageDialog(self.root, tr("No Selection"), tr("Select one or more events to delete."), "warning")
            return
        
        # Raccolta event_id e validazione ownership
        events_to_delete = []
        for row_idx in selected_rows:
            if row_idx >= len(sheet._event_metadata):
                continue
            
            metadata = sheet._event_metadata[row_idx]
            
            # Valida ownership
            if not metadata['is_mine']:
                SimpleMessageDialog(self.root, tr("Operation Not Allowed"), tr("You can only delete your own VSM events.\nSome selected events belong to other users."), "error")
                return
            
            events_to_delete.append(metadata['event_id'])
        
        if not events_to_delete:
            return
        
        # Conferma eliminazione
        count = len(events_to_delete)
        if not SimpleYesNoDialog(self.root, tr("Delete Confirmation"), tr("Are you sure you want to delete {} VSM event(s)?\nThis operation cannot be undone.").format(count)).result:
            return
        
        # Determina event_type da status
        event_type = status_to_event_type(status)
        if not event_type:
            return  # Fail-safe
        
        from services.vsm_persistence import VSMError
        
        try:
            delete_vsm_events_by_ids(events_to_delete)
            
            SimpleMessageDialog(self.root, tr("Success"), tr("{} VSM event(s) successfully deleted.").format(count), "info")
            
            # Refresh
            self._load_vsm_events(event_type, sheet)
            logger.info(f"Eliminati {count} eventi VSM con successo")
        
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore eliminazione eventi VSM: {e}")
            SimpleMessageDialog(self.root, tr("Deletion Error"), tr("Unable to delete events:\n{}").format(e), "error")

    def _delete_supplier(self):
        """Handler per eliminazione fornitori potenziali dal tab Derisking.

        Separato da _delete_vsm_events: usa supplier_persistence, non vsm_persistence.
        Pattern coerente con _delete_vsm_events (confirm dialog, refresh, error handling).
        """
        sheet, _status = self.get_current_tree_and_status()

        selected_rows = self._get_selected_row_indices(sheet)
        if not selected_rows:
            SimpleMessageDialog(self.root, tr("No Selection"), tr("Select one or more suppliers to delete."), "warning")
            return

        # Raccolta supplier_id e validazione ownership
        suppliers_to_delete = []
        for row_idx in selected_rows:
            if not hasattr(sheet, '_supplier_metadata') or row_idx >= len(sheet._supplier_metadata):
                continue
            metadata = sheet._supplier_metadata[row_idx]
            if not metadata.get('is_mine', False):
                SimpleMessageDialog(
                    self.root,
                    tr("Operation Not Allowed"),
                    tr("You can delete only your suppliers.\nSome selected suppliers belong to other users."),
                    "error",
                )
                return
            sid = metadata.get('supplier_id')
            if sid:
                suppliers_to_delete.append(sid)

        if not suppliers_to_delete:
            return

        count = len(suppliers_to_delete)
        if not SimpleYesNoDialog(
            self.root,
            tr("Delete Confirmation"),
            tr("Are you sure you want to delete {} supplier(s)?\nThis operation cannot be undone.").format(count),
        ).result:
            return

        from services.supplier_persistence import SupplierError
        try:
            delete_suppliers_by_ids(suppliers_to_delete)

            SimpleMessageDialog(
                self.root,
                tr("Success"),
                tr("{} supplier(s) deleted successfully.").format(count),
                "info",
            )
            self._load_potential_suppliers(self.sheet_derisking)
            logger.info("Eliminati %d fornitori potenziali", count)

        except (DatabaseError, SupplierError) as e:
            logger.error("Errore eliminazione fornitori: %s", e)
            SimpleMessageDialog(
                self.root,
                tr("Deletion Error"),
                tr("Unable to delete suppliers:\n{}").format(e),
                "error",
            )

    def _on_vsm_sheet_double_click(self, sheet, event=None):
        """Gestisce doppio click su riga VSM per aprire edit evento.
        
        Step 4D.4: Handler double-click che delega a _edit_vsm_event().
        La selezione è gestita da tksheet al momento del click.
        UX pulita: click su area vuota → silent return (no popup).
        
        Args:
            sheet: Widget tksheet VSM
            event: Evento Tkinter (non utilizzato, tksheet gestisce selezione)
        """
        # Debounce (pattern RFQ: evita aperture multiple rapide)
        if hasattr(self, '_opening_vsm_edit') and self._opening_vsm_edit:
            return
        
        # UX pulita: verifica selezione PRIMA di impostare flag
        selected_rows = self._get_selected_row_indices(sheet)
        if not selected_rows:
            return  # Silent return, no popup warning
        
        self._opening_vsm_edit = True
        
        try:
            # Delega a handler edit (gestisce validazioni, ownership, dialog)
            self._edit_vsm_event()
        finally:
            # Reset flag dopo breve delay
            self.root.after(300, lambda: setattr(self, '_opening_vsm_edit', False))
    
    def _duplicate_vsm_event(self):
        """Duplica evento VSM selezionato creando una copia identica.
        
        Step 4D.5: Duplicazione 1:1 di evento VSM.
        Pattern RFQ: validazione selezione singola + ownership.
        Logica VSM: usa backend persistence per recupero e salvataggio.
        
        Flow:
        1. Validazione selezione (singola riga, ownership)
        2. Recupero evento completo con get_event_with_impacts()
        3. Creazione copia 1:1 (stesso evento, id=None per auto-increment)
        4. Salvataggio con save_event_with_impacts() (genera impatti automaticamente)
        5. Auto-refresh sheet
        
        NO dialog, NO conferma, duplicazione immediata (come RFQ).
        """
        sheet, status = self.get_current_tree_and_status()
        
        # Guard: verifica che siamo su tab VSM
        if not status.startswith('vsm_'):
            logger.warning("_duplicate_vsm_event chiamato su tab non-VSM")
            return
        
        # Validazione selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            SimpleMessageDialog(self.root, tr("Missing selection"), tr("Select a VSM event to duplicate."), "warning")
            return
        
        if len(selected_rows) > 1:
            SimpleMessageDialog(self.root, tr("Invalid selection"), tr("Select only one VSM event to duplicate it."), "warning")
            return
        
        row_idx = selected_rows[0]
        
        # Validazione ownership
        if row_idx >= len(sheet._event_metadata):
            logger.error(f"Indice VSM {row_idx} fuori range metadata")
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to identify the selected event."), "error")
            return
        
        metadata = sheet._event_metadata[row_idx]
        is_mine = metadata.get('is_mine', False)
        
        if not is_mine:
            SimpleMessageDialog(self.root, tr("Operation Not Allowed"), tr("You cannot duplicate VSM events from other users.\nYou can only work on your own events."), "error")
            logger.warning(f"Tentativo duplicazione evento VSM altrui bloccato: utente={self.current_username}")
            return
        
        event_id = metadata.get('event_id')
        if not event_id:
            logger.error(f"event_id mancante in metadata per riga {row_idx}")
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to identify the selected event."), "error")
            return
        
        # Recupero evento completo dal backend
        try:
            from services.vsm_persistence import VSMError

            new_event_id = duplicate_vsm_event_by_id(event_id)
            logger.info(f"Evento VSM duplicato: {event_id} → {new_event_id}")
            
            # Status mapping per refresh (fallback safe)
            event_type = status_to_event_type(status)
            if status == "vsm_derisking":
                event_type = "Derisking"
            
            if event_type:
                # Auto-refresh sheet
                self._load_vsm_events(event_type, sheet)
                
                # Success feedback
                SimpleMessageDialog(self.root, tr("Success"), tr("VSM event duplicated."), "info")
            else:
                logger.warning(f"Tipo evento non riconosciuto per refresh: {status}")
                SimpleMessageDialog(self.root, tr("Success"), tr("VSM event duplicated. Refresh manually to see the copy."), "info")
        
        except VSMError as e:
            logger.error(f"Errore VSM durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("VSM Error"), tr("Unable to duplicate the event:\n{}").format(e), "error")
        except DatabaseError as e:
            logger.error(f"Errore database durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Database Error"), tr("Unable to duplicate the event:\n{}").format(e), "error")
        except Exception as e:
            logger.error(f"Errore imprevisto durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to duplicate the event:\n{}").format(e), "error")

    def _get_selected_row_indices(self, sheet):
        """Ritorna indici riga selezionati delegando alla policy condivisa."""
        return policy_get_selected_row_indices(sheet)
    
    def _check_if_all_selected_are_mine(self, sheet, selected_indices):
        """Verifica ownership RFQ con fallback fail-safe."""
        return policy_check_all_selected_are_mine(
            sheet=sheet,
            selected_indices=selected_indices,
            metadata_attr="_sheet_rows_metadata",
            entity_label="rfq",
        )
    
    def _check_if_all_vsm_events_are_mine(self, sheet, selected_indices):
        """Verifica ownership eventi VSM con fallback fail-safe."""
        return policy_check_all_selected_are_mine(
            sheet=sheet,
            selected_indices=selected_indices,
            metadata_attr="_event_metadata",
            entity_label="vsm",
        )

    def _check_if_all_suppliers_are_mine(self, sheet, selected_indices):
        """Verifica ownership fornitori Derisking con fallback fail-safe."""
        return policy_check_all_selected_are_mine(
            sheet=sheet,
            selected_indices=selected_indices,
            metadata_attr="_supplier_metadata",
            entity_label="supplier",
        )

    def archive_selected_request(self): self._change_request_status('archiviata')
    def reactivate_selected_request(self): self._change_request_status('attiva')
    def _change_request_status(self, new_status):
        sheet, _status = self.get_current_tree_and_status()
        
        # Ottieni le righe selezionate usando il metodo helper
        selected_rows_indices = self._get_selected_row_indices(sheet)
        if not selected_rows_indices:
            return
        
        # VALIDAZIONE SICUREZZA: Verifica che tutte le RfQ selezionate siano dell'utente corrente
        if not self._check_if_all_selected_are_mine(sheet, selected_rows_indices):
            SimpleMessageDialog(self.root, tr("Operation Not Allowed"), tr("You cannot modify the status of other users' RfQs.\nYou can only operate on your own RfQs."), "error")
            logger.warning(f"Tentativo di modifica stato RfQ altrui bloccato: utente={self.current_username}")
            return
        
        # Ottieni gli ID dalle righe selezionate
        ids = []
        for row_idx in selected_rows_indices:
            try:
                row_data = sheet.get_row_data(row_idx)
                if row_data and len(row_data) > 0:
                    ids.append(row_data[0])  # Primo elemento è l'ID
            except Exception as e:
                logger.error(f"Errore nel recupero dati riga {row_idx}: {e}", exc_info=True)
        
        if not ids:
            return
        
        try:
            update_request_status(request_ids=ids, new_status=new_status)
        except DatabaseError as e:
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to update status: {}").format(e), "error")
        else:
            self.refresh_data()

    def on_tab_changed(self, event):
        self.update_button_visibility()
        self.clear_selection()
        self._update_filter_panel_for_current_tab()
        self._update_advanced_filters_toggle()

    def _update_filter_panel_for_current_tab(self):
        self.dashboard_controller._update_filter_panel_for_current_tab()

    def _update_advanced_filters_toggle(self):
        """Disabilita visivamente Advanced Filters sul tab Derisking.

        Derisking è supplier-based: i filtri avanzati VSM/RFQ non sono applicabili.
        Chiude il pannello se era aperto prima dell'arrivo su Derisking.
        """
        _, status = self.get_current_tree_and_status()
        toolbar = getattr(self, 'main_dashboard_toolbar', None)
        if toolbar is None:
            return
        is_derisking = (status == 'vsm_derisking')
        toolbar.set_advanced_filters_enabled(not is_derisking)
        if is_derisking and hasattr(self, 'collapsible_filters') and self.collapsible_filters.is_expanded():
            self.collapsible_filters.toggle()
            toolbar.filters_toggle_label.config(text=f"⌄ {tr('Advanced Filters')}")
    def update_button_visibility(self):
        """Aggiorna lo stato del pulsante Actions in base alla selezione e proprietà delle RfQ"""
        sheet, status = self.get_current_tree_and_status()
        
        if sheet is None:
            self.btn_actions.config(state="disabled")
            return
        
        # Step 4D.1/4D.2: Gestione abilitazione pulsante Actions per VSM
        if status.startswith('vsm_'):
            selected_rows_indices = self._get_selected_row_indices(sheet)
            selected_count = len(selected_rows_indices) if selected_rows_indices else 0
            
            # Verifica ownership: Derisking usa _supplier_metadata, altri VSM usano _event_metadata
            if status == 'vsm_derisking':
                all_mine = self._check_if_all_suppliers_are_mine(sheet, selected_rows_indices) if selected_count else False
            else:
                all_mine = self._check_if_all_vsm_events_are_mine(sheet, selected_rows_indices) if selected_count else False

            caps = compute_actions_capabilities(
                status=status,
                selected_count=selected_count,
                all_mine=all_mine,
            )
            self.btn_actions.config(state="normal" if caps["can_act"] else "disabled")
            
            # Step 4D.2/4D.5: Popola menu Actions con opzioni VSM
            self._populate_actions_menu(
                status,
                caps["can_delete"],
                caps["can_duplicate"],
                caps["can_change_status"],
            )
            return
        
        # RFQ logic (invariata)
        selected_rows_indices = self._get_selected_row_indices(sheet)
        selected_count = len(selected_rows_indices) if selected_rows_indices else 0
        
        # Verifica se tutte le RfQ selezionate appartengono all'utente corrente
        all_mine = self._check_if_all_selected_are_mine(sheet, selected_rows_indices) if selected_count else False

        caps = compute_actions_capabilities(
            status=status,
            selected_count=selected_count,
            all_mine=all_mine,
        )
        self.btn_actions.config(state="normal" if caps["can_act"] else "disabled")
        
        # Popola il menu Actions dinamicamente in base al tab corrente
        self._populate_actions_menu(
            status,
            caps["can_delete"],
            caps["can_duplicate"],
            caps["can_change_status"],
        )

    def _populate_actions_menu(self, status, can_delete=False, can_duplicate=False, can_change_status=False):
        """Popola il menu Actions in base al tab corrente e capacità utente.
        
        Collega le voci del menu ai metodi esistenti della toolbar.
        Nessuna nuova logica: riusa al 100% i metodi già implementati.
        
        Args:
            status: 'attiva', 'archiviata', o 'vsm_*' (saving/cost_avoidance/derisking)
            can_delete: bool, se può eliminare
            can_duplicate: bool, se può duplicare (1 sola selezione) - per RFQ
                           oppure se può editare (1 sola selezione) - per VSM (riuso stesso param)
            can_change_status: bool, se può archiviare/riattivare (solo RFQ)
        """
        command_map = {
            "delete": (tr("🗑 Delete"), self._delete_vsm_events if status.startswith("vsm_") else self.delete_selected_request),
            "duplicate": (tr("🔁 Duplicate"), self._duplicate_vsm_event if status.startswith("vsm_") else self.duplicate_selected_request),
            "archive": (tr("📦 Archive"), self.archive_selected_request),
            "reactivate": (tr("↩️ Reactivate"), self.reactivate_selected_request),
        }

        menu_spec = build_actions_menu_spec(
            status=status,
            can_delete=can_delete,
            can_duplicate=can_duplicate,
            can_change_status=can_change_status,
        )

        self.actions_menu.delete(0, 'end')
        for item in menu_spec:
            if item[0] == "separator":
                self.actions_menu.add_separator()
                continue

            _kind, key, enabled = item
            label, cmd = command_map[key]
            self.actions_menu.add_command(
                label=label,
                command=cmd,
                state="normal" if enabled else "disabled",
            )

    def _on_root_click(self, event):
        """Gestisce i click sul root per deselezionare quando si clicca fuori dalle griglie.
        
        Args:
            event: Evento click di Tkinter
        """
        # Verifica se il click è avvenuto su uno dei sheet, sul pulsante Actions o sui loro widget figli
        widget = event.widget
        
        # Naviga verso l'alto nella gerarchia widget per vedere se siamo dentro un sheet o sul pulsante Actions
        is_inside_protected_area = False
        current_widget = widget
        while current_widget:
            # Proteggi i sheet e il pulsante Actions (e il suo menu)
            if (current_widget == self.tree_attive or 
                current_widget == self.tree_archiviate or
                current_widget == self.sheet_saving or
                current_widget == self.sheet_cost_avoidance or
                current_widget == self.sheet_derisking or
                current_widget == self.btn_actions or
                current_widget == self.actions_menu):
                is_inside_protected_area = True
                break
            # Prova a salire al parent
            try:
                current_widget = current_widget.master
            except:
                break
        
        # Se il click è fuori dalle aree protette, deseleziona tutto
        if not is_inside_protected_area:
            self.clear_selection()

    def clear_selection(self):
        """Deseleziona tutte le righe in entrambi i sheet"""
        self.tree_attive.deselect("all")
        self.tree_archiviate.deselect("all")
        self.sheet_saving.deselect("all")
        self.sheet_cost_avoidance.deselect("all")
        self.sheet_derisking.deselect("all")
        self.update_button_visibility()

    def get_current_tree_and_status(self):
        tab_index = self.notebook.index(self.notebook.select())
        # Tab 0: RdO Attive, Tab 1: RdO Archiviate
        if tab_index == 0:
            return (self.tree_attive, 'attiva')
        elif tab_index == 1:
            return (self.tree_archiviate, 'archiviata')
        # Step 4B: VSM tabs con sheet reali (riutilizzati da VSMManagementWindow)
        elif tab_index == 2:
            return (self.sheet_saving, 'vsm_saving')
        elif tab_index == 3:
            return (self.sheet_cost_avoidance, 'vsm_cost_avoidance')
        elif tab_index == 4:
            return (self.sheet_derisking, 'vsm_derisking')
        else:
            # Fallback per tab non previsti
            return (None, 'unknown')

    def refresh_data(self):
        self.dashboard_controller.refresh_data()

    def _load_requests_by_status(self, tree, status, *, pre_fetched_rows=None):
        """Carica richieste per stato specifico con supporto multi-database."""
        try:
            username_filter = self._get_active_username_filter()
            tipo_filter = self.search_tipo.get()
            tipo_canonico = None if tipo_filter == tr("All") else normalize_rfq_type(tipo_filter)
            filtered_rows = service_load_requests_by_status(
                status=status,
                username_filter=username_filter,
                tipo_canonico=tipo_canonico,
                pre_fetched_rows=pre_fetched_rows,
            )
            self.update_treeview(tree, filtered_rows)
                
        except DatabaseError as e:
            logger.error(f"Errore database in _load_requests_by_status: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to load list: {}").format(e), "error")

    def update_treeview(self, sheet, requests):
        """Aggiorna il foglio tksheet con i dati delle richieste"""
        today = date.today()
        data_rows, metadata_rows, max_ref_length = build_rfq_sheet_payload(
            requests=requests,
            translate_rfq_type=translate_rfq_type,
            format_date_for_display=self._format_date_for_display,
        )

        # Inizializza lista metadati se non esiste
        if not hasattr(sheet, "_sheet_rows_metadata"):
            sheet._sheet_rows_metadata = []
        sheet._sheet_rows_metadata = metadata_rows
        
        # Carica i dati nel sheet
        sheet.set_sheet_data(data_rows)
        
        # Salva i dati per uso successivo (ad esempio per l'ordinamento)
        sheet._sheet_data = data_rows
        sheet._sheet_requests = requests  # Salva anche i dati completi del DB
        
        # Calcola larghezza ottimale per la colonna Riferimento
        try:
            import tkinter.font as tkfont
            
            # Font usato nel sheet
            content_font = tkfont.Font(family="Calibri", size=11, weight="normal")
            header_font = tkfont.Font(family="Calibri", size=11, weight="bold")
            
            # Larghezza minima (basata sull'header "Riferimento")
            header_text = tr("Reference")
            min_width = header_font.measure(header_text) + 30  # +30 per padding
            
            # Se ci sono riferimenti, calcola la larghezza in base al più lungo
            if max_ref_length > 0:
                # Trova il riferimento più lungo per misurarlo con precisione
                longest_ref = max((row[4] for row in data_rows if row[4]), key=len, default="")
                content_width = content_font.measure(longest_ref) + 40  # +40 per padding e margini
                optimal_width = max(min_width, content_width)
            else:
                optimal_width = min_width
            
            # Limita la larghezza massima per evitare colonne eccessivamente larghe
            MAX_WIDTH = 600  # Massimo 600 pixel
            optimal_width = min(optimal_width, MAX_WIDTH)
            
            # Applica la larghezza ottimale alla colonna Riferimento (indice 4)
            sheet.column_width(column=4, width=int(optimal_width))
            
        except Exception as e:
            logger.warning(f"Errore calcolo larghezza colonna Riferimento: {e}. Uso larghezza default.")
            # Fallback a larghezza fissa se il calcolo fallisce
            sheet.column_width(column=4, width=300)
        
        # Reset completo di tutte le evidenziazioni precedenti
        sheet.dehighlight_all()
        
        # Applica colorazione per righe scadute (solo per tab attive)
        if sheet is self.tree_attive:
            for i, req in enumerate(requests):
                if req[3]:
                    try:
                        # BUG #22 FIX: Parse con strptime e confronto date robusto
                        expiry_date = datetime.strptime(req[3], '%Y-%m-%d').date()
                        if expiry_date < today:
                            # Evidenzia la riga in rosso per scadenza
                            sheet.highlight_rows([i], bg='#FFE6E6', fg='red')
                    except (ValueError, TypeError) as e:
                        # BUG #22 FIX: Log dettagliato con ID RdO per troubleshooting
                        logger.warning(f"Formato data scadenza non valido per RdO {req[0]}: '{req[3]}' - {e}")
        
        # BUG #27 FIX: Usa enumerate invece di range(len())
        # Applica strisce alternate per le righe
        for i, _row in enumerate(data_rows):
            if i % 2 != 0:
                if sheet is self.tree_attive:
                    # BUG #29 FIX: Catch specifico invece di bare except
                    # Controlla se già evidenziata come scaduta
                    try:
                        req = requests[i]
                        if req[3] and datetime.strptime(req[3], '%Y-%m-%d').date() < today:
                            continue  # Non sovrascrivere l'evidenziazione rossa
                    except (ValueError, TypeError, IndexError) as e:
                        logger.debug(f"Errore parsing data per stripe row {i}: {e}")
                sheet.highlight_rows([i], bg='#F0F0F0', fg='black')
            else:
                # Righe pari: assicura che abbiano fg='black' se non sono scadute
                if sheet is self.tree_attive:
                    # BUG #29 FIX: Catch specifico invece di bare except
                    try:
                        req = requests[i]
                        if req[3] and datetime.strptime(req[3], '%Y-%m-%d').date() < today:
                            continue  # Già evidenziata in rosso
                    except (ValueError, TypeError, IndexError) as e:
                        logger.debug(f"Errore parsing data per stripe row {i}: {e}")
                # Applica esplicitamente sfondo bianco con testo nero per righe pari non scadute
                sheet.highlight_rows([i], bg='white', fg='black')

    # ===========================
    # Global Search — VSM handler
    # ===========================

    _VSM_STATUS_TO_TYPE = {
        'vsm_saving': 'Saving',
        'vsm_cost_avoidance': 'Cost Avoidance',
        # 'vsm_derisking' escluso: tab supplier-based, gestito da _search_derisking_suppliers.
    }

    # ================================
    # Global Search — Derisking handler
    # ================================

    def _search_derisking_suppliers(self, sheet):
        """Handler ricerca globale per il tab Derisking (supplier-based).

        Filtra i fornitori potenziali per sottostringa in tutti i campi visibili.
        Query vuota = ripristina dataset completo (stesso comportamento degli altri tab).
        """
        from services.supplier_persistence import get_all_suppliers

        query = self.search_vars['global'].get().strip().lower()
        username_filter = self._get_active_username_filter(self.vsm_username_filter_var)

        try:
            with DatabaseManager(get_db_path()) as db_manager:
                suppliers = get_all_suppliers(db_manager, username=username_filter)
        except Exception as e:
            logger.error(f"[DerisSearch] Errore caricamento fornitori: {e}", exc_info=True)
            return

        if not query:
            # Query vuota: ripristina dataset completo
            self._populate_potential_suppliers_sheet(sheet, suppliers)
            return

        _FIELDS = (
            "supplier_name",
            "category",
            "supplier_status",
            "contact_name",
            "email",
            "phone",
            "website",
            "notes",
            "username",
        )
        results = filter_derisking_suppliers_by_query(
            suppliers=suppliers,
            query=query,
            fields=_FIELDS,
        )

        logger.info(f"[DerisSearch] query='{query}' risultati={len(results)}")
        self._populate_potential_suppliers_sheet(sheet, results, resize_columns=False)

    def _search_vsm_events(self, sheet, status):
        """Handler di ricerca globale per il modulo VSM.

        Dispatch point isolato per i tab VSM.
        Completamente separato dalla logica RFQ: nessuna condizione condivisa.
        Per aggiungere un nuovo modulo: creare _search_<modulo>() e aggiungerlo
        al dispatch in search_requests().

        Args:
            sheet: Widget tksheet del tab corrente
            status: Status stringa del tab (es. 'vsm_saving')
        """
        event_type = self._VSM_STATUS_TO_TYPE.get(status)
        if not event_type:
            return  # Fail-safe per stati non mappati

        query = self.search_vars['global'].get().strip().lower()

        if not query:
            # Query vuota: ripristina dataset completo rispettando lo scope filtri attivi
            self._load_vsm_events(event_type, sheet)
            return

        # Scope utente determinato esclusivamente dal filtro UI (mai da current_username diretto)
        vsm_username_filter = self._get_active_username_filter(self.vsm_username_filter_var)

        try:
            raw_events, raw_meta = self._get_vsm_dataset(vsm_username_filter)
        except DatabaseError as e:
            logger.error(f"[VSMSearch] Errore caricamento eventi: {e}")
            return

        # Filtra per event_type mantenendo allineamento con raw_meta
        results, result_meta = split_vsm_events_by_type(
            events=raw_events,
            metadata=raw_meta,
            event_type=event_type,
        )

        # Applica Advanced Filters (stesso scope di _load_vsm_events)
        results, result_meta = self._apply_vsm_filters(results, event_type, extra_meta=result_meta)

        # Applica query testuale globale
        _VSM_SEARCH_FIELDS = (
            'description', 'reference', 'buyer', 'driver',
            'action', 'event_type', 'new_supplier', 'note',
        )
        results, result_meta = filter_vsm_events_by_query(
            events=results,
            metadata=result_meta,
            query=query,
            fields=_VSM_SEARCH_FIELDS,
        )

        logger.info(f"[VSMSearch] query='{query}' event_type='{event_type}' risultati={len(results)}")
        self._populate_vsm_sheet(sheet, results, event_type=event_type, extra_metadata=result_meta)

    def search_requests(self):
        self.dashboard_controller.search_requests()

    def delete_selected_request(self):
        sheet, _status = self.get_current_tree_and_status()
        
        # Ottieni le righe selezionate (supporta sia selezione cella che riga)
        selected_rows_indices = self._get_selected_row_indices(sheet)
        if not selected_rows_indices:
            return
        
        # VALIDAZIONE SICUREZZA: Verifica che tutte le RfQ selezionate siano dell'utente corrente
        if not self._check_if_all_selected_are_mine(sheet, selected_rows_indices):
            SimpleMessageDialog(self.root, tr("Operation Not Allowed"), tr("You cannot delete other users' RfQs.\nYou can only operate on your own RfQs."), "error")
            logger.warning(f"Tentativo di eliminazione RfQ altrui bloccato: utente={self.current_username}")
            return
        
        # Ottieni gli ID dalle righe selezionate
        request_ids = []
        for row_idx in selected_rows_indices:
            try:
                row_data = sheet.get_row_data(row_idx)
                if row_data and len(row_data) > 0:
                    request_ids.append(row_data[0])
            except Exception as e:
                logger.error(f"Errore nel recupero dati riga {row_idx}: {e}", exc_info=True)
        
        if not request_ids:
            return
        count = len(request_ids)
        
        if count == 1:
            rdo_num = request_ids[0]
            msg = tr("Are you sure you want to delete RfQ N° {}?\nThe operation is permanent.").format(rdo_num)
        else:
            msg = tr("Are you sure you want to delete the {} selected RfQs?\nThe operation is permanent.").format(count)
        if not SimpleYesNoDialog(self.root, tr("Delete Confirmation"), msg).result: return
        
        try:
            archive_path = get_fixed_attachments_dir()
            count = delete_requests_with_attachments(
                request_ids=request_ids,
                archive_path=archive_path,
            )

            # Ricarica i dati invece di cancellare elementi dalla view
            self.refresh_data()
            if count == 1:
                msg = tr("1 RfQ deleted.")
            else:
                msg = tr("{} RfQs deleted.").format(count)
            SimpleMessageDialog(self.root, tr("Success"), msg, "info")
        except DatabaseError as e:
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to delete: {}").format(e), "error")

    def duplicate_selected_request(self):
        sheet, _status = self.get_current_tree_and_status()
        
        # Prova a ottenere la riga selezionata in vari modi
        row_index = None
        
        # Metodo 1: Prova con get_currently_selected (per selezione cella)
        currently_selected = sheet.get_currently_selected()
        if currently_selected:
            if hasattr(currently_selected, 'row') and currently_selected.row is not None:
                row_index = currently_selected.row
            elif isinstance(currently_selected, tuple) and len(currently_selected) >= 1:
                row_index = currently_selected[0]
        
        # Metodo 2: Prova con get_selected_rows (per selezione riga)
        if row_index is None:
            selected_rows = sheet.get_selected_rows()
            if selected_rows:
                if len(selected_rows) > 1:
                    SimpleMessageDialog(self.root, tr("Invalid selection"), tr("Select only one RfQ to duplicate."), "warning")
                    return
                row_index = selected_rows[0] if isinstance(selected_rows, (list, set, tuple)) else selected_rows
        
        # VALIDAZIONE SICUREZZA: Verifica che la RfQ selezionata sia dell'utente corrente
        if row_index is not None:
            if not self._check_if_all_selected_are_mine(sheet, [row_index]):
                SimpleMessageDialog(self.root, tr("Operation Not Allowed"), tr("You cannot duplicate other users' RfQs.\nYou can only operate on your own RfQs."), "error")
                logger.warning(f"Tentativo di duplicazione RfQ altrui bloccato: utente={self.current_username}")
                return
        
        # Se non c'è nessuna selezione
        if row_index is None:
            SimpleMessageDialog(self.root, tr("Missing selection"), tr("Select an RFQ to duplicate."), "warning")
            return

        # Ottieni i dati della riga
        try:
            row_data = sheet.get_row_data(row_index)
            if not row_data or len(row_data) == 0:
                SimpleMessageDialog(self.root, tr("Error"), tr("Unable to determine the selected RfQ."), "error")
                return
            original_id = int(row_data[0])
        except (ValueError, TypeError, IndexError) as e:
            logger.error(f"Errore nel recupero dati riga per duplicazione: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to determine the selected RfQ."), "error")
            return

        try:
            new_request_id = duplicate_request_full(original_id=original_id)
            if new_request_id is None:
                raise ValueError("Duplicazione fallita: ID nuova RdO non ottenuto")
            logger.info(f"RdO duplicata: {original_id} -> {new_request_id}")
        except ValueError as ve:
            SimpleMessageDialog(self.root, tr("Error"), str(ve), "error")
            return
        except DatabaseError as e:
            logger.error(f"Errore duplicazione RdO {original_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to duplicate the RfQ: {}").format(e), "error")
            return
        except Exception as e:
            logger.error(f"Errore duplicazione RdO {original_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Unable to duplicate: {}").format(e), "error")
            return

        # Ora new_request_id è garantito essere valido

        self._assign_request_to_current_user(new_request_id)
        self.refresh_data()
        self.notebook.select(self.tab_attive)
        
        # Cerca e seleziona la riga con il nuovo ID nel sheet
        total_rows = self.tree_attive.get_total_rows()
        for row_idx in range(total_rows):
            row_data = self.tree_attive.get_row_data(row_idx)
            if row_data and str(row_data[0]) == str(new_request_id):
                self.tree_attive.select_row(row_idx)
                self.tree_attive.see(row_idx)
                break
        
        self.update_button_visibility()
        SimpleMessageDialog(self.root, tr("Success"), tr("RFQ duplicated as N° {}.").format(new_request_id), "info")

    def clear_filters(self):
        self.dashboard_controller.clear_filters()
    
    def toggle_filters(self):
        """Toggle visibilità filtri avanzati (Step 5: Collapsible Filters).
        
        Chiamato dal trigger nella Global Search toolbar.
        Delega a CollapsibleFilters per gestione expand/collapse.
        """
        if hasattr(self, 'collapsible_filters'):
            self.collapsible_filters.toggle()

    def on_sheet_double_click(self, sheet, event=None):
        """Gestisce il doppio click su una riga del sheet per aprire la RdO con debounce"""
        try:
            # Verifica che non ci sia già una finestra in apertura (debounce)
            if hasattr(self, '_opening_request') and self._opening_request:
                return
            
            self._opening_request = True
            
            try:
                # Ottieni la selezione corrente (riga, colonna)
                # Al momento del doppio click, tksheet ha già selezionato la cella
                currently_selected = sheet.get_currently_selected()
                row_index = None
                
                # Prova a determinare la riga dal get_currently_selected
                if currently_selected:
                    # get_currently_selected restituisce un oggetto con vari attributi
                    if hasattr(currently_selected, 'row') and currently_selected.row is not None:
                        row_index = currently_selected.row
                    elif isinstance(currently_selected, tuple) and len(currently_selected) >= 2:
                        row_index = currently_selected[0]
                
                # Se non abbiamo trovato la riga, proviamo con get_selected_rows
                if row_index is None:
                    selected = sheet.get_selected_rows()
                    if not selected:
                        return
                    row_index = selected[0] if isinstance(selected, (list, set, tuple)) else selected
                
                # Verifica che l'indice sia valido
                if row_index is None or row_index < 0 or row_index >= sheet.get_total_rows():
                    return
                
                # Ottieni i dati della riga
                data = sheet.get_row_data(row_index)
                if data and len(data) > 0:
                    request_id = data[0]  # Primo elemento è l'ID
                    
                    # Controllo se la RdO è mia (per multi-utente)
                    is_mine = True  # Default per compatibilità
                    source_db_path = None  # Percorso del DB sorgente
                    
                    # BUG #3 FIX: Validazione robusta con gestione errori completa
                    if hasattr(sheet, '_sheet_rows_metadata'):
                        try:
                            if 0 <= row_index < len(sheet._sheet_rows_metadata):
                                metadata = sheet._sheet_rows_metadata[row_index]
                                is_mine = metadata.get('is_mine', True)
                                source_db_path = metadata.get('source_file', None)
                                logger.debug(f"RdO {request_id}: is_mine={is_mine}, source={source_db_path}")
                            else:
                                logger.warning(f"Indice {row_index} fuori range metadati (len={len(sheet._sheet_rows_metadata)}), uso default is_mine=True")
                        except (AttributeError, KeyError, TypeError) as e:
                            logger.error(f"Errore accesso metadati riga {row_index}: {e}")
                    
                    # Apri la finestra di dettaglio della RdO (con flag read_only se non è mia)
                    self.root.wait_window(ViewRequestWindow(
                        self.root, 
                        request_id, 
                        read_only=not is_mine,
                        source_db_path=source_db_path if not is_mine else None
                    ))
                    # Aggiorna i dati dopo la chiusura della finestra
                    self.refresh_data()
            finally:
                # BUG #30 FIX: Usa weakref per evitare memory leak
                # Rilascia il lock dopo un breve delay per evitare doppi click rapidi
                import weakref
                weak_self = weakref.ref(self)
                def release_lock():
                    obj = weak_self()
                    if obj is not None:
                        obj._opening_request = False
                self.root.after(300, release_lock)
                
        except Exception as e:
            logger.error(f"Errore nell'apertura della RdO: {e}", exc_info=True)
            self._opening_request = False
    
    def open_new_event(self):
        """Handler dinamico per pulsante + Nuovo Evento.
        
        Step 4D.6: Routing intelligente basato sul tab attivo:
        - RFQ (attiva/archiviata): crea nuova RdO
        - VSM Saving/Cost Avoidance: crea nuovo evento VSM
        - Derisking: apre PotentialSupplierDialog (non VSMEventDialog)
        """
        _, status = self.get_current_tree_and_status()
        
        # Branch 1: RFQ - usa la logica esistente
        if status in ('attiva', 'archiviata'):
            self.open_new_request_window()
        
        # Branch 2: VSM - apri dialog CREATE
        elif status.startswith('vsm_'):
            # Derisking: usa PotentialSupplierDialog (non VSMEventDialog)
            if status == 'vsm_derisking':
                from ui.dialogs.potential_supplier_dialog import PotentialSupplierDialog
                try:
                    dlg = PotentialSupplierDialog(
                        self.root,
                        self.current_username,
                        refresh_derisking_cb=lambda: self._load_potential_suppliers(self.sheet_derisking),
                    )
                    self.root.wait_window(dlg)
                    if dlg.result:
                        self._load_potential_suppliers(self.sheet_derisking)
                except Exception as e:
                    logger.error("Errore creazione fornitore: %s", e, exc_info=True)
                    SimpleMessageDialog(self.root, tr("Error"), tr("Unable to open the form: {}").format(e), "error")
                return

            # Mappa status → event_type
            event_type_map = {
                'vsm_saving': 'Saving',
                'vsm_cost_avoidance': 'Cost Avoidance',
            }
            event_type = event_type_map.get(status)
            
            if not event_type:
                return  # Fail-safe
            
            # Lazy import (come in _edit_vsm_event)
            from ui.dialogs.vsm_event_dialog import VSMEventDialog
            
            try:
                # Apri dialog in modalità CREATE (event_id=None)
                dialog = VSMEventDialog(
                    self.root,
                    current_username=self.current_username,
                    event_type=event_type,
                    event_id=None  # CREATE mode
                )
                self.root.wait_window(dialog)
                
                # Refresh se salvato
                if hasattr(dialog, 'result') and dialog.result:
                    # Ottieni sheet corrente
                    sheet, _ = self.get_current_tree_and_status()
                    self._load_vsm_events(event_type, sheet)
                    logger.info(f"Nuovo evento VSM {event_type} creato con successo")
            
            except Exception as e:
                logger.error(f"Errore creazione evento VSM: {e}", exc_info=True)
                SimpleMessageDialog(self.root, tr("Error"), tr("Unable to open the form: {}").format(e), "error")

    def open_new_request_window(self):
        """Crea una nuova RdO 'guscio' e apre l'editor"""
        # Mostra dialog per scelta tipo
        dialog = NewRdOTypeDialog(self.root)
        self.root.wait_window(dialog)
        
        # Se l'utente ha annullato, esci
        if not dialog.result:
            return
        
        tipo_rdo = normalize_rfq_type(dialog.result)
        
        try:
            data_oggi = datetime.now().strftime('%Y-%m-%d')
            id_nuova = create_request_shell(
                tipo_rdo=tipo_rdo,
                status="attiva",
                issue_date=data_oggi,
                username=self.current_username,
            )
            
            logger.info(f"Creata nuova RdO guscio N° {id_nuova} (tipo: {tipo_rdo})")
            
            # Apri immediatamente l'editor
            self.root.wait_window(ViewRequestWindow(self.root, id_nuova))
            
            # Aggiorna la lista dopo la chiusura
            self.refresh_data()
            
        except DatabaseError as e:
            logger.error(f"Errore creazione RdO guscio: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Database Error"), tr("Unable to create the new RfQ: {}").format(e), "error")

    def mega_export_excel(self):
        """
        Esporta tutte le RfQ attualmente visibili nella lista (filtrate) in un unico file Excel.
        Genera un report a blocchi verticali, adattandosi al tipo di ogni singola RfQ.
        """
        # 1. Identifica quale tabella è attiva e recupera lo stato corrente
        current_tree, status = self.get_current_tree_and_status()
        
        # VSM tabs: dispatch to dedicated export handlers
        if status == 'vsm_derisking':
            self._export_derisking_excel()
            return
        if status.startswith('vsm_'):
            self._export_vsm_excel(status, current_tree)
            return
        
        if current_tree is None:
            return
        
        # 2. Recupera TUTTI gli ID che corrispondono ai filtri attivi (non solo quelli visualizzati nel sheet)
        # Questo è necessario perché il sheet potrebbe avere un limite di righe visualizzate
        # IMPORTANTE: recupera anche il percorso del database sorgente per ogni RfQ
        request_data = []  # Lista di tuple (request_id, source_db_path)
        
        try:
            # Verifica se ci sono filtri di ricerca attivi
            if self._has_active_search_filters():
                # CI SONO FILTRI ATTIVI: esegui la stessa query di search_requests per ottenere TUTTI i risultati
                logger.info("[export_excel] Filtri di ricerca attivi - recupero tutti i risultati filtrati")
                
                username_filter = self._get_active_username_filter()
                crit = {k: v.get().strip() for k, v in self.search_vars.items()}
                dates = {k: format_date_for_db(v.get().strip()) for k, v in self.date_entries.items()}
                
                # Gestione tipo RdO
                tipo_rdo = None
                if self.search_tipo.get() != tr("All"):
                    tipo_rdo = normalize_rfq_type(self.search_tipo.get())
                
                # Usa SEMPRE aggregazione multi-database per avere dati da tutti gli utenti
                # Poi filtra in memoria per username specifico se necessario
                with DatabaseManager(get_db_path()) as db_manager:
                    all_results = db_manager.get_all_richieste_aggregated(get_db_path())
                
                # Filtra in memoria applicando TUTTI i criteri
                for row in all_results:
                    # Filtro stato
                    if row[6] != status:
                        continue
                    # Filtro username
                    if username_filter and (not row[5] or row[5].lower() != username_filter.lower()):
                        continue
                    # Filtro tipo RdO
                    if tipo_rdo and row[1] != tipo_rdo:
                        continue
                    # Filtri di testo
                    if crit['num'] and crit['num'] not in str(row[0]):
                        continue
                    if crit['ref'] and (not row[4] or crit['ref'].lower() not in row[4].lower()):
                        continue
                    # Filtri data
                    if dates['emm_da'] and (not row[2] or row[2] < dates['emm_da']):
                        continue
                    if dates['emm_a'] and (not row[2] or row[2] > dates['emm_a']):
                        continue
                    if dates['scad_da'] and (not row[3] or row[3] < dates['scad_da']):
                        continue
                    if dates['scad_a'] and (not row[3] or row[3] > dates['scad_a']):
                        continue
                    
                    # Filtri su dettagli (fornitore, materiale, ecc.)
                    if any([crit['forn'], crit['cod'], crit['desc'], crit['ord'], 
                           crit['cod_grezzo'], crit['dis_grezzo'], crit['mat_cl']]):
                        source_db_path = row[8] if len(row) > 8 else 'local'
                        if source_db_path == 'local':
                            source_db_path = get_db_path()
                        try:
                            with DatabaseManager(source_db_path) as source_db_mgr:
                                detail_match = source_db_mgr.check_richiesta_detail_criteria(
                                    row[0],
                                    {
                                        'forn': crit['forn'], 'cod': crit['cod'], 'desc': crit['desc'],
                                        'ord': crit['ord'], 'cod_grezzo': crit['cod_grezzo'],
                                        'dis_grezzo': crit['dis_grezzo'], 'mat_cl': crit['mat_cl']
                                    }
                                )
                            if not detail_match:
                                continue
                        except Exception as e:
                            logger.warning(f"Errore verifica criteri dettaglio per RdO {row[0]}: {e}")
                            continue
                    
                    # Tutti i filtri passati - salva ID e percorso database
                    source_db_path = row[8] if len(row) > 8 else 'local'
                    if source_db_path == 'local':
                        source_db_path = get_db_path()
                    request_data.append((row[0], source_db_path))
                
            else:
                # NESSUN FILTRO ATTIVO: carica tutte le RfQ nello stato corrente
                logger.info("[export_excel] Nessun filtro attivo - carico tutte le RfQ nello stato corrente")
                
                username_filter = self._get_active_username_filter()
                
                with DatabaseManager(get_db_path()) as db_manager:
                    all_rows = db_manager.get_all_richieste_aggregated(get_db_path())
                
                # Filtra per stato corrente
                filtered_rows = [row for row in all_rows if row[6] == status]
                
                # Applica filtro username se presente
                if username_filter is not None:
                    filtered_rows = [row for row in filtered_rows if row[5] and row[5].lower() == username_filter.lower()]
                
                # Applica filtro tipo RdO se presente (non "Tutte")
                tipo_filter = self.search_tipo.get()
                if tipo_filter != tr("All"):
                    tipo_canonico = normalize_rfq_type(tipo_filter)
                    filtered_rows = [row for row in filtered_rows if row[1] == tipo_canonico]
                
                # Salva ID e percorso database per ogni RfQ
                for row in filtered_rows:
                    source_db_path = row[8] if len(row) > 8 else 'local'
                    if source_db_path == 'local':
                        source_db_path = get_db_path()
                    request_data.append((row[0], source_db_path))
            
            if not request_data:
                SimpleMessageDialog(self.root, tr("Warning"), tr("No RfQ to export in the current view."), "warning")
                return
            
            logger.info(f"[export_excel] Trovate {len(request_data)} RfQ da esportare: {[r[0] for r in request_data[:10]]}{'...' if len(request_data) > 10 else ''}")
            
        except Exception as e:
            logger.error(f"[export_excel] Errore nel recupero degli ID: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Error nel recupero delle RfQ da esportare: {}").format(e), "error")
            return

        export_rfq_requests_excel(
            parent=self.root,
            request_data=request_data,
            format_date_for_display=self._format_date_for_display,
        )

    def _export_vsm_excel(self, status, sheet):
        """Esporta i dati VSM (Saving / Cost Avoidance) del tab corrente in un file Excel.

        Gestisce solo status in {'vsm_saving', 'vsm_cost_avoidance'}.
        Il tab Derisking è gestito separatamente da _export_derisking_excel().
        """
        status_to_event_type = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
        }
        event_type = status_to_event_type.get(status, status)

        # Re-load eventi dal DB rispettando scope utente e filtri attivi (allineato a _load_vsm_events)
        try:
            vsm_username_filter = self._get_active_username_filter(self.vsm_username_filter_var)
            all_events, extra_meta = self._get_vsm_dataset(vsm_username_filter)
            if extra_meta is not None:
                pairs = [(ev, m) for ev, m in zip(all_events, extra_meta) if ev.event_type == event_type]
                events = [p[0] for p in pairs]
            else:
                events = [e for e in all_events if e.event_type == event_type]
            events, _unused_meta = self._apply_vsm_filters(events, event_type)
        except Exception as e:
            logger.error(f"[export_vsm] Errore recupero eventi: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Error retrieving data: {}").format(e), "error")
            return

        if not events:
            SimpleMessageDialog(self.root, tr("Warning"), tr("No data to export in the current view."), "warning")
            return

        export_vsm_events_excel(
            parent=self.root,
            status=status,
            sheet_col_widths=getattr(sheet, "_vsm_col_widths", None),
            events=events,
        )

    def _export_derisking_excel(self):
        """Esporta i fornitori potenziali del tab Derisking in un file Excel.

        Routing separato da _export_vsm_excel: usa PotentialSupplier, non VSMEvent.
        """
        # Carica tutti i fornitori potenziali dal DB (stesso pattern di _load_potential_suppliers)
        try:
            username_filter = self._get_active_username_filter(self.vsm_username_filter_var)
            suppliers = load_derisking_suppliers_for_export(username_filter=username_filter)
        except Exception as e:
            logger.error(f"[export_derisking] Errore recupero fornitori: {e}", exc_info=True)
            SimpleMessageDialog(self.root, tr("Error"), tr("Error retrieving data: {}").format(e), "error")
            return

        export_derisking_suppliers_excel(
            parent=self.root,
            suppliers=suppliers,
        )

    def _format_date_for_display(self, db_date):
        if not db_date: return ""
        try: return datetime.strptime(db_date, '%Y-%m-%d').strftime('%d/%m/%Y')
        except (ValueError, TypeError): return db_date
# ------------------------------------------------------------------------------------
# SCALING DPI GESTITO AUTOMATICAMENTE DA WINDOWS + TKINTER
# ------------------------------------------------------------------------------------
# Il manifest app.manifest dichiara PerMonitorV2, quindi Tkinter gestisce
# automaticamente il DPI scaling senza bisogno di intervento manuale.


# Le dimensioni delle finestre sono ora gestite automaticamente da Tkinter
# in base al DPI di Windows, senza bisogno di scaling manuale.

# ------------------------------------------------------------------------------------
# ESECUZIONE PRINCIPALE
# ------------------------------------------------------------------------------------
_pending_restart: tuple | None = None  # (cmd, cwd) impostato da restart_program() pre-mainloop

if __name__ == '__main__':
    # Inizializza il sistema di internazionalizzazione PRIMA di creare qualsiasi finestra
    logger.info("=" * 70)
    logger.info("INIZIALIZZAZIONE I18N")
    logger.info("=" * 70)
    language_code = init_i18n()
    logger.info(f"Lingua inizializzata: {language_code}")
    logger.info("=" * 70)
    
    root = tk.Tk()
    
    # --- DPI scaling automatico per monitor ad alta risoluzione ---
    try:
        dpi = root.winfo_fpixels('1i')
        scaling = dpi / 72
        root.tk.call('tk', 'scaling', scaling)
    except Exception:
        pass

    # --- Migliora rendering grafico su tutti i sistemi ---
    try:
        from tkinter import ttk
        import platform
        style = ttk.Style()
        system = platform.system()
        
        if system == 'Windows':
            # Su Windows, usa il tema nativo per aspetto coerente con l'OS
            if 'vista' in style.theme_names():
                style.theme_use('vista')
            elif 'winnative' in style.theme_names():
                style.theme_use('winnative')
        elif system == 'Linux':
            # Su Linux, usa clam per rendering moderno
            if 'clam' in style.theme_names():
                style.theme_use('clam')
        # macOS usa 'aqua' di default, non serve configurare
    except Exception:
        pass
    
    root.withdraw()
    splash = None
    
    def main_task():
        # Leggi config esistente (se presente)
        config = configparser.ConfigParser(interpolation=None)
        config_file = get_config_file()
        license_was_accepted = False
        if os.path.exists(config_file):
            try:
                config.read(config_file)
                license_was_accepted = config.getboolean('Settings', 'license_accepted', fallback=False)
            except Exception:
                license_was_accepted = False

        # 1) Mostra la licenza PRIMA di qualsiasi creazione DB
        if not license_was_accepted:
            license_prompt = LicenseAcceptanceDialog(root, url="https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE")
            root.wait_window(license_prompt)
            if not getattr(license_prompt, 'accepted', False):
                try:
                    root.destroy()
                except:
                    pass
                return

            # Salva subito l'accettazione
            try:
                if 'Settings' not in config:
                    config['Settings'] = {}
                config['Settings']['license_accepted'] = 'True'
                with open(config_file, 'w', encoding='utf-8') as f:
                    config.write(f)
            except Exception as e:
                logger.warning(f"Impossibile salvare stato licenza: {e}")

        # 2) Verifica se l'identità utente è già presente, altrimenti richiedila
        identity = load_user_identity()
        if not identity.get('username'):
            # L'identità non è presente o incompleta, mostra il dialogo
            dialog = UserIdentityDialog(root)
            root.wait_window(dialog)
            identity = getattr(dialog, 'result', None)
            if not identity:
                try:
                    root.destroy()
                except:
                    pass
                return
            # Salva subito l'identità nel config
            save_user_identity(identity['first_name'], identity['last_name'], identity['username'])
            # Ricarica l'identità appena salvata
            identity = load_user_identity()
        else:
            # L'identità è già presente, logga e continua
            logger.info(f"Identità utente già presente: {identity['username']}")

        # Salva identità e imposta percorso DB utente (solo se non già presente o diverso)
        try:
            if 'Settings' not in config:
                config['Settings'] = {}
            existing_identity = load_user_identity()
            if not existing_identity.get('username'):
                config['User'] = {
                    'first_name': identity['first_name'],
                    'last_name': identity['last_name'],
                    'username': identity['username']
                }
                # Crea la struttura DataFlow_{username} solo ora
                user_dataflow_dir = get_user_documents_dataflow_dir()
                if not user_dataflow_dir:
                    logger.error("Impossibile creare la cartella utente: username mancante.")
                    SimpleMessageDialog(root, tr("Error"), tr("Unable to create user folder."), "error")
                    try:
                        root.destroy()
                    except:
                        pass
                    return
                user_db_dir = os.path.join(user_dataflow_dir, 'Database')
                os.makedirs(user_db_dir, exist_ok=True)
                user_db_name = f"dataflow_db_{identity['username']}.db"
                user_db_path = os.path.join(user_db_dir, user_db_name)
                config['Settings']['custom_db_path'] = user_db_path
                with open(config_file, 'w', encoding='utf-8') as f:
                    config.write(f)
                logger.info(f"Salvata identità utente e percorso DB: {user_db_path}")
            else:
                logger.info(f"Identità utente già salvata: {existing_identity['username']}")
        except Exception as e:
            logger.error(f"Impossibile salvare identità nel config: {e}", exc_info=True)

        # Invalida la cache e crea il DB specifico per l'utente
        reset_db_cache()
        try:
            crea_database_v4()
        except Exception as e:
            logger.error(f"Errore creazione DB utente: {e}", exc_info=True)
            SimpleMessageDialog(root, tr("Error"), tr("Unable to create user database: {}").format(e), "error")
            try:
                root.destroy()
            except:
                pass
            return

        # 3) Ora mostriamo lo splash (dopo la creazione DB) e carichiamo l'interfaccia
        splash_local = SplashScreen(root)
        splash_local.update_progress(90, tr("Loading interface..."))
        splash_local.update()

        app = MainWindow(root)
        time.sleep(0.3)

        splash_local.update_progress(100, tr("Completed!"))
        time.sleep(0.25)

        # Prepara e mostra la finestra principale
        geometry = calculate_center_position(root)
        root.geometry(geometry)
        root.deiconify()
        root.lift()
        root.attributes('-topmost', True)

        try:
            splash_local.destroy()
        except:
            pass

        # Rimuovi il topmost forzato dopo che la finestra ha il focus
        # BUG #41 FIX: usa funzione nominata invece di lambda per evitare reference leak
        def remove_topmost():
            try:
                root.attributes('-topmost', False)
            except:
                pass
        root.after(50, remove_topmost)

        root.focus_set()

    root.after(200, main_task)
    root.mainloop()

    # Riavvio post-mainloop: lancia il nuovo processo solo dopo la chiusura completa della GUI
    if _pending_restart is not None:
        _cmd, _cwd = _pending_restart
        try:
            launch_post_mainloop_restart(
                cmd=_cmd,
                cwd=_cwd,
                platform_name=sys.platform,
            )
        except Exception as e:
            logger.error(f"Errore nel riavvio post-mainloop: {e}")
