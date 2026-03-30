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
from tkinter import ttk, messagebox, filedialog, simpledialog
from tksheet import Sheet
import os
from database_manager import DatabaseManager, DatabaseError
import tempfile
from tkcalendar import DateEntry
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
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
import builtins
if not hasattr(builtins, '_'):
    builtins._ = lambda x: x
import webbrowser
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
from utils.format_utils import parse_float_from_comma_string, format_quantity_display, format_currency_display
from utils.window_utils import calculate_center_position, calculate_optimal_window_size, center_window
from utils.user_utils import get_app_data_dir, get_config_file, load_user_identity, save_user_identity
from utils.resource_utils import resource_path, set_window_icon
from utils.i18n_utils import (
    _,
    init_i18n,
    get_current_language,
    get_pos_column_text,
    get_qty_column_text,
    normalize_rfq_type,
    translate_rfq_type
)
from utils.validation_utils import sanitize_filename, format_date_for_db, format_price_display

# !!!!! IMPORTANTE: Inizializza le traduzioni PRIMA di importare moduli UI !!!!!
# I moduli UI usano _() durante l'import, quindi init_i18n() DEVE essere chiamato prima
init_i18n()

# Importa UI components (DOPO init_i18n per avere _() disponibile)
from ui.help_window import HelpWindow
from ui.kpi_window import KpiWindow
from ui.window_launchers import open_help_window, on_kpi_click
from ui.windows.view_request_window import ViewRequestWindow
from ui.components.main_dashboard_toolbar import MainDashboardToolbar
from ui.components.collapsible_filters import CollapsibleFilters
from ui.main_dashboard_builder import build_main_dashboard
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
from database.db_helpers import crea_database_v4
from ui.license_window import LicenseWindow
from ui.dialogs.common_dialogs import (
    LanguagePrompt,
    NewRdOTypeDialog,
    UserIdentityDialog,
    CopyProgressWindow,
    SplashScreen,
    SimpleYesNoDialog,
    SimpleMessageDialog
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
                self.title(_("Impostazioni e Manutenzione"))
            except Exception as e:
                logger.error(f"Errore nel settare il titolo: {e}")
                self.title(_("Impostazioni e Manutenzione"))
            self.transient(parent)
            self.grab_set()
            
            self.autobackup_enabled = tk.BooleanVar()
            self.autobackup_hour = tk.StringVar()
            self.autobackup_path = tk.StringVar()
            self.language_var = tk.StringVar()
            # Imposta un valore di default per la lingua (verrà aggiornato da load_settings)
            self.language_var.set("English")

            # Le impostazioni di visualizzazione sono ora gestite automaticamente da Windows DPI

            main_frame = ttk.Frame(self, padding="20")
            main_frame.pack(fill="both", expand=True)

            # --- Sezione Posizione DataFlow Standard ---
            dataflow_frame = ttk.LabelFrame(main_frame, text=_("Posizione DataFlow Standard"), padding=10)
            dataflow_frame.pack(fill="x", pady=(0, 15), padx=5)
            
            dataflow_label = ttk.Label(
                dataflow_frame, 
                text=_("Scegli dove salvare la cartella DataFlow (richiede riavvio)."),
                font=(None, 10),
                wraplength=480,
                justify="left"
            )
            dataflow_label.pack(anchor="w", pady=(0, 10))
            
            ttk.Button(
                dataflow_frame, 
                text=_("📁 Cambia Posizione DataFlow..."), 
                command=self.select_standard_dataflow_location
            ).pack(anchor="w")
            
            try:
                current_dataflow = get_user_documents_dataflow_dir()
                ttk.Label(
                    dataflow_frame,
                    text=_("Cartella DataFlow attuale: {}").format(current_dataflow),
                    font=(None, 9),
                    foreground="gray",
                    wraplength=480,
                    justify="left"
                ).pack(anchor="w", pady=(10, 0))
            except Exception as e:
                logger.error(f"Errore visualizzazione posizione DataFlow corrente: {e}")

            # --- Sezione Backup Manuale ---
            backup_frame = ttk.LabelFrame(main_frame, text=_("Backup Manuale"), padding="10")
            backup_frame.pack(fill="x", pady=(0, 15), padx=5)
            ttk.Label(backup_frame, text=_("Crea una copia di sicurezza immediata del database."), font=(None, 10), wraplength=500).pack(anchor="w", pady=(0, 10))
            ttk.Button(backup_frame, text=_("💾 Backup Manuale..."), command=self.backup_database).pack(anchor="w")

            # --- Sezione Backup Automatico ---
            autobackup_frame = ttk.LabelFrame(main_frame, text=_("Backup Automatico Giornaliero"), padding="10")
            autobackup_frame.pack(fill="x", pady=(0, 15), padx=5)

            ttk.Checkbutton(autobackup_frame, text=_("Abilita backup automatico giornaliero (max 3 copie)"), variable=self.autobackup_enabled).pack(anchor="w", pady=(0, 10))
            
            hour_frame = ttk.Frame(autobackup_frame)
            hour_frame.pack(fill="x", pady=5)
            ttk.Label(hour_frame, text=_("Ora:")).pack(side="left", padx=(0, 5))
            ttk.Combobox(hour_frame, textvariable=self.autobackup_hour, values=[f"{h:02}" for h in range(24)], width=5, state="readonly").pack(side="left")

            path_frame = ttk.Frame(autobackup_frame)
            path_frame.pack(fill="x", pady=5)
            ttk.Label(path_frame, text=_("Salva in:")).pack(anchor="w")
            
            path_entry_frame = ttk.Frame(autobackup_frame)
            path_entry_frame.pack(fill="x")
            ttk.Entry(path_entry_frame, textvariable=self.autobackup_path, state="readonly", width=50).pack(side="left", fill="x", expand=True, pady=(0, 5))
            ttk.Button(path_entry_frame, text=_("📁 Scegli..."), command=self.select_autobackup_path).pack(side="left", padx=(5,0), pady=(0,5))

            ttk.Button(autobackup_frame, text=_("💾 Salva Impostazioni Backup"), command=self.save_autobackup_settings).pack(pady=(10,0))

            # --- Sezione Lingua ---
            language_frame = ttk.LabelFrame(main_frame, text=_("Lingua"), padding="10")
            language_frame.pack(fill="x", pady=(0, 15), padx=5)
            
            ttk.Label(language_frame, text=_("Seleziona la lingua dell'interfaccia. Il cambio richiede il riavvio dell'applicazione."), font=(None, 10), wraplength=500).pack(anchor="w", pady=(0, 15))
            
            # Riga per il controllo della lingua
            lang_row = ttk.Frame(language_frame)
            lang_row.pack(fill="x", pady=(0, 5))
            
            ttk.Label(lang_row, text=_("Lingua:")).pack(side="left", padx=(0, 10))
            language_combo = ttk.Combobox(lang_row, textvariable=self.language_var, values=["English", "Italiano"], state="readonly", width=20)
            language_combo.pack(side="left", padx=(0, 10))
            self.language_combo = language_combo  # Salva riferimento per aggiornamento successivo
            ttk.Button(lang_row, text=_("💾 Salva Lingua"), command=self.save_language_settings).pack(side="left")
            
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
            config = configparser.ConfigParser(interpolation=None)
            config_file = get_config_file()
            config.read(config_file)
            
            # Carica impostazioni AutoBackup
            if 'AutoBackup' in config:
                try:
                    self.autobackup_enabled.set(config['AutoBackup'].getboolean('enabled', False))
                    self.autobackup_hour.set(config['AutoBackup'].get('hour', '12'))
                    self.autobackup_path.set(config['AutoBackup'].get('path', ''))
                except Exception as e:
                    logger.warning(f"Errore nel caricare impostazioni AutoBackup: {e}")
                    self.autobackup_enabled.set(False)
                    self.autobackup_hour.set("12")
                    self.autobackup_path.set("")
            else:
                self.autobackup_enabled.set(False)
                self.autobackup_hour.set("12")
                self.autobackup_path.set("")
            
            # Carica impostazioni generali
            if 'Settings' in config:
                # Carica la lingua (default 'en' per primo avvio)
                try:
                    current_lang = config.get('Settings', 'language', fallback='en')
                    # Validazione: accetta solo 'en' o 'it'
                    if current_lang not in ['en', 'it']:
                        current_lang = 'en'
                    self.language_var.set("English" if current_lang == 'en' else "Italiano")
                except Exception as e:
                    logger.warning(f"Errore nel caricare lingua: {e}")
                    self.language_var.set("English")
            else:
                # Se non c'è la sezione Settings, usa default inglese
                self.language_var.set("English")
        except Exception as e:
            logger.error(f"Errore critico nel caricare impostazioni: {e}", exc_info=True)
            # Imposta valori di default in caso di errore
            self.autobackup_enabled.set(False)
            self.autobackup_hour.set("12")
            self.autobackup_path.set("")
            self.language_var.set("English")

    # La funzione save_display_settings() è stata rimossa perché le impostazioni
    # di visualizzazione sono ora gestite automaticamente da Windows DPI

    def save_language_settings(self):
        """Salva la lingua selezionata nel config.ini."""
        try:
            config = configparser.ConfigParser(interpolation=None)
            config_file = get_config_file()
            if os.path.exists(config_file):
                config.read(config_file)
            
            if 'Settings' not in config:
                config['Settings'] = {}
            
            # Converte "English"/"Italiano" in "en"/"it"
            selected_lang = self.language_var.get()
            if not selected_lang:
                SimpleMessageDialog(self, _("Attenzione"), _("Seleziona una lingua."), "warning")
                return
            
            lang_code = "en" if selected_lang == "English" else "it"
            config['Settings']['language'] = lang_code
            
            # BUG #49 FIX: Usa encoding UTF-8 per gestire caratteri speciali
            with open(config_file, 'w', encoding='utf-8') as f:
                config.write(f)
            
            dialog = SimpleYesNoDialog(
                self,
                _("Successo"),
                _("Impostazione lingua salvata.\nRiavviare ora l'applicazione per applicare le modifiche?")
            )
            if dialog.result:
                # Riavvia l'applicazione
                self.main_app.restart_program()
        except Exception as e:
            logger.error(f"Errore nel salvare la lingua: {e}", exc_info=True)
            SimpleMessageDialog(self, _("Errore"), _("Impossibile salvare l'impostazione della lingua: {}").format(e), "error")

    def select_autobackup_path(self):
        path = filedialog.askdirectory(title=_("Seleziona cartella per backup automatici"), parent=self)
        if path: self.autobackup_path.set(path)

    def save_autobackup_settings(self):
        config = configparser.ConfigParser(interpolation=None); config.read(get_config_file())
        if 'AutoBackup' not in config: config['AutoBackup'] = {}
        config['AutoBackup']['enabled'] = str(self.autobackup_enabled.get())
        config['AutoBackup']['hour'] = self.autobackup_hour.get()
        config['AutoBackup']['path'] = self.autobackup_path.get()
        if self.autobackup_enabled.get() and not self.autobackup_path.get():
            SimpleMessageDialog(self, _("Attenzione"), _("Per abilitare il backup automatico specificare un percorso."), "warning")
            return
        try:
            # BUG #49 FIX: Usa encoding UTF-8 per gestire caratteri speciali
            with open(get_config_file(), 'w', encoding='utf-8') as f: config.write(f)
            SimpleMessageDialog(self, _("Successo"), _("Impostazioni backup salvate."), "info")
        except Exception as e:
            SimpleMessageDialog(self, _("Errore"), _("Impossibile salvare: {}").format(e), "error")

    def backup_database(self):
        """Crea backup manuale con VACUUM INTO per garantire consistenza."""
        db_file = get_db_path()
        if not os.path.exists(db_file):
            SimpleMessageDialog(self, _("Errore"), _("File database '{}' non trovato!").format(db_file), "error")
            return
        
        dest = filedialog.asksaveasfilename(
            title=_("Salva backup come..."), 
            initialfile=f"backup_manuale_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db", 
            defaultextension=".db", 
            filetypes=[(_("Database SQLite"), "*.db"), (_("Tutti i file"), "*.*")], 
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
        
        # BUG #24 FIX: Verifica che tutte le connessioni database siano chiuse prima della copia
        # Su Windows, file database con handle aperti possono causare corruzione durante copia
        try:
            # Attendi che il DB rilasci tutti i lock (max 1 secondo)
            import time
            for attempt in range(5):
                try:
                    # Test se possiamo aprire il file in modalità esclusiva
                    with open(db_file, 'r+b') as test_handle:
                        pass  # File accessibile senza lock
                    break  # Successo, esci dal loop
                except (PermissionError, IOError) as lock_error:
                    if attempt < 4:  # Non l'ultimo tentativo
                        logger.debug(f"Database ancora locked, tentativo {attempt+1}/5: {lock_error}")
                        time.sleep(0.2)  # Attendi 200ms
                    else:
                        logger.warning(f"Database potrebbe avere lock attivi dopo 5 tentativi")
            
            # ✅ COPIA FILE PRINCIPALE
            shutil.copy2(db_file, dest)
            logger.info(f"Backup DB principale: {dest}")
            
            # ✅ COPIA FILE WAL (se esiste)
            wal_file = db_file.replace('.db', '.db-wal')
            if os.path.exists(wal_file):
                wal_dest = dest.replace('.db', '.db-wal')
                shutil.copy2(wal_file, wal_dest)
                logger.info(f"Backup WAL copiato: {wal_dest}")
            else:
                logger.info("File WAL non presente (normale se DB appena chiuso)")
            
            # ✅ COPIA FILE SHM (se esiste)
            shm_file = db_file.replace('.db', '.db-shm')
            if os.path.exists(shm_file):
                shm_dest = dest.replace('.db', '.db-shm')
                shutil.copy2(shm_file, shm_dest)
                logger.info(f"Backup SHM copiato: {shm_dest}")
            else:
                logger.info("File SHM non presente (normale se DB appena chiuso)")
            
            # Verifica dimensione backup principale (sanity check)
            original_size = os.path.getsize(db_file)
            backup_size = os.path.getsize(dest)
            
            if backup_size < original_size * 0.5:
                logger.warning(f"Backup manuale potenzialmente incompleto: {backup_size} vs {original_size} bytes")
                dialog = SimpleYesNoDialog(
                    self,
                    _("Attenzione Dimensione"), 
                    _("Il backup creato è significativamente più piccolo del database originale.\n\nOriginale: {:.2f} MB\nBackup: {:.2f} MB\n\nVuoi conservarlo comunque?").format(original_size / (1024*1024), backup_size / (1024*1024))
                )
                if not dialog.result:
                    try:
                        os.remove(dest)
                        # Rimuovi anche WAL e SHM se esistono
                        wal_dest = dest.replace('.db', '.db-wal')
                        shm_dest = dest.replace('.db', '.db-shm')
                        if os.path.exists(wal_dest):
                            os.remove(wal_dest)
                        if os.path.exists(shm_dest):
                            os.remove(shm_dest)
                    except:
                        pass
                    return
            
            # Messaggio di successo con info sui file copiati
            files_copied = [os.path.basename(dest)]
            wal_dest = dest.replace('.db', '.db-wal')
            shm_dest = dest.replace('.db', '.db-shm')
            if os.path.exists(wal_dest):
                files_copied.append(os.path.basename(wal_dest))
            if os.path.exists(shm_dest):
                files_copied.append(os.path.basename(shm_dest))
            
            SimpleMessageDialog(
                self,
                _("Successo"), 
                _("Backup creato con successo:\n\nFile copiati:\n{}\n\nDimensione totale: {:.2f} MB").format(
                    '\n'.join(f'  • {f}' for f in files_copied),
                    sum(os.path.getsize(f) for f in [dest] + 
                        ([wal_dest] if os.path.exists(wal_dest) else []) + 
                        ([shm_dest] if os.path.exists(shm_dest) else [])) / (1024*1024)
                ),
                "info"
            )
            logger.info(f"Backup manuale completato: {len(files_copied)} file copiati")
            
        except Exception as e:
            logger.error(f"Errore backup manuale: {e}", exc_info=True)
            SimpleMessageDialog(
                self,
                _("Errore"), 
                _("Impossibile creare backup:\n{}").format(e),
                "error"
            )
            # Rimuovi backup parziale/corrotto
            if os.path.exists(dest):
                try:
                    os.remove(dest)
                    # Rimuovi anche WAL e SHM parziali
                    wal_dest = dest.replace('.db', '.db-wal')
                    shm_dest = dest.replace('.db', '.db-shm')
                    if os.path.exists(wal_dest):
                        os.remove(wal_dest)
                    if os.path.exists(shm_dest):
                        os.remove(shm_dest)
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
                        _("Attenzione"),
                        _("Il backup è stato completato, ma non è stato possibile riaprire la connessione principale.\nSi consiglia di riavviare l'applicazione."),
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
        
        warning_text = _(
            "⚠️ ATTENZIONE: stai per cambiare la posizione della cartella DataFlow.\n\n"
            "IMPORTANTE:\n"
            "- La cartella attuale non verrà spostata automaticamente\n"
            "- L'app verrà riavviata per applicare la modifica\n\n"
            "Posizione attuale:\n{}\n\n"
            "Vuoi procedere?"
        ).format(current_dataflow_dir)
        
        dialog = SimpleYesNoDialog(
            self,
            _("Conferma Cambio Posizione"), 
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
                title=_("Seleziona la nuova posizione della cartella DataFlow"),
                initialdir=initial_dir,
                parent=self
            )
        except Exception as e:
            logger.error(f"Errore apertura dialog selezione cartella: {e}")
            SimpleMessageDialog(
                self,
                _("Errore"),
                _("Errore durante la selezione della cartella: {}").format(e),
                "error"
            )
            return
        
        if not selected_dir:
            logger.info("Utente ha annullato la selezione della nuova posizione")
            return
        
        normalized_dir = os.path.normpath(os.path.abspath(selected_dir.strip()))
        if not normalized_dir:
            SimpleMessageDialog(self, _("Errore"), _("Percorso non valido."), "error")
            return
        
        # ✅ CORREZIONE: NON aggiungere "DataFlow" - useremo DataFlow_{username}
        # Il percorso selezionato dall'utente è la directory PARENT dove verrà creata DataFlow_{username}
        logger.info(f"Cartella parent selezionata per DataFlow: {normalized_dir}")
        
        # Verifica che la directory parent esista o possa essere creata
        try:
            os.makedirs(normalized_dir, exist_ok=True)
        except OSError as e:
            logger.error(f"Impossibile creare/accedere alla cartella parent: {e}")
            SimpleMessageDialog(
                self,
                _("Errore"),
                _("Impossibile accedere alla cartella selezionata:\n{}\n\nDettagli: {}").format(normalized_dir, e),
                "error"
            )
            return
        
        # Validazione permessi scrittura nella directory parent
        try:
            # BUG #28 FIX: Risolto TOCTOU usando try-except invece di check esistenza
            # Test scrittura nella directory parent (già esistente o appena creata)
            test_file = os.path.join(normalized_dir, ".dataflow_test_write")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
            finally:
                # BUG #28 FIX: Cleanup in finally per garantire rimozione anche se write fallisce
                try:
                    os.remove(test_file)
                except FileNotFoundError:
                    pass  # File già rimosso, va bene
            logger.info(f"Permessi verifica OK per {normalized_dir}")
        except (OSError, PermissionError) as e:
            logger.error(f"Test permessi fallito per {normalized_dir}: {e}")
            SimpleMessageDialog(
                self,
                _("Errore Permessi"),
                _("Impossibile scrivere nella cartella selezionata:\n{}\n\nDettagli: {}").format(normalized_dir, e),
                "error"
            )
            return
        
        # Controllo lunghezza
        if len(normalized_dir) > 240:
            logger.warning(f"Percorso DataFlow troppo lungo ({len(normalized_dir)} caratteri)")
            length_warning = _(
                "Il percorso selezionato è molto lungo ({} caratteri).\n"
                "Windows potrebbe avere problemi nell'accesso ai file.\n"
                "Vuoi procedere comunque?"
            ).format(len(normalized_dir))
            dialog = SimpleYesNoDialog(
                self,
                _("Percorso Molto Lungo"),
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
                removable_warning = _(
                    "⚠️ L'unità selezionata ({}) potrebbe essere rimovibile.\n"
                    "Se viene scollegata, DataFlow non potrà accedere ai dati."
                ).format(drive_letter)
                SimpleMessageDialog(self, _("Unità Rimovibile?"), removable_warning, "warning")
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
                _("Errore"),
                _("Impossibile determinare l'utente corrente. Riavvia DataFlow."),
                "error"
            )
            return
        
        # Variabili per gestione cambio username
        final_username = current_username
        username_changed = False
        
        # Loop controllo conflitto username
        while True:
            # Controlla se esiste già un database con questo username nella destinazione
            potential_folder = os.path.join(normalized_dir, f"DataFlow_{final_username}")
            potential_db = os.path.join(potential_folder, 'Database', f'dataflow_db_{final_username}.db')
            
            folder_exists = os.path.exists(potential_folder)
            db_exists = False
            
            # Controllo robusto dell'esistenza del DB (gestisce file locked)
            if folder_exists:
                try:
                    # Verifica esistenza DB in modo più robusto
                    db_exists = os.path.exists(potential_db)
                    
                    # Se il DB esiste, prova ad aprirlo per verificare che sia accessibile
                    if db_exists:
                        try:
                            # Test di accesso in lettura (non modifica il file)
                            with open(potential_db, 'rb') as f:
                                f.read(1)  # Leggi solo 1 byte per verificare accesso
                            logger.info(f"Controllo conflitto: DB '{potential_db}' esiste ed è accessibile")
                        except (PermissionError, OSError) as e:
                            # File locked o inaccessibile: CONSIDERA COME ESISTENTE
                            logger.warning(f"DB '{potential_db}' esistente ma locked/inaccessibile: {e}")
                            db_exists = True
                except Exception as e:
                    logger.error(f"Errore nel controllo esistenza DB: {e}")
                    # In caso di errore, ASSUME CHE ESISTA (principio di precauzione)
                    db_exists = True
            
            logger.info(f"Controllo conflitto per username '{final_username}': folder={folder_exists}, db={db_exists}")
            
            # ✅ CORREZIONE LOGICA: Se ESISTE cartella O database, è un CONFLITTO
            if folder_exists or db_exists:
                # Conflitto rilevato: chiedi se vuole cambiare username
                conflict_message = _(
                    "⚠️ CONFLITTO UTENTE RILEVATO\n\n"
                    "Nella cartella di destinazione selezionata esiste già un database \n"
                    "associato all'utente '{}'.\n\n"
                    "Per evitare conflitti e perdita dati, è necessario cambiare \n"
                    "il tuo username prima di procedere.\n\n"
                    "Vuoi procedere con il cambio username?"
                ).format(final_username)
                
                dialog = SimpleYesNoDialog(
                    self,
                    _("Conflitto Username"),
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
                _("Errore"),
                _("Cartella DataFlow di origine non trovata:\n{}").format(source_folder),
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
        progress_win = CopyProgressWindow(self, title=_("Copia DataFlow in corso..."))
        progress_win.update_progress(0, _("Preparazione copia..."))
        
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
            progress_win.update_progress(5, _("Analisi file da copiare..."))
            total_files = 0
            for root, dirs, files in os.walk(source_folder):
                total_files += len(files)
            
            logger.info(f"File totali da copiare: {total_files}")
            
            if total_files == 0:
                raise Exception(_("Nessun file da copiare nella cartella sorgente"))
            
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
                            _("Copia file {}/{}: {}").format(files_copied, total_files, file_name[:40])
                        )
            
            copy_with_progress(source_folder, dest_folder)
            
            logger.info(f"Copia file completata: {files_copied} file copiati")
            progress_win.update_progress(85, _("Copia completata, aggiornamento configurazione..."))
            
            # === AGGIORNA USERNAME NEL DATABASE (SOLO SE CAMBIATO) ===
            if username_changed:
                logger.info(f"Username cambiato da '{current_username}' a '{final_username}', aggiorno database")
                progress_win.update_progress(90, _("Aggiornamento username nel database..."))
                
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
            progress_win.update_progress(95, _("Salvataggio configurazione..."))
            
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
            
            progress_win.update_progress(100, _("Operazione completata!"))
            time.sleep(0.5)
            progress_win.destroy()
            
            # === MESSAGGIO SUCCESSO ===
            username_info = ""
            if username_changed:
                username_info = _("\n\n✓ Username aggiornato da '{}' a '{}'").format(current_username, final_username)
            
            success_msg = _(
                "✓ OPERAZIONE COMPLETATA CON SUCCESSO\n\n"
                "La cartella DataFlow è stata copiata con successo in:\n"
                "{dest}\n"
                "\nFile copiati: {count}{username_change}\n\n"
                "⚠️ IMPORTANTE:\n"
                "- La cartella ORIGINALE in '{src}' NON è stata eliminata.\n"
                "- Prima di eliminarla manualmente, TESTA il corretto funzionamento \n"
                "  del database copiato.\n"
                "- DataFlow verrà riavviato automaticamente."
            ).format(
                dest=dest_folder,
                count=files_copied,
                username_change=username_info,
                src=source_folder
            )
            
            SimpleMessageDialog(self, _("Operazione Completata"), success_msg, "info")
            
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
            
            error_msg = _(
                "❌ OPERAZIONE FALLITA\n\n"
                "Impossibile completare lo spostamento della cartella DataFlow.\n\n"
                "Dettaglio errore:\n{error}\n\n"
                "Le impostazioni originali sono state ripristinate.\n"
                "Consulta il file di log per maggiori dettagli."
            ).format(error=str(e))
            
            SimpleMessageDialog(self, _("Errore Spostamento"), error_msg, "error")

# ------------------------------------------------------------------------------------
# --- NUOVA FINESTRA LICENZA ---
# ------------------------------------------------------------------------------------
class LicenseWindow(tk.Toplevel):
    def __init__(self, parent, first_run=False):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("Licenza d'Uso - DataFlow Procurement Software"))
        self.transient(parent)
        self.grab_set()
        
        # Frame pulsanti (sempre in fondo)
        button_frame = ttk.Frame(self)
        button_frame.pack(side="bottom", fill="x", padx=10, pady=10)

        if first_run:
            self.accepted = False # Stato di default
            ttk.Button(button_frame, text=_("❌ Esci"), command=self.on_exit).pack(side="right")
            ttk.Button(button_frame, text=_("✅ Accetto"), command=self.on_accept).pack(side="right", padx=10)
            # Gestisce la chiusura della finestra con la 'X' come un "Esci"
            self.protocol("WM_DELETE_WINDOW", self.on_exit) 
        else:
            ttk.Button(button_frame, text=_("❌ Chiudi"), command=self.destroy).pack(side="right")
        
        # Frame contenuto (espandibile)
        main_frame = ttk.Frame(self)
        main_frame.pack(side="top", fill="both", expand=True)

        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(content_frame)
        self.text_content = tk.Text(content_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=15, pady=10, relief="flat", background="#FFFFFF", font=(None, 10))
        scrollbar.config(command=self.text_content.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self._populate_content()
        self.text_content.config(state="disabled")
        
        # Centra la finestra dopo aver aggiunto tutti i widget
        # Usiamo un after per assicurarci che il contenuto sia disegnato
        center_window(self)

    def on_accept(self):
        self.accepted = True
        self.destroy()

    def on_exit(self):
        self.accepted = False
        self.destroy()

    def _populate_content(self):
        # Configurazione degli stili di testo
        self.text_content.tag_configure("h1", font=(None, 13, "bold", "underline"), justify="center")
        self.text_content.tag_configure("h2", font=(None, 10, "bold"))
        self.text_content.tag_configure("normal", font=(None, 10))
        self.text_content.tag_configure("code", font=("Courier New", 9))
        
        # Configurazione tag per link cliccabile
        self.text_content.tag_configure("link", foreground="blue", underline=True)
        self.text_content.tag_bind("link", "<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/guido-soraru-buyer/"))
        self.text_content.tag_bind("link", "<Enter>", lambda e: self.text_content.config(cursor="hand2"))
        self.text_content.tag_bind("link", "<Leave>", lambda e: self.text_content.config(cursor=""))
        
        def add(txt, tag_keys):
            tag_tuple = tag_keys if isinstance(tag_keys, tuple) else (tag_keys,)
            self.text_content.insert(tk.END, txt, tag_tuple)

        # --- INIZIO CONTENUTO LICENZA ---
        
        add(_("Contratto di Licenza per l'Utente Finale (GNU GPLv3) - DataFlow Procurement Software\n\n"), "h1")
        
        add(_("Sviluppatore: "), "h2"); add("Guido Sorarù", ("normal", "link")); add("\n", "normal")
        add(_("E-mail: "), "h2"); add("sorguido@gmail.com\n", "normal")
        add(_("Copyright © 2025 Guido Sorarù.\n\n"), "h2")
        
        add("--------------------------------------------------\n\n", "normal")
        
        add(_("Questo software, \"DataFlow\" (di seguito \"il Software\"), è rilasciato come software open source sotto la licenza GNU General Public License versione 3 (GPLv3).\n\n"), "normal")
        
        add(_("1. CONCESSIONE DELLA LICENZA\n"), "h2")
        add(_("Lo sviluppatore concede all'utente una licenza non esclusiva per scaricare, installare, utilizzare, studiare, modificare e ridistribuire il Software in conformità con i termini della GNU General Public License versione 3.\n\n"), "normal")
        add(_("Il codice sorgente completo del Software è disponibile pubblicamente.\n\n"), "normal")
        add(_("Una copia della licenza GNU GPLv3 dovrebbe essere distribuita insieme a questo Software.\nIn caso contrario consultare: https://www.gnu.org/licenses/\n\n"), "normal")
        
        add(_("2. DISTRIBUZIONE E MODIFICA\n"), "h2")
        add(_("Il Software può essere utilizzato, studiato, modificato e ridistribuito liberamente secondo i termini della GNU General Public License versione 3.\n\n"), "normal")
        add(_("Qualsiasi ridistribuzione del Software, modificato o non modificato, deve mantenere l'avviso di copyright ed essere distribuita sotto la stessa licenza GNU GPLv3.\n\n"), "normal")
        
        add(_("3. ESCLUSIONE DI GARANZIA\n"), "h2")
        add(_("IL SOFTWARE È FORNITO \"COSÌ COM'È\" (AS IS), SENZA ALCUNA GARANZIA, ESPRESSA O IMPLICITA. LO SVILUPPATORE NON FORNISCE ALCUNA GARANZIA RIGUARDO LA COMMERCIABILITÀ, L'IDONEITÀ PER UNO SCOPO PARTICOLARE O LA NON VIOLAZIONE DI DIRITTI DI TERZI.\n"), "normal")
        add(_("L'INTERO RISCHIO DERIVANTE DALL'USO O DALLE PRESTAZIONI DEL SOFTWARE RIMANE A CARICO DELL'UTENTE.\n\n"), "normal")
        
        add(_("4. LIMITAZIONE DI RESPONSABILITÀ\n"), "h2")
        add(_("IN NESSUN CASO LO SVILUPPATORE (GUIDO SORARÙ) POTRÀ ESSERE RITENUTO RESPONSABILE PER QUALSIASI DANNO DIRETTO, INDIRETTO, INCIDENTALE, SPECIALE, ESEMPLARE O CONSEQUENZIALE (INCLUSI, A TITOLO ESEMPLIFICATIVO MA NON ESAUSIVO, DANNI PER PERDITA DI DATI, PERDITA DI PROFITTI O INTERRUZIONE DELL'ATTIVITÀ) DERIVANTE DALL'USO, DALL'USO IMPROPRIO O DALL'IMPOSSIBILITÀ DI UTILIZZARE IL SOFTWARE, ANCHE SE LO SVILUPPATORE È STATO AVVISATO DELLA POSSIBILITÀ DI TALI DANNI.\n\n"), "normal")
        
        # --- INIZIO TESTO AGGIUNTO ---
        add(_("Il Software utilizza un database SQLite con modalità WAL per ogni utente. DataFlow 2.0.0 supporta l'utilizzo multi-utente con database separati per ciascun utente, permettendo la condivisione sicura dei dati in sola lettura.\n"), "normal")
        add(_("L'utente si assume la piena responsabilità per la perdita o corruzione dei dati derivante dall'uso improprio del software.\n"), "normal")
        add(_("L'accesso simultaneo in scrittura da parte di più utenti allo stesso file di database non è supportato e causerà con alta probabilità la corruzione irreversibile dei dati. Tuttavia, l'architettura multi-utente di DataFlow garantisce che ogni utente abbia il proprio database separato, eliminando questo rischio.\n\n"), "normal")
        # --- FINE TESTO AGGIUNTO ---
        
        add(_("Utilizzando questo Software, l'utente accetta i termini e le condizioni di questa licenza.\n"), "normal")
        
        # Disabilita il widget dopo il caricamento
        self.text_content.config(state="disabled")

# ------------------------------------------------------------------------------------
# DIALOG SELEZIONE TIPO RDO
# ------------------------------------------------------------------------------------
class NewRdOTypeDialog(tk.Toplevel):
    """Dialog minimale per scegliere il tipo di RdO da creare"""
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        
        self.title(_("Nuova Richiesta di Offerta"))
        # NON usare transient() e grab_set() per evitare che la chiusura chiuda anche il parent
        self.result = None
        
        # Frame principale
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Etichetta domanda
        ttk.Label(
            main_frame, 
            text=_("Che tipo di RdO vuoi creare?"), 
            font=(None, 10)
        ).pack(pady=(0, 15))
        
        # Frame pulsanti tipo
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Button(
            btn_frame,
            text=_("📦 Fornitura piena"),
            command=lambda: self.set_result("Fornitura piena"),
            width=20
        ).pack(side="left", padx=5)
        
        ttk.Button(
            btn_frame,
            text=_("🔧 Conto lavoro"),
            command=lambda: self.set_result("Conto lavoro"),
            width=20
        ).pack(side="left", padx=5)
        
        # Pulsante annulla
        ttk.Button(
            main_frame,
            text=_("❌ Annulla"),
            command=self.destroy
        ).pack()
        
        # Gestione chiusura con X
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        
        center_window(self)
    
    def set_result(self, tipo):
        """Salva la scelta e chiude il dialog"""
        self.result = tipo
        self.destroy()

# ------------------------------------------------------------------------------------
# FINESTRA PRINCIPALE
# ------------------------------------------------------------------------------------
class MainWindow:
    def __init__(self, root):
        self.root = root;
        set_window_icon(self.root)
        self.root.title(_("DataFlow Procurement Software - Cruscotto Principale"))
        
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
        
        self.all_users_placeholder = _("Tutti gli utenti")
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
            ("Derisking", self.sheet_derisking),
        ]:
            self._load_vsm_events(event_type, sheet)
        self.populate_vsm_username_filter()

        self.refresh_data(); self.update_button_visibility(); self.check_for_autobackup()

    # --- INIZIO NUOVI METODI LICENZA ---
    def open_license_window(self):
        """Apre la finestra della licenza in modalità sola lettura."""
        # --- CORREZIONE OBBLIGATORIA ---
        # Rimuovi temporaneamente topmost per permettere alla finestra 
        # di licenza di apparire SOPRA.
        self.root.attributes('-topmost', False)
        # --- FINE CORREZIONE ---
        
        LicenseWindow(self.root, first_run=False)

    def show_first_run_license(self):
        """
        Mostra la finestra modale della licenza al primo avvio.
        Blocca l'esecuzione finché l'utente non accetta o esce.
        """
        # --- CORREZIONE OBBLIGATORIA ---
        # Rimuovi temporaneamente topmost per permettere alla finestra 
        # di licenza di apparire SOPRA.
        self.root.attributes('-topmost', False)
        # --- FINE CORREZIONE ---
        
        license_prompt = LicenseWindow(self.root, first_run=True)
        self.root.wait_window(license_prompt) # Attende che la finestra di licenza venga chiusa
        
        if not license_prompt.accepted:
            # L'utente ha cliccato "Esci" o ha chiuso la finestra
            # Usa after per evitare problemi di timing con la distruzione della finestra
            self.root.after(100, self.root.destroy)
            return False
        else:
            # L'utente ha cliccato "Accetto", salva l'impostazione
            try:
                config = configparser.ConfigParser(interpolation=None)
                config_file = get_config_file()
                if os.path.exists(config_file):
                    config.read(config_file)
                if 'Settings' not in config:
                    config['Settings'] = {}
                config['Settings']['license_accepted'] = 'True'
                # BUG #49 FIX: Usa encoding UTF-8 per gestire caratteri speciali
                with open(config_file, 'w', encoding='utf-8') as f:
                    config.write(f)
            except Exception as e:
                SimpleMessageDialog(self.root, _("Errore"), _("Impossibile salvare l'impostazione della licenza: {}\n\nIl programma continuerà, ma la licenza potrebbe riapparire al prossimo avvio.").format(e), "error")
            
            # Assicura che la finestra principale sia visibile e attiva
            self.root.deiconify()
            self.root.focus_force()
            return True
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
                SimpleMessageDialog(self.root, _("Dati mancanti"), _("Per utilizzare DataFlow devi inserire nome e cognome."), "warning")
                continue
            try:
                save_user_identity(result['first_name'], result['last_name'], result['username'])
                self._load_identity_from_config()
                self.apply_user_identity_to_ui()
                return True
            except Exception as e:
                logger.error(f"Errore salvataggio identità utente: {e}", exc_info=True)
                SimpleMessageDialog(self.root, _("Errore"), _("Impossibile salvare i dati utente: {}").format(e), "error")
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
            ("Derisking", self.sheet_derisking),
        ]:
            self._load_vsm_events(event_type, sheet)

    def _has_active_search_filters(self):
        """Verifica se ci sono filtri di ricerca attivi (escludendo username e stato)"""
        # Controlla filtri di testo
        for var in self.search_vars.values():
            if var.get().strip():
                return True
        
        # Controlla filtro tipo RdO
        if self.search_tipo.get() != _("Tutte"):
            return True
        
        # Controlla filtri data
        for entry in self.date_entries.values():
            if entry.get().strip():
                return True
        
        return False

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
        
        # Determina il percorso corretto del file Python
        script_path = None
        
        # Prova prima con __file__ (sempre disponibile quando eseguito come script)
        try:
            # __file__ è sempre disponibile quando il file viene eseguito come script
            current_file = __file__
            if current_file:
                script_path = os.path.abspath(current_file)
                if os.path.exists(script_path) and script_path.endswith('.py'):
                    # Percorso valido trovato
                    pass
        except (NameError, AttributeError):
            # __file__ non disponibile (raro, ma può accadere in alcuni contesti)
            pass
        
        # Se __file__ non è disponibile o non valido, prova sys.argv[0]
        if not script_path or not os.path.exists(script_path):
            if sys.argv[0]:
                # Se sys.argv[0] è un percorso relativo, prova a risolverlo
                if os.path.exists(sys.argv[0]):
                    script_path = os.path.abspath(sys.argv[0])
                else:
                    # Se non esiste, prova a costruire il percorso assoluto
                    # basandosi sulla directory corrente
                    possible_path = os.path.join(os.getcwd(), sys.argv[0])
                    if os.path.exists(possible_path):
                        script_path = os.path.abspath(possible_path)
                    else:
                        # Ultimo tentativo: usa il nome del file nella directory dello script
                        # (se siamo in modalità PyInstaller o MSIX)
                        if hasattr(sys, '_MEIPASS'):
                            # PyInstaller: usa sys.executable
                            script_path = sys.executable
                        else:
                            script_path = sys.argv[0]
        
        # Se ancora non abbiamo un percorso valido, usa sys.executable
        if not script_path or (not os.path.exists(script_path) and not hasattr(sys, '_MEIPASS')):
            script_path = sys.executable
        
        # Riavvia l'applicazione usando subprocess invece di os.execl
        # Questo gestisce correttamente i percorsi con spazi
        try:
            # Costruisci il comando da eseguire
            # Usa subprocess.Popen con lista di argomenti per gestire correttamente gli spazi
            if script_path.endswith('.py') or (not hasattr(sys, '_MEIPASS') and script_path != sys.executable):
                # Esecuzione come script Python
                cmd = [python, script_path]
            else:
                # Eseguibile (PyInstaller o MSIX)
                cmd = [script_path]
            
            # Imposta la working directory
            if os.path.dirname(script_path):
                cwd = os.path.dirname(script_path)
            else:
                cwd = os.getcwd()
            
            # Funzione per chiudere tutto e avviare il nuovo processo
            def do_restart():
                try:
                    # Invalida la cache del DB prima del riavvio
                    reset_db_cache()
                    
                    # Avvia il nuovo processo PRIMA di chiudere quello corrente
                    # Usa DETACHED_PROCESS su Windows per evitare che apra una nuova console
                    if sys.platform == 'win32':
                        new_process = subprocess.Popen(
                            cmd, 
                            cwd=cwd,
                            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
                            close_fds=True
                        )
                    else:
                        new_process = subprocess.Popen(cmd, cwd=cwd, start_new_session=True)
                    
                    # Attendi fino a 2 secondi che il processo si stabilizzi
                    import time
                    for _ in range(20):
                        if new_process.poll() is None:  # Processo ancora in esecuzione
                            break
                        time.sleep(0.1)
                    
                    # Piccolo delay aggiuntivo per sicurezza
                    time.sleep(0.2)
                    
                    # Chiudi tutte le finestre Tkinter
                    if hasattr(self, 'root') and self.root:
                        try:
                            # Distruggi tutte le finestre Toplevel
                            for widget in self.root.winfo_children():
                                if isinstance(widget, tk.Toplevel):
                                    try:
                                        widget.destroy()
                                    except:
                                        pass
                            # Esci dal mainloop
                            self.root.quit()
                            # Distruggi la root
                            self.root.destroy()
                        except:
                            pass
                    
                    # Forza la terminazione immediata del processo
                    # Usa os._exit() invece di sys.exit() per evitare che il cleanup blocchi
                    os._exit(0)
                    
                except Exception as e:
                    logger.error(f"Errore nel riavvio dell'applicazione: {e}")
                    try:
                        messagebox.showerror(
                            _("Errore"),
                            _("Impossibile riavviare l'applicazione automaticamente.\n\nPerfavore, chiudi e riapri manualmente l'applicazione per applicare le modifiche.\n\nPercorso tentato: {}").format(script_path),
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
            messagebox.showerror(
                _("Errore"),
                _("Impossibile riavviare l'applicazione automaticamente.\n\nPerfavore, chiudi e riapri manualmente l'applicazione per applicare le modifiche.\n\nPercorso tentato: {}").format(script_path),
                parent=self.root if hasattr(self, 'root') else None
            )

    def check_for_autobackup(self):
        config = configparser.ConfigParser(interpolation=None); config.read(get_config_file())
        if config.getboolean('AutoBackup', 'enabled', fallback=False):
            # BUG #38 FIX: Strip whitespace da valori config per evitare path con spazi invisibili
            path = config.get('AutoBackup', 'path', fallback='').strip()
            hour = config.get('AutoBackup', 'hour', fallback='').strip()
            if path and hour:
                try:
                    now = datetime.now()
                    if now.hour == int(hour) and now.date() != self.last_backup_date:
                        self.perform_autobackup(path); self.last_backup_date = now.date()
                except Exception as e: print(f"ERRORE AUTOBACKUP: {e}")
        
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
            return
        
        self._backup_in_progress = True
        
        try:
            if not os.path.exists(db_file): 
                logger.warning(f"File database non trovato per backup: {db_file}")
                return
            
            # Attendi un momento per permettere sync su disco
            import time
            time.sleep(0.2)
            
            # Gestione vecchi backup (mantieni solo gli ultimi 3 SET completi)
            # Un SET = .db + .db-wal + .db-shm con lo stesso timestamp
            backup_sets = {}  # timestamp -> [file_path1, file_path2, ...]
            
            for ext in ['*.db', '*.db-wal', '*.db-shm']:
                pattern = os.path.join(dest_folder, f"*_backup_auto_{ext.replace('*', '')}")
                for filepath in glob.glob(pattern):
                    # Estrai timestamp dal nome file (es: gestione_offerte_backup_auto_20250102_143000.db)
                    basename = os.path.basename(filepath)
                    try:
                        # Pattern: *_backup_auto_YYYYMMDD_HHMMSS.ext
                        timestamp_part = basename.split('_backup_auto_')[1].rsplit('.', 1)[0]
                        if timestamp_part not in backup_sets:
                            backup_sets[timestamp_part] = []
                        backup_sets[timestamp_part].append(filepath)
                    except (IndexError, ValueError):
                        logger.warning(f"Formato nome backup non riconosciuto: {basename}")
            
            # Ordina i set per timestamp e mantieni solo gli ultimi 3
            sorted_timestamps = sorted(backup_sets.keys())
            while len(sorted_timestamps) > 3:
                old_timestamp = sorted_timestamps.pop(0)
                for old_file in backup_sets[old_timestamp]:
                    try:
                        os.remove(old_file)
                        logger.info(f"Rimosso vecchio backup: {old_file}")
                    except Exception as e:
                        logger.warning(f"Impossibile eliminare vecchio backup {old_file}: {e}")
            
            # Genera timestamp per il nuovo set di backup
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_name = f"gestione_offerte_backup_auto_{timestamp}"
            
            # ✅ COPIA FILE PRINCIPALE
            dest_path = os.path.join(dest_folder, f"{base_name}.db")
            
            # Copia con retry su errori temporanei
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    shutil.copy2(db_file, dest_path)
                    break
                except (PermissionError, OSError) as e:
                    if attempt < max_retries - 1:
                        logger.warning(f"Tentativo backup {attempt+1} fallito: {e}, riprovo...")
                        time.sleep(1)
                    else:
                        raise
            
            logger.info(f"Backup automatico DB principale: {dest_path}")
            
            # ✅ COPIA FILE WAL (se esiste)
            wal_file = db_file.replace('.db', '.db-wal')
            if os.path.exists(wal_file):
                wal_dest = os.path.join(dest_folder, f"{base_name}.db-wal")
                try:
                    shutil.copy2(wal_file, wal_dest)
                    logger.info(f"Backup WAL copiato: {wal_dest}")
                except Exception as e:
                    logger.warning(f"Impossibile copiare WAL: {e}")
            else:
                logger.info("File WAL non presente per autobackup (normale se DB chiuso)")
            
            # ✅ COPIA FILE SHM (se esiste)
            shm_file = db_file.replace('.db', '.db-shm')
            if os.path.exists(shm_file):
                shm_dest = os.path.join(dest_folder, f"{base_name}.db-shm")
                try:
                    shutil.copy2(shm_file, shm_dest)
                    logger.info(f"Backup SHM copiato: {shm_dest}")
                except Exception as e:
                    logger.warning(f"Impossibile copiare SHM: {e}")
            else:
                logger.info("File SHM non presente per autobackup (normale se DB chiuso)")
            
            # Verifica integrità backup principale
            original_size = os.path.getsize(db_file)
            backup_size = os.path.getsize(dest_path)
            
            if backup_size < original_size * 0.5:
                logger.error(f"Backup automatico potenzialmente corrotto: {backup_size} vs {original_size} bytes")
                # Elimina backup corrotto (tutti i file del set)
                try:
                    os.remove(dest_path)
                    wal_dest = os.path.join(dest_folder, f"{base_name}.db-wal")
                    shm_dest = os.path.join(dest_folder, f"{base_name}.db-shm")
                    if os.path.exists(wal_dest):
                        os.remove(wal_dest)
                    if os.path.exists(shm_dest):
                        os.remove(shm_dest)
                except:
                    pass
            else:
                # Conta i file effettivamente copiati
                files_copied = 1  # DB principale
                wal_dest = os.path.join(dest_folder, f"{base_name}.db-wal")
                shm_dest = os.path.join(dest_folder, f"{base_name}.db-shm")
                if os.path.exists(wal_dest):
                    files_copied += 1
                if os.path.exists(shm_dest):
                    files_copied += 1
                
                total_size = sum(os.path.getsize(f) for f in [dest_path] + 
                               ([wal_dest] if os.path.exists(wal_dest) else []) + 
                               ([shm_dest] if os.path.exists(shm_dest) else []))
                
                logger.info(f"Backup automatico completato: {files_copied} file copiati, {total_size} bytes totali ({total_size/original_size*100:.1f}% dimensione originale)")
            
        except Exception as e:
            logger.error(f"Errore backup automatico: {e}", exc_info=True)
            print(f"ERRORE AUTOBACKUP: {e}")
        finally:
            self._backup_in_progress = False

    def open_help_window(self): open_help_window(self)
    def open_settings_window(self): self.root.wait_window(SettingsWindow(self.root, self))

    def on_kpi_click(self): on_kpi_click(self)
    
    def create_request_treeview(self, parent):
        # Frame per contenere il Sheet
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)
        
        # Crea il widget tksheet invece di Treeview
        sheet = Sheet(tree_frame,
                     theme="light blue",
                     header_font=("Calibri", 11, "bold"),
                     font=("Calibri", 11, "normal"),
                     headers=[_("Num RdO"), _("Tipo RdO"), _("Data Emiss."), _("Data Scad."), _("Riferimento"), _("Utente")],
                     show_header=True,
                     show_row_index=False)
        
        # Configura le larghezze delle colonne
        sheet.set_column_widths([80, 120, 120, 120, 300, 140])
        
        # Centra tutte le colonne tranne "Riferimento"
        sheet.align_columns(columns=[0, 1, 2, 3, 5], align="center")
        
        # Abilita tutti i binding
        sheet.enable_bindings()
        
        # Rendi il sheet completamente in sola lettura (nessuna cella editabile)
        for col_idx in range(6):
            sheet.readonly_columns(columns=[col_idx], readonly=True)
        
        # Configura il binding per doppio click su cella (metodo nativo tksheet)
        # Questo si attiva quando si fa doppio click su qualsiasi cella della riga
        sheet.extra_bindings("cell_select", self.create_cell_select_handler(sheet))
        sheet.extra_bindings("row_select", self.create_row_select_handler(sheet))
        
        # Variabile per tracciare il tempo dell'ultimo click (per doppio click)
        sheet._last_click_time = 0
        sheet._last_click_row = None
        
        # Binding generico per gestire il doppio click
        sheet.bind("<Double-Button-1>", lambda event: self.on_sheet_double_click(sheet, event))
        
        sheet.pack(fill="both", expand=True)
        
        # Salva riferimento per uso successivo
        sheet._sheet_data = []  # Per memorizzare i dati attuali
        
        return sheet
    
    def create_cell_select_handler(self, sheet):
        """Crea un handler per il doppio click su celle"""
        def handler(event_data):
            # Quando viene selezionata una cella, aggiorna i pulsanti
            self.update_button_visibility()
        return handler
    
    def create_row_select_handler(self, sheet):
        """Crea un handler per la selezione di righe"""
        def handler(event_data):
            # Quando viene selezionata una riga, aggiorna i pulsanti
            self.update_button_visibility()
        return handler

    def _create_vsm_event_sheet(self, parent, event_type=None):
        """
        Crea un tksheet per visualizzare eventi VSM.
        
        ESTRATTO da VSMManagementWindow._create_event_sheet() (Step 4B).
        Pattern identico a create_request_treeview() per coerenza visiva.
        
        Args:
            parent: Widget parent (tab frame)
            event_type: Tipo evento ("Saving"|"Cost Avoidance"|None)
                        Determina le intestazioni della colonna valore.
        
        Returns:
            Sheet: Widget tksheet configurato con colonne VSM
        """
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        
        # Intestazioni e layout dipendono dal tab
        if event_type == "Saving":
            headers = [
                _("Data"), _("Tipo"), _("Azione"), _("Descrizione"),
                _("Theoretical Savings"), _("Actual Savings"),
                _("Realizzo %"), _("Variance %"), _("Ripetitivo"), _("Utente")
            ]
            align_cols = [0, 1, 2, 4, 5, 6, 7, 8, 9]
            n_cols = 10
        elif event_type == "Cost Avoidance":
            headers = [
                _("Data"), _("Tipo"), _("Azione"), _("Descrizione"),
                _("CA Theoretical"), _("CA Actual"),
                _("Realizzo %"), _("Variance %"), _("Ripetitivo"), _("Utente")
            ]
            align_cols = [0, 1, 2, 4, 5, 6, 7, 8, 9]
            n_cols = 10
        else:
            headers = [
                _("Data"), _("Nuovo Fornitore"), _("Descrizione"),
                _("Ripetitivo"), _("Utente")
            ]
            align_cols = [0, 3, 4]
            n_cols = 5

        # Calcola larghezze colonne dinamicamente dall'header (eccetto Descrizione)
        # Usa il font degli header (Calibri 11 bold) per misurare il testo reale visualizzato
        _HEADER_PADDING = 30  # px extra per evitare header troppo "tirati"
        # Derisking: Descrizione a col 2; Saving/CA: Descrizione a col 3
        _DESC_COL_IDX = 2 if event_type is None else 3
        _DESC_COL_WIDTH = 400
        _DATE_COL_IDX = 0
        # Saving/CA: Azione a col 2, Tipo a col 1; Derisking: nessuna di queste colonne
        _ACTION_COL_IDX = 2 if event_type is not None else None
        _ACTION_MIN_WIDTH = 150  # "Negoziazione" è il valore più lungo atteso (~115px + padding)
        _TYPE_COL_IDX = 1 if event_type is not None else None
        # New Supplier solo nel tab Derisking (event_type=None), in colonna 1
        _NEW_SUPPLIER_COL_IDX = 1 if event_type is None else None
        try:
            import tkinter.font as tkfont
            _hfont = tkfont.Font(family="Calibri", size=11, weight="bold")
            _cfont = tkfont.Font(family="Calibri", size=11)  # font celle (normal)
            # Larghezze minime ricavate da colonne di riferimento già ben calibrate:
            # - "Data": spazio per "dd/mm/YYYY" (contenuto celle) + padding
            # - "Tipo" deve contenere "Derisking" → larghezza ≥ colonna "Realizzo %"
            # - "Nuovo Fornitore" ha contenuto medio → larghezza ≥ colonna "Valore Teorico"
            _date_min = _cfont.measure("dd/mm/YYYY") + _HEADER_PADDING
            _type_min = _hfont.measure(_("Realizzo %")) + _HEADER_PADDING
            _new_supplier_min = _hfont.measure(_("Valore Teorico")) + _HEADER_PADDING
            col_widths = [
                _DESC_COL_WIDTH if i == _DESC_COL_IDX
                else max(_date_min, _hfont.measure(h) + _HEADER_PADDING) if i == _DATE_COL_IDX
                else max(_ACTION_MIN_WIDTH, _hfont.measure(h) + _HEADER_PADDING) if i == _ACTION_COL_IDX
                else max(_type_min, _hfont.measure(h) + _HEADER_PADDING) if i == _TYPE_COL_IDX
                else max(_new_supplier_min, _hfont.measure(h) + _HEADER_PADDING) if i == _NEW_SUPPLIER_COL_IDX
                else max(60, _hfont.measure(h) + _HEADER_PADDING)
                for i, h in enumerate(headers)
            ]
        except Exception:
            # Fallback conservativo alle larghezze originali se tkfont non disponibile
            col_widths = [400 if i == _DESC_COL_IDX else 150 if i in (_ACTION_COL_IDX, _TYPE_COL_IDX) else 120 for i in range(len(headers))]

        # Crea widget tksheet con colonne VSM
        sheet = Sheet(
            frame,
            theme="light blue",
            header_font=("Calibri", 11, "bold"),
            font=("Calibri", 11, "normal"),
            headers=headers,
            show_header=True,
            show_row_index=False
        )
        
        # Salva il tipo di tab per uso in _populate_vsm_sheet e _export_vsm_excel
        sheet._vsm_event_type = event_type
        
        # Salva intestazioni tradotte per uso in _export_vsm_excel
        sheet._vsm_headers = headers
        
        # Salva larghezze calcolate per riapplicarle dopo set_sheet_data()
        sheet._vsm_col_widths = col_widths
        
        # Salva colonne centrate per riapplicarle dopo set_sheet_data() / set_column_widths()
        sheet._vsm_align_cols = align_cols
        
        # Configura larghezze colonne
        sheet.set_column_widths(col_widths)
        
        # Centra colonne numeriche e date (Descrizione rimane left-aligned)
        sheet.align_columns(columns=align_cols, align="center")
        
        # Abilita bindings
        sheet.enable_bindings()
        
        # Step 4D.1: Binding per aggiornamento stato pulsante Actions
        sheet.extra_bindings("cell_select", self.create_cell_select_handler(sheet))
        sheet.extra_bindings("row_select", self.create_row_select_handler(sheet))
        
        # Step 4D.4: Binding per doppio click (apre edit evento VSM)
        sheet.bind("<Double-Button-1>", lambda event: self._on_vsm_sheet_double_click(sheet, event))
        
        # Rendi readonly
        for col_idx in range(n_cols):
            sheet.readonly_columns(columns=[col_idx], readonly=True)
        
        sheet.pack(fill="both", expand=True)
        
        # Metadata storage (come nell'originale)
        sheet._event_metadata = []  # Lista di dict con event_id, username, is_mine
        
        return sheet

    # NOTA: I metodi sort_treeview_column e update_sort_indicators sono stati rimossi
    # perché tksheet ha funzionalità di ordinamento integrate che si abilitano automaticamente
    # con enable_bindings(). L'utente può cliccare sugli header delle colonne per ordinare.

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
            SimpleMessageDialog(self.root, _("Errore Database"), _("Impossibile caricare gli eventi VSM: {}\n").format(e), "error")

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
        with DatabaseManager(get_db_path()) as db_manager:
            if vsm_username_filter is None:
                # Tutti gli utenti: aggregazione multi-DB
                raw = db_manager.get_all_vsm_events_aggregated(get_db_path())
                all_events = [ev for ev, _im, _src in raw]
                extra_meta = [{'is_mine': im, 'source_file': src} for _, im, src in raw]
            elif vsm_username_filter == (self.current_username or '').lower():
                # Utente corrente: path locale ottimizzato (nessuna aggregazione)
                all_events = db_manager.get_all_vsm_events(username=self.current_username)
                extra_meta = None
            else:
                # Altro utente specifico: aggregazione con filtro username
                raw = db_manager.get_all_vsm_events_aggregated(get_db_path(), username=vsm_username_filter)
                all_events = [ev for ev, _im, _src in raw]
                extra_meta = [{'is_mine': im, 'source_file': src} for _, im, src in raw]
        return all_events, extra_meta

    def _apply_vsm_filters(self, events, event_type, extra_meta=None):
        """Applica i filtri VSM avanzati (data, azione, ripetitivo, importi) a una lista di eventi.

        Chiamato dopo il filtro per event_type in _load_vsm_events e _search_vsm_events.
        I filtri vuoti vengono ignorati. Restituisce una tupla (filtered_events, filtered_meta)
        con extra_meta allineato agli eventi filtrati (None se non era fornito).
        """
        # Raccolta valori filtro (con guard per inizializzazione parziale)
        _de_from = getattr(self, 'vsm_date_from_entry', None)
        date_from_str = _de_from.get().strip() if _de_from else ""
        _de_to = getattr(self, 'vsm_date_to_entry', None)
        date_to_str = _de_to.get().strip() if _de_to else ""
        _av = getattr(self, 'vsm_action_var', None)
        action_filter = _av.get().strip() if _av else ""
        _rv = getattr(self, 'vsm_repetitive_var', None)
        repetitive_filter = _rv.get().strip() if _rv else ""
        _tfv = getattr(self, 'vsm_theoretical_from_var', None)
        theoretical_from_str = _tfv.get().strip() if _tfv else ""
        _ttv = getattr(self, 'vsm_theoretical_to_var', None)
        theoretical_to_str = _ttv.get().strip() if _ttv else ""
        _afv = getattr(self, 'vsm_actual_from_var', None)
        actual_from_str = _afv.get().strip() if _afv else ""
        _atv = getattr(self, 'vsm_actual_to_var', None)
        actual_to_str = _atv.get().strip() if _atv else ""

        # Short-circuit: nessun filtro attivo
        if not any([date_from_str, date_to_str, action_filter, repetitive_filter,
                    theoretical_from_str, theoretical_to_str, actual_from_str, actual_to_str]):
            return events, extra_meta

        # Parse date range
        date_from = date_to = None
        _FMT = '%d/%m/%Y'
        try:
            if date_from_str:
                date_from = datetime.strptime(date_from_str, _FMT).date()
        except ValueError:
            pass
        try:
            if date_to_str:
                date_to = datetime.strptime(date_to_str, _FMT).date()
        except ValueError:
            pass

        # Parse importi (accetta sia "10.000,50" che "10000.50")
        def _parse_amount(s):
            if not s:
                return None
            s = s.strip()
            if ',' in s:
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
            try:
                return float(s)
            except ValueError:
                return None

        theoretical_from = _parse_amount(theoretical_from_str)
        theoretical_to = _parse_amount(theoretical_to_str)
        actual_from = _parse_amount(actual_from_str)
        actual_to = _parse_amount(actual_to_str)

        use_dual_value = event_type in ("Saving", "Cost Avoidance")
        meta_iter = extra_meta if extra_meta is not None else [None] * len(events)
        filtered_pairs = []

        for event, meta in zip(events, meta_iter):
            # Filtro data
            if event.event_date:
                ev_date = event.event_date.date() if hasattr(event.event_date, 'date') else event.event_date
                if date_from and ev_date < date_from:
                    continue
                if date_to and ev_date > date_to:
                    continue
            elif date_from or date_to:
                continue  # Evento senza data escluso se filtro data attivo

            # Filtro azione (solo Saving/CA; in Derisking l'azione è sempre "Derisking")
            if action_filter and use_dual_value:
                if _(event.action) != action_filter:
                    continue

            # Filtro ripetitivo
            if repetitive_filter:
                want = repetitive_filter == _("Sì")
                if event.opex_ripetitivo != want:
                    continue

            # Filtro importo teorico
            if theoretical_from is not None or theoretical_to is not None:
                tval = event.calculate_theoretical_value()
                if theoretical_from is not None and tval < theoretical_from:
                    continue
                if theoretical_to is not None and tval > theoretical_to:
                    continue

            # Filtro importo effettivo (solo Saving/CA)
            if use_dual_value and (actual_from is not None or actual_to is not None):
                aval = event.calculate_effective_value()
                if actual_from is not None and aval < actual_from:
                    continue
                if actual_to is not None and aval > actual_to:
                    continue

            filtered_pairs.append((event, meta))

        filtered_events = [p[0] for p in filtered_pairs]
        filtered_meta = [p[1] for p in filtered_pairs] if extra_meta is not None else None
        return filtered_events, filtered_meta

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
                    _(event.action),
                    (event.description or event.reference or "")[:50],
                    format_currency_display(valore_teorico),
                    format_currency_display(valore_effettivo),
                    f"{event.percent_realizzo:.2f}".replace('.', ',') + "%",
                    _variance_str,
                    "✓" if event.opex_ripetitivo else "",
                    event.username
                ]
            else:
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
                _header_w = _hfont.measure(_("Nuovo Fornitore")) + 30
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
            SimpleMessageDialog(self.root, _("Nessuna Selezione"), _("Seleziona un evento da modificare."), "warning")
            return
        
        if len(selected_rows) > 1:
            SimpleMessageDialog(self.root, _("Selezione Multipla"), _("Seleziona un solo evento per la modifica."), "warning")
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
            event_type_map = {
                'vsm_saving': 'Saving',
                'vsm_cost_avoidance': 'Cost Avoidance',
                'vsm_derisking': 'Derisking'
            }
            event_type = event_type_map.get(status)
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
        event_type_map = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking'
        }
        event_type = event_type_map.get(status)
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
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile aprire il form: {}").format(e), "error")
    
    def _delete_vsm_events(self):
        """Handler per eliminazione eventi VSM.
        
        Step 4D.3: Implementazione completa con delete_event_and_impacts.
        Pattern estratto da VSMManagementWindow.on_delete_event().
        """
        sheet, status = self.get_current_tree_and_status()
        if not status.startswith('vsm_'):
            return
        
        # Ottieni selezione
        selected_rows = self._get_selected_row_indices(sheet)
        
        if not selected_rows:
            SimpleMessageDialog(self.root, _("Nessuna Selezione"), _("Seleziona uno o più eventi da eliminare."), "warning")
            return
        
        # Raccolta event_id e validazione ownership
        events_to_delete = []
        for row_idx in selected_rows:
            if row_idx >= len(sheet._event_metadata):
                continue
            
            metadata = sheet._event_metadata[row_idx]
            
            # Valida ownership
            if not metadata['is_mine']:
                SimpleMessageDialog(self.root, _("Operazione Non Consentita"), _("Puoi eliminare solo i tuoi eventi VSM.\nAlcuni eventi selezionati appartengono ad altri utenti."), "error")
                return
            
            events_to_delete.append(metadata['event_id'])
        
        if not events_to_delete:
            return
        
        # Conferma eliminazione
        count = len(events_to_delete)
        if not SimpleYesNoDialog(self.root, _("Conferma Eliminazione"), _("Sei sicuro di voler eliminare {} evento(i) VSM?\nQuesta operazione non può essere annullata.").format(count)).result:
            return
        
        # Determina event_type da status
        event_type_map = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking'
        }
        event_type = event_type_map.get(status)
        if not event_type:
            return  # Fail-safe
        
        # Elimina eventi
        from services.vsm_persistence import delete_event_and_impacts, VSMError
        
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                for event_id in events_to_delete:
                    delete_event_and_impacts(db_manager, event_id)
            
            SimpleMessageDialog(self.root, _("Successo"), _("{} evento(i) VSM eliminato(i) con successo.").format(count), "info")
            
            # Refresh
            self._load_vsm_events(event_type, sheet)
            logger.info(f"Eliminati {count} eventi VSM con successo")
        
        except (DatabaseError, VSMError) as e:
            logger.error(f"Errore eliminazione eventi VSM: {e}")
            SimpleMessageDialog(self.root, _("Errore Eliminazione"), _("Impossibile eliminare gli eventi:\n{}").format(e), "error")

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
            SimpleMessageDialog(self.root, _("Selezione mancante"), _("Selezionare un evento VSM da duplicare."), "warning")
            return
        
        if len(selected_rows) > 1:
            SimpleMessageDialog(self.root, _("Selezione non valida"), _("Seleziona un solo evento VSM per duplicarlo."), "warning")
            return
        
        row_idx = selected_rows[0]
        
        # Validazione ownership
        if row_idx >= len(sheet._event_metadata):
            logger.error(f"Indice VSM {row_idx} fuori range metadata")
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile identificare l'evento selezionato."), "error")
            return
        
        metadata = sheet._event_metadata[row_idx]
        is_mine = metadata.get('is_mine', False)
        
        if not is_mine:
            SimpleMessageDialog(self.root, _("Operazione Non Consentita"), _("Non puoi duplicare eventi VSM di altri utenti.\nPuoi operare solo sui tuoi eventi."), "error")
            logger.warning(f"Tentativo duplicazione evento VSM altrui bloccato: utente={self.current_username}")
            return
        
        event_id = metadata.get('event_id')
        if not event_id:
            logger.error(f"event_id mancante in metadata per riga {row_idx}")
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile identificare l'evento selezionato."), "error")
            return
        
        # Recupero evento completo dal backend
        try:
            # Lazy import per evitare dipendenze circolari
            from services.vsm_persistence import (
                get_event_with_impacts,
                save_event_with_impacts,
                VSMError
            )
            
            with DatabaseManager(get_db_path()) as db_manager:
                # Recupera evento originale (con impatti, ma useremo solo l'evento)
                original_event, _impacts = get_event_with_impacts(db_manager, event_id)
                
                logger.info(
                    f"Duplicazione evento VSM {event_id}: "
                    f"tipo={original_event.event_type}, data={original_event.event_date}"
                )
                
                # Crea copia 1:1: stesso evento, id=None per nuovo insert
                # Il dataclass VSMEvent supporta costruzione da attributi
                from models.vsm_event import VSMEvent
                duplicate_event = VSMEvent(
                    id=None,  # Nuovo ID verrà assegnato dal DB
                    event_date=original_event.event_date,
                    username=original_event.username,
                    buyer=original_event.buyer,
                    event_type=original_event.event_type,
                    action=original_event.action,
                    description=original_event.description,
                    reference=original_event.reference,
                    importo_bdg=original_event.importo_bdg,
                    importo_negoziato=original_event.importo_negoziato,
                    importo_richiesto_iniziale=original_event.importo_richiesto_iniziale,
                    quantita_annua=original_event.quantita_annua,
                    percent_realizzo=original_event.percent_realizzo,
                    driver=original_event.driver,
                    giorni_pagamento_attuali=original_event.giorni_pagamento_attuali,
                    giorni_pagamento_negoziati=original_event.giorni_pagamento_negoziati,
                    spending_annuo=original_event.spending_annuo,
                    opex_ripetitivo=original_event.opex_ripetitivo,
                    note=original_event.note,
                    # created_at e updated_at saranno impostati automaticamente dal DB
                )
                
                # Salva copia (genera impatti automaticamente)
                new_event_id = save_event_with_impacts(db_manager, duplicate_event)
                
                logger.info(f"Evento VSM duplicato: {event_id} → {new_event_id}")
            
            # Status mapping per refresh (fallback safe)
            event_type_map = {
                'vsm_saving': 'Saving',
                'vsm_cost_avoidance': 'Cost Avoidance',
                'vsm_derisking': 'Derisking'
            }
            event_type = event_type_map.get(status)
            
            if event_type:
                # Auto-refresh sheet
                self._load_vsm_events(event_type, sheet)
                
                # Success feedback
                SimpleMessageDialog(self.root, _("Successo"), _("Evento VSM duplicato."), "info")
            else:
                logger.warning(f"Tipo evento non riconosciuto per refresh: {status}")
                SimpleMessageDialog(self.root, _("Successo"), _("Evento VSM duplicato. Aggiorna manualmente per vedere la copia."), "info")
        
        except VSMError as e:
            logger.error(f"Errore VSM durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore VSM"), _("Impossibile duplicare l'evento:\n{}").format(e), "error")
        except DatabaseError as e:
            logger.error(f"Errore database durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore Database"), _("Impossibile duplicare l'evento:\n{}").format(e), "error")
        except Exception as e:
            logger.error(f"Errore imprevisto durante duplicazione evento {event_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile duplicare l'evento:\n{}").format(e), "error")

    def _get_selected_row_indices(self, sheet):
        """
        Metodo helper per ottenere gli indici delle righe selezionate dal sheet.
        Gestisce sia la selezione di celle che di righe complete.
        Restituisce una lista di indici di riga.
        """
        row_indices = []
        
        # Metodo 1: Prova con get_currently_selected (per selezione cella singola)
        currently_selected = sheet.get_currently_selected()
        if currently_selected:
            if hasattr(currently_selected, 'row') and currently_selected.row is not None:
                row_indices.append(currently_selected.row)
            elif isinstance(currently_selected, tuple) and len(currently_selected) >= 1:
                row_indices.append(currently_selected[0])
        
        # Metodo 2: Prova con get_selected_rows (per selezione righe multiple)
        if not row_indices:
            selected_rows = sheet.get_selected_rows()
            if selected_rows:
                if isinstance(selected_rows, (list, set, tuple)):
                    row_indices.extend(selected_rows)
                else:
                    row_indices.append(selected_rows)
        
        return row_indices
    
    def _check_if_all_selected_are_mine(self, sheet, selected_indices):
        """Verifica se tutte le RfQ selezionate appartengono all'utente corrente.
        
        Args:
            sheet: Il widget Sheet da controllare
            selected_indices: Lista di indici riga selezionati
        
        Returns:
            bool: True se tutte le RfQ selezionate sono dell'utente corrente, False altrimenti
        """
        if not selected_indices:
            return False
        
        # Se i metadati non sono disponibili, per sicurezza blocca le operazioni su RfQ legacy
        if not hasattr(sheet, '_sheet_rows_metadata'):
            logger.warning("Metadati sheet non disponibili - blocco operazioni per sicurezza")
            return False
        
        for idx in selected_indices:
            # Salta indici fuori range
            if idx >= len(sheet._sheet_rows_metadata):
                logger.warning(f"Indice {idx} fuori range metadati (len={len(sheet._sheet_rows_metadata)})")
                continue
            
            metadata = sheet._sheet_rows_metadata[idx]
            is_mine = metadata.get('is_mine', False)  # Default False per sicurezza
            
            if not is_mine:
                return False  # Almeno una RfQ non è mia
        
        return True  # Tutte le RfQ selezionate sono mie
    
    def _check_if_all_vsm_events_are_mine(self, sheet, selected_indices):
        """Verifica se tutti gli eventi VSM selezionati appartengono all'utente corrente.
        
        Args:
            sheet: Il widget Sheet VSM da controllare
            selected_indices: Lista di indici riga selezionati
        
        Returns:
            bool: True se tutti gli eventi VSM selezionati sono dell'utente corrente, False altrimenti
        """
        if not selected_indices:
            return False
        
        # Se i metadati VSM non sono disponibili, blocca le operazioni
        if not hasattr(sheet, '_event_metadata'):
            logger.warning("Metadati VSM non disponibili - blocco operazioni per sicurezza")
            return False
        
        for idx in selected_indices:
            # Salta indici fuori range
            if idx >= len(sheet._event_metadata):
                logger.warning(f"Indice VSM {idx} fuori range metadati (len={len(sheet._event_metadata)})")
                continue
            
            metadata = sheet._event_metadata[idx]
            is_mine = metadata.get('is_mine', False)  # Default False per sicurezza
            
            if not is_mine:
                return False  # Almeno un evento non è mio
        
        return True  # Tutti gli eventi selezionati sono miei
    
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
            SimpleMessageDialog(self.root, _("Operazione Non Consentita"), _("Non puoi modificare lo stato di RfO di altri utenti.\nPuoi operare solo sulle tue RdO."), "error")
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
            # Usa db_manager per aggiornare lo stato
            params = [(new_status, req_id) for req_id in ids]
            # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
            with DatabaseManager(get_db_path()) as db_manager:
                db_manager.update_stato_richieste(params)
        except DatabaseError as e:
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile aggiornare stato: {}").format(e), "error")
        else:
            self.refresh_data()

    def on_tab_changed(self, event):
        self.update_button_visibility()
        self.clear_selection()
        self._update_filter_panel_for_current_tab()

    def _update_filter_panel_for_current_tab(self):
        self.dashboard_controller._update_filter_panel_for_current_tab()
    def update_button_visibility(self):
        """Aggiorna lo stato del pulsante Actions in base alla selezione e proprietà delle RfQ"""
        sheet, status = self.get_current_tree_and_status()
        
        if sheet is None:
            self.btn_actions.config(state="disabled")
            return
        
        # Step 4D.1/4D.2: Gestione abilitazione pulsante Actions per VSM
        if status.startswith('vsm_'):
            # Per VSM: abilita Actions se c'è almeno una riga selezionata
            selected_rows_indices = self._get_selected_row_indices(sheet)
            has_selection = bool(selected_rows_indices)
            num_selected = len(selected_rows_indices) if selected_rows_indices else 0
            
            # Step 4D.2: Verifica ownership per calcolare capacità
            all_mine = self._check_if_all_vsm_events_are_mine(sheet, selected_rows_indices) if has_selection else False
            
            # Calcola capacità per ogni tipo di azione VSM
            can_delete = has_selection and all_mine  # Delete su uno o più eventi propri
            can_duplicate = (num_selected == 1) and all_mine  # Duplicate solo su singolo evento proprio
            
            # Abilita Actions solo se c'è selezione valida (tutte mie)
            can_act = has_selection and all_mine
            self.btn_actions.config(state="normal" if can_act else "disabled")
            
            # Step 4D.2/4D.5: Popola menu Actions con opzioni VSM
            self._populate_actions_menu(status, can_delete, can_duplicate)
            return
        
        # RFQ logic (invariata)
        selected_rows_indices = self._get_selected_row_indices(sheet)
        has_sel = bool(selected_rows_indices)
        num_selected = len(selected_rows_indices) if selected_rows_indices else 0
        
        # Verifica se tutte le RfQ selezionate appartengono all'utente corrente
        all_mine = self._check_if_all_selected_are_mine(sheet, selected_rows_indices) if has_sel else False
        
        # Calcola capacità per ogni tipo di azione
        can_delete = has_sel and all_mine
        can_duplicate = (num_selected == 1) and all_mine
        can_change_status = has_sel and all_mine
        
        # Abilita Actions solo se c'è almeno una selezione valida (tutte mie)
        can_act = has_sel and all_mine
        self.btn_actions.config(state="normal" if can_act else "disabled")
        
        # Popola il menu Actions dinamicamente in base al tab corrente
        self._populate_actions_menu(status, can_delete, can_duplicate, can_change_status)

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
        # Pulisci menu esistente
        self.actions_menu.delete(0, 'end')
        
        # Step 4D.2/4D.4/4D.5: Branch VSM
        if status.startswith('vsm_'):
            # Menu VSM: Delete + Duplicate (Edit tramite double-click)
            # Ordine identico a RFQ per coerenza UX
            self.actions_menu.add_command(
                label=_("🗑 Elimina"),
                command=self._delete_vsm_events,
                state="normal" if can_delete else "disabled"
            )
            
            self.actions_menu.add_command(
                label=_("🔁 Duplica"),
                command=self._duplicate_vsm_event,
                state="normal" if can_duplicate else "disabled"
            )
            return  # Early return per VSM
        
        # RFQ logic (invariata)
        # Azioni comuni a entrambi i tab (riuso metodi esistenti)
        self.actions_menu.add_command(
            label=_("🗑 Elimina"),
            command=self.delete_selected_request,
            state="normal" if can_delete else "disabled"
        )
        
        self.actions_menu.add_command(
            label=_("🔁 Duplica"),
            command=self.duplicate_selected_request,
            state="normal" if can_duplicate else "disabled"
        )
        
        self.actions_menu.add_separator()
        
        # Azione specifica per tab (riuso metodi esistenti)
        if status == 'attiva':
            self.actions_menu.add_command(
                label=_("📦 Archivia"),
                command=self.archive_selected_request,
                state="normal" if can_change_status else "disabled"
            )
        else:  # archiviata
            self.actions_menu.add_command(
                label=_("↩️ Riattiva"),
                command=self.reactivate_selected_request,
                state="normal" if can_change_status else "disabled"
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

    def _load_requests_by_status(self, tree, status):
        """Carica richieste per stato specifico con supporto multi-database."""
        try:
            username_filter = self._get_active_username_filter()
            
            # SEMPRE usa aggregazione multi-database per avere accesso a tutti gli utenti
            logger.info(f"[MULTI-DB] Caricamento da tutti i database (filtro utente: {username_filter})...")
            
            # BUG #47 FIX: Usa context manager per garantire chiusura DB
            with DatabaseManager(get_db_path()) as db_manager:
                # Chiama il metodo aggregato che legge TUTTI i database
                all_rows = db_manager.get_all_richieste_aggregated(get_db_path())
            
            # Filtra per stato richiesto
            # Struttura: [0] id_richiesta, [1] tipo_rdo, [2] data_emissione,
            # [3] data_scadenza, [4] riferimento, [5] username, [6] stato, 
            # [7] is_mine, [8] source_file
            filtered_rows = [row for row in all_rows if row[6] == status]
            
            # SE C'È UN FILTRO UTENTE SPECIFICO, filtra anche per username
            if username_filter is not None:
                filtered_rows = [row for row in filtered_rows if row[5] and row[5].lower() == username_filter.lower()]
                logger.info(f"[MULTI-DB] Trovate {len(filtered_rows)} RdO in stato '{status}' per utente '{username_filter}'")
            else:
                logger.info(f"[MULTI-DB] Trovate {len(filtered_rows)} RdO in stato '{status}' da tutti gli utenti")
            
            # BUGFIX: Applica filtro tipo RdO se presente (non solo "Tutte")
            tipo_filter = self.search_tipo.get()
            if tipo_filter != _("Tutte"):
                tipo_canonico = normalize_rfq_type(tipo_filter)
                filtered_rows = [row for row in filtered_rows if row[1] == tipo_canonico]
                logger.info(f"[MULTI-DB] Filtro tipo RdO '{tipo_filter}' applicato: {len(filtered_rows)} risultati")
            
            self.update_treeview(tree, filtered_rows)
                
        except DatabaseError as e:
            logger.error(f"Errore database in _load_requests_by_status: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile caricare elenco: {}").format(e), "error")

    def update_treeview(self, sheet, requests):
        """Aggiorna il foglio tksheet con i dati delle richieste"""
        today = date.today()
        data_rows = []
        
        # Variabile per tracciare la lunghezza massima del riferimento
        max_ref_length = 0
        
        # Inizializza lista metadati se non esiste
        if not hasattr(sheet, '_sheet_rows_metadata'):
            sheet._sheet_rows_metadata = []
        
        sheet._sheet_rows_metadata = []  # Reset metadati
        
        for i, req in enumerate(requests):
            # Traduci il tipo RFQ prima di inserirlo nel sheet
            tipo_rdo_tradotto = translate_rfq_type(req[1])
            riferimento = req[4] if req[4] else ""
            username_value = ""
            if len(req) > 5 and req[5]:
                username_value = str(req[5]).strip()
            
            # BUG #3 FIX: Validazione robusta per metadati con logging dettagliato
            # Salva metadati per questa riga (is_mine e source_file)
            # La struttura aggregate è: [..., stato, is_mine, source_file]
            if len(req) > 8:
                is_mine = req[7]
                source_file = req[8]
                logger.debug(f"Riga {i} (ID {req[0]}): is_mine={is_mine}, source={source_file}")
            else:
                # Fallback per dati non aggregati
                is_mine = True
                source_file = 'local'
                if len(req) < 6:
                    logger.warning(f"Riga {i}: tuple troppo corta ({len(req)} elementi), dati incompleti. Usando default is_mine=True")
            
            sheet._sheet_rows_metadata.append({
                'is_mine': is_mine,
                'source_file': source_file
            })
            
            # Aggiorna la lunghezza massima del riferimento
            if riferimento:
                max_ref_length = max(max_ref_length, len(riferimento))
            
            row = [
                str(req[0]),  # ID
                tipo_rdo_tradotto,  # Tipo
                self._format_date_for_display(req[2]),  # Data emissione
                self._format_date_for_display(req[3]),  # Data scadenza
                riferimento,  # Riferimento
                username_value  # Username
            ]
            data_rows.append(row)
        
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
            header_text = _("Riferimento")
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
        'vsm_derisking': 'Derisking',
    }

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
        if raw_meta is not None:
            pairs = [(ev, m) for ev, m in zip(raw_events, raw_meta) if ev.event_type == event_type]
            results = [p[0] for p in pairs]
            result_meta = [p[1] for p in pairs]
        else:
            results = [ev for ev in raw_events if ev.event_type == event_type]
            result_meta = None

        # Applica Advanced Filters (stesso scope di _load_vsm_events)
        results, result_meta = self._apply_vsm_filters(results, event_type, extra_meta=result_meta)

        # Applica query testuale globale
        _VSM_SEARCH_FIELDS = (
            'description', 'reference', 'buyer', 'driver',
            'action', 'event_type', 'new_supplier', 'note',
        )
        if result_meta is not None:
            pairs = [
                (ev, m) for ev, m in zip(results, result_meta)
                if any(query in (getattr(ev, f) or "").lower() for f in _VSM_SEARCH_FIELDS)
            ]
            results = [p[0] for p in pairs]
            result_meta = [p[1] for p in pairs]
        else:
            results = [
                ev for ev in results
                if any(query in (getattr(ev, f) or "").lower() for f in _VSM_SEARCH_FIELDS)
            ]

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
            SimpleMessageDialog(self.root, _("Operazione Non Consentita"), _("Non puoi eliminare RdO di altri utenti.\nPuoi operare solo sulle tue RdO."), "error")
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
            msg = _("Sei sicuro di voler eliminare la RdO N° {}?\nL'operazione è permanente.").format(rdo_num)
        else:
            msg = _("Sei sicuro di voler eliminare le {} RdO selezionate?\nL'operazione è permanente.").format(count)
        if not SimpleYesNoDialog(self.root, _("Conferma Eliminazione"), msg).result: return
        
        try:
            print(f"[MainWindow.delete_selected_request] Eliminazione di {len(request_ids)} richieste: {request_ids}")

            # Rimuovi i file fisici degli allegati per ogni richiesta prima di eliminare le righe DB
            archive_path = get_fixed_attachments_dir()
            try:
                with DatabaseManager(get_db_path()) as db_manager:
                    # Per ogni richiesta, recupera i percorsi_esterni e prova a rimuovere i file
                    for req_id in request_ids:
                        try:
                            rows = db_manager.conn.execute(
                                "SELECT percorso_esterno FROM allegati_richiesta WHERE id_richiesta = ? AND percorso_esterno IS NOT NULL",
                                (req_id,)
                            ).fetchall()
                        except Exception:
                            rows = []

                        for row in rows:
                            percorso = row[0]
                            if not percorso:
                                continue
                            # Se percorso è relativo, cerca nella cartella Attachments
                            if archive_path and not os.path.isabs(percorso):
                                file_to_delete = os.path.join(archive_path, percorso)
                            else:
                                file_to_delete = percorso

                            try:
                                if os.path.exists(file_to_delete):
                                    os.remove(file_to_delete)
                                    logger.info(f"Allegato eliminato dal disco durante cancellazione RdO: {file_to_delete}")
                                else:
                                    logger.info(f"File allegato non trovato durante cancellazione RdO: {file_to_delete}")
                            except Exception as disk_error:
                                logger.warning(f"Impossibile eliminare il file allegato {file_to_delete}: {disk_error}")

                    # Ora elimina le richieste e i record correlati nel DB
                    count = db_manager.delete_richieste_batch(request_ids)

            except DatabaseError as e:
                raise

            print(f"[MainWindow.delete_selected_request] Eliminate {count} richieste dal database")

            # Ricarica i dati invece di cancellare elementi dalla view
            self.refresh_data()
            if count == 1:
                msg = _("1 RdO eliminata.")
            else:
                msg = _("{} RdO eliminate.").format(count)
            SimpleMessageDialog(self.root, _("Successo"), msg, "info")
        except DatabaseError as e:
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile eliminare: {}").format(e), "error")

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
                    SimpleMessageDialog(self.root, _("Selezione non valida"), _("Seleziona una sola RdO per duplicarla."), "warning")
                    return
                row_index = selected_rows[0] if isinstance(selected_rows, (list, set, tuple)) else selected_rows
        
        # VALIDAZIONE SICUREZZA: Verifica che la RfQ selezionata sia dell'utente corrente
        if row_index is not None:
            if not self._check_if_all_selected_are_mine(sheet, [row_index]):
                SimpleMessageDialog(self.root, _("Operazione Non Consentita"), _("Non puoi duplicare RdO di altri utenti.\nPuoi operare solo sulle tue RdO."), "error")
                logger.warning(f"Tentativo di duplicazione RfQ altrui bloccato: utente={self.current_username}")
                return
        
        # Se non c'è nessuna selezione
        if row_index is None:
            SimpleMessageDialog(self.root, _("Selezione mancante"), _("Selezionare una RdO da duplicare."), "warning")
            return

        # Ottieni i dati della riga
        try:
            row_data = sheet.get_row_data(row_index)
            if not row_data or len(row_data) == 0:
                SimpleMessageDialog(self.root, _("Errore"), _("Impossibile determinare la RdO selezionata."), "error")
                return
            original_id = int(row_data[0])
        except (ValueError, TypeError, IndexError) as e:
            logger.error(f"Errore nel recupero dati riga per duplicazione: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile determinare la RdO selezionata."), "error")
            return

        new_request_id = None

        try:
            # Helper function per ottenere colonne
            def get_columns(table_name, exclude):
                # BUG #47 FIX: Usa context manager per garantire chiusura DB anche su eccezione
                with DatabaseManager(get_db_path()) as db_mgr:
                    cols_info = db_mgr.get_table_columns(table_name)
                excluded = set(exclude)
                # SQLite PRAGMA table_info restituisce: colonna[0] = cid, colonna[1] = nome colonna, colonna[2] = tipo
                # Usa colonna[1] per estrarre il nome della colonna
                columns = [row[1] for row in cols_info if row[1] not in excluded]
                print(f"[get_columns] Tabella {table_name}: colonne recuperate = {columns}")
                print(f"[get_columns] Dettagli PRAGMA: {cols_info[:3] if cols_info else []}")  # Prime 3 righe per debug
                return columns

            # BUG #47 FIX: Usa context manager anche per duplicazione
            with DatabaseManager(get_db_path()) as db_manager:
                new_request_id = db_manager.duplicate_richiesta_full(original_id, get_columns)
            
            # BUG #3 FIX: Verifica SUBITO dopo duplicazione
            if new_request_id is None:
                raise ValueError("Duplicazione fallita: ID nuova RdO non ottenuto")
            
            logger.info(f"RdO duplicata: {original_id} -> {new_request_id}")

        except ValueError as ve:
            SimpleMessageDialog(self.root, _("Errore"), str(ve), "error")
            return
        except DatabaseError as e:
            logger.error(f"Errore duplicazione RdO {original_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile duplicare la RdO: {}").format(e), "error")
            return
        except Exception as e:
            logger.error(f"Errore duplicazione RdO {original_id}: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Impossibile duplicare: {}").format(e), "error")
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
        SimpleMessageDialog(self.root, _("Successo"), _("RdO duplicata come N° {}.").format(new_request_id), "info")

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
        - VSM (Saving/Cost Avoidance/Derisking): crea nuovo evento VSM
        """
        _, status = self.get_current_tree_and_status()
        
        # Branch 1: RFQ - usa la logica esistente
        if status in ('attiva', 'archiviata'):
            self.open_new_request_window()
        
        # Branch 2: VSM - apri dialog CREATE
        elif status.startswith('vsm_'):
            # Mappa status → event_type (pattern già usato in _edit_vsm_event)
            event_type_map = {
                'vsm_saving': 'Saving',
                'vsm_cost_avoidance': 'Cost Avoidance',
                'vsm_derisking': 'Derisking'
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
                SimpleMessageDialog(self.root, _("Errore"), _("Impossibile aprire il form: {}").format(e), "error")

    def open_new_request_window(self):
        """Crea una nuova RdO 'guscio' e apre l'editor"""
        # Mostra dialog per scelta tipo
        dialog = NewRdOTypeDialog(self.root)
        self.root.wait_window(dialog)
        
        # Se l'utente ha annullato, esci
        if not dialog.result:
            return
        
        tipo_rdo = normalize_rfq_type(dialog.result)
        
        # BUG #32 FIX: Usa try-finally per garantire chiusura DB
        db_manager = None
        try:
            # Inserisce testata minima usando db_manager
            data_oggi = datetime.now().strftime('%Y-%m-%d')
            db_manager = DatabaseManager(get_db_path())
            id_nuova = db_manager.insert_richiesta_offerta(tipo_rdo, 'attiva', data_oggi, username=self.current_username)
            
            logger.info(f"Creata nuova RdO guscio N° {id_nuova} (tipo: {tipo_rdo})")
            
            # Apri immediatamente l'editor
            self.root.wait_window(ViewRequestWindow(self.root, id_nuova))
            
            # Aggiorna la lista dopo la chiusura
            self.refresh_data()
            
        except DatabaseError as e:
            logger.error(f"Errore creazione RdO guscio: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore Database"), _("Impossibile creare la nuova RdO: {}").format(e), "error")
        finally:
            # BUG #32 FIX: Garantisce chiusura connessione anche in caso di eccezione
            if db_manager is not None:
                try:
                    db_manager.close()
                except Exception as close_error:
                    logger.warning(f"Errore chiusura database in open_new_request_window: {close_error}")

    def mega_export_excel(self):
        """
        Esporta tutte le RfQ attualmente visibili nella lista (filtrate) in un unico file Excel.
        Genera un report a blocchi verticali, adattandosi al tipo di ogni singola RfQ.
        """
        # 1. Identifica quale tabella è attiva e recupera lo stato corrente
        current_tree, status = self.get_current_tree_and_status()
        
        # VSM tabs: dispatch to dedicated VSM export
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
                if self.search_tipo.get() != _("Tutte"):
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
                if tipo_filter != _("Tutte"):
                    tipo_canonico = normalize_rfq_type(tipo_filter)
                    filtered_rows = [row for row in filtered_rows if row[1] == tipo_canonico]
                
                # Salva ID e percorso database per ogni RfQ
                for row in filtered_rows:
                    source_db_path = row[8] if len(row) > 8 else 'local'
                    if source_db_path == 'local':
                        source_db_path = get_db_path()
                    request_data.append((row[0], source_db_path))
            
            if not request_data:
                SimpleMessageDialog(self.root, _("Attenzione"), _("Nessuna RfQ da esportare nella vista corrente."), "warning")
                return
            
            logger.info(f"[export_excel] Trovate {len(request_data)} RfQ da esportare: {[r[0] for r in request_data[:10]]}{'...' if len(request_data) > 10 else ''}")
            
        except Exception as e:
            logger.error(f"[export_excel] Errore nel recupero degli ID: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Errore nel recupero delle RfQ da esportare: {}").format(e), "error")
            return

        # 2. Chiedi Lingua
        prompt = LanguagePrompt(self.root)
        self.root.wait_window(prompt)
        lang = prompt.choice  # 'ita' o 'eng'
        if not lang:
            return

        # 3. Configurazione Testi e Header in base alla lingua
        is_ita = (lang == 'ita')
        headers_map = {
            'cod': "Codice" if is_ita else "Code",
            'att': "Allegato" if is_ita else "Attachment",
            'desc': "Descrizione" if is_ita else "Description",
            'qty': "Q.tà" if is_ita else "Q.ty",
            'cod_g': "Cod. Grezzo" if is_ita else "Raw Code",
            'dis_g': "Dis. Grezzo" if is_ita else "Raw Dwg",
            'mat_cl': "Mat. C/L" if is_ita else "Work Order Mat.",
            'vs_best': "VS. MIGLIORE" if is_ita else "YOUR BEST",
            'rdo_num': "Richiesta N°" if is_ita else "RfQ N°",
            'date': "Del" if is_ita else "Date",
            'type': "Tipo" if is_ita else "Type"
        }

        # 4. Setup Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Export DataFlow"
        
        # Stili
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)
        header_fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')  # Grigio chiaro
        best_price_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')  # Verde
        
        # Setup larghezze colonne
        ws.column_dimensions['A'].width = 15  # Codice
        ws.column_dimensions['B'].width = 15  # Allegato
        ws.column_dimensions['C'].width = 10  # Qta
        ws.column_dimensions['D'].width = 35  # Descrizione
        ws.column_dimensions['E'].width = 15  # Cod Grezzo
        ws.column_dimensions['F'].width = 15  # Dis Grezzo
        ws.column_dimensions['G'].width = 20  # Mat CL
        
        current_row = 1
        
        try:
            # 5. CICLO SULLE RDO - usa il database corretto per ogni RfQ
            for req_id, source_db_path in request_data:
                # Apri il database corretto per questa RfQ
                db_manager = DatabaseManager(source_db_path)
                
                try:
                    # Recupera dati testata
                    rdo_data = db_manager.get_richiesta_full_data(req_id)
                    if not rdo_data:
                        continue
                    de_db, ds_db, rif, tipo_raw = rdo_data
                    
                    # Normalizza tipo
                    tipo_normalizzato = normalize_rfq_type(tipo_raw)
                    is_cl = (tipo_normalizzato == 'Conto lavoro')
                    
                    # Recupera dettagli e fornitori
                    items = db_manager.get_dettagli_by_richiesta(req_id)
                    suppliers_rows = db_manager.get_fornitori_by_richiesta(req_id, order_by=True)
                    suppliers = [r[0] for r in suppliers_rows]
                    prices_rows = db_manager.get_offerte_by_richiesta(req_id)
                    prices = {(id_d, nf): p for id_d, nf, p in prices_rows}
                    
                finally:
                    # Chiudi il database manager dopo ogni RfQ
                    db_manager.close()

                # --- SCRITTURA BLOCCO TESTATA ---
                ws.cell(row=current_row, column=1, value=f"{headers_map['rdo_num']} {req_id}").font = Font(size=12, bold=True)
                ws.cell(row=current_row, column=4, value=f"{headers_map['date']}: {self._format_date_for_display(de_db)}")
                ws.cell(row=current_row, column=7, value=f"Ref: {rif}")
                current_row += 1
                ws.cell(row=current_row, column=1, value=f"{headers_map['type']}: {translate_rfq_type(tipo_normalizzato)}")
                current_row += 2

                # --- SCRITTURA HEADER TABELLA ---
                col_headers = [
                    headers_map['cod'], headers_map['att'], headers_map['qty'], headers_map['desc'],
                    headers_map['cod_g'], headers_map['dis_g'], headers_map['mat_cl']
                ]
                
                for i, h_text in enumerate(col_headers, start=1):
                    c = ws.cell(row=current_row, column=i, value=h_text)
                    c.font = bold_font
                    c.border = thin_border
                    c.fill = header_fill
                    c.alignment = Alignment(horizontal='center')

                # Colonna separatore
                c_sep = ws.cell(row=current_row, column=8, value=headers_map['vs_best'])
                c_sep.font = bold_font
                c_sep.border = thin_border
                c_sep.alignment = Alignment(horizontal='center')

                # Colonne Fornitori
                start_supplier_col = 9
                for i, sup in enumerate(suppliers):
                    c = ws.cell(row=current_row, column=start_supplier_col + i, value=sup)
                    c.font = bold_font
                    c.border = thin_border
                    c.alignment = Alignment(horizontal='center')
                
                current_row += 1

                # --- SCRITTURA RIGHE ARTICOLI ---
                for item in items:
                    id_d, cod, all_file, desc, qta, c_g, d_g, m_cl = item
                    
                    ws.cell(row=current_row, column=1, value=cod).border = thin_border
                    ws.cell(row=current_row, column=2, value=all_file).border = thin_border
                    ws.cell(row=current_row, column=3, value=format_quantity_display(qta)).border = thin_border
                    ws.cell(row=current_row, column=4, value=desc).border = thin_border
                    
                    ws.cell(row=current_row, column=5, value=c_g if is_cl else "").border = thin_border
                    ws.cell(row=current_row, column=6, value=d_g if is_cl else "").border = thin_border
                    ws.cell(row=current_row, column=7, value=m_cl if is_cl else "").border = thin_border
                    
                    ws.cell(row=current_row, column=8, value="").border = thin_border
                    # Prezzi
                    min_price = None
                    row_prices = []
                    for sup in suppliers:
                        p_val = prices.get((id_d, sup))
                        if p_val:
                            try:
                                row_prices.append(float(str(p_val).replace(',', '.')))
                            except:
                                pass
                    if row_prices:
                        min_price = min(row_prices)
                    for i, sup in enumerate(suppliers):
                        col_idx = start_supplier_col + i
                        cell = ws.cell(row=current_row, column=col_idx)
                        price_val = prices.get((id_d, sup))
                        
                        if price_val is not None:
                            try:
                                val_float = float(str(price_val).replace(',', '.'))
                                cell.value = val_float
                                cell.number_format = '0.0000'
                                if min_price is not None and val_float == min_price and val_float > 0:
                                    cell.fill = best_price_fill
                            except:
                                cell.value = price_val
                                cell.alignment = Alignment(horizontal='right')
                        cell.border = thin_border
                    current_row += 1
                
                current_row += 3

            # 6. Salvataggio
            default_name = f"Export_DataFlow_{datetime.now().strftime('%Y%m%d')}.xlsx"
            save_path = filedialog.asksaveasfilename(
                title=_("Salva Export"),
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel Files", "*.xlsx")]
            )
            
            if save_path:
                wb.save(save_path)
                SimpleMessageDialog(self.root, _("Successo"), _("Export completato con successo:\n{}").format(save_path), "info")
                logger.info(f"Export Excel salvato in: {save_path}")
        except Exception as e:
            logger.error(f"Errore Export Excel: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Errore durante l'esportazione: {}").format(e), "error")

    def _export_vsm_excel(self, status, sheet):
        """Esporta i dati VSM del tab corrente in un file Excel.

        Flusso identico a mega_export_excel:
        1. Dialog scelta lingua (LanguagePrompt)
        2. Re-query DB → eventi raw (valori numerici puliti, senza simbolo €)
        3. Intestazioni basate sulla lingua scelta
        4. Scrittura Excel con numeri float, non stringhe formattate
        """
        status_to_event_type = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking',
        }
        event_type = status_to_event_type.get(status, status)

        # 1. Scelta lingua — identico a mega_export_excel
        prompt = LanguagePrompt(self.root)
        self.root.wait_window(prompt)
        lang = prompt.choice  # 'ita' o 'eng'
        if not lang:
            return
        is_ita = (lang == 'ita')

        # 2. Re-load eventi dal DB (stessa query di _load_vsm_events → dati raw)
        try:
            with DatabaseManager(get_db_path()) as db_manager:
                all_events = db_manager.get_all_vsm_events(username=self.current_username)
            events = [e for e in all_events if e.event_type == event_type]
        except Exception as e:
            logger.error(f"[export_vsm] Errore recupero eventi: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Errore nel recupero dati: {}").format(e), "error")
            return

        if not events:
            SimpleMessageDialog(self.root, _("Attenzione"), _("Nessun dato da esportare nella vista corrente."), "warning")
            return

        # 3. Intestazioni in base a lingua e tipo tab (hardcoded IT/EN come mega_export_excel)
        use_dual = event_type in ("Saving", "Cost Avoidance")
        action_map_en = {"Negoziazione": "Negotiation", "Altro": "Other"}
        if is_ita:
            if event_type == "Saving":
                headers = ["Data", "Tipo", "Azione", "Descrizione",
                           "Saving Teorico", "Saving Effettivo", "Realizzo %", "Variance %", "Ripetitivo", "Utente"]
            elif event_type == "Cost Avoidance":
                headers = ["Data", "Tipo", "Azione", "Descrizione",
                           "CA Teorico", "CA Effettivo", "Realizzo %", "Variance %", "Ripetitivo", "Utente"]
            else:  # Derisking
                headers = ["Data", "Tipo", "Azione", "Nuovo Fornitore", "Descrizione",
                           "Valore Teorico", "Realizzo %", "Ripetitivo", "Utente"]
        else:
            if event_type == "Saving":
                headers = ["Date", "Type", "Action", "Description",
                           "Theoretical Savings", "Actual Savings", "Realization %", "Variance %", "Repetitive", "User"]
            elif event_type == "Cost Avoidance":
                headers = ["Date", "Type", "Action", "Description",
                           "CA Theoretical", "CA Actual", "Realization %", "Variance %", "Repetitive", "User"]
            else:  # Derisking
                headers = ["Date", "Type", "Action", "New Supplier", "Description",
                           "Theoretical Value", "Realization %", "Repetitive", "User"]

        # 4. Costruzione righe con valori numerici raw (nessun simbolo €, nessuna formattazione display)
        data_rows = []
        for event in events:
            valore_teorico = event.calculate_theoretical_value() or 0.0
            date_str = event.event_date.strftime("%d/%m/%Y") if event.event_date else ""
            desc = (event.description or event.reference or "")[:50]
            action_str = event.action if is_ita else action_map_en.get(event.action, event.action)

            if use_dual:
                valore_effettivo = event.calculate_effective_value() or 0.0
                if event.driver == "Pagamenti" and event.giorni_pagamento_attuali is not None and event.giorni_pagamento_negoziati is not None:
                    _delta = event.giorni_pagamento_negoziati - event.giorni_pagamento_attuali
                    _variance_pct = round((_delta / 30.0) * event.effective_payments_rate_pct, 2)
                elif event_type == "Cost Avoidance":
                    _baseline = event.importo_richiesto_iniziale or 0.0
                    _variance_pct = round(
                        (_baseline - (event.importo_negoziato or 0.0)) / _baseline * 100, 1
                    ) if _baseline != 0.0 else 0.0
                else:
                    _baseline = event.importo_bdg or 0.0
                    _variance_pct = round(
                        (_baseline - (event.importo_negoziato or 0.0)) / _baseline * 100, 1
                    ) if _baseline != 0.0 else 0.0
                row = [
                    date_str, event.event_type, action_str, desc,
                    round(valore_teorico, 2), round(valore_effettivo, 2),
                    round(event.percent_realizzo, 1), _variance_pct,
                    "✓" if event.opex_ripetitivo else "", event.username
                ]
            else:  # Derisking
                row = [
                    date_str, event.event_type, action_str,
                    event.new_supplier or "",
                    desc,
                    round(valore_teorico, 2),
                    round(event.percent_realizzo, 1),
                    "✓" if event.opex_ripetitivo else "", event.username
                ]
            data_rows.append(row)

        # 5. Setup Excel — stessi stili di mega_export_excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = event_type[:31]

        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        bold_font = Font(bold=True)
        header_fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = bold_font
            cell.border = thin_border
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

        # Indici colonne monetarie e percentuali per number_format (1-based)
        monetary_cols = {5, 6} if use_dual else {6}
        pct_cols = {7, 8} if use_dual else {7}
        rep_col = 9 if use_dual else 8  # Colonna "Ripetitivo/Repetitive" (1-based)

        for row_idx, row_data in enumerate(data_rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = thin_border
                if col_idx in monetary_cols:
                    cell.number_format = '#,##0.00'
                elif col_idx in pct_cols:
                    cell.number_format = '0.0'
                elif col_idx == rep_col:
                    cell.alignment = Alignment(horizontal='center')

        # Larghezze colonne adattive (px tkinter ÷ 7 → unità Excel)
        col_widths_px = getattr(sheet, '_vsm_col_widths', None)
        if col_widths_px:
            for i, px_width in enumerate(col_widths_px):
                col_letter = ws.cell(row=1, column=i + 1).column_letter
                ws.column_dimensions[col_letter].width = max(10, px_width / 7)

        # 6. Salvataggio — identico a mega_export_excel
        default_name = f"Export_VSM_{event_type.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        try:
            save_path = filedialog.asksaveasfilename(
                title=_("Salva Export"),
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if save_path:
                wb.save(save_path)
                SimpleMessageDialog(self.root, _("Successo"), _("Export completato con successo:\n{}").format(save_path), "info")
                logger.info(f"Export VSM Excel salvato in: {save_path}")
        except Exception as e:
            logger.error(f"Errore Export VSM Excel: {e}", exc_info=True)
            SimpleMessageDialog(self.root, _("Errore"), _("Errore durante l'esportazione: {}").format(e), "error")

    def _format_date_for_display(self, db_date):
        if not db_date: return ""
        try: return datetime.strptime(db_date, '%Y-%m-%d').strftime('%d/%m/%Y')
        except (ValueError, TypeError): return db_date

class UserIdentityDialog(tk.Toplevel):
    """Finestra modale che forza l'inserimento di nome e cognome."""
    def __init__(self, parent, first_name='', last_name=''):
        super().__init__(parent)
        self.withdraw()
        self.title(_("Dati Utente Richiesti"))
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        set_window_icon(self)
        self.protocol("WM_DELETE_WINDOW", self._prevent_close)
        
        self.first_var = tk.StringVar(value=first_name)
        self.last_var = tk.StringVar(value=last_name)
        self.username_var = tk.StringVar(value=_("(in attesa dati)"))
        
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(
            frame,
            text=_("Per procedere è necessario indicare il tuo nome e cognome."),
            font=(None, 10),
            wraplength=320,
            justify="left"
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        
        ttk.Label(frame, text=_("Nome:")).grid(row=1, column=0, sticky="w", pady=5)
        first_entry = ttk.Entry(frame, textvariable=self.first_var, width=30)
        first_entry.grid(row=1, column=1, sticky="ew", pady=5)
        
        ttk.Label(frame, text=_("Cognome:")).grid(row=2, column=0, sticky="w", pady=5)
        last_entry = ttk.Entry(frame, textvariable=self.last_var, width=30)
        last_entry.grid(row=2, column=1, sticky="ew", pady=5)
        
        ttk.Label(frame, text=_("Username generato:")).grid(row=3, column=0, sticky="w", pady=(10, 0))
        username_display = ttk.Label(frame, textvariable=self.username_var, font=("Calibri", 12, "bold"), foreground="#005AA0")
        username_display.grid(row=3, column=1, sticky="w", pady=(10, 0))
        
        confirm_btn = ttk.Button(frame, text=_("Conferma"), command=self._on_confirm)
        confirm_btn.grid(row=4, column=0, columnspan=2, pady=(20, 0), sticky="ew")
        
        frame.columnconfigure(1, weight=1)
        
        self.first_var.trace_add("write", self._update_preview)
        self.last_var.trace_add("write", self._update_preview)
        self._update_preview()
        self._center_window()
        first_entry.focus_set()

    def _update_preview(self, *_args):
        first = self.first_var.get().strip()
        last = self.last_var.get().strip()
        if not first or not last:
            self.username_var.set(_("(in attesa dati)"))
            return
        try:
            username = generate_username(first, last)
            self.username_var.set(username)
        except ValueError:
            self.username_var.set(_("Dati non validi"))

    def _on_confirm(self):
        first = self.first_var.get().strip()
        last = self.last_var.get().strip()
        if not first or not last:
            SimpleMessageDialog(self, _("Campi obbligatori"), _("Inserisci sia il nome sia il cognome."), "error")
            return
        try:
            username = generate_username(first, last)
        except ValueError as e:
            SimpleMessageDialog(self, _("Formato non valido"), str(e), "error")
            return
        self.result = {
            'first_name': first,
            'last_name': last,
            'username': username
        }
        self.grab_release()
        self.destroy()

    def _prevent_close(self):
        SimpleMessageDialog(self, _("Operazione necessaria"), _("Per utilizzare DataFlow è necessario completare i dati richiesti."), "warning")

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        if not w or not h:
            w, h = 360, 220
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w // 2) - (w // 2)
        y = (screen_h // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()

# ------------------------------------------------------------------------------------
# FINESTRA PROGRESSO COPIA (PER SPOSTAMENTO CARTELLA)
# ------------------------------------------------------------------------------------
class CopyProgressWindow(tk.Toplevel):
    """Finestra di progresso per operazioni di copia file (stile splash screen)."""
    def __init__(self, parent, title="Copia in corso..."):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(title)
        self.overrideredirect(True)
        
        frame = ttk.Frame(self, borderwidth=2, relief="raised")
        frame.pack(fill="both", expand=True)
        
        # Logo (opzionale)
        try:
            logo_path = resource_path(os.path.join("add_data", "logo_dataflow.png"))
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                # BUG #51 FIX: Check dimensioni valide prima di divisione per evitare ZeroDivisionError
                if img.width > 0 and img.height > 0:
                    img.thumbnail((200, int(200 * (img.height/img.width))), Image.Resampling.LANCZOS)
                    self.logo_photo = ImageTk.PhotoImage(img)
                    ttk.Label(frame, image=self.logo_photo).pack(pady=(20, 10))
        except Exception as e:
            print(f"Errore logo: {e}")
            ttk.Label(frame, text="DataFlow", font=("Helvetica", 18, "bold")).pack(pady=(20, 10))
        
        self.status_label = ttk.Label(
            frame,
            text="Preparazione...",
            font=("Helvetica", 10),
            width=50,
            anchor="center"
        )
        self.status_label.pack(pady=(10, 5))
        
        self.progress = ttk.Progressbar(frame, orient="horizontal", length=400, mode='determinate')
        self.progress.pack(pady=(0, 20))
        
        # Calcola dimensioni e posizione
        self.update_idletasks()
        w = 500
        h = 250
        x = (self.winfo_screenwidth()//2) - (w//2)
        y = (self.winfo_screenheight()//2) - (h//2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
    
    def update_progress(self, val, txt):
        """Aggiorna barra e testo."""
        self.progress['value'] = val
        self.status_label['text'] = ""
        self.update_idletasks()
        self.status_label['text'] = txt
        self.update_idletasks()

# ------------------------------------------------------------------------------------
# FINESTRA DI AVVIO (SPLASH SCREEN)
# ------------------------------------------------------------------------------------
class SplashScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()  # 1. Nascondi subito
        set_window_icon(self)
        self.title(_("Avvio DataFlow")); self.overrideredirect(True) # 2. Rendi senza bordi
        
        # 3. Aggiungi TUTTI i widget (mentre è ancora nascosta)
        frame = ttk.Frame(self, borderwidth=2, relief="raised"); frame.pack(fill="both", expand=True)
        try:
            logo_path = resource_path(os.path.join("add_data", "logo_dataflow.png"))
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                # BUG #51 FIX: Check dimensioni valide prima di divisione per evitare ZeroDivisionError
                if img.width > 0 and img.height > 0:
                    img.thumbnail((273, int(273 * (img.height/img.width))), Image.Resampling.LANCZOS)
                    self.logo_photo = ImageTk.PhotoImage(img)
                    ttk.Label(frame, image=self.logo_photo).pack(pady=(30, 20))
        except Exception as e:
            print(f"Errore logo splash: {e}"); ttk.Label(frame, text=_("DataFlow"), font=("Helvetica", 24, "bold")).pack(pady=(30, 20))
        
        # CORREZIONE: Aggiunti width e anchor per evitare sovrapposizioni
        self.status_label = ttk.Label(
            frame, 
            text=_("Avvio in corso..."), 
            font=("Helvetica", 10),
            width=40,
            anchor="center"
        )
        self.status_label.pack(pady=(10, 5))
        self.progress = ttk.Progressbar(frame, orient="horizontal", length=300, mode='determinate'); self.progress.pack(pady=(0, 20))
        
        # 4. Forza Tkinter a calcolare le dimensioni REALI
        self.update_idletasks() 
        
        # 5. Leggi le dimensioni REALI (non più 450x250 fisse)
        w = 450
        h = 250
        
        # 6. Calcola la posizione centrale
        x = (self.winfo_screenwidth()//2) - (w//2)
        y = (self.winfo_screenheight()//2) - (h//2)
        
        # 7. Applica la geometria corretta
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        # 8. Mostra la finestra (ora perfetta)
        self.deiconify()

    def update_progress(self, val, txt):
        self.progress['value'] = val
        # CORREZIONE: Pulisci il testo prima di aggiornarlo per evitare sovrapposizioni
        self.status_label['text'] = ""
        self.update_idletasks()
        self.status_label['text'] = txt
        self.update_idletasks()

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
            license_prompt = LicenseWindow(root, first_run=True)
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
                    SimpleMessageDialog(root, _("Errore"), _("Impossibile creare la cartella utente."), "error")
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
            SimpleMessageDialog(root, _("Errore"), _("Impossibile creare il database utente: {}").format(e), "error")
            try:
                root.destroy()
            except:
                pass
            return

        # 3) Ora mostriamo lo splash (dopo la creazione DB) e carichiamo l'interfaccia
        splash_local = SplashScreen(root)
        splash_local.update_progress(90, _("Caricamento interfaccia..."))
        splash_local.update()

        app = MainWindow(root)
        time.sleep(0.3)

        splash_local.update_progress(100, _("Completato!"))
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
