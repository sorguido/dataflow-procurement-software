"""
DataFlow Database Migration Tool - UI Dialogs
Tkinter-based GUI dialogs for user interaction
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys


# --- INIZIO CODICE AGGIUNTO PER PYINSTALLER ---
def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    
    Args:
        relative_path: Path relative to the script/bundle
    
    Returns:
        Absolute path to resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Normal Python execution
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
# --- FINE CODICE AGGIUNTO PER PYINSTALLER ---


def select_config_file(parent=None):
    """
    Show file dialog to select config.ini file.
    
    Args:
        parent: Parent window (optional)
    
    Returns:
        str: Selected file path or None if cancelled
    """
    if parent:
        # Bring parent window to front
        parent.attributes('-topmost', True)
        parent.update()
    
    file_path = filedialog.askopenfilename(
        parent=parent,
        title="Select DataFlow 2.0.0 config.ini file",
        filetypes=[("Configuration files", "*.ini"), ("All files", "*.*")],
        initialdir=os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DataFlow')
    )
    
    return file_path if file_path else None


def select_source_database(parent=None):
    """
    Show file dialog to select source v1.1.0 database file.
    
    Args:
        parent: Parent window (optional)
    
    Returns:
        str: Selected file path or None if cancelled
    """
    if parent:
        # Bring parent window to front
        parent.attributes('-topmost', True)
        parent.update()
    
    file_path = filedialog.askopenfilename(
        parent=parent,
        title="Select DataFlow 1.1.0 database file",
        filetypes=[("Database files", "*.db"), ("All files", "*.*")],
        initialdir=os.path.expanduser('~\\Documents')
    )
    
    return file_path if file_path else None


def show_error_dialog(title, message, parent=None):
    """
    Show error dialog.
    
    Args:
        title: Dialog title
        message: Error message
        parent: Parent window (optional)
    """
    if parent:
        # Bring parent window to front
        parent.attributes('-topmost', True)
        parent.update()
        messagebox.showerror(title, message, parent=parent)
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()


def show_info_dialog(title, message, parent=None):
    """
    Show information dialog.
    
    Args:
        title: Dialog title
        message: Information message
        parent: Parent window (optional)
    """
    if parent:
        # Bring parent window to front
        parent.attributes('-topmost', True)
        parent.update()
        messagebox.showinfo(title, message, parent=parent)
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()


class WelcomeDialog:
    """
    Shows welcome dialog with migration tool explanation.
    This dialog remains open during file selection to keep the taskbar icon visible.
    """
    
    def __init__(self):
        """Initialize welcome dialog."""
        self.result = False
        self.continue_with_selection = False
        
        self.root = tk.Tk()
        # Hide window IMMEDIATELY to prevent flicker
        self.root.withdraw()
        
        self.root.title("DataFlow Database Migration Tool")
        
        # Set window icon
        try:
            self.root.iconbitmap(resource_path(os.path.join("add_data", "DataFlow.ico")))
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        # Set size (increased height for buttons)
        window_width = 650
        window_height = 600
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)
        
        # Keep window on top
        self.root.attributes('-topmost', True)
        
        # Center window (while hidden)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (window_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        self._create_widgets()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Title section
        title_frame = ttk.Frame(self.root, padding="20")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame,
            text="🔄 DataFlow Database Migration Tool",
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack()
        
        version_label = ttk.Label(
            title_frame,
            text="v1.1.0 → v2.0.0",
            font=("Segoe UI", 11)
        )
        version_label.pack(pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(self.root, orient="horizontal")
        separator.pack(fill=tk.X, padx=20)
        
        # Scrollable text section
        text_frame = ttk.Frame(self.root, padding="20")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=15,
            pady=10,
            bg="#f5f5f5",
            height=15
        )
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Build description text
        description = self._build_description_text()
        text_widget.insert("1.0", description)
        text_widget.config(state=tk.DISABLED)
        
        # Buttons section
        btn_frame = ttk.Frame(self.root, padding="20")
        btn_frame.pack(fill=tk.X)
        
        cancel_btn = ttk.Button(
            btn_frame,
            text="❌ Cancel",
            command=self._on_cancel,
            width=15
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        proceed_btn = ttk.Button(
            btn_frame,
            text="▶️ Start Migration",
            command=self._on_proceed,
            width=25
        )
        proceed_btn.pack(side=tk.RIGHT, padx=5)
    
    def _build_description_text(self):
        """Build welcome message text"""
        text = """Welcome to the DataFlow Migration Tool!

This tool will help you migrate your database from DataFlow 1.1.0 to the new version 2.0.0.

🔹 WHAT THIS TOOL DOES:
   • Migrates all data (RfQs, articles, suppliers)
   • Converts attachments from BLOB to filesystem
   • Assigns a username to all RfQs
   • Regenerates IDs in year-based format
   • Creates a detailed migration report

⚠️ WARNING - DESTRUCTIVE OPERATION:
   This tool will COMPLETELY DELETE the destination folder before proceeding
   with the migration. Make sure you don't have important data in the new
   DataFlow 2.0.0 installation.

📁 WHAT YOU WILL BE ASKED FOR:

1. DataFlow 2.0.0 config.ini file
   • This file is normally located at:
     C:\\Users\\<YourName>\\AppData\\Local\\DataFlow\\config.ini
   • Must be already configured (run DataFlow 2.0.0 at least once)
   • Contains your username that will be assigned to RfQs

2. DataFlow 1.1.0 Database (file .db)
   • The database from the previous version you want to migrate
   • Usually located in the Documents folder
   • Example: gestione_offerte.db

💡 SUGGESTIONS:
   Before proceeding, make sure you have:
   ✓ Run DataFlow 2.0.0 at least once
   ✓ Configure your username in DataFlow 2.0.0
   ✓ A backup of your v1.1.0 database
   ✓ Enough time to complete the migration

Ready to begin?
"""
        return text
    
    def _on_proceed(self):
        """Handle proceed button"""
        self.result = True
        self.continue_with_selection = True
        # Iconify (minimize) instead of destroying to keep taskbar icon
        self.root.iconify()
    
    def _on_cancel(self):
        """Handle cancel button"""
        self.result = False
        self.continue_with_selection = True  # Exit the wait loop
        self.root.quit()
    
    def show(self):
        """
        Show dialog and wait for user to click Start or Cancel.
        Window is minimized (not destroyed) when user proceeds.
        
        Returns:
            bool: True if user clicked proceed, False if cancelled
        """
        # Show window (now centered and configured)
        self.root.deiconify()
        
        # Bring to front and focus
        self.root.lift()
        self.root.focus_force()
        
        # Wait for user to proceed or cancel
        while not self.continue_with_selection and self.result == False:
            self.root.update()
            if not self.root.winfo_exists():
                return False
        
        return self.result
    
    def get_root(self):
        """Get the Tkinter root window for use as parent"""
        return self.root
    
    def close(self):
        """Close the welcome dialog"""
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass


class MigrationSummaryDialog:
    """
    Shows migration summary and first confirmation dialog.
    """
    
    def __init__(self, source_db_path, target_paths, username, db_summary, parent=None):
        """
        Initialize summary dialog.
        
        Args:
            source_db_path: Path to source database
            target_paths: Dict with target paths
            username: Username from config
            db_summary: Dict with database summary (from schema_validator)
            parent: Parent window (optional, will use main window if not provided)
        """
        self.result = False
        self.source_db_path = source_db_path
        self.target_paths = target_paths
        self.username = username
        self.db_summary = db_summary
        
        # Use provided parent or create new Tk root
        if parent is None:
            self.root = tk.Tk()
            self.is_own_root = True
        else:
            self.root = tk.Toplevel(parent)
            self.is_own_root = False
        
        # Hide window IMMEDIATELY to prevent flicker
        self.root.withdraw()
        
        self.root.title("Migration Summary")
        
        # Set window icon
        try:
            self.root.iconbitmap(resource_path(os.path.join("add_data", "DataFlow.ico")))
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        window_width = 700
        window_height = 550
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)
        
        # Keep window on top
        self.root.attributes('-topmost', True)
        
        # Center window (while hidden)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (window_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        self._create_widgets()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Title section
        title_frame = ttk.Frame(self.root, padding="20")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame,
            text="Migration Summary",
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack()
        
        # Scrollable text section
        text_frame = ttk.Frame(self.root, padding="20")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
            relief=tk.SOLID,
            borderwidth=1,
            padx=10,
            pady=10,
            height=15
        )
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Build summary text
        summary_text = self._build_summary_text()
        text_widget.insert("1.0", summary_text)
        text_widget.config(state=tk.DISABLED)
        
        # Configure text tags for formatting
        text_widget.tag_configure("heading", font=("Consolas", 10, "bold"))
        text_widget.tag_configure("warning", foreground="red", font=("Consolas", 10, "bold"))
        
        # Apply tags
        self._apply_text_tags(text_widget, summary_text)
        
        # Buttons section
        btn_frame = ttk.Frame(self.root, padding="20")
        btn_frame.pack(fill=tk.X)
        
        cancel_btn = ttk.Button(
            btn_frame,
            text="❌ Cancel",
            command=self._on_cancel,
            width=15
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        proceed_btn = ttk.Button(
            btn_frame,
            text="⚠️ I UNDERSTAND - PROCEED",
            command=self._on_proceed,
            width=30
        )
        proceed_btn.pack(side=tk.RIGHT, padx=5)
    
    def _build_summary_text(self):
        """Build summary text content"""
        target_exists = os.path.exists(self.target_paths['base_dir'])
        
        text = "=" * 68 + "\n"
        text += "DATABASE MIGRATION: DataFlow 1.1.0 → 2.0.0\n"
        text += "=" * 68 + "\n\n"
        
        text += "SOURCE DATABASE:\n"
        text += f"  Path: {self.source_db_path}\n"
        text += f"  RfQs: {self.db_summary['rfqs']}\n"
        text += f"  Articles: {self.db_summary['articles']}\n"
        text += f"  Suppliers: {self.db_summary['suppliers']}\n"
        text += f"  Attachments: {self.db_summary['attachments']}\n"
        
        att_types = self.db_summary['attachment_types']
        if att_types['blob'] > 0:
            text += f"    - BLOB attachments: {att_types['blob']}\n"
        if att_types['external'] > 0:
            text += f"    - External file attachments: {att_types['external']}\n"
        if att_types['both'] > 0:
            text += f"    - Hybrid (BLOB + external): {att_types['both']}\n"
        
        text += "\n" + "-" * 68 + "\n\n"
        
        text += "TARGET LOCATION:\n"
        text += f"  Folder: {self.target_paths['base_dir']}\n"
        text += f"  Database: {self.target_paths['db_file']}\n"
        text += f"  Attachments: {self.target_paths['attachments_dir']}\n"
        text += f"  Username: {self.username}\n\n"
        
        if target_exists:
            text += "⚠️ WARNING: TARGET FOLDER EXISTS AND WILL BE COMPLETELY DELETED!\n\n"
            
            # Try to get info about existing content
            try:
                existing_db = self.target_paths['db_file']
                if os.path.exists(existing_db):
                    import sqlite3
                    conn = sqlite3.connect(existing_db, timeout=5.0)
                    cursor = conn.cursor()
                    cursor.execute("SELECT COUNT(*) FROM richieste_offerta")
                    existing_rfqs = cursor.fetchone()[0]
                    conn.close()
                    text += f"  Existing database contains {existing_rfqs} RfQs that will be LOST!\n"
            except:
                pass
            
            # Count existing attachments
            try:
                att_dir = self.target_paths['attachments_dir']
                if os.path.exists(att_dir):
                    att_count = len([f for f in os.listdir(att_dir) if os.path.isfile(os.path.join(att_dir, f))])
                    if att_count > 0:
                        text += f"  {att_count} existing attachment files will be DELETED!\n"
            except:
                pass
            
            text += "\n"
        else:
            text += "✓ Target folder does not exist yet (will be created)\n\n"
        
        text += "-" * 68 + "\n\n"
        
        text += "MIGRATION OPERATIONS:\n"
        text += f"  • Delete target folder (if exists): {self.target_paths['base_dir']}\n"
        text += "  • Create fresh folder structure\n"
        text += "  • Create new v2.0.0 database with WAL mode enabled\n"
        text += f"  • Migrate {self.db_summary['rfqs']} RfQs with ID remapping\n"
        text += f"  • Assign username '{self.username}' to all RfQs\n"
        text += f"  • Extract {self.db_summary['attachments']} attachments to filesystem\n"
        text += "  • Verify foreign key integrity\n"
        text += "  • Generate migration report\n\n"
        
        text += "=" * 68 + "\n\n"
        
        if target_exists:
            text += "⚠️ THIS OPERATION IS IRREVERSIBLE!\n"
            text += "⚠️ ALL EXISTING DATA IN TARGET FOLDER WILL BE PERMANENTLY DELETED!\n\n"
        
        text += "Click 'I UNDERSTAND - PROCEED' to continue or 'Cancel' to abort.\n"
        
        return text
    
    def _apply_text_tags(self, widget, text):
        """Apply formatting tags to text widget"""
        # This is a simple implementation - could be enhanced
        pass
    
    def _on_proceed(self):
        """Handle proceed button click"""
        self.result = True
        self.root.destroy()
    
    def _on_cancel(self):
        """Handle cancel button click"""
        self.result = False
        self.root.destroy()
    
    def show(self):
        """
        Show dialog and wait for user input.
        
        Returns:
            bool: True if user clicked Proceed, False otherwise
        """
        # Show window (now centered and configured)
        self.root.deiconify()
        
        # Bring to front and focus
        self.root.lift()
        self.root.focus_force()
        
        self.root.wait_window()
        return self.result


class FinalConfirmationDialog:
    """
    Shows final confirmation dialog with 'DELETE' typing requirement.
    """
    
    def __init__(self, target_base_dir, parent=None):
        """
        Initialize final confirmation dialog.
        
        Args:
            target_base_dir: Path to target folder that will be deleted
            parent: Parent window (optional, will use main window if not provided)
        """
        self.result = False
        self.target_base_dir = target_base_dir
        
        # Use provided parent or create new Tk root
        if parent is None:
            self.root = tk.Tk()
            self.is_own_root = True
        else:
            self.root = tk.Toplevel(parent)
            self.is_own_root = False
        
        # Hide window IMMEDIATELY to prevent flicker
        self.root.withdraw()
        
        self.root.title("⚠️ FINAL CONFIRMATION")
        
        # Set window icon
        try:
            self.root.iconbitmap(resource_path(os.path.join("add_data", "DataFlow.ico")))
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        window_width = 600
        window_height = 520
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)
        
        # Keep window on top
        self.root.attributes('-topmost', True)
        
        # Center window (while hidden)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (window_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        self._create_widgets()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Warning icon and title
        title_label = ttk.Label(
            main_frame,
            text="⚠️ FINAL CONFIRMATION ⚠️",
            font=("Segoe UI", 16, "bold"),
            foreground="red"
        )
        title_label.pack(pady=(0, 15))
        
        # Warning message
        warning_text = (
            "This operation is IRREVERSIBLE and will:\n\n"
            f"• DELETE the entire folder:\n"
            f"  {self.target_base_dir}\n\n"
            "• All existing data in this folder will be PERMANENTLY LOST\n\n"
            "• A new folder will be created with migrated data\n\n"
            "• The source database will remain UNTOUCHED\n\n"
        )
        
        warning_label = ttk.Label(
            main_frame,
            text=warning_text,
            font=("Segoe UI", 10),
            justify=tk.LEFT
        )
        warning_label.pack(pady=(0, 15))
        
        # Confirmation instruction
        confirm_label = ttk.Label(
            main_frame,
            text="Type 'DELETE' to confirm (case-sensitive):",
            font=("Segoe UI", 11, "bold")
        )
        confirm_label.pack(pady=(0, 10))
        
        # Entry field
        self.entry_var = tk.StringVar()
        self.entry = ttk.Entry(
            main_frame,
            textvariable=self.entry_var,
            font=("Segoe UI", 12),
            width=30,
            justify=tk.CENTER
        )
        self.entry.pack(pady=(0, 15))
        
        # Bind Enter key
        self.entry.bind("<Return>", lambda e: self._on_proceed())
        
        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        cancel_btn = ttk.Button(
            btn_frame,
            text="❌ Cancel",
            command=self._on_cancel,
            width=15
        )
        cancel_btn.pack(side=tk.RIGHT)
        
        self.proceed_btn = ttk.Button(
            btn_frame,
            text="✅ Proceed with Migration",
            command=self._on_proceed,
            state=tk.DISABLED,
            width=25
        )
        self.proceed_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # Enable proceed button only when 'DELETE' is typed
        self.entry_var.trace_add("write", self._on_entry_changed)
    
    def _on_entry_changed(self, *args):
        """Enable proceed button when 'DELETE' is typed"""
        if self.entry_var.get() == "DELETE":
            self.proceed_btn.config(state=tk.NORMAL)
        else:
            self.proceed_btn.config(state=tk.DISABLED)
    
    def _on_proceed(self):
        """Handle proceed button click"""
        if self.entry_var.get() == "DELETE":
            self.result = True
            self.root.destroy()
    
    def _on_cancel(self):
        """Handle cancel button click"""
        self.result = False
        self.root.destroy()
    
    def show(self):
        """
        Show dialog and wait for user response.
        
        Returns:
            bool: True if user confirmed (typed DELETE), False otherwise
        """
        # Show window (now centered and configured)
        self.root.deiconify()
        
        # Bring to front and focus
        self.root.lift()
        self.root.focus_force()
        
        # Focus on entry field
        self.entry.focus()
        
        self.root.wait_window()
        return self.result


class ProgressDialog:
    """
    Shows migration progress with progress bar and step messages.
    """
    
    def __init__(self, total_steps=10, parent=None):
        """
        Initialize progress dialog.
        
        Args:
            total_steps: Total number of migration steps
            parent: Parent window (optional, will use main window if not provided)
        """
        self.total_steps = total_steps
        self.current_step = 0
        
        # Use provided parent or create new Tk root
        if parent is None:
            self.root = tk.Tk()
            self.is_own_root = True
        else:
            self.root = tk.Toplevel(parent)
            self.is_own_root = False
        
        # Hide window IMMEDIATELY to prevent flicker
        self.root.withdraw()
        
        self.root.title("Migration Progress")
        
        # Set window icon
        try:
            self.root.iconbitmap(resource_path(os.path.join("add_data", "DataFlow.ico")))
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        window_width = 700
        window_height = 220
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)
        
        # Keep window on top
        self.root.attributes('-topmost', True)
        
        # Center window (while hidden)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (window_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        # Prevent closing during migration
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="⏳ Migrating DataFlow 1.1.0 → 2.0.0...",
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=640,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 20))
        
        # Step label
        self.step_label = ttk.Label(
            main_frame,
            text="Initializing...",
            font=("Segoe UI", 10)
        )
        self.step_label.pack()
        
        # Percentage label
        self.percent_label = ttk.Label(
            main_frame,
            text="0%",
            font=("Segoe UI", 12, "bold")
        )
        self.percent_label.pack(pady=(10, 0))
    
    def update(self, step, message):
        """
        Update progress dialog.
        
        Args:
            step: Current step number (1-based)
            message: Step message to display
        """
        self.current_step = step
        progress_percent = (step / self.total_steps) * 100
        
        self.progress_var.set(progress_percent)
        self.step_label.config(text=f"Step {step}/{self.total_steps}: {message}")
        self.percent_label.config(text=f"{progress_percent:.0f}%")
        
        self.root.update()
    
    def close(self):
        """Close progress dialog"""
        self.root.destroy()
    
    def show(self):
        """Show progress dialog (non-blocking)"""
        # Show window (now centered and configured)
        self.root.deiconify()
        
        # Bring to front and focus
        self.root.lift()
        self.root.focus_force()
        
        self.root.update()


class CompletionDialog:
    """
    Shows migration completion summary with statistics.
    """
    
    def __init__(self, statistics, log_file_path, target_db_path, parent=None):
        """
        Initialize completion dialog.
        
        Args:
            statistics: Migration statistics dict
            log_file_path: Path to detailed log file
            target_db_path: Path to migrated database
            parent: Parent window (optional, will use main window if not provided)
        """
        self.statistics = statistics
        self.log_file_path = log_file_path
        self.target_db_path = target_db_path
        self.open_folder = False
        
        # Use provided parent or create new Tk root
        if parent is None:
            self.root = tk.Tk()
            self.is_own_root = True
        else:
            self.root = tk.Toplevel(parent)
            self.is_own_root = False
        
        # Hide window IMMEDIATELY to prevent flicker
        self.root.withdraw()
        
        self.root.title("Migration Complete")
        
        # Set window icon
        try:
            self.root.iconbitmap(resource_path(os.path.join("add_data", "DataFlow.ico")))
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")
        
        window_width = 700
        window_height = 600
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)
        
        # Keep window on top
        self.root.attributes('-topmost', True)
        
        # Center window (while hidden)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (window_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        self._create_widgets()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Success icon and title
        title_label = ttk.Label(
            main_frame,
            text="✅ Migration Completed Successfully!",
            font=("Segoe UI", 16, "bold"),
            foreground="green"
        )
        title_label.pack(pady=(0, 20))
        
        # Statistics frame
        stats_frame = ttk.LabelFrame(main_frame, text="Migration Statistics", padding="10")
        stats_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Build statistics text
        stats_text = self._build_statistics_text()
        
        # Text widget for statistics
        text_widget = tk.Text(
            stats_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            relief=tk.FLAT,
            height=15
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", stats_text)
        text_widget.config(state=tk.DISABLED)
        
        # Log file info
        log_info = f"Detailed log file saved to:\n{self.log_file_path}"
        log_label = ttk.Label(
            main_frame,
            text=log_info,
            font=("Segoe UI", 9),
            foreground="gray"
        )
        log_label.pack(pady=(0, 20))
        
        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        close_btn = ttk.Button(
            btn_frame,
            text="Close",
            command=self._on_close
        )
        close_btn.pack(side=tk.RIGHT)
        
        open_folder_btn = ttk.Button(
            btn_frame,
            text="📁 Open Database Folder",
            command=self._on_open_folder
        )
        open_folder_btn.pack(side=tk.RIGHT, padx=(0, 10))
    
    def _build_statistics_text(self):
        """Build statistics text"""
        stats = self.statistics
        
        text = "=" * 70 + "\n"
        text += "MIGRATION SUMMARY\n"
        text += "=" * 70 + "\n\n"
        
        text += f"Duration: {stats.get('duration_seconds', 0):.1f} seconds\n\n"
        
        text += "Data Migrated:\n"
        text += f"  • RfQs: {stats.get('rfqs_migrated', 0)}\n"
        text += f"  • Articles: {stats.get('articles_migrated', 0)}\n"
        text += f"  • Suppliers: {stats.get('suppliers_migrated', 0)}\n"
        text += f"  • Prices: {stats.get('prices_migrated', 0)}\n"
        text += f"  • Attachments: {stats.get('attachments_migrated', 0)}\n\n"
        
        warnings = stats.get('warnings', [])
        if warnings:
            text += f"Warnings ({len(warnings)}):\n"
            for warning in warnings[:10]:  # Show first 10
                text += f"  ⚠ {warning}\n"
            if len(warnings) > 10:
                text += f"  ... and {len(warnings) - 10} more (see log file)\n"
            text += "\n"
        
        errors = stats.get('errors', [])
        if errors:
            text += f"Errors ({len(errors)}):\n"
            for error in errors:
                text += f"  ❌ {error}\n"
            text += "\n"
        
        text += "=" * 70 + "\n\n"
        text += "Migration completed successfully!\n"
        text += f"Database ready at:\n{self.target_db_path}\n"
        
        return text
    
    def _on_open_folder(self):
        """Open database folder in file explorer"""
        self.open_folder = True
        self.root.destroy()
    
    def _on_close(self):
        """Close dialog"""
        self.root.destroy()
    
    def show(self):
        """
        Show completion dialog.
        
        Returns:
            bool: True if user clicked 'Open Folder', False otherwise
        """
        # Show window (now centered and configured)
        self.root.deiconify()
        
        # Bring to front and focus
        self.root.lift()
        self.root.focus_force()
        
        self.root.wait_window()
        return self.open_folder
