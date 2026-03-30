"""
Dialog componenti per l'applicazione DataFlow.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser
from PIL import Image, ImageTk
import os

from utils.resource_utils import resource_path, set_window_icon
from utils.window_utils import center_window
from utils.string_utils import generate_username
from utils.i18n_utils import get_current_language, _


class SimpleMessageDialog(tk.Toplevel):
    """Dialog semplice per messaggi con font uniforme all'app."""
    def __init__(self, parent, title, message, msg_type="info"):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()
        
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill="both", expand=True)
        
        # Messaggio con font uniforme (stesso di LanguagePrompt)
        ttk.Label(
            frame,
            text=message,
            font=(None, 10),
            wraplength=400,
            justify="left"
        ).pack(pady=(0, 15))
        
        # Pulsante OK
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(
            btn_frame,
            text="OK",
            command=self.destroy,
            width=10
        ).pack()
        
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        center_window(self)
        self.deiconify()
        self.wait_window()


class SimpleYesNoDialog(tk.Toplevel):
    """Dialog semplice per domande Yes/No con font uniforme all'app."""
    def __init__(self, parent, title, message, icon='question'):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(title)
        self.result = False  # Default: No
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()
        
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill="both", expand=True)
        
        # Messaggio con font uniforme (stesso di LanguagePrompt)
        ttk.Label(
            frame,
            text=message,
            font=(None, 10),
            wraplength=400,
            justify="left"
        ).pack(pady=(0, 15))
        
        # Pulsanti Yes/No
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        
        ttk.Button(
            btn_frame,
            text=_("Sì"),
            command=self._on_yes,
            width=10
        ).pack(side="left", padx=5)
        
        ttk.Button(
            btn_frame,
            text=_("No"),
            command=self._on_no,
            width=10
        ).pack(side="left", padx=5)
        
        self.protocol("WM_DELETE_WINDOW", self._on_no)
        center_window(self)
        self.deiconify()
        self.wait_window()
    
    def _on_yes(self):
        self.result = True
        self.destroy()
    
    def _on_no(self):
        self.result = False
        self.destroy()


class LanguagePrompt(tk.Toplevel):
    """Dialog per scelta lingua esportazione Excel."""
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("Scegli Lingua"))
        self.choice = None
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=_("In quale lingua vuoi esportare il file Excel?"), font=(None, 10)).pack(pady=(0, 15))

        lang_frame = ttk.Frame(frame)
        lang_frame.pack(pady=10)
        
        ttk.Label(lang_frame, text=_("Lingua:")).pack(side="left", padx=(0, 10))
        
        # Determina la lingua corrente dell'app
        current_lang = get_current_language()
        default_language = "Italiano" if current_lang == 'it' else "English"
        
        # Ordina le opzioni ponendo per prima la lingua correntemente usata dal programma
        values = ["Italiano", "English"] if default_language == "Italiano" else ["English", "Italiano"]

        self.language_var = tk.StringVar(value=default_language)
        language_combo = ttk.Combobox(lang_frame, textvariable=self.language_var,
                                      values=values,
                                      state="readonly", width=20)
        language_combo.pack(side="left", padx=(0, 10))
        language_combo.current(0)
        language_combo.bind("<<ComboboxSelected>>", lambda e: self.on_language_selected())
        
        btn_ok = ttk.Button(lang_frame, text="OK", command=self.confirm_choice)
        btn_ok.pack(side="left", padx=5)
        
        btn_cancel = ttk.Button(lang_frame, text=_("❌ Annulla"), command=self.on_close)
        btn_cancel.pack(side="left", padx=5)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        center_window(self)
    
    def on_language_selected(self):
        """Gestisce la selezione della lingua"""
        pass  # Già gestita dalla variabile
    
    def confirm_choice(self):
        """Conferma la scelta e chiude la finestra"""
        selected = self.language_var.get()
        if selected == "Italiano":
            self.choice = "ita"
        elif selected == "English":
            self.choice = "eng"
        else:
            self.choice = None
        self.destroy()

    def on_close(self):
        self.choice = None
        self.destroy()


class NewRdOTypeDialog(tk.Toplevel):
    """Dialog minimale per scegliere il tipo di RdO da creare"""
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        
        self.title(_("Nuova Richiesta di Offerta"))
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
            messagebox.showerror(_("Campi obbligatori"), _("Inserisci sia il nome sia il cognome."), parent=self)
            return
        try:
            username = generate_username(first, last)
        except ValueError as e:
            messagebox.showerror(_("Formato non valido"), str(e), parent=self)
            return
        self.result = {
            'first_name': first,
            'last_name': last,
            'username': username
        }
        self.grab_release()
        self.destroy()

    def _prevent_close(self):
        messagebox.showwarning(_("Operazione necessaria"), _("Per utilizzare DataFlow è necessario completare i dati richiesti."), parent=self)

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


class SplashScreen(tk.Toplevel):
    """Finestra di avvio splash screen."""
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_("Avvio DataFlow"))
        self.overrideredirect(True)
        
        frame = ttk.Frame(self, borderwidth=2, relief="raised")
        frame.pack(fill="both", expand=True)
        
        try:
            logo_path = resource_path(os.path.join("add_data", "logo_dataflow.png"))
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                if img.width > 0 and img.height > 0:
                    img.thumbnail((273, int(273 * (img.height/img.width))), Image.Resampling.LANCZOS)
                    self.logo_photo = ImageTk.PhotoImage(img)
                    ttk.Label(frame, image=self.logo_photo).pack(pady=(30, 20))
        except Exception as e:
            print(f"Errore logo splash: {e}")
            ttk.Label(frame, text=_("DataFlow"), font=("Helvetica", 24, "bold")).pack(pady=(30, 20))
        
        self.status_label = ttk.Label(
            frame, 
            text=_("Avvio in corso..."), 
            font=("Helvetica", 10),
            width=40,
            anchor="center"
        )
        self.status_label.pack(pady=(10, 5))
        self.progress = ttk.Progressbar(frame, orient="horizontal", length=300, mode='determinate')
        self.progress.pack(pady=(0, 20))
        
        self.update_idletasks()
        
        w = 450
        h = 250
        
        x = (self.winfo_screenwidth()//2) - (w//2)
        y = (self.winfo_screenheight()//2) - (h//2)
        
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        self.deiconify()

    def update_progress(self, val, txt):
        """Aggiorna barra di progresso."""
        self.progress['value'] = val
        self.status_label['text'] = ""
        self.update_idletasks()
        self.status_label['text'] = txt
        self.update_idletasks()
