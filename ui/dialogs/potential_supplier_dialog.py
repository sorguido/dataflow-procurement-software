"""
Potential Supplier Dialog - Dialog per creazione/modifica fornitori potenziali.

Riusabile per modalità NEW (supplier_id=None) e EDIT (supplier_id valorizzato).
Nessuna dipendenza da VSMEvent o da logica economica.
"""

import tkinter as tk
from tkinter import ttk
import logging
import webbrowser

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from services.supplier_persistence import (
    create_supplier,
    update_supplier,
    get_supplier_by_id,
    SupplierError,
)
from services.supplier_name_suggestion_service import SupplierNameSuggestionService
from services.supplier_category_persistence import (
    get_all_supplier_categories,
    ensure_supplier_category_exists,
    CategoryError,
)
from models.potential_supplier import (
    PotentialSupplier,
    SUPPLIER_STATUS_CHOICES,
)

from utils.i18n_utils import (
    tr,
    normalize_derisking_status,
    translate_derisking_status,
)


from utils.resource_utils import set_window_icon
from utils.validation_utils import is_valid_email, is_valid_website
from ui.dialogs.common_dialogs import SimpleMessageDialog
from ui.components.supplier_name_suggest import SupplierNameSuggestController

logger = logging.getLogger(__name__)


class PotentialSupplierDialog(tk.Toplevel):
    """
    Dialog per creazione e modifica di un fornitore potenziale.

    Utilizzo:
        dlg = PotentialSupplierDialog(parent, current_username)          # NEW
        dlg = PotentialSupplierDialog(parent, current_username, sid)     # EDIT
        if dlg.result:
            # salvataggio avvenuto con successo
    """

    def __init__(self, parent, current_username, supplier_id=None, read_only=False, refresh_derisking_cb=None):
        """
        Args:
            parent:               Widget parent (root o altra finestra)
            current_username:     Username dell'utente corrente
            supplier_id:          None → modalità NEW; int → modalità EDIT
            read_only:            Se True, tutti i campi disabilitati (sola lettura)
            refresh_derisking_cb: Callback opzionale da passare a ManageSupplierCategoriesDialog
        """
        super().__init__(parent)

        self.current_username = current_username
        self.supplier_id = supplier_id
        self.is_edit_mode = supplier_id is not None
        self.read_only = read_only
        self._refresh_derisking_cb = refresh_derisking_cb
        self.result = None  # True dopo salvataggio riuscito
        self._supplier_index = None
        self._suggest_controller = None

        # Nascondi durante costruzione UI
        self.withdraw()
        set_window_icon(self)

        if read_only:
            self.title(tr("View Supplier"))
        elif self.is_edit_mode:
            self.title(tr("Edit Supplier"))
        else:
            self.title(tr("New Supplier"))

        # Evita transient per mantenere i controlli finestra standard del WM
        # (massimizzazione inclusa dove disponibile).
        self.resizable(True, True)

        # --- Variabili tk ---
        self.var_supplier_name = tk.StringVar()
        self.var_category = tk.StringVar()       # selezione da dropdown
        self.var_new_category = tk.StringVar()   # nuova categoria testo libero
        self.var_status = tk.StringVar(value=translate_derisking_status(SUPPLIER_STATUS_CHOICES[0]))
        self.var_contact = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_phone = tk.StringVar()
        self.var_website = tk.StringVar()

        # Carica categorie esistenti dal DB per la combobox
        self._known_categories = self._load_known_categories()

        # Costruisci UI
        self._build_ui()
        self._init_supplier_suggestions()

        # Popola campo Utente (sempre disabled, auto-valorizzato)
        self._entry_username.configure(state="normal")
        self._entry_username.insert(0, self.current_username)
        self._entry_username.configure(state="disabled")

        # Carica dati se modalità EDIT
        if self.is_edit_mode:
            self._load_supplier_data()

        # Sola lettura: disabilita tutto
        if self.read_only:
            self._apply_read_only()

        # Geometria iniziale centrata e definita prima di mostrare la finestra.
        self.update_idletasks()
        base_width = self.winfo_reqwidth()
        base_height = self.winfo_reqheight()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        target_width = min(max(base_width + 160, int(screen_w * 0.56)), int(screen_w * 0.82))
        target_height = min(max(base_height + 80, int(screen_h * 0.58)), int(screen_h * 0.86))
        x = max(0, (screen_w - target_width) // 2)
        y = max(0, (screen_h - target_height) // 2)
        self.geometry(f"{target_width}x{target_height}+{x}+{y}")
        self.minsize(base_width, base_height)

        self.deiconify()
        self.wait_visibility()
        self.grab_set()

    def destroy(self):
        if self._suggest_controller is not None:
            self._suggest_controller.destroy()
            self._suggest_controller = None
        super().destroy()

    # -----------------------------------------------------------------------
    # UI BUILDING
    # -----------------------------------------------------------------------

    def _build_ui(self):
        """Costruisce l'intera interfaccia del dialog."""
        main = ttk.Frame(self, padding="20")
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        # --- Sezione: Informazioni Generali ---
        general = ttk.LabelFrame(main, text=tr("General Information"), padding="10")
        general.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        general.columnconfigure(1, weight=1)

        row = 0

        ttk.Label(general, text=tr("Supplier: *"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._entry_supplier_name = ttk.Entry(
            general, textvariable=self.var_supplier_name, width=38
        )
        self._entry_supplier_name.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1

        ttk.Label(general, text=tr("Category:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._combo_category = ttk.Combobox(
            general,
            textvariable=self.var_category,
            values=self._known_categories,
            state="readonly",
            width=36,
        )
        self._combo_category.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1

        ttk.Label(general, text=tr("New category:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=(0, 8)
        )
        self._entry_new_category = ttk.Entry(
            general, textvariable=self.var_new_category, width=38
        )
        self._entry_new_category.grid(row=row, column=1, sticky="ew", pady=(0, 8))
        row += 1

        ttk.Label(general, text=tr("Status:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._combo_status = ttk.Combobox(
            general,
            textvariable=self.var_status,
            values=[translate_derisking_status(c) for c in SUPPLIER_STATUS_CHOICES],
            state="readonly",
            width=16,
        )
        self._combo_status.grid(row=row, column=1, sticky="w", pady=5)
        row += 1

        ttk.Label(general, text=tr("User:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._entry_username = ttk.Entry(general, width=22, state="disabled")
        self._entry_username.grid(row=row, column=1, sticky="w", pady=5)

        # --- Sezione: Contatti ---
        contacts = ttk.LabelFrame(main, text=tr("Contacts"), padding="10")
        contacts.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        contacts.columnconfigure(1, weight=1)

        row = 0

        ttk.Label(contacts, text=tr("Contact:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._entry_contact = ttk.Entry(
            contacts, textvariable=self.var_contact, width=38
        )
        self._entry_contact.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1

        self._lbl_email = ttk.Label(contacts, text=tr("E-mail:"), font=(None, 10))
        self._lbl_email.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        self._entry_email = ttk.Entry(
            contacts, textvariable=self.var_email, width=38
        )
        self._entry_email.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1

        ttk.Label(contacts, text=tr("Phone:"), font=(None, 10)).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self._entry_phone = ttk.Entry(
            contacts, textvariable=self.var_phone, width=28
        )
        self._entry_phone.grid(row=row, column=1, sticky="w", pady=5)
        row += 1

        self._lbl_web = ttk.Label(contacts, text=tr("Web:"), font=(None, 10))
        self._lbl_web.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        self._entry_website = ttk.Entry(
            contacts, textvariable=self.var_website, width=38
        )
        self._entry_website.grid(row=row, column=1, sticky="ew", pady=5)

        # --- Sezione: Note ---
        notes_frame = ttk.LabelFrame(main, text=tr("Notes"), padding="10")
        notes_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 15))
        notes_frame.columnconfigure(0, weight=1)
        notes_frame.rowconfigure(0, weight=1)

        self._text_notes = tk.Text(
            notes_frame, height=4, width=50, wrap="word", font=(None, 10)
        )
        self._text_notes.grid(row=0, column=0, sticky="nsew", pady=5)

        # --- Pulsanti ---
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=3, column=0, sticky="ew")

        if self.read_only:
            ttk.Button(
                btn_frame, text=tr("Close"), command=self.destroy, width=12
            ).pack(side="right", padx=(5, 0))
        else:
            ttk.Button(
                btn_frame, text=tr("❌ Cancel"), command=self.destroy, width=12
            ).pack(side="right", padx=(5, 0))
            ttk.Button(
                btn_frame,
                text=tr("💾 Save"),
                command=self._on_save,
                width=12,
            ).pack(side="right")
            self._btn_manage_categories = ttk.Button(
                btn_frame,
                text=tr("Manage Categories"),
                command=self._on_manage_categories,
                width=18,
            )
            self._btn_manage_categories.pack(side="left")

        # --- Stile label cliccabili (definito una volta, globale all'app) ---
        _s = ttk.Style()
        if "ClickLink.TLabel" not in _s.theme_names():
            pass  # avoid theme check issues; just configure
        _s.configure("ClickLink.TLabel", foreground="#0055aa")

        # Aggiorna stato cliccabile e registra trace sui campi contatto
        self._update_clickable_contact_labels()
        self.var_email.trace_add("write", lambda *_: self._update_clickable_contact_labels())
        self.var_website.trace_add("write", lambda *_: self._update_clickable_contact_labels())

        # Chiusura con X
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _init_supplier_suggestions(self):
        # In read-only non attiviamo interazioni suggerimento.
        if self.read_only:
            return
        try:
            self._supplier_index = SupplierNameSuggestionService.build_index(get_db_path())
            self._suggest_controller = SupplierNameSuggestController(
                self,
                self._entry_supplier_name,
                self._get_supplier_suggestions,
                min_chars=2,
                max_items=8,
            )
        except Exception as e:
            logger.warning("Suggerimenti fornitori non disponibili nel dialog Derisking: %s", e)
            self._supplier_index = None
            self._suggest_controller = None

    def _get_supplier_suggestions(self, query: str) -> list:
        if not self._supplier_index:
            return []
        return self._supplier_index.suggest(query, limit=8)

    def _show_soft_duplicate_warning_if_needed(self, supplier_name: str):
        if not self._supplier_index:
            return
        candidates = self._supplier_index.get_soft_duplicate_candidates(supplier_name, limit=3)
        if not candidates:
            return
        SimpleMessageDialog(
            self,
            tr("Possible Supplier Duplicate"),
            tr(
                "The entered name may refer to an existing supplier.\n"
                "Possible matches:\n\n{}"
            ).format("\n".join(candidates)),
            "warning",
        )

    # -----------------------------------------------------------------------
    # DATA LOADING (modalità EDIT)
    # -----------------------------------------------------------------------

    def _load_supplier_data(self):
        """Carica i dati del fornitore dal DB e popola i campi."""
        try:
            with DatabaseManager(get_db_path()) as db:
                supplier = get_supplier_by_id(db, self.supplier_id)
        except DatabaseError as e:
            logger.error("Errore caricamento fornitore ID %s: %s", self.supplier_id, e)
            SimpleMessageDialog(
                self,
                tr("Database Error"),
                tr("Unable to load supplier data: {}").format(e),
                "error",
            )
            return

        if supplier is None:
            SimpleMessageDialog(
                self,
                tr("Supplier not found"),
                tr("The supplier with ID {} does not exist in the database.").format(self.supplier_id),
                "error",
            )
            return

        self.var_supplier_name.set(supplier.supplier_name or "")
        cat = supplier.category or ""
        if cat in self._known_categories:
            self.var_category.set(cat)
            self.var_new_category.set("")
        else:
            self.var_category.set("")
            self.var_new_category.set(cat)
        _canonical = normalize_derisking_status(supplier.supplier_status)
        if _canonical not in SUPPLIER_STATUS_CHOICES:
            _canonical = SUPPLIER_STATUS_CHOICES[0]
        self.var_status.set(translate_derisking_status(_canonical))
        self.var_contact.set(supplier.contact_name or "")
        self.var_email.set(supplier.email or "")
        self.var_phone.set(supplier.phone or "")
        self.var_website.set(supplier.website or "")

        self._text_notes.delete("1.0", tk.END)
        self._text_notes.insert("1.0", supplier.notes or "")

        # Aggiorna campo utente con il valore originale del record
        self._entry_username.configure(state="normal")
        self._entry_username.delete(0, tk.END)
        self._entry_username.insert(0, supplier.username or self.current_username)
        self._entry_username.configure(state="disabled")

        # Rifletti stato cliccabile label dopo caricamento dati
        self._update_clickable_contact_labels()

    # -----------------------------------------------------------------------
    # SALVATAGGIO
    # -----------------------------------------------------------------------

    def _on_save(self):
        """Valida i campi e salva il fornitore (NEW o EDIT)."""
        supplier_name = self.var_supplier_name.get().strip()
        if not supplier_name:
            SimpleMessageDialog(
                self,
                tr("Validation"),
                tr("The 'Supplier' field is required."),
                "error",
            )
            self._entry_supplier_name.focus_set()
            return

        # Warning soft non bloccante per nomi semanticamente simili.
        self._show_soft_duplicate_warning_if_needed(supplier_name)

        email_value = self.var_email.get().strip()
        if not is_valid_email(email_value):
            SimpleMessageDialog(self, tr("Validation"), tr("Invalid email format."), "error")
            self._entry_email.focus_set()
            return

        web_value = self.var_website.get().strip()
        if not is_valid_website(web_value):
            SimpleMessageDialog(self, tr("Validation"), tr("Invalid web URL format."), "error")
            self._entry_website.focus_set()
            return

        # Recupera l'username dal campo (always disabled, ma contiene il valore)
        self._entry_username.configure(state="normal")
        saved_username = self._entry_username.get().strip() or self.current_username
        self._entry_username.configure(state="disabled")

        category = self.var_new_category.get().strip() or self.var_category.get().strip()

        supplier = PotentialSupplier(
            id=self.supplier_id,  # None per NEW, int per EDIT
            supplier_name=supplier_name,
            category=category,
            supplier_status=normalize_derisking_status(self.var_status.get()),
            contact_name=self.var_contact.get().strip(),
            email=self.var_email.get().strip(),
            phone=self.var_phone.get().strip(),
            website=self.var_website.get().strip(),
            notes=self._text_notes.get("1.0", tk.END).strip(),
            username=saved_username,
        )

        try:
            with DatabaseManager(get_db_path()) as db:
                if category:
                    ensure_supplier_category_exists(db, category)
                if self.is_edit_mode:
                    update_supplier(db, supplier)
                    logger.info("Fornitore ID %s aggiornato.", self.supplier_id)
                else:
                    new_id = create_supplier(db, supplier)
                    logger.info("Nuovo fornitore creato con ID %s.", new_id)
        except (SupplierError, DatabaseError) as e:
            logger.error("Errore salvataggio fornitore: %s", e)
            SimpleMessageDialog(
                self,
                tr("Save Error"),
                tr("Unable to save the supplier:\n{}").format(e),
                "error",
            )
            return

        self.result = True

        # Aggiorna indice locale per coerenza in edit consecutivi nello stesso dialog.
        if self._supplier_index is not None:
            try:
                self._supplier_index = SupplierNameSuggestionService.build_index(get_db_path())
                if self._suggest_controller is not None:
                    self._suggest_controller.refresh()
            except Exception as idx_err:
                logger.warning("Impossibile aggiornare indice suggerimenti Derisking: %s", idx_err)

        self.destroy()

    # -----------------------------------------------------------------------
    # SOLA LETTURA
    # -----------------------------------------------------------------------

    def _apply_read_only(self):
        """Disabilita tutti i campi di input (modalità read_only)."""
        for widget in (
            self._entry_supplier_name,
            self._entry_new_category,
            self._entry_contact,
            self._entry_email,
            self._entry_phone,
            self._entry_website,
        ):
            widget.configure(state="disabled")

        self._combo_category.configure(state="disabled")
        self._combo_status.configure(state="disabled")
        self._text_notes.configure(state="disabled")

    # -----------------------------------------------------------------------
    # HELPERS
    # -----------------------------------------------------------------------

    def _load_known_categories(self) -> list:
        """Carica la lista di categorie distinte dal catalogo ufficiale."""
        try:
            with DatabaseManager(get_db_path()) as db:
                return get_all_supplier_categories(db)
        except Exception:
            return []

    def _refresh_categories(self, prefer: str = None):
        """
        Ricarica i valori della combo categorie dal DB.

        Preserva la selezione corrente se ancora valida.
        Se il valore corrente non è più nel catalogo:
          - se `prefer` è fornito (es. nuovo nome dopo rinomina/unione): usarlo
          - altrimenti svuota la selezione come fallback sicuro

        Args:
            prefer: valore da pre-selezionare dopo un'operazione che cambia il nome
        """
        current = self.var_category.get()
        try:
            with DatabaseManager(get_db_path()) as db:
                new_categories = get_all_supplier_categories(db)
        except Exception:
            new_categories = []

        self._known_categories = new_categories
        self._combo_category.configure(values=new_categories)

        if prefer and prefer in new_categories:
            self.var_category.set(prefer)
        elif current in new_categories:
            self.var_category.set(current)   # mantieni selezione valida
        else:
            self.var_category.set("")         # fallback: svuota

    def _on_manage_categories(self):
        """Apre il dialog Gestisci Categorie e refresha la combo al ritorno."""
        from ui.dialogs.manage_supplier_categories_dialog import ManageSupplierCategoriesDialog
        dlg = ManageSupplierCategoriesDialog(self, refresh_derisking_cb=self._refresh_derisking_cb)
        self.wait_window(dlg)
        if dlg.changes_made:
            self._refresh_categories()

    def _update_clickable_contact_labels(self):
        """Aggiorna stile e binding delle label E-mail e Web in base al contenuto."""
        email = self.var_email.get().strip()
        if email:
            self._lbl_email.configure(style="ClickLink.TLabel", cursor="hand2")
            self._lbl_email.bind("<Button-1>", self._on_email_click)
        else:
            self._lbl_email.configure(style="TLabel", cursor="")
            self._lbl_email.unbind("<Button-1>")

        website = self.var_website.get().strip()
        if website:
            self._lbl_web.configure(style="ClickLink.TLabel", cursor="hand2")
            self._lbl_web.bind("<Button-1>", self._on_web_click)
        else:
            self._lbl_web.configure(style="TLabel", cursor="")
            self._lbl_web.unbind("<Button-1>")

    def _on_email_click(self, event=None):
        """Copia l'e-mail negli appunti e mostra conferma."""
        email = self.var_email.get().strip()
        if not email:
            return
        self.clipboard_clear()
        self.clipboard_append(email)
        self.update()
        SimpleMessageDialog(self, tr("Info"), tr("Mail copied"), "info")

    def _on_web_click(self, event=None):
        """Apre il sito web nel browser predefinito."""
        url = self.var_website.get().strip()
        if not url:
            return
        if not url.startswith(("http://", "https://")):
            url = f"https://{url}"
        webbrowser.open(url)
