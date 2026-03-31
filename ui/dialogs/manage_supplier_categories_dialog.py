"""
ManageSupplierCategoriesDialog — Dialog per la gestione centralizzata
delle categorie dei fornitori potenziali.

Operazioni supportate:
  - Rinomina categoria
  - Unisci categoria sorgente → destinazione
  - Elimina categoria se non usata da alcun supplier

Utilizzo:
    dlg = ManageSupplierCategoriesDialog(parent)
    parent.wait_window(dlg)
    if dlg.changes_made:
        # aggiornare combo categorie nel dialog chiamante
"""

import tkinter as tk
from tkinter import ttk
import logging

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from services.supplier_category_persistence import (
    get_all_supplier_categories,
    rename_supplier_category,
    merge_supplier_categories,
    delete_supplier_category_if_unused,
    count_suppliers_by_category,
    CategoryError,
)
from utils.i18n_utils import _
from utils.resource_utils import set_window_icon
from utils.window_utils import center_window
from ui.dialogs.common_dialogs import SimpleMessageDialog, SimpleYesNoDialog

logger = logging.getLogger(__name__)


class ManageSupplierCategoriesDialog(tk.Toplevel):
    """
    Dialog per rinominare, unire ed eliminare le categorie dei fornitori potenziali.

    self.changes_made è impostato a True se almeno un'operazione di scrittura
    è andata a buon fine — il parent può usarlo per refreshare la propria combo.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.changes_made = False

        self.withdraw()
        set_window_icon(self)
        self.title(_("Gestisci Categorie"))
        self.transient(parent)
        self.resizable(False, False)

        self._build_ui()
        self._refresh_list()

        center_window(self)
        self.wait_visibility()
        self.grab_set()
        self.deiconify()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    # -----------------------------------------------------------------------
    # UI BUILDING
    # -----------------------------------------------------------------------

    def _build_ui(self):
        """Costruisce il layout del dialog."""
        outer = ttk.Frame(self, padding="15")
        outer.pack(fill="both", expand=True)

        # ------------------------------------------------------------------ #
        # Riga superiore: lista categorie (sinistra) + sezioni azioni (destra)
        # ------------------------------------------------------------------ #
        top_frame = ttk.Frame(outer)
        top_frame.pack(fill="both", expand=True, pady=(0, 10))

        # --- Lista categorie ---
        list_frame = ttk.LabelFrame(top_frame, text=_("Categorie"), padding="8")
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        self._listbox = tk.Listbox(
            list_frame,
            font=(None, 10),
            selectmode="single",
            width=26,
            height=14,
            yscrollcommand=scrollbar.set,
            exportselection=False,
        )
        scrollbar.config(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")
        self._listbox.bind("<<ListboxSelect>>", self._on_selection_changed)

        # --- Sezioni azioni (a destra) ---
        actions_frame = ttk.Frame(top_frame)
        actions_frame.pack(side="left", fill="both", expand=True)

        # == Rinomina ==
        rename_frame = ttk.LabelFrame(actions_frame, text=_("Rinomina"), padding="8")
        rename_frame.pack(fill="x", pady=(0, 8))
        rename_frame.columnconfigure(1, weight=1)

        ttk.Label(rename_frame, text=_("Nuovo nome:"), font=(None, 10)).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.var_new_name = tk.StringVar()
        self._entry_new_name = ttk.Entry(rename_frame, textvariable=self.var_new_name, width=22)
        self._entry_new_name.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(
            rename_frame,
            text=_("Rinomina"),
            command=self._on_rename,
            width=12,
        ).grid(row=1, column=0, columnspan=2, sticky="e", pady=(4, 0))

        # == Unisci ==
        merge_frame = ttk.LabelFrame(actions_frame, text=_("Unisci"), padding="8")
        merge_frame.pack(fill="x", pady=(0, 8))
        merge_frame.columnconfigure(1, weight=1)

        ttk.Label(merge_frame, text=_("Unisci con:"), font=(None, 10)).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.var_merge_target = tk.StringVar()
        self._combo_merge_target = ttk.Combobox(
            merge_frame,
            textvariable=self.var_merge_target,
            state="readonly",
            width=20,
        )
        self._combo_merge_target.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(
            merge_frame,
            text=_("Unisci"),
            command=self._on_merge,
            width=12,
        ).grid(row=1, column=0, columnspan=2, sticky="e", pady=(4, 0))

        # == Elimina se non usata ==
        delete_frame = ttk.LabelFrame(
            actions_frame, text=_("Elimina se non usata"), padding="8"
        )
        delete_frame.pack(fill="x")

        self._lbl_supplier_count = ttk.Label(
            delete_frame, text="", font=(None, 10), foreground="#555555"
        )
        self._lbl_supplier_count.pack(anchor="w", pady=(0, 4))

        ttk.Button(
            delete_frame,
            text=_("Elimina se non usata"),
            command=self._on_delete,
            width=22,
        ).pack(anchor="e", pady=(4, 0))

        # ------------------------------------------------------------------ #
        # Pulsante Chiudi (in basso a destra)
        # ------------------------------------------------------------------ #
        btn_frame = ttk.Frame(outer)
        btn_frame.pack(fill="x")

        ttk.Button(
            btn_frame,
            text=_("Chiudi"),
            command=self.destroy,
            width=12,
        ).pack(side="right")

    # -----------------------------------------------------------------------
    # REFRESH
    # -----------------------------------------------------------------------

    def _refresh_list(self, keep_selection: str = None):
        """
        Ricarica la lista dal DB e aggiorna Listbox e combo merge.

        Args:
            keep_selection: se fornito, tenta di riselezionare questo valore.
        """
        try:
            with DatabaseManager(get_db_path()) as db:
                categories = get_all_supplier_categories(db)
        except (DatabaseError, CategoryError) as e:
            logger.error("Errore refresh categorie: %s", e)
            categories = []

        self._listbox.delete(0, tk.END)
        for cat in categories:
            self._listbox.insert(tk.END, cat)

        # Aggiorna combo merge
        self._combo_merge_target.configure(values=categories)

        # Ripristina selezione se possibile
        if keep_selection and keep_selection in categories:
            idx = categories.index(keep_selection)
            self._listbox.selection_set(idx)
            self._listbox.see(idx)
            self._update_count_label(keep_selection)
        else:
            self._lbl_supplier_count.configure(text="")
            self.var_new_name.set("")
            self.var_merge_target.set("")

    def _get_selected_category(self) -> str:
        """Restituisce la categoria selezionata nella Listbox, o stringa vuota."""
        sel = self._listbox.curselection()
        if not sel:
            return ""
        return self._listbox.get(sel[0])

    def _on_selection_changed(self, _event=None):
        """Aggiorna entry nuovo nome e contatore quando cambia selezione."""
        cat = self._get_selected_category()
        if cat:
            self.var_new_name.set(cat)
            self._update_count_label(cat)
        else:
            self._lbl_supplier_count.configure(text="")

    def _update_count_label(self, name: str):
        """Aggiorna il label con il numero di supplier associati."""
        try:
            with DatabaseManager(get_db_path()) as db:
                count = count_suppliers_by_category(db, name)
            self._lbl_supplier_count.configure(
                text=_("Fornitori associati: {}").format(count)
            )
        except Exception:
            self._lbl_supplier_count.configure(text="")

    # -----------------------------------------------------------------------
    # AZIONI
    # -----------------------------------------------------------------------

    def _on_rename(self):
        """Rinomina la categoria selezionata."""
        old_name = self._get_selected_category()
        new_name = self.var_new_name.get().strip()

        if not old_name:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Seleziona una categoria dalla lista."), "warning"
            )
            return
        if not new_name:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Il nuovo nome non può essere vuoto."), "warning"
            )
            return
        if old_name == new_name:
            return  # no-op silenzioso

        try:
            with DatabaseManager(get_db_path()) as db:
                rename_supplier_category(db, old_name, new_name)
        except CategoryError as e:
            SimpleMessageDialog(self, _("Operazione non consentita"), str(e), "error")
            return
        except DatabaseError as e:
            SimpleMessageDialog(self, _("Errore Database"), str(e), "error")
            return

        self.changes_made = True
        SimpleMessageDialog(
            self, _("Successo"),
            _("Categoria rinominata correttamente."), "info"
        )
        self._refresh_list(keep_selection=new_name)

    def _on_merge(self):
        """Unisce la categoria selezionata verso il target scelto nella combo."""
        source = self._get_selected_category()
        target = self.var_merge_target.get().strip()

        if not source:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Seleziona una categoria sorgente dalla lista."), "warning"
            )
            return
        if not target:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Seleziona la categoria destinazione dal menu."), "warning"
            )
            return
        if source == target:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Sorgente e destinazione devono essere diverse."), "warning"
            )
            return

        # Conferma
        dlg = SimpleYesNoDialog(
            self,
            _("Conferma Unione"),
            _("Tutti i fornitori con categoria '{src}' verranno spostati a '{tgt}'.\n"
              "La categoria '{src}' verrà eliminata.\n\nProcedere?").format(
                src=source, tgt=target
            ),
        )
        if not dlg.result:
            return

        try:
            with DatabaseManager(get_db_path()) as db:
                merge_supplier_categories(db, source, target)
        except CategoryError as e:
            SimpleMessageDialog(self, _("Operazione non consentita"), str(e), "error")
            return
        except DatabaseError as e:
            SimpleMessageDialog(self, _("Errore Database"), str(e), "error")
            return

        self.changes_made = True
        SimpleMessageDialog(
            self, _("Successo"),
            _("Unione completata correttamente."), "info"
        )
        self._refresh_list(keep_selection=target)

    def _on_delete(self):
        """Elimina la categoria selezionata solo se non usata."""
        name = self._get_selected_category()
        if not name:
            SimpleMessageDialog(
                self, _("Attenzione"),
                _("Seleziona una categoria dalla lista."), "warning"
            )
            return

        try:
            with DatabaseManager(get_db_path()) as db:
                count = delete_supplier_category_if_unused(db, name)
        except CategoryError as e:
            SimpleMessageDialog(self, _("Operazione non consentita"), str(e), "error")
            return
        except DatabaseError as e:
            SimpleMessageDialog(self, _("Errore Database"), str(e), "error")
            return

        if count > 0:
            SimpleMessageDialog(
                self,
                _("Operazione non consentita"),
                _("Impossibile eliminare: la categoria è ancora assegnata a uno o più fornitori."),
                "error",
            )
            return

        self.changes_made = True
        SimpleMessageDialog(
            self, _("Successo"),
            _("Categoria eliminata correttamente."), "info"
        )
        self._refresh_list()
