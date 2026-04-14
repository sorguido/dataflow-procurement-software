"""
ManageSupplierCategoriesDialog — Dialog per la gestione centralizzata
delle categorie dei fornitori potenziali.

Operazioni supportate (in-memory fino al click Salva):
  - Rinomina categoria
  - Unisci categoria sorgente → destinazione
  - Elimina categoria se non usata

Utilizzo:
    dlg = ManageSupplierCategoriesDialog(parent, refresh_derisking_cb=cb)
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
    CategoryError,
)
from utils.i18n_utils import tr
from utils.resource_utils import set_window_icon
from utils.window_utils import center_window
from ui.dialogs.common_dialogs import SimpleMessageDialog, SimpleYesNoDialog

logger = logging.getLogger(__name__)


class ManageSupplierCategoriesDialog(tk.Toplevel):
    """
    Dialog per rinominare, unire ed eliminare le categorie dei fornitori potenziali.

    Tutte le operazioni avvengono in memoria finché l'utente non clicca Salva.
    Annulla o chiusura finestra scartano ogni modifica.

    self.changes_made è impostato a True se il salvataggio ha avuto successo.
    """

    def __init__(self, parent, refresh_derisking_cb=None):
        super().__init__(parent)
        self.changes_made = False
        self._refresh_derisking_cb = refresh_derisking_cb
        self._original_categories: list = []
        self._working_categories: list = []
        self._pending_ops: list = []

        self.withdraw()
        set_window_icon(self)
        self.title(tr("Manage Categories"))
        self.transient(parent)
        self.resizable(False, False)

        self._load_initial_state()
        self._build_ui()
        self._refresh_list_from_memory()

        center_window(self)
        self.wait_visibility()
        self.grab_set()
        self.deiconify()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    # -----------------------------------------------------------------------
    # INITIALISATION
    # -----------------------------------------------------------------------

    def _load_initial_state(self):
        """Carica le categorie dal DB nello stato iniziale in memoria."""
        try:
            with DatabaseManager(get_db_path()) as db:
                categories = get_all_supplier_categories(db)
        except (DatabaseError, CategoryError) as e:
            logger.error("Errore caricamento categorie: %s", e)
            categories = []
        self._original_categories = list(categories)
        self._working_categories = list(categories)

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
        list_frame = ttk.LabelFrame(top_frame, text=tr("Categories"), padding="8")
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
        rename_frame = ttk.LabelFrame(actions_frame, text=tr("Rename"), padding="8")
        rename_frame.pack(fill="x", pady=(0, 8))
        rename_frame.columnconfigure(1, weight=1)

        ttk.Label(rename_frame, text=tr("New name:"), font=(None, 10)).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.var_new_name = tk.StringVar()
        self._entry_new_name = ttk.Entry(rename_frame, textvariable=self.var_new_name, width=22)
        self._entry_new_name.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(
            rename_frame,
            text=tr("Rename"),
            command=self._on_rename,
            width=12,
        ).grid(row=1, column=0, columnspan=2, sticky="e", pady=(4, 0))

        # == Unisci ==
        merge_frame = ttk.LabelFrame(actions_frame, text=tr("Merge"), padding="8")
        merge_frame.pack(fill="x", pady=(0, 8))
        merge_frame.columnconfigure(1, weight=1)

        ttk.Label(merge_frame, text=tr("Merge with:"), font=(None, 10)).grid(
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
            text=tr("Merge"),
            command=self._on_merge,
            width=12,
        ).grid(row=1, column=0, columnspan=2, sticky="e", pady=(4, 0))

        # == Elimina se non usata ==
        delete_frame = ttk.LabelFrame(
            actions_frame, text=tr("Delete if unused"), padding="8"
        )
        delete_frame.pack(fill="x")

        self._lbl_supplier_count = ttk.Label(
            delete_frame, text="", font=(None, 10), foreground="#555555"
        )
        self._lbl_supplier_count.pack(anchor="w", pady=(0, 4))

        ttk.Button(
            delete_frame,
            text=tr("Delete if unused"),
            command=self._on_delete,
            width=22,
        ).pack(anchor="e", pady=(4, 0))

        # ------------------------------------------------------------------ #
        # Pulsanti Annulla / Salva (in basso a destra)
        # ------------------------------------------------------------------ #
        btn_frame = ttk.Frame(outer)
        btn_frame.pack(fill="x")

        ttk.Button(
            btn_frame,
            text=tr("💾 Save"),
            command=self._on_save,
            width=12,
        ).pack(side="right")

        ttk.Button(
            btn_frame,
            text=tr("❌ Cancel"),
            command=self._on_cancel,
            width=12,
        ).pack(side="right", padx=(5, 0))

    # -----------------------------------------------------------------------
    # REFRESH (in-memory, no DB)
    # -----------------------------------------------------------------------

    def _refresh_list_from_memory(self, keep_selection: str = None):
        """
        Ripopola Listbox e combo merge dalla lista in-memory _working_categories.
        Nessuna query al DB.

        Args:
            keep_selection: se fornito, tenta di riselezionare questo valore.
        """
        categories = self._working_categories

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
        if name not in self._original_categories:
            self._lbl_supplier_count.configure(
                text=tr("Suppliers: — (verified on save)")
            )
            return
        try:
            with DatabaseManager(get_db_path()) as db:
                count = db.count_suppliers_by_category(name)
            self._lbl_supplier_count.configure(
                text=tr("Associated suppliers: {}").format(count)
            )
        except Exception:
            self._lbl_supplier_count.configure(text="")

    # -----------------------------------------------------------------------
    # PENDING OPS HELPERS
    # -----------------------------------------------------------------------

    def _apply_rename_in_ops(self, old_name: str, new_name: str):
        """
        Aggiunge/consolida un rename nelle pending ops e normalizza tutti
        i riferimenti ad old_name nelle ops esistenti.
        """
        # 1. Consolida: se esiste già un rename che produce old_name → aggiorna new
        found = False
        for op in self._pending_ops:
            if op["type"] == "rename" and op["new"] == old_name:
                op["new"] = new_name
                found = True
                break
        if not found:
            self._pending_ops.append({"type": "rename", "old": old_name, "new": new_name})

        # 2. Aggiorna tutti gli altri riferimenti a old_name nelle ops successive
        for op in self._pending_ops:
            if op["type"] == "merge":
                if op["source"] == old_name:
                    op["source"] = new_name
                if op["target"] == old_name:
                    op["target"] = new_name
            elif op["type"] == "delete_unused":
                if op["name"] == old_name:
                    op["name"] = new_name

        # 3. Rimuovi rename no-op (old == new, es. dopo catena circolare A→B→A)
        self._pending_ops = [
            op for op in self._pending_ops
            if not (op["type"] == "rename" and op["old"] == op["new"])
        ]

        # 4. Pulizia ops stale
        self._prune_stale_ops()

    def _prune_stale_ops(self):
        """
        Rimuove delete_unused ops i cui nomi non sono più presenti in
        _working_categories né come target di un rename pendente.
        """
        valid_names = set(self._working_categories)
        for op in self._pending_ops:
            if op["type"] == "rename":
                valid_names.add(op["new"])
        self._pending_ops = [
            op for op in self._pending_ops
            if not (
                op["type"] == "delete_unused"
                and op["name"] not in valid_names
            )
        ]

    # -----------------------------------------------------------------------
    # AZIONI (in-memory)
    # -----------------------------------------------------------------------

    def _on_rename(self):
        """Rinomina la categoria selezionata (in memoria)."""
        old_name = self._get_selected_category()
        new_name = self.var_new_name.get().strip()

        if not old_name:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("Select a category from the list."), "warning"
            )
            return
        if not new_name:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("The new name cannot be empty."), "warning"
            )
            return
        if old_name == new_name:
            return  # no-op silenzioso

        # Blocco early: rename != merge
        if new_name in self._working_categories:
            SimpleMessageDialog(
                self, tr("Operation not allowed"),
                tr("The category already exists. Use the Merge function."), "error"
            )
            return

        # Aggiorna lista in memoria
        idx = self._working_categories.index(old_name)
        self._working_categories[idx] = new_name
        self._working_categories = sorted(self._working_categories, key=lambda s: s.lower())

        self._apply_rename_in_ops(old_name, new_name)
        self._refresh_list_from_memory(keep_selection=new_name)

    def _on_merge(self):
        """Unisce la categoria selezionata verso il target scelto (in memoria)."""
        source = self._get_selected_category()
        target = self.var_merge_target.get().strip()

        if not source:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("Select a source category from the list."), "warning"
            )
            return
        if not target:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("Select the destination category from the menu."), "warning"
            )
            return
        if source == target:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("Source and destination must be different."), "warning"
            )
            return

        dlg = SimpleYesNoDialog(
            self,
            tr("Confirm Merge"),
            tr("All suppliers with category '{src}' will be moved to '{tgt}'.\nThe category '{src}' will be deleted.\n\nProceed?").format(
                src=source, tgt=target
            ),
        )
        if not dlg.result:
            return

        self._working_categories.remove(source)
        self._pending_ops.append({"type": "merge", "source": source, "target": target})
        self._prune_stale_ops()
        self._refresh_list_from_memory(keep_selection=target)

    def _on_delete(self):
        """Elimina la categoria selezionata (in memoria, verificata al salvataggio)."""
        name = self._get_selected_category()
        if not name:
            SimpleMessageDialog(
                self, tr("Warning"),
                tr("Select a category from the list."), "warning"
            )
            return

        self._working_categories.remove(name)
        self._pending_ops.append({"type": "delete_unused", "name": name})
        self._refresh_list_from_memory()

    # -----------------------------------------------------------------------
    # SALVA / ANNULLA
    # -----------------------------------------------------------------------

    def _on_save(self):
        """Applica tutte le operazioni pendenti al DB in un'unica transazione."""
        if not self._pending_ops:
            self.destroy()
            return

        try:
            with DatabaseManager(get_db_path()) as db:
                db.apply_category_ops_atomic(self._pending_ops)
        except DatabaseError as e:
            SimpleMessageDialog(self, tr("Database Error"), str(e), "error")
            return  # non chiudere: l'utente può correggere o annullare

        self.changes_made = True

        if self._refresh_derisking_cb:
            try:
                self._refresh_derisking_cb()
            except Exception as e:
                logger.warning("Errore refresh_derisking_cb: %s", e)

        self.destroy()

    def _on_cancel(self):
        """Chiude il dialog senza scrivere nulla nel DB."""
        self.destroy()
