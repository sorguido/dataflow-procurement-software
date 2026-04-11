"""Dialog export RFQ PDF con gestione logo persistente."""
import os
import tkinter as tk
from tkinter import filedialog, ttk

from services.rfq_pdf_export_service import export_rfq_pdf
from services.rfq_pdf_logo_service import (
    LogoValidationError,
    get_logo_status,
    remove_persisted_logo,
    save_logo_from_source,
)
from ui.dialogs.common_dialogs import show_confirm, show_error, show_info, show_warning
from utils.i18n_utils import tr
from utils.resource_utils import set_window_icon
from utils.window_utils import center_window


class RfqPdfExportDialog(tk.Toplevel):
    """Dialog modale per configurazione logo ed export RFQ PDF."""

    def __init__(self, parent, request_id: int, db_path: str, read_only: bool = False):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)

        self.parent = parent
        self.request_id = request_id
        self.db_path = db_path
        self.read_only = read_only
        self.result_path = None

        self.title(tr("Esporta RFQ PDF"))
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()

        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text=tr("Esporta la RFQ in PDF A4."),
            font=(None, 10),
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        logo_frame = ttk.LabelFrame(main, text=tr("Logo Aziendale"), padding=10)
        logo_frame.pack(fill="x", pady=(0, 12))

        self.logo_status_var = tk.StringVar(value="")
        self.logo_note_var = tk.StringVar(value="")

        ttk.Label(
            logo_frame,
            textvariable=self.logo_status_var,
            font=(None, 10),
            justify="left",
            wraplength=460,
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            logo_frame,
            textvariable=self.logo_note_var,
            font=(None, 9),
            justify="left",
            wraplength=460,
            foreground="#555555",
        ).pack(anchor="w", pady=(0, 10))

        logo_btn_row = ttk.Frame(logo_frame)
        logo_btn_row.pack(fill="x")

        self.btn_choose_logo = ttk.Button(
            logo_btn_row,
            text=tr("Seleziona logo"),
            command=self._choose_logo,
            width=18,
        )
        self.btn_choose_logo.pack(side="left")

        self.btn_remove_logo = ttk.Button(
            logo_btn_row,
            text=tr("Rimuovi logo"),
            command=self._remove_logo,
            width=18,
        )
        self.btn_remove_logo.pack(side="left", padx=(8, 0))

        action_row = ttk.Frame(main)
        action_row.pack(fill="x", pady=(4, 0))

        ttk.Button(
            action_row,
            text=tr("❌ Annulla"),
            command=self._on_cancel,
            width=14,
        ).pack(side="right")

        ttk.Button(
            action_row,
            text=tr("Conferma Export PDF"),
            command=self._on_export,
            width=20,
        ).pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._refresh_logo_status()

        center_window(self)
        self.deiconify()
        self.wait_visibility()

    def _refresh_logo_status(self):
        status = get_logo_status()
        self.logo_note_var.set(
            tr("Formati supportati: PNG/JPG. Consigliata immagine orizzontale, max 8 MB. Il logo verra ridimensionato automaticamente nel PDF senza deformazioni.")
        )

        if not status["configured"]:
            self.logo_status_var.set(tr("Logo non configurato."))
            self.btn_choose_logo.config(text=tr("Seleziona logo"))
            self.btn_remove_logo.config(state="disabled")
            return

        if status["available"]:
            self.logo_status_var.set(
                tr("Logo configurato: {}").format(status["filename"])
            )
            self.btn_choose_logo.config(text=tr("Sostituisci logo"))
            self.btn_remove_logo.config(state="normal")
            return

        self.logo_status_var.set(
            tr("Logo configurato ma non disponibile/corrotto. Puoi sostituirlo o rimuoverlo.")
        )
        self.btn_choose_logo.config(text=tr("Sostituisci logo"))
        self.btn_remove_logo.config(state="normal")

    def _choose_logo(self):
        self.lift()
        self.focus_force()
        self.update_idletasks()

        selected_path = filedialog.askopenfilename(
            parent=self,
            title=tr("Seleziona Logo Aziendale"),
            filetypes=[
                (tr("Immagini PNG/JPG"), "*.png *.jpg *.jpeg"),
                (tr("Tutti i file"), "*.*"),
            ],
        )
        if not selected_path:
            return

        try:
            save_logo_from_source(selected_path)
            self._refresh_logo_status()
            show_info(self, tr("Successo"), tr("Logo aziendale salvato con successo."))
        except LogoValidationError as exc:
            show_warning(self, tr("Logo non valido"), tr("Impossibile usare il logo selezionato: {}").format(exc))
        except Exception as exc:
            show_error(self, tr("Errore"), tr("Errore durante il salvataggio del logo: {}").format(exc))

    def _remove_logo(self):
        if not show_confirm(self, tr("Conferma"), tr("Vuoi rimuovere il logo aziendale salvato?")):
            return

        try:
            remove_persisted_logo()
            self._refresh_logo_status()
            show_info(self, tr("Successo"), tr("Logo rimosso."))
        except Exception as exc:
            show_error(self, tr("Errore"), tr("Impossibile rimuovere il logo: {}").format(exc))

    def _on_export(self):
        self.lift()
        self.focus_force()
        self.update_idletasks()

        initial_name = f"RFQ_{self.request_id}.pdf"
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title=tr("Salva RFQ PDF"),
            defaultextension=".pdf",
            initialfile=initial_name,
            filetypes=[(tr("File PDF"), "*.pdf")],
            confirmoverwrite=False,
        )
        if not save_path:
            return

        if os.path.exists(save_path):
            if not show_confirm(
                self,
                tr("Conferma"),
                tr("Il file selezionato esiste gia. Vuoi sovrascriverlo?"),
            ):
                return

        logo_path = None
        logo_status = get_logo_status()
        if logo_status.get("available"):
            logo_path = logo_status.get("path")

        try:
            export_result = export_rfq_pdf(
                db_path=self.db_path,
                request_id=self.request_id,
                output_path=save_path,
                logo_path=logo_path,
                read_only=self.read_only,
            )
        except Exception as exc:
            show_error(self, tr("Errore Esportazione"), tr("Errore durante la generazione PDF: {}").format(exc))
            return

        warnings = export_result.get("warnings", [])
        if warnings:
            warning_text = "\n".join(warnings)
            show_warning(
                self,
                tr("Export completato con avvisi"),
                tr("PDF generato in:\n{}\n\nAvvisi:\n{}").format(save_path, warning_text),
            )
        else:
            show_info(self, tr("Successo"), tr("PDF generato in:\n{}").format(save_path))

        self.result_path = save_path
        self.destroy()

    def _on_cancel(self):
        self.result_path = None
        self.destroy()
