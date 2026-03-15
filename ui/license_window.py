"""
Finestra gestione licenza DataFlow.
"""
import tkinter as tk
from tkinter import ttk
import webbrowser

from utils.resource_utils import set_window_icon
from utils.window_utils import center_window

def _():
    """Placeholder per traduzioni."""
    import builtins
    return builtins._ if hasattr(builtins, '_') else lambda x: x


class LicenseWindow(tk.Toplevel):
    def __init__(self, parent, first_run=False):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        self.title(_()("Licenza d'Uso - DataFlow Procurement Software"))
        self.transient(parent)
        self.grab_set()
        
        # Frame pulsanti (sempre in fondo)
        button_frame = ttk.Frame(self)
        button_frame.pack(side="bottom", fill="x", padx=10, pady=10)

        if first_run:
            self.accepted = False
            ttk.Button(button_frame, text=_()("❌ Esci"), command=self.on_exit).pack(side="right")
            ttk.Button(button_frame, text=_()("✅ Accetto"), command=self.on_accept).pack(side="right", padx=10)
            self.protocol("WM_DELETE_WINDOW", self.on_exit) 
        else:
            ttk.Button(button_frame, text=_()("❌ Chiudi"), command=self.destroy).pack(side="right")
        
        # Frame contenuto (espandibile)
        main_frame = ttk.Frame(self)
        main_frame.pack(side="top", fill="both", expand=True)

        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(content_frame)
        self.text_content = tk.Text(content_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, 
                                    padx=15, pady=10, relief="flat", background="#FFFFFF", font=("Arial", 10))
        scrollbar.config(command=self.text_content.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self._populate_content()
        self.text_content.config(state="disabled")
        
        center_window(self)

    def on_accept(self):
        self.accepted = True
        self.destroy()

    def on_exit(self):
        self.accepted = False
        self.destroy()

    def _populate_content(self):
        # Configurazione degli stili di testo
        self.text_content.tag_configure("h1", font=("Arial", 14, "bold", "underline"), justify="center")
        self.text_content.tag_configure("h2", font=("Arial", 11, "bold"))
        self.text_content.tag_configure("normal", font=("Arial", 10))
        self.text_content.tag_configure("code", font=("Arial", 9))
        
        # Configurazione tag per link cliccabile
        self.text_content.tag_configure("link", foreground="blue", underline=True)
        self.text_content.tag_bind("link", "<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/guido-soraru-buyer/"))
        self.text_content.tag_bind("link", "<Enter>", lambda e: self.text_content.config(cursor="hand2"))
        self.text_content.tag_bind("link", "<Leave>", lambda e: self.text_content.config(cursor=""))
        
        def add(txt, tag_keys):
            tag_tuple = tag_keys if isinstance(tag_keys, tuple) else (tag_keys,)
            self.text_content.insert(tk.END, txt, tag_tuple)

        # Contenuto licenza
        add(_()("Contratto di Licenza per l'Utente Finale (GNU GPLv3) - DataFlow Procurement Software\n\n"), "h1")
        
        add(_()("Sviluppatore: "), "h2"); add("Guido Sorarù", ("normal", "link")); add("\n", "normal")
        add(_()("E-mail: "), "h2"); add("sorguido@gmail.com\n", "normal")
        add(_()("Copyright © 2025 Guido Sorarù.\n\n"), "h2")
        
        add("--------------------------------------------------\n\n", "normal")
        
        add(_()("Questo software, \"DataFlow\" (di seguito \"il Software\"), è rilasciato come software open source sotto la licenza GNU General Public License versione 3 (GPLv3).\n\n"), "normal")
        
        add(_()("1. CONCESSIONE DELLA LICENZA\n"), "h2")
        add(_()("Lo sviluppatore concede all'utente una licenza non esclusiva per scaricare, installare, utilizzare, studiare, modificare e ridistribuire il Software in conformità con i termini della GNU General Public License versione 3.\n\n"), "normal")
        add(_()("Il codice sorgente completo del Software è disponibile pubblicamente.\n\n"), "normal")
        add(_()("Una copia della licenza GNU GPLv3 dovrebbe essere distribuita insieme a questo Software.\nIn caso contrario consultare: https://www.gnu.org/licenses/\n\n"), "normal")
        
        add(_()("2. DISTRIBUZIONE E MODIFICA\n"), "h2")
        add(_()("Il Software può essere utilizzato, studiato, modificato e ridistribuito liberamente secondo i termini della GNU General Public License versione 3.\n\n"), "normal")
        add(_()("Qualsiasi ridistribuzione del Software, modificato o non modificato, deve mantenere l'avviso di copyright ed essere distribuita sotto la stessa licenza GNU GPLv3.\n\n"), "normal")
        
        add(_()("3. ESCLUSIONE DI GARANZIA\n"), "h2")
        add(_()("IL SOFTWARE È FORNITO \"COSÌ COM'È\" (AS IS), SENZA ALCUNA GARANZIA, ESPRESSA O IMPLICITA. LO SVILUPPATORE NON FORNISCE ALCUNA GARANZIA RIGUARDO LA COMMERCIABILITÀ, L'IDONEITÀ PER UNO SCOPO PARTICOLARE O LA NON VIOLAZIONE DI DIRITTI DI TERZI.\n"), "normal")
        add(_()("L'INTERO RISCHIO DERIVANTE DALL'USO O DALLE PRESTAZIONI DEL SOFTWARE RIMANE A CARICO DELL'UTENTE.\n\n"), "normal")
        
        add(_()("4. LIMITAZIONE DI RESPONSABILITÀ\n"), "h2")
        add(_()("IN NESSUN CASO LO SVILUPPATORE (GUIDO SORARÙ) POTRÀ ESSERE RITENUTO RESPONSABILE PER QUALSIASI DANNO DIRETTO, INDIRETTO, INCIDENTALE, SPECIALE, ESEMPLARE O CONSEQUENZIALE (INCLUSI, A TITOLO ESEMPLIFICATIVO MA NON ESAUSIVO, DANNI PER PERDITA DI DATI, PERDITA DI PROFITTI O INTERRUZIONE DELL'ATTIVITÀ) DERIVANTE DALL'USO, DALL'USO IMPROPRIO O DALL'IMPOSSIBILITÀ DI UTILIZZARE IL SOFTWARE, ANCHE SE LO SVILUPPATORE È STATO AVVISATO DELLA POSSIBILITÀ DI TALI DANNI.\n\n"), "normal")
        
        add(_()("Il Software utilizza un database SQLite con modalità WAL per ogni utente. DataFlow 2.0.0 supporta l'utilizzo multi-utente con database separati per ciascun utente, permettendo la condivisione sicura dei dati in sola lettura.\n"), "normal")
        add(_()("L'utente si assume la piena responsabilità per la perdita o corruzione dei dati derivante dall'uso improprio del software.\n"), "normal")
        add(_()("L'accesso simultaneo in scrittura da parte di più utenti allo stesso file di database non è supportato e causerà con alta probabilità la corruzione irreversibile dei dati. Tuttavia, l'architettura multi-utente di DataFlow garantisce che ogni utente abbia il proprio database separato, eliminando questo rischio.\n\n"), "normal")
        
        add(_()("Utilizzando questo Software, l'utente accetta i termini e le condizioni di questa licenza.\n"), "normal")
        
        self.text_content.config(state="disabled")
