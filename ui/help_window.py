# -*- coding: utf-8 -*-
"""
HelpWindow - Finestra Guida Utente per DataFlow Procurement Software
Fornisce documentazione interattiva con sommario navigabile, ricerca full-text e formattazione avanzata.

ESTRATTO DAL MONOLITE: Versione originale funzionante prima del refactoring.
"""

import tkinter as tk
from tkinter import ttk
import webbrowser
import configparser
import os
import builtins

# Import utility functions dal progetto
from utils.window_utils import center_window
from utils.user_utils import get_config_file
from utils.resource_utils import resource_path, set_window_icon

# IMPORTANTE: La funzione _() è installata in builtins da init_i18n() nel main
# Se non esiste (ad esempio nei test), usa una funzione dummy
if not hasattr(builtins, '_'):
    builtins._ = lambda x: x


class HelpWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.withdraw()
        set_window_icon(self)
        
        # Leggi la lingua corrente per la traduzione del sommario
        config = configparser.ConfigParser(interpolation=None)
        config_file = get_config_file()
        current_language = "en"  # default
        if os.path.exists(config_file):
            config.read(config_file, encoding='utf-8')
            if 'Settings' in config and config.has_option('Settings', 'language'):
                current_language = config.get('Settings', 'language', fallback='en')
        if current_language not in ['en', 'it']:
            current_language = "en"
        
        # Determina la traduzione corretta per "Analisi SQDC"
        sqdc_text = "   - Analysis SQDC" if current_language == "en" else _("   - Analisi SQDC")
        # Determina la traduzione corretta per "File Paths"
        file_paths_text = "   - File Paths and Environment" if current_language == "en" else _("   - File Paths e Ambiente")

        self.title(_("Guida Utente - DataFlow Procurement Software"))
        self.transient(parent)
        self.grab_set()
        
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True)
        
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame sommario (TOC) con Canvas scrollabile
        toc_outer_frame = ttk.Frame(paned, padding=10)
        paned.add(toc_outer_frame, weight=1)

        ttk.Label(toc_outer_frame, text=_("Sommario"), font=("Helvetica", 12, "bold")).pack(anchor="w", pady=(0, 5))

        toc_canvas = tk.Canvas(toc_outer_frame, highlightthickness=0)
        toc_vscrollbar = ttk.Scrollbar(toc_outer_frame, orient="vertical", command=toc_canvas.yview)
        toc_canvas.configure(yscrollcommand=toc_vscrollbar.set)
        toc_vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        toc_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        toc_frame = ttk.Frame(toc_canvas)
        _toc_canvas_window = toc_canvas.create_window((0, 0), window=toc_frame, anchor="nw")

        def _toc_frame_configure(event):
            toc_canvas.configure(scrollregion=toc_canvas.bbox("all"))

        def _toc_canvas_configure(event):
            toc_canvas.itemconfig(_toc_canvas_window, width=event.width)

        toc_frame.bind("<Configure>", _toc_frame_configure)
        toc_canvas.bind("<Configure>", _toc_canvas_configure)

        def _toc_mousewheel(event):
            toc_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _toc_scroll_up(event):
            toc_canvas.yview_scroll(-1, "units")

        def _toc_scroll_down(event):
            toc_canvas.yview_scroll(1, "units")

        for _w in (toc_canvas, toc_frame):
            _w.bind("<MouseWheel>", _toc_mousewheel)
            _w.bind("<Button-4>", _toc_scroll_up)
            _w.bind("<Button-5>", _toc_scroll_down)

        
        self.topics = [
            (_("0. Primi Passi"), "quick_start"),
            (_("   - Benvenuto in DataFlow"), "welcome"),
            (_("   - Primo Avvio"), "first_run"),
            (_("   - Prima RdO di Prova"), "first_rdo"),
            (_("1. Schermata Principale"), "main_screen"),
            (_("   - Interfaccia e Pulsanti Principali"), "main_interface"),
            (_("   - Scorciatoie da Tastiera"), "keyboard_shortcuts"),
            (_("   - Ordinamento delle Colonne"), "column_sorting"),
            (_("   - Filtri di Ricerca"), "main_filters"),
            (_("   - Global Search Bar"), "global_search_bar"),
            (_("   - Search vs Filters"), "search_vs_filters"),
            (_("2. Creare una Nuova RdO"), "new_rdo"),
            (_("   - Inserimento Manuale degli Articoli"), "new_rdo_data"),
            (_("   - Importazione da Excel"), "new_rdo_excel"),
            (_("   - Data Validation Rules"), "data_validation"),
            (_("3. Gestire una RdO Esistente"), "manage_rdo"),
            (_("   - Quick Actions (New / Delete / Duplicate)"), "quick_actions"),
            (_("   - La Griglia Prezzi"), "manage_grid"),
            (_("   - Modifica Dati e Aggiunta Note"), "manage_edit"),
            (_("   - Gestione Numeri Ordine (PO)"), "manage_po"),
            (_("   - Gestione Allegati"), "manage_attachments"),
            (sqdc_text, "manage_sqdc"),
            (_("   - Esportazione Excel"), "manage_export"),
            (_("   - Value Stream Mapping"), "vsm_overview"),
            (_("4. Impostazioni e Manutenzione"), "settings"),
            (_("   - Gestione Database"), "settings_db"),
            (_("   - Backup"), "settings_backup"),
            (_("   - Avanzate"), "settings_advanced"),
            (file_paths_text, "file_paths"),
            (_("5. Problemi Comuni e Soluzioni"), "troubleshooting"),
            (_("   - Database Bloccato"), "ts_db_locked"),
            (_("   - Errori Importazione Excel"), "ts_import"),
            (_("   - Allegati Non Trovati"), "ts_attachments"),
            (_("   - Recupero da Backup"), "ts_backup"),
            (_("6. Requisiti di Sistema e Limiti"), "requirements"),
            (_("7. Glossario"), "glossary"),
            (_("8. Contatti e Supporto"), "support")
        ]
        
        # Creiamo una mappa di ricerca veloce (Testo del titolo -> tag_ancoraggio)
        self.topic_anchor_map = {}
        for text, tag in self.topics:
            clean_text = text.strip()
            # Rimuovi il prefisso "   - " per le sottovoci
            if clean_text.startswith("- "):
                clean_text = clean_text[2:]
            self.topic_anchor_map[clean_text] = tag
        
        # Aggiungi anche le chiavi alternative per "Analisi SQDC" quando la lingua è inglese
        if current_language == "en":
            self.topic_anchor_map["Analisi SQDC"] = "manage_sqdc"
            self.topic_anchor_map["SQDC Analysis"] = "manage_sqdc"
            self.topic_anchor_map["Purchase Order Number Management"] = "manage_po"
            self.topic_anchor_map["3. Managing an Existing RFQ"] = "manage_rdo"
            self.topic_anchor_map["First Test RFQ"] = "first_rdo"
            self.topic_anchor_map["2. Creating a New RFQ"] = "new_rdo"
        
        # Aggiungi mappatura per titolo completo della sezione PO (italiano)
        self.topic_anchor_map["Gestione Numeri Ordine di Acquisto (PO)"] = "manage_po"
        
        # Crea i link del sommario
        for text, tag in self.topics:
            link = ttk.Label(toc_frame, text=text, foreground="blue", cursor="hand2")
            if text.strip().startswith(('0','1','2','3','4','5','6','7','8')):
                link.pack(anchor="w", pady=2)
            else:
                link.pack(anchor="w", pady=1)
            link.bind("<Button-1>", lambda e, t=tag: self.text_content.see(f"{t}.first"))
            link.bind("<MouseWheel>", _toc_mousewheel)
            link.bind("<Button-4>", _toc_scroll_up)
            link.bind("<Button-5>", _toc_scroll_down)
        
        # Frame contenuto
        content_frame = ttk.Frame(paned)
        paned.add(content_frame, weight=4)
        
        # --- SEARCH BAR ---
        # Inizializza variabili per la ricerca
        self.search_var = tk.StringVar()
        self.search_matches = []  # Lista di posizioni dei match trovati
        self.current_match_index = -1  # Indice del match corrente
        self.search_result_label = None  # Label per mostrare "Risultato X di Y"
        
        # Frame per la barra di ricerca (sopra il widget Text)
        search_frame = ttk.Frame(content_frame)
        search_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        self.setup_search_functionality(search_frame)
        # --- FINE SEARCH BAR ---
        
        # Widget Text per il contenuto
        scrollbar = ttk.Scrollbar(content_frame)
        self.text_content = tk.Text(
            content_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            padx=15,
            pady=10,
            relief="flat",
            background="#FFFFFF",
            font=("Arial", 10)
        )
        
        scrollbar.config(command=self.text_content.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.text_content.insert(tk.END, _("Caricamento della guida in corso...\n\n"), "normal")
        self.text_content.config(state="disabled")
        
        center_window(self)
        # Carica il contenuto in modo asincrono
        self.after(100, self.populate_content)

    def populate_content(self):
        try:
            # Riabilita il widget per inserire il contenuto
            self.text_content.config(state="normal")
            # Pulisci il messaggio di caricamento
            self.text_content.delete("1.0", tk.END)
            
            # Configurazione degli stili di testo
            tags = {
                "h1": ("Arial", 14, "bold", "underline"),
                "h2": ("Arial", 12, "bold"),
                "h3": ("Arial", 10, "bold", "underline"),
                "bold": ("Arial", 10, "bold"),
                "normal": ("Arial", 10),
                "list": ("Arial", 10),
                "warning": ("Arial", 10, "italic"),
                "code": ("Arial", 9),
                "info": ("Arial", 9, "italic"),
                "warning_red": ("Arial", 10, "bold")
            }
            for t, f_tuple in tags.items():
                self.text_content.tag_configure(t, font=f_tuple)
            
            # Configurazione colori
            self.text_content.tag_configure("warning_red", foreground="red")
            self.text_content.tag_configure("info", foreground="gray")
            
            # Configurazione tag di ancoraggio
            for _, tag in self.topics:
                self.text_content.tag_configure(tag, underline=False)
            
            # Configurazione tag per ricerca
            self.text_content.tag_configure("search_highlight", background="yellow", foreground="black")
            
            # Carica il contenuto dal file esterno in base alla lingua
            config = configparser.ConfigParser(interpolation=None)
            config_file = get_config_file()
            current_language = "en"  # default
            if os.path.exists(config_file):
                config.read(config_file, encoding='utf-8')
                if 'Settings' in config and config.has_option('Settings', 'language'):
                    current_language = config.get('Settings', 'language', fallback='en')
            
            # Validazione: accetta solo 'en' o 'it'
            if current_language not in ['en', 'it']:
                current_language = "en"
            
            # Carica il file guida corretto in base alla lingua
            if current_language == "en":
                guida_path = resource_path(os.path.join("add_data", "guida_en.txt"))
            else:
                guida_path = resource_path(os.path.join("add_data", "guida.txt"))
            
            # Debug: log quale file viene caricato
            print(f"[HelpWindow] Lingua corrente: {current_language}")
            print(f"[HelpWindow] Tentativo di caricare: {guida_path}")
            print(f"[HelpWindow] File esiste: {os.path.exists(guida_path)}")
            
            if os.path.exists(guida_path):
                with open(guida_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                print(f"[HelpWindow] File caricato con successo, lunghezza contenuto: {len(content)} caratteri")
                self._parse_and_insert_content(content)
            else:
                # Fallback se il file non esiste
                error_msg = _("File guida non trovato. Contatta l'amministratore.") + f"\n\nPath cercato: {guida_path}"
                self.text_content.insert(tk.END, error_msg, "normal")
                print(f"[HelpWindow] ERRORE: File non trovato: {guida_path}")
            
            # Disabilita il widget dopo il caricamento
            self.text_content.config(state="disabled")
            
        except Exception as e:
            # In caso di errore, mostra un messaggio di errore
            self.text_content.config(state="normal")
            self.text_content.delete("1.0", tk.END)
            self.text_content.insert(tk.END, _("Errore caricamento guida: {}\n\nContatta l'amministratore.").format(e), "normal")
            self.text_content.config(state="disabled")
    
    def _parse_and_insert_content(self, content):
        """Analizza il contenuto del file e lo inserisce nel widget con la formattazione corretta"""
        lines = content.split('\n')
        
        for line in lines:
            if not line.strip():
                self.text_content.insert(tk.END, "\n")
                continue
                
            # Trova il tag di stile (es. 'h1') e il contenuto (es. '0. Primi Passi')
            style_tag = "normal"
            content_to_insert = line
            is_section_tag = False

            if line.startswith('[H1]'):
                style_tag = "h1"
                content_to_insert = line[4:]
                is_section_tag = True
            elif line.startswith('[H2]'):
                style_tag = "h2"
                content_to_insert = line[4:]
                is_section_tag = True
            elif line.startswith('[H3]'):
                style_tag = "h3"
                content_to_insert = line[4:]
                is_section_tag = True
            elif line.startswith('[INFO]'):
                style_tag = "info"
                content_to_insert = line[6:]
                is_section_tag = True
            elif line.startswith('[LIST]'):
                style_tag = "list"
                content_to_insert = line[6:]
            elif line.startswith('[WARNING_RED]'):
                content_to_insert = line
            elif line.startswith('[BOLD]'):
                style_tag = "bold"
                content_to_insert = line[6:]
            elif line.startswith('[NORMAL]'):
                style_tag = "normal"
                content_to_insert = line[8:]
            elif line.startswith('[CODE]'):
                style_tag = "code"
                content_to_insert = line[6:]
            elif line.startswith('[WARNING]'):
                content_to_insert = line

            # Se è un tag di sezione (H1, H2, H3, INFO) applica la formattazione e cerca il link
            if is_section_tag:
                clean_content = content_to_insert.strip()
                anchor_tag = self.topic_anchor_map.get(clean_content)
                tags_to_apply = [style_tag]
                if anchor_tag:
                    tags_to_apply.append(anchor_tag)
                # Usa il parser inline per gestire eventuali tag inline nel contenuto
                self._insert_formatted_line_with_anchor(content_to_insert, tuple(tags_to_apply))
                continue
            
            # Per tutte le altre righe (inclusi tag come [BOLD], [LIST], etc.), usa il parser inline
            if line != content_to_insert:
                # Se abbiamo estratto del testo dopo un tag, analizzalo con il parser inline
                self._insert_formatted_line(content_to_insert)
            else:
                # Nessun tag speciale, è testo normale con possibili tag inline
                self._insert_formatted_line(line)
    
    def _insert_formatted_line(self, line):
        """Inserisce una riga con formattazione inline"""
        parts = []
        current_text = ""
        current_tag = "normal"
        i = 0
        
        while i < len(line):
            # Cerca tag LINK con URL: [LINK:url]testo[/LINK]
            if line.startswith('[LINK:', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                # Estrae l'URL
                url_start = i + 6  # dopo '[LINK:'
                url_end = line.find(']', url_start)
                if url_end != -1:
                    url = line[url_start:url_end]
                    # Trova il testo del link
                    text_start = url_end + 1
                    text_end = line.find('[/LINK]', text_start)
                    if text_end != -1:
                        link_text = line[text_start:text_end]
                        parts.append((link_text, "hyperlink", url))
                        i = text_end + 7  # dopo '[/LINK]'
                        continue
                current_text += line[i]
                i += 1
            # Cerca tag di apertura/chiusura
            elif line[i:i+6] == '[BOLD]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "bold"
                i += 6
            elif line[i:i+8] == '[/BOLD]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line[i:i+8] == '[NORMAL]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line[i:i+10] == '[/NORMAL]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 10
            elif line[i:i+6] == '[CODE]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "code"
                i += 6
            elif line[i:i+8] == '[/CODE]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line.startswith('[WARNING]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "warning"
                i += len('[WARNING]')
            elif line.startswith('[/WARNING]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += len('[/WARNING]')
            elif line.startswith('[WARNING_RED]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "warning_red"
                i += len('[WARNING_RED]')
            elif line.startswith('[/WARNING_RED]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += len('[/WARNING_RED]')
            else:
                current_text += line[i]
                i += 1
        
        # Aggiunge l'ultimo pezzo
        if current_text:
            parts.append((current_text, current_tag, None))
        elif not parts:
            # Se non c'è nessun testo, aggiungi una stringa vuota
            parts.append(("", "normal", None))
        
        # Inserisce tutti i pezzi
        for item in parts:
            if len(item) == 3:
                text, tag, url = item
                if tag == "hyperlink" and url:
                    # Crea un tag univoco per questo link
                    link_tag = f"link_{id(url)}"
                    # Configura il tag con lo stile del link
                    self.text_content.tag_configure(link_tag, foreground="blue", underline=True)
                    # Bind del click per aprire l'URL
                    self.text_content.tag_bind(link_tag, "<Button-1>", lambda e, u=url: webbrowser.open(u))
                    self.text_content.tag_bind(link_tag, "<Enter>", lambda e: self.text_content.config(cursor="hand2"))
                    self.text_content.tag_bind(link_tag, "<Leave>", lambda e: self.text_content.config(cursor=""))
                    # Inserisce il testo con il tag del link
                    self.text_content.insert(tk.END, text, link_tag)
                else:
                    self.text_content.insert(tk.END, text, tag)
            else:
                # Fallback per compatibilità
                self.text_content.insert(tk.END, item[0], item[1])
        
        self.text_content.insert(tk.END, "\n")
    
    def _insert_formatted_line_with_anchor(self, line, anchor_tags):
        """Inserisce una riga con formattazione inline e tag di ancoraggio"""
        parts = []
        current_text = ""
        current_tag = "normal"
        i = 0
        
        while i < len(line):
            # Cerca tag LINK con URL: [LINK:url]testo[/LINK]
            if line.startswith('[LINK:', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                # Estrae l'URL
                url_start = i + 6  # dopo '[LINK:'
                url_end = line.find(']', url_start)
                if url_end != -1:
                    url = line[url_start:url_end]
                    # Trova il testo del link
                    text_start = url_end + 1
                    text_end = line.find('[/LINK]', text_start)
                    if text_end != -1:
                        link_text = line[text_start:text_end]
                        parts.append((link_text, "hyperlink", url))
                        i = text_end + 7  # dopo '[/LINK]'
                        continue
                current_text += line[i]
                i += 1
            # Cerca tag di apertura/chiusura
            elif line[i:i+6] == '[BOLD]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "bold"
                i += 6
            elif line[i:i+8] == '[/BOLD]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line[i:i+8] == '[NORMAL]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line[i:i+10] == '[/NORMAL]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 10
            elif line[i:i+6] == '[CODE]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "code"
                i += 6
            elif line[i:i+8] == '[/CODE]':
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += 8
            elif line.startswith('[WARNING]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "warning"
                i += len('[WARNING]')
            elif line.startswith('[/WARNING]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += len('[/WARNING]')
            elif line.startswith('[WARNING_RED]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "warning_red"
                i += len('[WARNING_RED]')
            elif line.startswith('[/WARNING_RED]', i):
                if current_text:
                    parts.append((current_text, current_tag, None))
                    current_text = ""
                current_tag = "normal"
                i += len('[/WARNING_RED]')
            else:
                current_text += line[i]
                i += 1
        
        # Aggiunge l'ultimo pezzo
        if current_text:
            parts.append((current_text, current_tag, None))
        elif not parts:
            parts.append(("", "normal", None))
        
        # Inserisce tutti i pezzi con i tag di ancoraggio
        for item in parts:
            if len(item) == 3:
                text, tag, url = item
                if tag == "hyperlink" and url:
                    # Crea un tag univoco per questo link
                    link_tag = f"link_{id(url)}"
                    # Configura il tag con lo stile del link
                    self.text_content.tag_configure(link_tag, foreground="blue", underline=True)
                    # Bind del click per aprire l'URL
                    self.text_content.tag_bind(link_tag, "<Button-1>", lambda e, u=url: webbrowser.open(u))
                    self.text_content.tag_bind(link_tag, "<Enter>", lambda e: self.text_content.config(cursor="hand2"))
                    self.text_content.tag_bind(link_tag, "<Leave>", lambda e: self.text_content.config(cursor=""))
                    # Combina anchor_tags con il tag del link
                    combined_tags = list(anchor_tags) + [link_tag]
                    self.text_content.insert(tk.END, text, tuple(combined_tags))
                else:
                    # Combina il tag di formattazione con i tag di ancoraggio
                    combined_tags = list(anchor_tags) + [tag]
                    self.text_content.insert(tk.END, text, tuple(combined_tags))
            else:
                # Fallback per compatibilità
                combined_tags = list(anchor_tags) + [item[1]]
                self.text_content.insert(tk.END, item[0], tuple(combined_tags))
        
        self.text_content.insert(tk.END, "\n")
    
    def setup_search_functionality(self, parent_frame):
        """Crea l'interfaccia di ricerca con Entry, pulsanti e contatore risultati"""
        # Determina la lingua corrente per le traduzioni
        config = configparser.ConfigParser(interpolation=None)
        config_file = get_config_file()
        current_language = "en"  # default
        if os.path.exists(config_file):
            config.read(config_file, encoding='utf-8')
            if 'Settings' in config and config.has_option('Settings', 'language'):
                current_language = config.get('Settings', 'language', fallback='en')
        if current_language not in ['en', 'it']:
            current_language = "en"
        
        # Testi tradotti
        search_label_text = "Trova:" if current_language == 'it' else "Search:"
        search_button_text = "🔍 Trova" if current_language == 'it' else "🔍 Search"
        next_button_text = "⏩ Successivo" if current_language == 'it' else "⏩ Next"
        
        # Label "Trova:" / "Search:"
        ttk.Label(parent_frame, text=search_label_text).pack(side=tk.LEFT, padx=(5, 5))
        
        # Entry per digitare la parola da cercare
        search_entry = ttk.Entry(parent_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=(0, 5))
        search_entry.bind("<Return>", lambda e: self.search_text())
        search_entry.bind("<Escape>", lambda e: self.clear_search())
        
        # Pulsante "Trova" / "Search"
        ttk.Button(parent_frame, text=search_button_text, command=self.search_text).pack(side=tk.LEFT, padx=(0, 5))
        
        # Pulsante "Successivo" / "Next"
        ttk.Button(parent_frame, text=next_button_text, command=self.search_next).pack(side=tk.LEFT, padx=(0, 5))
        
        # Label per mostrare "Risultato X di Y" / "Result X of Y"
        self.search_result_label = ttk.Label(parent_frame, text="", foreground="blue")
        self.search_result_label.pack(side=tk.LEFT, padx=(10, 5))
    
    def search_text(self):
        """Cerca il testo nel widget e evidenzia tutti i match (case-insensitive)
        
        STRATEGIA ANTI-SEGFAULT: Non usa Text.search() (nativo Tcl/Tk) ma ricerca Python pura.
        """
        # Pulisci ricerca precedente
        self.clear_search()
        
        search_term = self.search_var.get().strip()
        if not search_term:
            return
        
        # FASE 1: Estrai contenuto completo e cerca in Python puro (NO Text.search()!)
        full_text = self.text_content.get("1.0", tk.END)
        search_lower = search_term.lower()
        text_lower = full_text.lower()
        
        # Trova tutti gli offset dei match usando str.find() in loop
        matches_offsets = []
        start_offset = 0
        max_matches = 1000
        
        while len(matches_offsets) < max_matches:
            offset = text_lower.find(search_lower, start_offset)
            if offset == -1:
                break
            matches_offsets.append(offset)
            start_offset = offset + 1  # Continua dopo il carattere corrente per trovare overlapping matches
        
        # FASE 2: Converti offset Python -> indici Tkinter e applica tag
        if matches_offsets:
            self.text_content.config(state="normal")
            try:
                for offset in matches_offsets:
                    # Converti offset in indice Tk: "1.0" + N caratteri
                    start_idx = f"1.0+{offset}c"
                    end_idx = f"1.0+{offset + len(search_term)}c"
                    
                    self.text_content.tag_add("search_highlight", start_idx, end_idx)
                    self.search_matches.append(start_idx)
            finally:
                self.text_content.config(state="disabled")
            
            # Vai al primo match
            self.current_match_index = 0
            self.text_content.see(self.search_matches[0])
            self.update_search_counter()
        else:
            # Nessun risultato trovato
            config = configparser.ConfigParser(interpolation=None)
            config_file = get_config_file()
            current_language = "en"
            if os.path.exists(config_file):
                config.read(config_file, encoding='utf-8')
                if 'Settings' in config and config.has_option('Settings', 'language'):
                    current_language = config.get('Settings', 'language', fallback='en')
            if current_language not in ['en', 'it']:
                current_language = "en"
            
            no_result_text = "Nessun risultato" if current_language == 'it' else "No results"
            self.search_result_label.config(text=no_result_text)
    
    def search_next(self):
        """Naviga al risultato successivo nella lista dei match"""
        if not self.search_matches:
            # Se non ci sono match, esegui una nuova ricerca
            self.search_text()
            return
        
        # Incrementa l'indice con wrap-around
        self.current_match_index = (self.current_match_index + 1) % len(self.search_matches)
        
        # Scrolla al match corrente
        current_pos = self.search_matches[self.current_match_index]
        self.text_content.see(current_pos)
        
        # Aggiorna il contatore
        self.update_search_counter()
    
    def update_search_counter(self):
        """Aggiorna la label con il contatore dei risultati"""
        if not self.search_matches:
            self.search_result_label.config(text="")
            return
        
        # Determina la lingua per il formato del contatore
        config = configparser.ConfigParser(interpolation=None)
        config_file = get_config_file()
        current_language = "en"
        if os.path.exists(config_file):
            config.read(config_file, encoding='utf-8')
            if 'Settings' in config and config.has_option('Settings', 'language'):
                current_language = config.get('Settings', 'language', fallback='en')
        if current_language not in ['en', 'it']:
            current_language = "en"
        
        # Formato: "Risultato 1 di 5" (italiano) o "Result 1 of 5" (inglese)
        total = len(self.search_matches)
        current = self.current_match_index + 1
        
        if current_language == 'it':
            counter_text = f"Risultato {current} di {total}"
        else:
            counter_text = f"Result {current} of {total}"
        
        self.search_result_label.config(text=counter_text)
    
    def clear_search(self):
        """Pulisce le evidenziazioni della ricerca precedente"""
        # Tag remove richiede ancora il toggle state per funzionare su widget disabled
        self.text_content.config(state="normal")
        self.text_content.tag_remove("search_highlight", "1.0", tk.END)
        self.text_content.config(state="disabled")
        
        # Reset variabili
        self.search_matches = []
        self.current_match_index = -1
        if self.search_result_label:
            self.search_result_label.config(text="")
