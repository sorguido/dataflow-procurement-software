"""
Main Dashboard Toolbar Component

Componente UI per la Global Search bar della Main Dashboard.
Parte del redesign v2.1.0 per ridurre sovraccarico visivo e migliorare UX.

Step 1 (Foundation): Struttura UI statica
Step 2 (Integrazione): Inserimento in MainWindow
Step 3 (Placeholder): Comportamento placeholder dinamico con gestione focus
Step 4 (Ricerca): Collegamento a logica search_requests() esistente
Step 5 (Global Search): Ricerca multi-campo con OR logic su 6 campi principali
"""

import tkinter as tk
from tkinter import ttk

from utils.i18n_utils import _


class MainDashboardToolbar(ttk.Frame):
    """Toolbar principale con Global Search bar centrale.
    
    Questo componente rappresenta la barra di ricerca globale che sarà
    il punto centrale dell'esperienza utente nella nuova dashboard.
    
    Caratteristiche implementate:
    - Entry centrale per ricerca globale
    - Placeholder text descrittivo con stile differenziato
    - Comportamento placeholder dinamico (clear on focus, restore on blur)
    - Binding Enter → ricerca tramite search_requests() esistente
    - Layout responsive orizzontale
    - Ricerca multi-campo globale (OR su: num, ref, forn, cod, desc, ord)
    - Coesistenza con filtri avanzati (global OR + filtri AND)
    
    Comportamento ricerca:
    - Campo vuoto + Enter = Clear Filters
    - Campo pieno + Enter = Global search (OR) + filtri attivi (AND)
    - Global search NON cancella filtri esistenti (OPZIONE A)
    """
    
    # Costante per testo placeholder
    PLACEHOLDER_TEXT = "Search anything... (RFQ, Supplier, Code, Project...)"
    
    def __init__(self, parent, main_window):
        """Inizializza la toolbar.
        
        Args:
            parent: Widget parent (tipicamente root di MainWindow)
            main_window: Istanza di MainWindow per accesso a metodi di ricerca
        """
        super().__init__(parent)
        self.main_window = main_window
        self._is_placeholder = True  # Flag per tracciare stato placeholder
        self._setup_ui()
    
    def _setup_ui(self):
        """Crea layout della toolbar con Global Search Entry.
        
        Implementazione Step 4:
        - Entry con placeholder dinamico
        - Binding focus in/out per gestione placeholder
        - Binding Enter per invio ricerca
        - Stile visivo differenziato (colore grigio per placeholder)
        - Toggle filtri avanzati (Step 5)
        
        UI Enhancement:
        - Altezza search bar aumentata (~50%) tramite font più grande e padding
        - "Advanced Filters" integrato a destra nella search bar
        """
        # Frame wrapper per gestire padding e layout interno
        content_frame = ttk.Frame(self)
        content_frame.pack(fill="x", expand=True, padx=10, pady=5)
        
        # Container per search bar unificata (Entry + Advanced Filters)
        # Frame con bordo per sembrare un unico componente visivo
        search_container = tk.Frame(
            content_frame,
            relief='solid',
            borderwidth=1,
            background='white'
        )
        search_container.pack(fill="x", expand=True)
        
        # Entry per Global Search - elemento centrale della nuova UX
        # Usa tk.Entry invece di ttk.Entry per supporto nativo foreground color
        # Font aumentato a 12 e ipady=8 per altezza ~50% maggiore
        self.search_entry = tk.Entry(
            search_container,
            width=60,
            relief='flat',
            borderwidth=0,
            font=('TkDefaultFont', 12),
            background='white'
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=8)
        
        # Toggle filtri avanzati integrato a destra nella search bar
        # Label cliccabile con chevron per indicare stato expand/collapse
        self.filters_toggle_label = tk.Label(
            search_container,
            text=f"⌄ {_('Advanced Filters')}",
            cursor="hand2",
            foreground="#0066cc",
            font=('TkDefaultFont', 10),
            background='white',
            padx=10
        )
        self.filters_toggle_label.pack(side="right", pady=2)
        self.filters_toggle_label.bind("<Button-1>", self._on_toggle_filters)
        
        # Imposta placeholder iniziale
        self._set_placeholder()
        
        # Binding per gestione dinamica placeholder
        self.search_entry.bind("<FocusIn>", self._on_focus_in)
        self.search_entry.bind("<FocusOut>", self._on_focus_out)
        
        # Binding per invio ricerca (Step 4)
        self.search_entry.bind("<Return>", self._on_search)
    
    def _set_placeholder(self):
        """Mostra il placeholder text con stile visivo differenziato.
        
        Imposta il testo placeholder con colore grigio per indicare visivamente
        che non è input utente reale.
        """
        self._is_placeholder = True
        self.search_entry.delete(0, 'end')
        self.search_entry.insert(0, self.PLACEHOLDER_TEXT)
        self.search_entry.config(foreground='gray')
    
    def _clear_placeholder(self):
        """Rimuove il placeholder se presente.
        
        Pulisce il campo e ripristina lo stile normale (testo nero) per
        permettere input utente.
        """
        if self._is_placeholder:
            self._is_placeholder = False
            self.search_entry.delete(0, 'end')
            self.search_entry.config(foreground='black')
    
    def _on_focus_in(self, event=None):
        """Chiamato quando l'Entry riceve focus.
        
        Comportamento: rimuove il placeholder se presente, lasciando il campo
        vuoto e pronto per input utente.
        
        Args:
            event: Evento Tkinter (opzionale)
        """
        self._clear_placeholder()
    
    def _on_focus_out(self, event=None):
        """Chiamato quando l'Entry perde focus.
        
        Comportamento: se il campo è vuoto (utente non ha scritto nulla),
        ripristina il placeholder. Altrimenti mantiene il testo utente.
        
        Args:
            event: Evento Tkinter (opzionale)
        """
        if not self.search_entry.get():
            self._set_placeholder()
    
    def _on_search(self, event=None):
        """Chiamato quando l'utente preme Enter nella search bar.
        
        Step 5 Implementation: Global Search con OR Logic (OPZIONE A)
        
        Comportamento:
        1. Verifica che non sia placeholder
        2. Se campo vuoto → clear_filters() (reset completo)
        3. Se campo pieno → scrive in search_vars['global']
        4. Chiama search_requests() che applica ricerca globale OR + filtri AND
        
        Global Search OR su 6 campi principali:
        - id_richiesta (numero RfQ)
        - riferimento (progetto)
        - nome_fornitore
        - codice_materiale
        - descrizione_materiale
        - numeri_ordine
        
        IMPORTANTE: Global search + filtri avanzati COESISTONO (OPZIONE A)
        - Global search NON cancella i filtri esistenti
        - Combinazione: (global OR) AND (filtri AND)
        
        Args:
            event: Evento Tkinter (opzionale)
        """
        # Ignora se è ancora placeholder (campo effettivamente vuoto)
        if self._is_placeholder:
            return
        
        # Leggi testo dalla Global Search
        search_text = self.search_entry.get().strip()
        
        # Se vuoto (solo spazi), comportati come "Clear Filters"
        # Riusa la logica esistente del pulsante reset per coerenza UX
        if not search_text:
            if hasattr(self.main_window, 'clear_filters'):
                self.main_window.clear_filters()
            return
        
        # Popola il campo 'global' per ricerca multi-campo con OR
        # OPZIONE A: global search + filtri avanzati coesistono
        # La global search NON cancella i filtri esistenti
        if hasattr(self.main_window, 'search_vars') and 'global' in self.main_window.search_vars:
            self.main_window.search_vars['global'].set(search_text)
            # DEBUG: Verifica che il valore venga scritto
            print(f"[GlobalSearch] input='{search_text}'")
            print(f"[GlobalSearch] search_vars['global']='{self.main_window.search_vars['global'].get()}'")
        
        # Chiama il metodo di ricerca esistente
        # Tutta la logica (validazione, SQL, update UI) è gestita lì
        if hasattr(self.main_window, 'search_requests'):
            self.main_window.search_requests()
    
    def _on_toggle_filters(self, event=None):
        """Chiamato quando l'utente clicca sul toggle filtri avanzati.
        
        Step 5 Implementation: Collapsible Filters
        
        Comportamento:
        - Chiama MainWindow.toggle_filters() che gestisce CollapsibleFilters
        - Aggiorna l'icona del chevron (⌄ → ⌃ e viceversa)
        
        Args:
            event: Evento Tkinter (opzionale)
        """
        if hasattr(self.main_window, 'toggle_filters'):
            self.main_window.toggle_filters()
            
            # Aggiorna icona chevron in base allo stato
            if hasattr(self.main_window, 'collapsible_filters'):
                if self.main_window.collapsible_filters.is_expanded():
                    self.filters_toggle_label.config(text=f"⌃ {_('Advanced Filters')}")
                else:
                    self.filters_toggle_label.config(text=f"⌄ {_('Advanced Filters')}")
