"""
Collapsible Filters Component

Wrapper per rendere il blocco "Filtri di Ricerca" collassabile.
Parte del redesign v2.1.0 per ridurre il sovraccarico visivo.

Approccio: True Container Pattern
- Crea internamente il LabelFrame che conterrà i filtri
- I widget dei filtri vengono aggiunti dall'esterno al filters_frame
- Gestisce solo la visibilità (show/hide) tramite pack_forget()/pack()
- Il wrapper rimane sempre packed, solo il contenuto interno viene nascosto
"""

from tkinter import ttk
from utils.i18n_utils import tr


class CollapsibleFilters(ttk.Frame):
    """Wrapper collassabile per il frame filtri di ricerca.
    
    Questo componente crea un container interno (filters_frame) dove
    MainWindow inserirà i widget dei filtri. Il wrapper gestisce solo
    la visibilità del contenuto tramite toggle.
    
    Caratteristiche:
    - Default: collapsed (nascosto)
    - Toggle tramite metodo toggle()
    - Zero layout shift: il wrapper rimane sempre packed
    - Reparenting reale: filters_frame è figlio diretto del wrapper
    """
    
    def __init__(self, parent, label_text=None):
        """Inizializza il wrapper collassabile.
        
        Args:
            parent: Widget parent (tipicamente root di MainWindow)
            label_text: Testo del LabelFrame (i18n gestito dal chiamante)
        """
        super().__init__(parent)
        if label_text is None:
            label_text = tr("Search Filters")
        
        self._is_expanded = False  # Default: nascosto
        
        # Assicura che il wrapper si ridimensioni in base al contenuto
        # Quando filters_frame è nascosto, il wrapper si compatta a zero
        self.pack_propagate(True)
        
        # Crea il LabelFrame interno che conterrà i filtri
        # Pack UNA volta sola dentro il wrapper - rimane sempre visible internamente
        self.filters_frame = ttk.LabelFrame(self, text=f"🔍 {label_text}", padding=(10, 5))
        self.filters_frame.pack(fill="x", padx=0, pady=0)
        
        # Default collapsed: il wrapper stesso sarà rimosso con grid_remove()
        # Il filters_frame rimane packed dentro, ma il wrapper non è nella griglia
    
    def set_grid_config(self, **kwargs):
        """Salva la configurazione grid per ripristino dopo grid_remove().
        
        Args:
            **kwargs: Parametri grid (row, column, sticky, padx, pady, etc.)
        """
        self._grid_config = kwargs
    
    def expand(self):
        """Mostra i filtri di ricerca.
        
        Usa grid() per ripristinare il wrapper nella griglia.
        Grid è più robusto di pack() per show/hide dinamici.
        """
        if not self._is_expanded:
            self._is_expanded = True
            # Ripristina il wrapper nella griglia con i parametri salvati
            if self._grid_config:
                self.grid(**self._grid_config)
    
    def collapse(self):
        """Nasconde i filtri di ricerca.
        
        Usa grid_remove() per rimuovere il wrapper dalla griglia.
        Grid_remove() rimuove completamente senza lasciare gap.
        Mantiene i parametri grid per ripristino successivo.
        """
        if self._is_expanded:
            self._is_expanded = False
            # Rimuove il wrapper dalla griglia (ma mantiene configurazione)
            self.grid_remove()
    
    def toggle(self):
        """Alterna tra expanded e collapsed.
        
        Chiamato dal trigger nella Global Search toolbar.
        """
        if self._is_expanded:
            self.collapse()
        else:
            self.expand()
    
    def is_expanded(self):
        """Ritorna lo stato corrente (expanded/collapsed).
        
        Returns:
            bool: True se filtri visibili, False se nascosti
        """
        return self._is_expanded
