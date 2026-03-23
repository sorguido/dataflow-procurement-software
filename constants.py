"""
Costanti di configurazione per DataFlow.
Layout, dimensioni finestre e configurazione UI.
"""

# --- COSTANTI LAYOUT E GEOMETRIA FINESTRE ---

# Buffer per evitare sovrapposizione con taskbar (pixel)
TASKBAR_BUFFER = 100

# Dimensioni colonne articoli (pixel)
BASE_ARTICLE_WIDTH = 470  # Codice (80) + Allegato (80) + Descrizione (250) + Q.tà (60)
CONTO_LAVORO_WIDTH = 350  # Cod.Grezzo (100) + Dis.Grezzo (100) + Mat.C/L (150)
SUPPLIER_COLUMN_WIDTH = 120  # Larghezza colonna fornitore

# Margini e padding finestre (pixel)
PADDING = 140  # Margini laterali, scrollbar, bordi finestra e safety margin per DPI scaling

# Dimensioni minime finestre (pixel)
BUTTONS_MIN_WIDTH = 1150  # 6 pulsanti × 180px + spaziatura (testi tradotti lunghi + DPI scaling)
MIN_WINDOW_WIDTH = 850

# Percentuali dimensionamento automatico finestre
SCREEN_WIDTH_PERCENTAGE = 0.95   # 95% larghezza schermo (lascia spazio ai bordi)
SCREEN_HEIGHT_PERCENTAGE = 0.80  # 80% altezza schermo (lascia spazio per taskbar)
