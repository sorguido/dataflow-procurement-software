"""
Utility per posizionamento, centratura e dimensionamento finestre Tkinter.
Helper functions per DataFlow.
"""

from constants import (
    TASKBAR_BUFFER,
    BASE_ARTICLE_WIDTH,
    CONTO_LAVORO_WIDTH,
    SUPPLIER_COLUMN_WIDTH,
    PADDING,
    BUTTONS_MIN_WIDTH,
    MIN_WINDOW_WIDTH,
    SCREEN_WIDTH_PERCENTAGE,
    SCREEN_HEIGHT_PERCENTAGE
)


def calculate_center_position(win):
    """Calcola la posizione centrale per la finestra senza renderla visibile."""
    # Forza il ricalcolo della geometria per ottenere le dimensioni corrette
    win.update() 
    
    width = win.winfo_reqwidth()
    height = win.winfo_reqheight()

    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()

    # Limita le dimensioni alle dimensioni dello schermo
    if width > screen_w:
        width = screen_w
    if height > screen_h - TASKBAR_BUFFER:
        height = screen_h - TASKBAR_BUFFER

    # Calcola le coordinate per centrare la finestra
    x = max(0, (screen_w - width) // 2)
    y = max(0, (screen_h - height) // 2)

    # --- INIZIO BLOCCO DI CONTROLLO ANTI-TASKBAR ---
    if y + height > screen_h - TASKBAR_BUFFER:
        y = screen_h - height - TASKBAR_BUFFER
    if y < 0:
        y = 0
    # --- FINE BLOCCO DI CONTROLLO ---

    return f'{width}x{height}+{x}+{y}'


def calculate_optimal_window_size(win, num_suppliers, is_conto_lavoro=False):
    """Calcola la larghezza ottimale per ViewRequestWindow in base al numero di fornitori."""
    # Calcola larghezza necessaria
    article_width = BASE_ARTICLE_WIDTH
    if is_conto_lavoro:
        article_width += CONTO_LAVORO_WIDTH
    
    suppliers_width = num_suppliers * SUPPLIER_COLUMN_WIDTH
    total_width = article_width + suppliers_width + PADDING
    
    # Ottieni dimensioni schermo
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    
    # Limita la larghezza al 95% dello schermo (lascia spazio ai bordi)
    max_width = int(screen_w * SCREEN_WIDTH_PERCENTAGE)
    optimal_width = min(total_width, max_width)
    
    # Larghezza minima: il maggiore tra larghezza pulsanti e larghezza base
    min_width = max(BUTTONS_MIN_WIDTH, MIN_WINDOW_WIDTH)
    optimal_width = max(optimal_width, min_width)
    
    # Altezza ottimale (80% dello schermo, lasciando spazio per taskbar)
    optimal_height = int(screen_h * SCREEN_HEIGHT_PERCENTAGE)
    
    # Calcola posizione centrale
    x = max(0, (screen_w - optimal_width) // 2)
    y = max(0, (screen_h - optimal_height) // 2)
    
    # Anti-taskbar buffer
    if y + optimal_height > screen_h - TASKBAR_BUFFER:
        y = screen_h - optimal_height - TASKBAR_BUFFER
    if y < 0:
        y = 0
    
    return f'{optimal_width}x{optimal_height}+{x}+{y}'


def center_window(win):
    """Centra la finestra e la rende visibile."""
    geometry = calculate_center_position(win)
    win.geometry(geometry)
    win.deiconify()
