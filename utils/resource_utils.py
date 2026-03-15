"""
Gestione percorsi risorse per l'applicazione DataFlow.
Compatibile con PyInstaller e modalità sviluppo.
"""
import os
import sys
import logging
from PIL import Image, ImageTk


def resource_path(relative_path):
    """Ottiene il percorso assoluto della risorsa, funzionante sia in sviluppo che con PyInstaller.
    
    Args:
        relative_path: Percorso relativo alla root del progetto (es: "add_data/file.txt")
    
    Returns:
        Percorso assoluto alla risorsa
    """
    try:
        # PyInstaller crea una cartella temporanea e ci memorizza il percorso in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # BUG #21 FIX: Catch specifico invece di Exception generico
        # In sviluppo, usa la directory dello script
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            # NOTA: Doppio dirname() perché questo modulo è in utils/
            # Risale da utils/resource_utils.py → utils/ → root/
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def set_window_icon(window):
    """Imposta l'icona della finestra in modo cross-platform."""
    icon_path = resource_path(os.path.join("add_data", "DataFlow.ico"))
    logger = logging.getLogger(__name__)
    try:
        if sys.platform == 'win32':
            window.iconbitmap(icon_path)
        else:
            img = Image.open(icon_path)
            photo = ImageTk.PhotoImage(img)
            window._icon_photo = photo  # Evita garbage collection
            window.iconphoto(True, photo)
    except Exception as e:
        logger.debug(f"Icona non caricabile (normale in alcuni ambienti): {e}")
