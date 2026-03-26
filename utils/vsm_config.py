"""
Utility per gestione configurazione VSM.
Gestisce parametri configurabili per il modulo VSM (coefficienti, settings, etc.).
"""

import configparser
import logging
from utils.user_utils import get_config_file

logger = logging.getLogger(__name__)

# Default values
DEFAULT_PAGAMENTI_COEFFICIENT = 0.005  # 0.5% mensile costo opportunità capitale


def get_pagamenti_coefficient() -> float:
    """
    Recupera il coefficiente per il driver Pagamenti dal config.ini.
    
    Il coefficiente rappresenta il costo opportunità del capitale per ogni 30 giorni
    di dilazione pagamento. Valore default: 0.005 (0.5% mensile).
    
    Returns:
        float: Coefficiente Pagamenti (default 0.005)
    """
    try:
        config_file = get_config_file()
        config = configparser.ConfigParser(interpolation=None)
        config.read(config_file, encoding='utf-8')
        
        # Leggi dalla sezione Settings
        if config.has_section('Settings'):
            coeff_str = config.get('Settings', 'vsm_pagamenti_coefficient', fallback=None)
            if coeff_str:
                return float(coeff_str)
        
        # Se non esiste, inizializza con valore default
        _initialize_pagamenti_coefficient()
        return DEFAULT_PAGAMENTI_COEFFICIENT
        
    except Exception as e:
        logger.warning(f"Errore lettura coefficiente Pagamenti da config: {e}. Uso default {DEFAULT_PAGAMENTI_COEFFICIENT}")
        return DEFAULT_PAGAMENTI_COEFFICIENT


def set_pagamenti_coefficient(value: float) -> bool:
    """
    Imposta il coefficiente per il driver Pagamenti nel config.ini.
    
    Args:
        value: Nuovo valore del coefficiente (es. 0.005 per 0.5% mensile)
        
    Returns:
        bool: True se salvataggio riuscito, False altrimenti
    """
    try:
        config_file = get_config_file()
        config = configparser.ConfigParser(interpolation=None)
        config.read(config_file, encoding='utf-8')
        
        # Assicura esistenza sezione Settings
        if not config.has_section('Settings'):
            config.add_section('Settings')
        
        # Salva valore
        config.set('Settings', 'vsm_pagamenti_coefficient', str(value))
        
        with open(config_file, 'w', encoding='utf-8') as f:
            config.write(f)
        
        logger.info(f"Coefficiente Pagamenti aggiornato: {value}")
        return True
        
    except Exception as e:
        logger.error(f"Errore scrittura coefficiente Pagamenti: {e}")
        return False


def _initialize_pagamenti_coefficient():
    """Inizializza il coefficiente Pagamenti con valore default se non esiste."""
    try:
        config_file = get_config_file()
        config = configparser.ConfigParser(interpolation=None)
        config.read(config_file, encoding='utf-8')
        
        if not config.has_section('Settings'):
            config.add_section('Settings')
        
        # Inizializza solo se non esiste già
        if not config.has_option('Settings', 'vsm_pagamenti_coefficient'):
            config.set('Settings', 'vsm_pagamenti_coefficient', str(DEFAULT_PAGAMENTI_COEFFICIENT))
            
            with open(config_file, 'w', encoding='utf-8') as f:
                config.write(f)
            
            logger.info(f"Coefficiente Pagamenti inizializzato: {DEFAULT_PAGAMENTI_COEFFICIENT}")
            
    except Exception as e:
        logger.warning(f"Errore inizializzazione coefficiente Pagamenti: {e}")
