"""
Utility per gestione identità utente e configurazione applicazione.
Gestisce percorsi config cross-platform e lettura/scrittura identità utente.
"""

import os
import sys
import configparser


def get_app_data_dir():
    """Restituisce la directory dati dell'applicazione."""
    if getattr(sys, 'frozen', False):
        # Se eseguito come EXE/MSIX, usa la directory locale dell'utente
        if sys.platform == 'win32':
            return os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DataFlow')
        else:
            xdg_data_home = os.environ.get('XDG_DATA_HOME')
            if xdg_data_home:
                return os.path.join(xdg_data_home, 'DataFlow')
            else:
                return os.path.join(os.path.expanduser('~'), '.local', 'share', 'DataFlow')
    else:
        # Se eseguito come script Python, NON usare la directory del file .py
        xdg_data_home = os.environ.get('XDG_DATA_HOME')
        if xdg_data_home:
            return os.path.join(xdg_data_home, 'DataFlow')
        else:
            return os.path.join(os.path.expanduser('~'), '.local', 'share', 'DataFlow')


def get_config_file():
    """Restituisce il percorso del file config.ini."""
    # Assicura che la directory esista
    app_dir = get_app_data_dir()
    os.makedirs(app_dir, exist_ok=True)
    return os.path.join(app_dir, 'config.ini')


def load_user_identity():
    """Carica nome, cognome e username dell'utente dal config."""
    identity = {
        'first_name': '',
        'last_name': '',
        'username': '',
        'full_name': ''
    }
    config_file = get_config_file()
    config = configparser.ConfigParser(interpolation=None)
    if os.path.exists(config_file):
        config.read(config_file, encoding='utf-8')
        if config.has_section('User'):
            identity['first_name'] = config.get('User', 'first_name', fallback='').strip()
            identity['last_name'] = config.get('User', 'last_name', fallback='').strip()
            identity['username'] = config.get('User', 'username', fallback='').strip().lower()
            full_name = f"{identity['first_name']} {identity['last_name']}".strip()
            identity['full_name'] = full_name
    return identity


def save_user_identity(first_name, last_name, username):
    """Salva nel config il nome completo dell'utente e lo username derivato."""
    config_file = get_config_file()
    config = configparser.ConfigParser(interpolation=None)
    if os.path.exists(config_file):
        config.read(config_file, encoding='utf-8')
    if 'User' not in config:
        config['User'] = {}
    config['User']['first_name'] = first_name.strip()
    config['User']['last_name'] = last_name.strip()
    config['User']['username'] = username.strip().lower()
    with open(config_file, 'w', encoding='utf-8') as f:
        config.write(f)
