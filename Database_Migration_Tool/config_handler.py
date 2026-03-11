"""
DataFlow Database Migration Tool - Config.ini Handler
Reads and validates config.ini from DataFlow 2.0.0
"""

import configparser
import os
import unicodedata
import re


class ConfigError(Exception):
    """Custom exception for configuration errors"""
    pass


def normalize_string(s):
    """Remove accents and special characters from string"""
    if not s:
        return s
    # Decompose unicode characters and filter out accents
    nfkd = unicodedata.normalize('NFKD', s)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])


def generate_username(first_name, last_name):
    """
    Generate username from first name and last name.
    Format: first letter of first name + last name, lowercase, no accents.
    
    Example: Mario Rossi -> mrossi
    
    Args:
        first_name: User's first name
        last_name: User's last name
        
    Returns:
        str: Generated username
    """
    if not first_name or not last_name:
        raise ConfigError("First name and last name required to generate username")
    
    # Normalize and remove accents
    first = normalize_string(first_name.strip())
    last = normalize_string(last_name.strip())
    
    if not first or not last:
        raise ConfigError("Invalid names: cannot generate username")
    
    # Take first letter of first name + full last name
    username = (first[0] + last).lower()
    
    # Remove any remaining special characters
    username = re.sub(r'[^a-z0-9]', '', username)
    
    if not username:
        raise ConfigError("Generated username is empty")
    
    return username


def read_config_ini(config_path):
    """
    Read and parse config.ini file.
    
    Args:
        config_path: Full path to config.ini file
        
    Returns:
        dict: Configuration data with keys:
            - username (str)
            - first_name (str, optional)
            - last_name (str, optional)
            - language (str, optional)
            - dataflow_base_dir (str, optional)
            
    Raises:
        ConfigError: If config is invalid or missing required data
    """
    if not os.path.exists(config_path):
        raise ConfigError(f"Config file not found: {config_path}")
    
    config = configparser.ConfigParser()
    
    try:
        config.read(config_path, encoding='utf-8')
    except Exception as e:
        raise ConfigError(f"Failed to read config file: {e}")
    
    result = {}
    
    # Try to get username directly
    if config.has_section('User') and config.has_option('User', 'username'):
        username = config.get('User', 'username').strip()
        if username:
            result['username'] = username
            result['first_name'] = config.get('User', 'first_name', fallback='').strip()
            result['last_name'] = config.get('User', 'last_name', fallback='').strip()
        else:
            # Username exists but is empty - try to generate from names
            first_name = config.get('User', 'first_name', fallback='').strip()
            last_name = config.get('User', 'last_name', fallback='').strip()
            
            if first_name and last_name:
                result['username'] = generate_username(first_name, last_name)
                result['first_name'] = first_name
                result['last_name'] = last_name
            else:
                raise ConfigError(
                    "Username is empty in config.ini and first_name/last_name are missing.\n"
                    "Please configure user identity in DataFlow 2.0.0 before running migration."
                )
    else:
        # No username section - try to generate from names
        if config.has_section('User'):
            first_name = config.get('User', 'first_name', fallback='').strip()
            last_name = config.get('User', 'last_name', fallback='').strip()
            
            if first_name and last_name:
                result['username'] = generate_username(first_name, last_name)
                result['first_name'] = first_name
                result['last_name'] = last_name
            else:
                raise ConfigError(
                    "User section exists but first_name/last_name are missing.\n"
                    "Please configure user identity in DataFlow 2.0.0 before running migration."
                )
        else:
            raise ConfigError(
                "Config.ini is missing [User] section.\n"
                "Please run DataFlow 2.0.0 at least once to configure user identity."
            )
    
    # Get optional settings
    if config.has_section('Settings'):
        result['language'] = config.get('Settings', 'language', fallback='en')
        result['dataflow_base_dir'] = config.get('Settings', 'dataflow_base_dir', fallback='')
    
    return result


def get_target_paths(config_data):
    """
    Calculate target database and attachments paths based on config.
    
    Args:
        config_data: Dict returned by read_config_ini()
        
    Returns:
        dict: Paths with keys:
            - base_dir: Base DataFlow_{username} folder
            - db_dir: Database folder
            - db_file: Full path to database file
            - attachments_dir: Attachments folder
    """
    username = config_data['username']
    
    # Determine base directory
    if config_data.get('dataflow_base_dir'):
        base_dir = os.path.join(
            config_data['dataflow_base_dir'],
            f'DataFlow_{username}'
        )
    else:
        # Default: Documents folder
        base_dir = os.path.join(
            os.path.expanduser('~\\Documents'),
            f'DataFlow_{username}'
        )
    
    return {
        'base_dir': base_dir,
        'db_dir': os.path.join(base_dir, 'Database'),
        'db_file': os.path.join(base_dir, 'Database', f'dataflow_db_{username}.db'),
        'attachments_dir': os.path.join(base_dir, 'Attachments')
    }
