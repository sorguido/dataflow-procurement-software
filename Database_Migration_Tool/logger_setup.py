"""
DataFlow Database Migration Tool - Logging Setup
Configures dual logging: console (INFO) and file (DEBUG)
"""

import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler


def setup_migration_logger(log_dir=None):
    """
    Configure logging for migration tool.
    
    Console: INFO level (clean output for user)
    File: DEBUG level (detailed for troubleshooting)
    
    Args:
        log_dir: Directory for log files. If None, uses Documents/DataFlow/Migration_Logs/
        
    Returns:
        logging.Logger: Configured logger instance
    """
    # Determine log directory
    if log_dir is None:
        log_dir = os.path.join(
            os.path.expanduser('~\\Documents'),
            'DataFlow',
            'Migration_Logs'
        )
    
    os.makedirs(log_dir, exist_ok=True)
    
    # Create log filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'migration_{timestamp}.log')
    
    # Create logger
    logger = logging.getLogger('DataFlowMigration')
    logger.setLevel(logging.DEBUG)  # Capture everything
    
    # Remove existing handlers (avoid duplicates)
    if logger.handlers:
        logger.handlers.clear()
    
    # Console handler (INFO level - clean output)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter(
        '%(levelname)s: %(message)s'
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # File handler (DEBUG level - detailed)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    # Log the log file location
    logger.info(f"Detailed log file: {log_file}")
    
    return logger, log_file
