"""
DataFlow Database Migration Tool
Migrates DataFlow v1.1.0 database to v2.0.0 format

This tool:
1. Reads config.ini from DataFlow 2.0.0 to get username
2. Selects source v1.1.0 database
3. Validates source database schema
4. Shows migration summary with warnings
5. Performs destructive migration (deletes target folder)
6. Extracts BLOB attachments to filesystem
7. Remaps IDs to year-based format
8. Assigns username to all RfQs
9. Generates detailed migration report

Author: DataFlow Migration Tool
Version: 1.0.0
"""

import sys
import os
import subprocess

# Enable DPI awareness for sharp rendering on high-DPI displays
try:
    from ctypes import windll
except ImportError:
    windll = None

if windll is not None:
    try:
        windll.shcore.SetProcessDpiAwareness(1)  # 1 = System DPI aware
    except:
        try:
            # Fallback for older Windows versions
            windll.user32.SetProcessDPIAware()
        except:
            pass  # If both fail, continue without DPI awareness

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# --- INIZIO CODICE AGGIUNTO PER PYINSTALLER ---
def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    
    Args:
        relative_path: Path relative to the script/bundle
    
    Returns:
        Absolute path to resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Normal Python execution
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
# --- FINE CODICE AGGIUNTO PER PYINSTALLER ---

from logger_setup import setup_migration_logger
from config_handler import read_config_ini, get_target_paths, ConfigError
from schema_validator import validate_v1_schema, get_database_summary, SchemaError
from migration_engine import MigrationEngine, MigrationError
from ui_dialogs import (
    WelcomeDialog,
    select_config_file, 
    select_source_database,
    show_error_dialog,
    show_info_dialog,
    MigrationSummaryDialog,
    FinalConfirmationDialog,
    ProgressDialog,
    CompletionDialog
)


def main():
    """Main entry point for migration tool"""
    
    print("=" * 70)
    print("DataFlow Database Migration Tool")
    print("Version 1.0.0 - Migrates v1.1.0 to v2.0.0")
    print("=" * 70)
    print()
    
    # Setup logging
    logger, log_file_path = setup_migration_logger()
    logger.info("Migration tool started")
    
    # Create welcome dialog (will be main window)
    welcome_dialog = None
    
    try:
        # ===================================================================
        # STEP 0: Show welcome dialog
        # ===================================================================
        print("Showing welcome dialog...")
        welcome_dialog = WelcomeDialog()
        if not welcome_dialog.show():
            logger.info("User cancelled from welcome dialog")
            print("Migration cancelled by user.")
            if welcome_dialog:
                welcome_dialog.close()
            return
        
        logger.info("User accepted welcome dialog")
        print("✓ User ready to proceed")
        print()
        
        # Get root window for use as parent
        parent_window = welcome_dialog.get_root()
        
        # ===================================================================
        # STEP 1: Select and read config.ini
        # ===================================================================
        print("Step 1: Select DataFlow 2.0.0 config.ini file...")
        config_path = select_config_file(parent_window)
        
        if not config_path:
            logger.info("User cancelled config.ini selection")
            print("Migration cancelled by user.")
            if welcome_dialog:
                welcome_dialog.close()
            return
        
        logger.info(f"Selected config.ini: {config_path}")
        print(f"Selected: {config_path}")
        
        # Read and validate config
        try:
            config_data = read_config_ini(config_path)
            username = config_data['username']
            logger.info(f"Username from config: {username}")
            print(f"Username: {username}")
        except ConfigError as e:
            logger.error(f"Config error: {e}")
            show_error_dialog(
                "Configuration Error",
                f"Failed to read configuration:\n\n{e}\n\n"
                "Please run DataFlow 2.0.0 at least once to configure user identity.",
                parent_window
            )
            if welcome_dialog:
                welcome_dialog.close()
            return
        
        # Calculate target paths
        target_paths = get_target_paths(config_data)
        logger.info(f"Target database: {target_paths['db_file']}")
        print(f"Target location: {target_paths['base_dir']}")
        print()
        
        # ===================================================================
        # STEP 2: Select source database
        # ===================================================================
        print("Step 2: Select DataFlow 1.1.0 source database...")
        source_db_path = select_source_database(parent_window)
        
        if not source_db_path:
            logger.info("User cancelled database selection")
            print("Migration cancelled by user.")
            if welcome_dialog:
                welcome_dialog.close()
            return
        
        logger.info(f"Selected source database: {source_db_path}")
        print(f"Selected: {source_db_path}")
        print()
        
        # ===================================================================
        # STEP 3: Validate source database schema
        # ===================================================================
        print("Step 3: Validating source database schema...")
        try:
            validation_results = validate_v1_schema(source_db_path)
            logger.info("Source database schema validation passed")
            
            # Log warnings if any
            if validation_results['warnings']:
                for warning in validation_results['warnings']:
                    logger.warning(warning)
            
            # Get database summary
            db_summary = get_database_summary(source_db_path)
            logger.info(f"Source database summary: {db_summary}")
            print(f"✓ Valid v1.1.0 database")
            print(f"  - {db_summary['rfqs']} RfQs")
            print(f"  - {db_summary['articles']} articles")
            print(f"  - {db_summary['suppliers']} suppliers")
            print(f"  - {db_summary['attachments']} attachments")
            print()
            
        except SchemaError as e:
            logger.error(f"Schema validation failed: {e}")
            show_error_dialog(
                "Database Validation Error",
                f"The selected database is not compatible:\n\n{e}\n\n"
                "Please ensure you're selecting a valid DataFlow v1.1.0 database.",
                parent_window
            )
            if welcome_dialog:
                welcome_dialog.close()
            return
        
        # Close welcome dialog now that file selection is complete
        if welcome_dialog:
            welcome_dialog.close()
            welcome_dialog = None
        
        # ===================================================================
        # STEP 4: Show migration summary and first confirmation
        # ===================================================================
        print("Step 4: Showing migration summary...")
        summary_dialog = MigrationSummaryDialog(
            source_db_path,
            target_paths,
            username,
            db_summary
        )
        
        if not summary_dialog.show():
            logger.info("User cancelled migration at summary dialog")
            print("Migration cancelled by user.")
            return
        
        logger.info("User confirmed migration summary")
        print("✓ User confirmed migration summary")
        print()
        
        # ===================================================================
        # STEP 5: Final confirmation with DELETE typing
        # ===================================================================
        print("Step 5: Final confirmation...")
        final_dialog = FinalConfirmationDialog(target_paths['base_dir'])
        
        if not final_dialog.show():
            logger.info("User cancelled migration at final confirmation")
            print("Migration cancelled by user.")
            return
        
        logger.info("User confirmed final confirmation (typed DELETE)")
        print("✓ User typed DELETE and confirmed")
        print()
        
        # ===================================================================
        # STEP 6: Execute migration with progress dialog
        # ===================================================================
        print("Step 6: Executing migration...")
        print("(A progress window will show migration status)")
        print()
        
        # Create progress dialog
        progress_dialog = ProgressDialog(total_steps=10)
        progress_dialog.show()
        
        # Progress callback for migration engine
        def progress_callback(step, message):
            progress_dialog.update(step, message)
        
        # Create migration engine and execute
        try:
            engine = MigrationEngine(
                source_db_path,
                target_paths,
                username,
                logger,
                progress_callback=progress_callback
            )
            
            statistics = engine.execute_migration()
            
            progress_dialog.close()
            logger.info("Migration engine completed successfully")
            print("✓ Migration completed successfully!")
            print()
            
        except MigrationError as e:
            progress_dialog.close()
            logger.error(f"Migration failed: {e}")
            show_error_dialog(
                "Migration Failed",
                f"Migration process failed:\n\n{e}\n\n"
                f"Please check the log file for details:\n{log_file_path}"
            )
            return
        
        # ===================================================================
        # STEP 7: Show completion dialog with statistics
        # ===================================================================
        print("Step 7: Showing completion summary...")
        completion_dialog = CompletionDialog(
            statistics,
            log_file_path,
            target_paths['db_file']
        )
        
        open_folder = completion_dialog.show()
        
        # Open folder in explorer if requested
        if open_folder:
            try:
                folder = os.path.dirname(target_paths['db_file'])
                subprocess.Popen(f'explorer /select,"{target_paths["db_file"]}"')
                logger.info(f"Opened folder: {folder}")
            except Exception as e:
                logger.warning(f"Failed to open folder: {e}")
        
        print()
        print("=" * 70)
        print("MIGRATION COMPLETED SUCCESSFULLY!")
        print(f"Migrated database: {target_paths['db_file']}")
        print(f"Log file: {log_file_path}")
        print("=" * 70)
        
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        show_error_dialog(
            "Unexpected Error",
            f"An unexpected error occurred:\n\n{e}\n\n"
            f"Please check the log file for details:\n{log_file_path}"
        )
        if welcome_dialog:
            welcome_dialog.close()
        raise


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nMigration interrupted by user.")
        sys.exit(1)
    except Exception:
        sys.exit(1)
