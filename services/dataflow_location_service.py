"""Helpers for DataFlow folder migration validation flows."""

from __future__ import annotations

import os


def normalize_parent_directory(selected_dir):
    """Normalize selected parent folder path."""
    if not selected_dir:
        return None
    normalized = os.path.normpath(os.path.abspath(selected_dir.strip()))
    return normalized or None


def ensure_parent_directory_writable(path):
    """Ensure folder exists and can be written."""
    os.makedirs(path, exist_ok=True)
    test_file = os.path.join(path, ".dataflow_test_write")
    try:
        with open(test_file, "w", encoding="utf-8") as file_obj:
            file_obj.write("test")
    finally:
        try:
            os.remove(test_file)
        except FileNotFoundError:
            pass


def detect_username_conflict(*, parent_dir, username, logger):
    """Detect whether destination already contains DataFlow assets for username."""
    potential_folder = os.path.join(parent_dir, f"DataFlow_{username}")
    potential_db = os.path.join(potential_folder, "Database", f"dataflow_db_{username}.db")

    folder_exists = os.path.exists(potential_folder)
    db_exists = False

    if folder_exists:
        try:
            db_exists = os.path.exists(potential_db)
            if db_exists:
                try:
                    with open(potential_db, "rb") as file_obj:
                        file_obj.read(1)
                    logger.info("Controllo conflitto: DB '%s' esiste ed è accessibile", potential_db)
                except (PermissionError, OSError) as error:
                    logger.warning("DB '%s' esistente ma locked/inaccessibile: %s", potential_db, error)
                    db_exists = True
        except Exception as error:
            logger.error("Errore nel controllo esistenza DB: %s", error)
            db_exists = True

    return {
        "potential_folder": potential_folder,
        "potential_db": potential_db,
        "folder_exists": folder_exists,
        "db_exists": db_exists,
    }
