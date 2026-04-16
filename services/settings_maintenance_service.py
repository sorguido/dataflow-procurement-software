"""Maintenance helpers for backup and autobackup flows."""

from __future__ import annotations

import configparser
import glob
import os
import shutil
import time
from datetime import datetime


def read_autobackup_config(config_file):
    """Read autobackup schedule settings from config."""
    config = configparser.ConfigParser(interpolation=None)
    config.read(config_file)
    return (
        config.getboolean("AutoBackup", "enabled", fallback=False),
        config.get("AutoBackup", "path", fallback="").strip(),
        config.get("AutoBackup", "hour", fallback="").strip(),
    )


def copy_manual_backup_bundle(*, db_file, dest, logger):
    """Copy DB + optional WAL/SHM for manual backup destination."""
    for attempt in range(5):
        try:
            with open(db_file, "r+b"):
                pass
            break
        except (PermissionError, IOError) as lock_error:
            if attempt < 4:
                logger.debug("Database ancora locked, tentativo %d/5: %s", attempt + 1, lock_error)
                time.sleep(0.2)
            else:
                logger.warning("Database potrebbe avere lock attivi dopo 5 tentativi")

    shutil.copy2(db_file, dest)
    logger.info("Backup DB principale: %s", dest)

    copied_files = [dest]

    wal_file = db_file.replace(".db", ".db-wal")
    if os.path.exists(wal_file):
        wal_dest = dest.replace(".db", ".db-wal")
        shutil.copy2(wal_file, wal_dest)
        logger.info("Backup WAL copiato: %s", wal_dest)
        copied_files.append(wal_dest)
    else:
        logger.info("File WAL non presente (normale se DB appena chiuso)")

    shm_file = db_file.replace(".db", ".db-shm")
    if os.path.exists(shm_file):
        shm_dest = dest.replace(".db", ".db-shm")
        shutil.copy2(shm_file, shm_dest)
        logger.info("Backup SHM copiato: %s", shm_dest)
        copied_files.append(shm_dest)
    else:
        logger.info("File SHM non presente (normale se DB appena chiuso)")

    original_size = os.path.getsize(db_file)
    backup_size = os.path.getsize(dest)
    total_size = sum(os.path.getsize(path) for path in copied_files if os.path.exists(path))
    return {
        "copied_files": copied_files,
        "original_size": original_size,
        "backup_size": backup_size,
        "total_size": total_size,
    }


def perform_autobackup_copy(*, db_file, dest_folder, logger):
    """Execute autobackup copy preserving current retention and retry rules."""
    if not os.path.exists(db_file):
        logger.warning("File database non trovato per backup: %s", db_file)
        return {"copied": False, "reason": "db_missing"}

    time.sleep(0.2)

    backup_sets = {}
    for ext in ["*.db", "*.db-wal", "*.db-shm"]:
        pattern = os.path.join(dest_folder, f"*_backup_auto_{ext.replace('*', '')}")
        for filepath in glob.glob(pattern):
            basename = os.path.basename(filepath)
            try:
                timestamp_part = basename.split("_backup_auto_")[1].rsplit(".", 1)[0]
                backup_sets.setdefault(timestamp_part, []).append(filepath)
            except (IndexError, ValueError):
                logger.warning("Formato nome backup non riconosciuto: %s", basename)

    sorted_timestamps = sorted(backup_sets.keys())
    while len(sorted_timestamps) >= 3:
        old_timestamp = sorted_timestamps.pop(0)
        for old_file in backup_sets[old_timestamp]:
            try:
                os.remove(old_file)
                logger.info("Rimosso vecchio backup: %s", old_file)
            except Exception as error:
                logger.warning("Impossibile eliminare vecchio backup %s: %s", old_file, error)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"gestione_offerte_backup_auto_{timestamp}"
    dest_path = os.path.join(dest_folder, f"{base_name}.db")

    max_retries = 3
    for attempt in range(max_retries):
        try:
            shutil.copy2(db_file, dest_path)
            break
        except (PermissionError, OSError) as error:
            if attempt < max_retries - 1:
                logger.warning("Tentativo backup %d fallito: %s, riprovo...", attempt + 1, error)
                time.sleep(1)
            else:
                raise

    logger.info("Backup automatico DB principale: %s", dest_path)
    copied_files = [dest_path]

    wal_file = db_file.replace(".db", ".db-wal")
    if os.path.exists(wal_file):
        wal_dest = os.path.join(dest_folder, f"{base_name}.db-wal")
        try:
            shutil.copy2(wal_file, wal_dest)
            logger.info("Backup WAL copiato: %s", wal_dest)
            copied_files.append(wal_dest)
        except Exception as error:
            logger.warning("Impossibile copiare WAL: %s", error)
    else:
        logger.info("File WAL non presente per autobackup (normale se DB chiuso)")

    shm_file = db_file.replace(".db", ".db-shm")
    if os.path.exists(shm_file):
        shm_dest = os.path.join(dest_folder, f"{base_name}.db-shm")
        try:
            shutil.copy2(shm_file, shm_dest)
            logger.info("Backup SHM copiato: %s", shm_dest)
            copied_files.append(shm_dest)
        except Exception as error:
            logger.warning("Impossibile copiare SHM: %s", error)
    else:
        logger.info("File SHM non presente per autobackup (normale se DB chiuso)")

    original_size = os.path.getsize(db_file)
    backup_size = os.path.getsize(dest_path)
    if backup_size < original_size * 0.5:
        logger.error("Backup automatico potenzialmente corrotto: %s vs %s bytes", backup_size, original_size)
        for file_path in copied_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception:
                pass
        return {"copied": False, "reason": "size_check_failed", "backup_size": backup_size, "original_size": original_size}

    total_size = sum(os.path.getsize(path) for path in copied_files if os.path.exists(path))
    return {
        "copied": True,
        "copied_files": copied_files,
        "dest_path": dest_path,
        "backup_size": backup_size,
        "original_size": original_size,
        "total_size": total_size,
    }
