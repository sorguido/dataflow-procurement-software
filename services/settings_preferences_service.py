"""Settings preferences persistence helpers."""

from __future__ import annotations

import configparser
import os


ALLOWED_CURRENCIES = {"NONE", "EUR", "USD", "GBP", "CHF"}


def load_settings_snapshot(config_file):
    """Load settings snapshot with conservative defaults."""
    config = configparser.ConfigParser(interpolation=None)
    config.read(config_file)

    if "AutoBackup" in config:
        try:
            autobackup_enabled = config["AutoBackup"].getboolean("enabled", False)
            autobackup_hour = config["AutoBackup"].get("hour", "12")
            autobackup_path = config["AutoBackup"].get("path", "")
        except Exception:
            autobackup_enabled = False
            autobackup_hour = "12"
            autobackup_path = ""
    else:
        autobackup_enabled = False
        autobackup_hour = "12"
        autobackup_path = ""

    if "Settings" in config:
        language_code = config.get("Settings", "language", fallback="en")
        if language_code not in {"en", "it"}:
            language_code = "en"
        currency_code = config.get("Settings", "currency_code", fallback="NONE").strip().upper()
        if currency_code not in ALLOWED_CURRENCIES:
            currency_code = "NONE"
    else:
        language_code = "en"
        currency_code = "NONE"

    return {
        "autobackup_enabled": autobackup_enabled,
        "autobackup_hour": autobackup_hour,
        "autobackup_path": autobackup_path,
        "language_code": language_code,
        "currency_code": currency_code,
    }


def save_language_preference(config_file, selected_language_label):
    """Persist selected language label ('English'|'Italiano') into config."""
    if not selected_language_label:
        raise ValueError("missing_language")

    lang_code = "en" if selected_language_label == "English" else "it"

    config = configparser.ConfigParser(interpolation=None)
    if os.path.exists(config_file):
        config.read(config_file)
    if "Settings" not in config:
        config["Settings"] = {}
    config["Settings"]["language"] = lang_code

    with open(config_file, "w", encoding="utf-8") as file_obj:
        config.write(file_obj)

    return lang_code


def save_currency_preference(config_file, selected_currency_ui, none_label):
    """Persist selected currency with NONE fallback/validation."""
    selected_currency_ui = (selected_currency_ui or "").strip()
    selected_currency = "NONE" if selected_currency_ui in {none_label, "NONE"} else selected_currency_ui.upper()
    if selected_currency not in ALLOWED_CURRENCIES:
        selected_currency = "NONE"

    config = configparser.ConfigParser(interpolation=None)
    if os.path.exists(config_file):
        config.read(config_file, encoding="utf-8")
    if "Settings" not in config:
        config["Settings"] = {}
    config["Settings"]["currency_code"] = selected_currency

    with open(config_file, "w", encoding="utf-8") as file_obj:
        config.write(file_obj)

    return selected_currency


def save_autobackup_preferences(config_file, *, enabled, hour, path):
    """Persist autobackup config; require destination path when enabled."""
    if enabled and not path:
        raise ValueError("missing_autobackup_path")

    config = configparser.ConfigParser(interpolation=None)
    config.read(config_file)
    if "AutoBackup" not in config:
        config["AutoBackup"] = {}
    config["AutoBackup"]["enabled"] = str(bool(enabled))
    config["AutoBackup"]["hour"] = str(hour)
    config["AutoBackup"]["path"] = path

    with open(config_file, "w", encoding="utf-8") as file_obj:
        config.write(file_obj)
