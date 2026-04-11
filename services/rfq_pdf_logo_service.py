"""Servizi per gestione logo aziendale persistente usato nell'export RFQ PDF."""

import configparser
import os
import shutil
from typing import Optional, Tuple

from PIL import Image, UnidentifiedImageError

from services.app_paths import get_user_documents_dataflow_dir
from utils.user_utils import get_app_data_dir, get_config_file

_ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg"}
_MAX_LOGO_FILE_SIZE_BYTES = 8 * 1024 * 1024
_CONFIG_SECTION = "Settings"
_CONFIG_KEY = "rfq_pdf_logo_file"


class LogoValidationError(ValueError):
    """Errore validazione logo."""


def _get_logo_storage_dir() -> str:
    """Restituisce la directory interna dove salvare il logo persistente."""
    base_dir = get_user_documents_dataflow_dir()
    if not base_dir:
        base_dir = get_app_data_dir()

    logo_dir = os.path.join(base_dir, "Assets", "RFQ_PDF")
    os.makedirs(logo_dir, exist_ok=True)
    return logo_dir


def _load_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser(interpolation=None)
    config_path = get_config_file()
    if os.path.exists(config_path):
        config.read(config_path, encoding="utf-8")
    return config


def _save_config(config: configparser.ConfigParser) -> None:
    config_path = get_config_file()
    with open(config_path, "w", encoding="utf-8") as config_file:
        config.write(config_file)


def _get_configured_logo_filename() -> Optional[str]:
    config = _load_config()
    if not config.has_section(_CONFIG_SECTION):
        return None
    value = config.get(_CONFIG_SECTION, _CONFIG_KEY, fallback="").strip()
    return value or None


def _set_configured_logo_filename(filename: Optional[str]) -> None:
    config = _load_config()
    if not config.has_section(_CONFIG_SECTION):
        config.add_section(_CONFIG_SECTION)

    if filename:
        config.set(_CONFIG_SECTION, _CONFIG_KEY, filename)
    elif config.has_option(_CONFIG_SECTION, _CONFIG_KEY):
        config.remove_option(_CONFIG_SECTION, _CONFIG_KEY)

    _save_config(config)


def _validate_logo_image(source_path: str) -> Tuple[int, int]:
    """Valida estensione/file logo e ritorna dimensioni (width, height)."""
    if not source_path:
        raise LogoValidationError("Percorso logo non valido.")

    if not os.path.isfile(source_path):
        raise LogoValidationError("Il file selezionato non esiste.")

    extension = os.path.splitext(source_path)[1].lower()
    if extension not in _ALLOWED_EXTENSIONS:
        raise LogoValidationError("Formato non supportato. Usa PNG o JPG/JPEG.")

    file_size = os.path.getsize(source_path)
    if file_size <= 0:
        raise LogoValidationError("Il file selezionato e' vuoto.")
    if file_size > _MAX_LOGO_FILE_SIZE_BYTES:
        raise LogoValidationError("File troppo grande. Dimensione massima consigliata: 8 MB.")

    try:
        with Image.open(source_path) as img:
            img.verify()
        with Image.open(source_path) as img:
            width, height = img.size
    except (UnidentifiedImageError, OSError) as exc:
        raise LogoValidationError(f"File immagine non valido o corrotto: {exc}") from exc

    if width < 50 or height < 20:
        raise LogoValidationError("Immagine troppo piccola. Usa almeno 50x20 px.")

    return width, height


def get_persisted_logo_path() -> Optional[str]:
    """Restituisce il percorso assoluto del logo persistito se disponibile e valido."""
    filename = _get_configured_logo_filename()
    if not filename:
        return None

    logo_path = os.path.join(_get_logo_storage_dir(), filename)
    if not os.path.isfile(logo_path):
        return None

    try:
        _validate_logo_image(logo_path)
    except LogoValidationError:
        return None

    return logo_path


def save_logo_from_source(source_path: str) -> str:
    """Valida e copia il logo selezionato nella directory interna DataFlow."""
    _validate_logo_image(source_path)

    storage_dir = _get_logo_storage_dir()
    source_ext = os.path.splitext(source_path)[1].lower()
    target_name = f"company_logo{source_ext}"
    target_path = os.path.join(storage_dir, target_name)

    # Rimuove eventuali logo precedenti con estensioni diverse.
    for existing in os.listdir(storage_dir):
        if existing.startswith("company_logo"):
            try:
                os.remove(os.path.join(storage_dir, existing))
            except OSError:
                pass

    shutil.copy2(source_path, target_path)

    # Ricontrollo finale del file copiato
    _validate_logo_image(target_path)
    _set_configured_logo_filename(target_name)
    return target_path


def remove_persisted_logo() -> None:
    """Rimuove il logo persistito e pulisce il riferimento in config."""
    filename = _get_configured_logo_filename()
    if filename:
        logo_path = os.path.join(_get_logo_storage_dir(), filename)
        if os.path.exists(logo_path):
            try:
                os.remove(logo_path)
            except OSError:
                pass

    _set_configured_logo_filename(None)


def get_logo_status() -> dict:
    """Ritorna stato logo corrente per UI dialog."""
    filename = _get_configured_logo_filename()
    if not filename:
        return {"configured": False, "available": False, "path": None, "filename": None}

    logo_path = os.path.join(_get_logo_storage_dir(), filename)
    if not os.path.isfile(logo_path):
        return {"configured": True, "available": False, "path": None, "filename": filename}

    try:
        _validate_logo_image(logo_path)
    except LogoValidationError:
        return {"configured": True, "available": False, "path": None, "filename": filename}

    return {
        "configured": True,
        "available": True,
        "path": logo_path,
        "filename": os.path.basename(logo_path),
    }
