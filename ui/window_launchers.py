import webbrowser

from ui.kpi_window import KpiWindow
from utils.i18n_utils import get_current_language

_WIKI_URLS = {
    "it": "https://github.com/sorguido/dataflow-procurement-software/wiki/IT-Home",
    "en": "https://github.com/sorguido/dataflow-procurement-software/wiki/EN-Home",
}
_WIKI_FALLBACK = "https://github.com/sorguido/dataflow-procurement-software/wiki"

_LICENSE_URL = "https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE"


def open_license_window(app):
    webbrowser.open(_LICENSE_URL)


def open_help_window(app):
    lang = get_current_language()
    url = _WIKI_URLS.get(lang, _WIKI_FALLBACK)
    webbrowser.open(url)


def on_kpi_click(app):
    """Apre la finestra KPI Analysis."""
    KpiWindow(app.root)
