from ui.help_window import HelpWindow
from ui.kpi_window import KpiWindow


def open_help_window(app):
    HelpWindow(app.root)


def on_kpi_click(app):
    """Apre la finestra KPI Analysis."""
    KpiWindow(app.root)
