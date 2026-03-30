# Plan: Step 3 — Window Launchers Extraction — ui/window_launchers.py

## Status: APPROVED (pending user sign-off)

## TL;DR
Estrai 2 metodi launcher puri da `MainWindow` in `ui/window_launchers.py` con delegation-stub pattern (identico a step 1-2). I 4 metodi `open_license_window`, `open_settings_window`, `show_first_run_license`, `open_new_event`, `open_new_request_window` restano in dataflow.py perché complessi o con circular-import.

## Scope confermato dall'utente
- **Estratti**: `open_help_window` (riga 4138), `on_kpi_click` (riga 4141) — naming invariato
- **Non estratti**:
  - `open_license_window`, `open_settings_window` → circular import (LicenseWindow riga 3390, SettingsWindow riga 2529 sono in dataflow.py)
  - `show_first_run_license` → complesso (config I/O, root lifecycle)
  - `open_new_event`, `open_new_request_window` → complessi (DB/VSM/routing)

## Steps

### Step 1 — Crea `ui/window_launchers.py`
```python
from ui.help_window import HelpWindow
from ui.kpi_window import KpiWindow

def open_help_window(app):
    HelpWindow(app.root)

def on_kpi_click(app):
    """Apre la finestra KPI Analysis."""
    KpiWindow(app.root)
```

### Step 2 — Modifica `dataflow.py`
- Aggiungi import dopo gli import esistenti:
  `from ui.window_launchers import open_help_window, on_kpi_click`
- Riga 4138 — sostituisci body `open_help_window`:
  `def open_help_window(self): open_help_window(self)`
- Righe 4141-4143 — sostituisci body `on_kpi_click`:
  `def on_kpi_click(self): on_kpi_click(self)`

### Step 3 — Verifica
- `get_errors` su `ui/window_launchers.py` → 0 errori attesi
- `get_errors` su `dataflow.py` → no nuovi errori vs baseline

## File modificati
- `ui/window_launchers.py` — NUOVO
- `dataflow.py` — 1 import + 2 stub

## Decisions
- on_kpi_click naming invariato (utente ha scelto "No rinomina")
- open_license_window e open_settings_window restano in dataflow.py
- SettingsWindow/LicenseWindow extraction → future step separato
