# Plan: Help Button → GitHub Wiki (browser)

## 1) MAPPA IMPATTO

| File | Classe / Funzione | Ruolo |
|---|---|---|
| `ui/window_launchers.py` | `open_help_window(app)` | **PIVOT POINT — l'unico file da modificare** |
| `ui/main_dashboard_builder.py:90` | `build_main_dashboard` | Crea `btn_guida` — NON toccare |
| `dataflow.py:1573` | `MainWindow.open_help_window` | Thin delegante — NON toccare |
| `dataflow.py:96` | import `open_help_window` | NON toccare |
| `utils/i18n_utils.py:93` | `get_current_language()` | Funzione già esistente da usare |
| `ui/help_window.py` | `HelpWindow` | Legacy, lasciare intatta |
| `add_data/guida.txt`, `add_data/guida_en.txt` | — | File locali legacy, lasciare intatti |

---

## 2) FLUSSO ATTUALE

```
btn_guida.click()
  └─ app.open_help_window()               [dataflow.py:1573 — delegante]
       └─ open_help_window(app)            [window_launchers.py:5 — PUNTO DI INTERCETTAZIONE]
            └─ HelpWindow(app.root)        [help_window.py:25 — finestra Toplevel]
                 └─ populate_content()
                      ├─ get_current_language()   → legge config.ini
                      ├─ EN: add_data/guida_en.txt
                      └─ IT: add_data/guida.txt
                           └─ render in tk.Text con TOC e ricerca
```

**Lingua attiva**: letta da `config.ini` → `[Settings] language` tramite `get_current_language()` in `utils/i18n_utils.py`. Valori possibili: `'en'`, `'it'`. Fallback già integrato: ritorna `'en'` se config assente o corrotto.
**Menu Help**: NON esiste. Nessun binding tastiera (`F1` ecc.).

---

## 3) RISCHI E MITIGAZIONI

| Rischio | Causa | Prob. | Impatto | Mitigazione |
|---|---|---|---|---|
| Browser non configurato (Linux headless) | `webbrowser.open()` non trova browser | Bassa | Basso | Caso non supportato — nessuna azione necessaria |
| `get_current_language()` ritorna valore inatteso | config.ini assente/corrotto | Bassa | Basso | La funzione ha già fallback `'en'` integrato; aggiungere `else` che punta a `EN-Home` |
| Import `HelpWindow` rimane inutilizzato | Corpo sostituito, import non rimosso | Nulla | Nulla | Rimuovere l'import nella stessa patch |
| Regressione menu Help | Non esiste menu — rischio assente | — | — | — |
| Cross-platform Win/Linux | `webbrowser` è stdlib cross-platform | Nulla | Nulla | — |

---

## 4) STRATEGIA RACCOMANDATA

**Strategia 2 — Modifica del corpo della funzione esistente** in `ui/window_launchers.py`.

Motivazione: `open_help_window()` è già il punto di delega naturale tra il pulsante e il sistema Help. Sostituire solo il suo corpo (3 righe) in un file isolato:
- lascia intatti tutti gli altri file (pulsante, `MainWindow`, imports in `dataflow.py`)
- è immediatamente reversibile (ripristinare il corpo originale)
- non introduce alcun livello di indirezione aggiuntivo
- la firma della funzione rimane invariata

---

## 5) PIANO STEP-BY-STEP

**STEP A — Aggiornare gli import in `ui/window_launchers.py`**
- Rimuovere: `from ui.help_window import HelpWindow` (diventa inutilizzato)
- Aggiungere: `import webbrowser` (stdlib)
- Aggiungere: `from utils.i18n_utils import get_current_language`

**STEP B — Sostituire il corpo di `open_help_window(app)`** in `ui/window_launchers.py`
- Firma invariata: `def open_help_window(app):`
- Logica nuovo corpo:
  1. Chiamare `get_current_language()` → valore `'it'`, `'en'`, o altro
  2. Mappare `'it'` → `https://github.com/sorguido/dataflow-procurement-software/wiki/IT-Home`
  3. Mappare `'en'` → `https://github.com/sorguido/dataflow-procurement-software/wiki/EN-Home`
  4. Fallback (qualsiasi altro valore) → `https://github.com/sorguido/dataflow-procurement-software/wiki/EN-Home`
  5. Chiamare `webbrowser.open(url)`

**STEP C — Verifica manuale**
- Lingua impostata su IT → click Help → browser apre `IT-Home`
- Lingua impostata su EN → click Help → browser apre `EN-Home`
- Nessun altro comportamento visuale alterato (pulsante invariato, nessuna finestra aperta)

---

## 6) SCOPE PATCH FUTURA

| Categoria | Elemento |
|---|---|
| **Modificato** | `ui/window_launchers.py` — import section + corpo `open_help_window` |
| **NON toccato** | `ui/main_dashboard_builder.py` |
| **NON toccato** | `dataflow.py` (né line 96 né line 1573) |
| **NON toccato** | `utils/i18n_utils.py` |
| **Legacy temporaneo** | `ui/help_window.py` — `HelpWindow` lasciata intatta |
| **Legacy temporaneo** | `add_data/guida.txt`, `add_data/guida_en.txt` — lasciati intatti |
| **Sicuro rimuovere (in questa patch)** | `from ui.help_window import HelpWindow` in `ui/window_launchers.py` |
| **NON rimuovere** | La classe `HelpWindow` stessa in `ui/help_window.py` |
