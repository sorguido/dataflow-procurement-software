# Plan: Help Window — Massimizzata + Larghezza TOC Dinamica

**TL;DR**: Tre piccole aggiunte localizzate in `ui/help_window.py`: 1) due import in cima, 2) calcolo della larghezza TOC + scheduling di `sashpos`, 3) finestra massimizzata con bordi standard dopo `center_window`.

---

## Steps

### Phase 1 – Import

1. Aggiungere `import sys` vicino agli altri import standard (riga ~9 circa, dopo `import os`)
2. Aggiungere `from tkinter import font as tkfont` accanto agli import tkinter esistenti (riga ~10)

### Phase 2 – Larghezza dinamica sommario

3. Dopo il ciclo `for text, tag in self.topics:` che crea i link del sommario (~riga 195), aggiungere:
   ```python
   _toc_font = tkfont.nametofont("TkDefaultFont")
   _toc_width = max(_toc_font.measure(t) for t, _ in self.topics) + 45
   self.after(50, lambda: paned.sashpos(0, _toc_width))
   ```
   - `nametofont("TkDefaultFont")` è stabile e compatibile cross-platform
   - Il `+45` copre il padding di `toc_outer_frame` (padding=10 × 2 lati) + margine scrollbar
   - `sashpos(0, ...)` è API supportata di `ttk.PanedWindow`; viene schedulata a 50ms per garantire che il layout del pannello sia completato prima che venga applicata

### Phase 3 – Finestra massimizzata con bordi standard

4. Subito dopo `center_window(self)` (~riga 227) e prima di `self.after(100, self.populate_content)`, aggiungere:
   ```python
   if sys.platform == "win32":
       self.state("zoomed")
   else:
       self.attributes("-zoomed", True)
   ```
   - `state("zoomed")` per Windows (decorazioni standard, title bar intatta)
   - `attributes("-zoomed", True)` per Linux/X11 — stessa semantica, stessi controlli di sistema
   - Nessun `overrideredirect`, nessun `fullscreen`, nessuna perdita di pulsanti di sistema

---

## File toccati

- `ui/help_window.py` — unico file modificato

---

## Verifica

1. Cliccare Help → la finestra si apre maximizzata con title bar e pulsanti minimizza/massimizza/chiudi visibili
2. Il sommario a sinistra ha una larghezza calcolata sul testo più lungo (es. "4. Value Stream Mapping: Saving, Cost Avoidance e Derisking")
3. Scroll del sommario: muovere la rotella sul pannello sinistro → solo il TOC scrolla
4. Scroll del testo: muovere la rotella sul pannello destro → solo il testo scrolla
5. Nessuna eccezione in console
6. Nessun'altra finestra dell'app risulta modificata

---

## Decisioni

- `sashpos` è un'API `ttk.PanedWindow` ufficialmente supportata → usata come da specifica ("SOLO se già supportato")
- Scroll indipendenti: già corretti nel codice attuale (bind solo su `toc_canvas`/`toc_frame`/label, non `bind_all`). Nessuna modifica necessaria
- `window_launchers.py`: nessuna modifica necessaria — il launcher crea solo `HelpWindow(app.root)` e non gestisce posizione/stato
