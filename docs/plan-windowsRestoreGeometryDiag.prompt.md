# Plan: Windows Restore Geometry — Diagnostic Logging

**TL;DR**: La root cause più probabile è che il check `root.state() == 'zoomed'` alla L4502
**sia sempre `False`** — perché la finestra è ancora nello stato `"withdrawn"` quando
`state("zoomed")` viene chiamato dentro `MainWindow.__init__`. Il fix precedente cade nel
ramo `else`, chiama `calculate_center_position` (che fa `update()` + legge
`winfo_reqwidth/Height` = dimensioni minime widget) e scrive quella geometry minuscola come
restore position di Windows (`WINDOWPLACEMENT.rcNormalPosition`). Poi `deiconify()` mostra
correttamente la finestra massimizzata, ma il danno al restore size è già fatto.

---

## Sequenza annotata (Windows)

| Step | Dove | Stato successivo |
|------|------|-----------------|
| `root.withdraw()` | dataflow.py L4381 | state = `"withdrawn"` |
| `attributes("-zoomed", True)` | dataflow.py L1018 | → eccezione su Windows |
| `state("zoomed")` | dataflow.py L1022 | zoomed pending, ma root è withdrawn |
| `if root.state() == 'zoomed':` | dataflow.py L4502 | **probabilmente False → ramo else** |
| `calculate_center_position(root)` → `update()` + `winfo_reqwidth/Height` | dataflow.py L4511 | geometry = dimensione minima widget |
| `root.geometry(tiny_geom)` | dataflow.py L4512 | scrive il restore piccolo in WINDOWPLACEMENT |
| `root.deiconify()` | dataflow.py L4513 | finestra appare massimizzata (OK visivo) |
| **drag dalla title bar** | OS Windows | ripristina a WINDOWPLACEMENT.rcNormalPosition = tiny → bug |

---

## Piano diagnostico: 5 punti di logging, solo Windows, tutti reversibili

### Punto A — subito dopo `state("zoomed")` a L1022 (dentro `MainWindow.__init__`)

```python
if sys.platform == 'win32':
    logger.info("[DIAG-A] state=%s geometry=%s winfo=%dx%d req=%dx%d",
        self.root.state(), self.root.geometry(),
        self.root.winfo_width(), self.root.winfo_height(),
        self.root.winfo_reqwidth(), self.root.winfo_reqheight())
```

### Punto B — appena prima dell'`if` a L4502 (dentro `main_task()`)

```python
if sys.platform == 'win32':
    logger.info("[DIAG-B] PRE-CHECK state=%s geometry=%s", root.state(), root.geometry())
```

### Punto C — dopo la `geometry()` in **entrambi** i rami (L4509 e L4512)

```python
if sys.platform == 'win32':
    logger.info("[DIAG-C] POST-GEOM state=%s geometry=%s winfo=%dx%d req=%dx%d",
        root.state(), root.geometry(),
        root.winfo_width(), root.winfo_height(),
        root.winfo_reqwidth(), root.winfo_reqheight())
```

### Punto D — dopo `root.deiconify()` a L4513

```python
if sys.platform == 'win32':
    logger.info("[DIAG-D] POST-DEIC state=%s geometry=%s", root.state(), root.geometry())
```

### Punto E — dentro `remove_topmost()` a L4526, prima di `attributes('-topmost', False)`

```python
if sys.platform == 'win32':
    logger.info("[DIAG-E] POST-TOPMOST state=%s geometry=%s", root.state(), root.geometry())
```

---

## Cosa ci dicono i log

- **Ipotesi confermata**: DIAG-B mostra `state='withdrawn'` → condizione False → DIAG-C
  mostra geometry piccola (es. `800x600` o simile).
  **Fix**: non usare `state()` per il check, oppure impostare la geometry su root *prima*
  di chiamare `state("zoomed")`.

- **Ipotesi smentita**: DIAG-B mostra `state='zoomed'` → condizione True → DIAG-C mostra
  geometry grande, ma il bug persiste. In quel caso la root cause è altrove (es.
  `geometry()` chiamata su finestra già zoomed non aggiorna WINDOWPLACEMENT su Windows).

---

## File coinvolti

- `dataflow.py` — 5 inserimenti diagnostici, tutti dentro `if sys.platform == 'win32':`

## Garanzia Linux

Ogni riga di log è dentro `if sys.platform == 'win32':` — zero impatto su Linux.
`calculate_center_position` e il ramo `else` rimangono invariati.

## Vincoli rispettati

- Nessuna modifica al comportamento Linux
- Nessuna modifica permanente
- Nessun refactor
- Nessuna fix cosmetica
- Nessun `after()` workaround
- Nessuna doppia geometry "tentativa"
