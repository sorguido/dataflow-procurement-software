## ANALISI COERENZA PYINSTALLER SPEC

> **Nota preliminare**: I due file spec (`dataflow.spec` e `Tools per build WIN/Creare EXE/DataFlow.spec`) hanno **`hiddenimports` identici**. La sola differenza è `manifest='app.manifest.xml'` presente nel file di Tools ma assente in `dataflow.spec`. L'analisi seguente vale quindi per entrambi.

---

### 1. Moduli nel `.spec` NON presenti nel codebase

| Modulo | File fisico | Import rilevati | Conclusione |
|---|---|---|---|
| `ui.help_window` | **NON ESISTE** — `ui/help_window.py` assente dal filesystem | Nessuno — ricerca su tutto il codebase: 0 risultati | **ORFANO** |

Tutti gli altri moduli custom (`database.*`, `services.*`, `ui.dialogs.*`, `ui.windows.*`, `utils.*`) hanno file fisici corrispondenti. Nessun altro orfano rilevato.

---

### 2. Moduli presenti ma NON utilizzati / con anomalia

| Modulo | Stato file | Note |
|---|---|---|
| `ui` | **Nessun `__init__.py`** nella directory `ui/` | `ui/` è trattata come *namespace package* (Python 3.3+). `ui/__init__.py` non esiste. Inserire `'ui'` in hiddenimports è tecnicamente irrilevante per un namespace package. Non è un orfano, ma è una voce ambigua. |

**Valutazione**: La voce `'ui'` nel `.spec` non causa danni ma non ha effetto reale, dato che PyInstaller gestisce i namespace package tramite i moduli figli che lo referenziano. Candidato a revisione, ma rischio zero di problemi se lasciato.

---

### 3. Verifica specifica Help e License

#### `ui.help_window`
- **file**: `ui/help_window.py` — **NON ESISTE** sul filesystem
- **import**: **NO** — nessuna occorrenza di `import ui.help_window` o `from ui.help_window` in tutto il codebase
- **utilizzo runtime**: **NO** — la funzionalità è stata migrata in `ui/window_launchers.py`: `open_help_window()` chiama `webbrowser.open(url)` direttamente, senza finestra Tk dedicata
- **azione consigliata**: **RIMUOVERE** dal `.spec` — modulo orfano, il file non esiste, nessun import, nessun utilizzo runtime. Rischio rimozione: **zero**

#### `ui.license_window`
- **file**: `ui/license_window.py` — **NON ESISTE** sul filesystem
- **import**: **NO** — nessuna occorrenza in tutto il codebase
- **utilizzo runtime**: **NO** — la funzionalità è in `ui/window_launchers.py`: `open_license_window()` chiama `webbrowser.open(_LICENSE_URL)` direttamente
- **presenza nel `.spec`**: **NON È nel `.spec`** — `ui.license_window` non compare in nessuno dei due `hiddenimports`. Non è necessario aggiungerlo. Nessuna azione richiesta.

---

### 4. Conclusione operativa

**Moduli da rimuovere dal `.spec` (entrambi i file):**

| Voce | Motivo |
|---|---|
| `'ui.help_window'` | File non esiste, nessun import, funzionalità migrata a `ui.window_launchers` |

**Moduli da mantenere:** tutti gli altri 27 hiddenimports — ogni file fisico corrispondente è confermato presente sul filesystem.

**Voce ambigua (da non rimuovere, ma da segnalare):**
- `'ui'` — non ha `__init__.py`, è namespace package. Tecnicamente ininfluente, ma non causa danni.

**Differenza tra i due `.spec`:**
- `Tools per build WIN/.../DataFlow.spec` ha `manifest='app.manifest.xml'` nell'EXE, `dataflow.spec` no. Non è legato agli hiddenimports, ma è una **divergenza strutturale** tra i due file da verificare intenzionalmente.

---

**Livello rischio pulizia**: **BASSO**
L'unica modifica necessaria è rimuovere `'ui.help_window'` da entrambi i file `.spec`. Il modulo non esiste sul disco, non viene importato, non è usato a runtime. PyInstaller non può includerlo (non c'è nulla da includere), quindi la voce è già inerte — la sua rimozione è puramente di manutenzione.
