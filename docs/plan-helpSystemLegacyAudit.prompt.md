## REPORT — Analisi Statica Sistema Help Legacy

---

### 1. ELENCO FILE COINVOLTI

| File | Ruolo |
|---|---|
| `ui/help_window.py` | Il vecchio sistema Help completo — classe Tkinter con ~800 righe |
| `ui/window_launchers.py` | Il nuovo sistema Help attivo — apre la Wiki nel browser |
| `dataflow.py` | Import sia del legacy (HelpWindow) sia del nuovo (open_help_window); delega al nuovo |
| `ui/main_dashboard_builder.py` | Il pulsante "❓ Guida" che attiva il flusso Help |

---

### 2. ELENCO ELEMENTI LEGACY

#### 2a. CLASSE

| Tipo | Nome | File | Linea | Stato |
|---|---|---|---|---|
| Classe | `HelpWindow` | `ui/help_window.py` | 27 | **ORFANA** |

**Motivazione:** `HelpWindow(tk.Toplevel)` è definita ma non viene istanziata da nessuna parte nel progetto. Il file è importato in `dataflow.py` (L94) ma il simbolo `HelpWindow` non viene mai usato.

#### 2b. METODI DI HelpWindow (tutti ORFANI per transitività)

| Metodo | Linea | Stato |
|---|---|---|
| `__init__` | L28 | ORFANO |
| `populate_content` | L251 | ORFANO |
| `_parse_and_insert_content` | L330 | ORFANO |
| `_insert_formatted_line` | L396 | ORFANO |
| `_insert_formatted_line_with_anchor` | L519 | ORFANO |
| `setup_search_functionality` | L645 | ORFANO |
| `search_text` | L682 | ORFANO |
| `search_next` | L744 | ORFANO |
| `update_search_counter` | L761 | ORFANO |
| `clear_search` | L789 | ORFANO |

#### 2c. IMPORT

| Tipo | Nome | File | Linea | Stato |
|---|---|---|---|---|
| Import | `from ui.help_window import HelpWindow` | `dataflow.py` | 94 | **ORFANO** |

**Motivazione:** `HelpWindow` viene importata ma il simbolo non compare in nessun altro punto di `dataflow.py`. Il modulo viene caricato inutilmente a ogni avvio dell'applicazione.

#### 2d. CHIAMATE `webbrowser.open` NON RAGGIUNGIBILI

| File | Linea | Contesto | Stato |
|---|---|---|---|
| `ui/help_window.py` | ~506 | Tag bind per link `[LINK:url]` in `guida.txt` | **ORFANO** |
| `ui/help_window.py` | ~583 | Variante con anchor in `_insert_formatted_line_with_anchor` | **ORFANO** |

**Motivazione:** Queste righe sono metodi interni di `HelpWindow`, che non viene mai istanziata.

#### 2e. FILE DI CONTENUTO LEGACY

| Tipo | Path | Stato |
|---|---|---|
| File testo | `add_data/guida.txt` | **ASSENTE** — non esiste nel filesystem |
| File testo | `add_data/guida_en.txt` | **ASSENTE** — non esiste nel filesystem |

**Motivazione:** Il codice in `ui/help_window.py` (L300–302) tenta di caricarli tramite `resource_path()`, ma entrambi i file sono stati rimossi dal repository. Anche se `HelpWindow` venisse istanziata, fallirebbe con `FileNotFoundError` (gestita in modo silenzioso con uno stato di errore mostrato nella finestra).

---

### 3. MAPPATURA FLUSSO ATTUALE HELP

```
Utente clicca "❓ Guida"
    ↓
ui/main_dashboard_builder.py:90
    command=app.open_help_window
    ↓
dataflow.py:1574
    def open_help_window(self): open_help_window(self)
      ↓ (chiama la funzione importata da window_launchers)
ui/window_launchers.py:13
    def open_help_window(app):
        lang = get_current_language()   # → "it" o "en"
        url = _WIKI_URLS.get(lang, _WIKI_FALLBACK)
        webbrowser.open(url)            # → apre browser del SO
    ↓
Browser del sistema operativo apre:
    IT → https://github.com/sorguido/dataflow-procurement-software/wiki/IT-Home
    EN → https://github.com/sorguido/dataflow-procurement-software/wiki/EN-Home
```

`HelpWindow` **non compare in nessun punto di questo flusso**.

---

### 4. RISCHI

| Elemento | Rischio | Impatto |
|---|---|---|
| Import di `HelpWindow` in `dataflow.py:94` | Il modulo `ui/help_window.py` (~800 righe) viene importato e compilato a ogni avvio, occupando memoria inutilmente e allungando il tempo di avvio | **BASSO** |
| File `guida.txt` / `guida_en.txt` assenti | Se per qualsiasi motivo `HelpWindow` venisse istanziata (es. regressione di codice, copia errata da vecchio branch), si otterrebbe un errore runtime silenzioso con finestra parziale | **BASSO** |
| Codice non raggiungibile in `help_window.py` | Nessun rischio funzionale attivo, ma crea confusione manutentiva: chi legge il codice potrebbe credere che il vecchio sistema sia ancora in uso | **BASSO** |
| Presenza di `print()` di debug in `HelpWindow.populate_content` (L305–318) | Se il codice fosse mai riattivato, stamperebbe path interni e stato del filesystem sulla console/log | **BASSO** |

Nessun rischio di sicurezza o integrità dati identificato.

---

### 5. SUGGERIMENTI (SOLO ANALISI)

#### Teoricamente eliminabile (sicuro):
- L'intero file `ui/help_window.py` (~800 righe) — non è istanziato, i file di contenuto non esistono
- La riga `from ui.help_window import HelpWindow` in `dataflow.py:94` — import inutilizzato

#### Da NON toccare per sicurezza:
- `ui/window_launchers.py` — contiene il sistema Help **attivo**
- `dataflow.py:96` — import di `open_help_window` da `window_launchers` — **necessario** per il funzionamento
- `dataflow.py:1574` — metodo delegante `open_help_window(self)` — è il punto di collegamento UI → logica
- `ui/main_dashboard_builder.py:90` — il pulsante "❓ Guida" — **attivo**

---

**ANALISI COMPLETATA – NESSUNA MODIFICA APPLICATA**
