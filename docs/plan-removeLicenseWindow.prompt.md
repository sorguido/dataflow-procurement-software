# ANALISI FORENSE — RIMOZIONE FINESTRA LICENSE

---

## 1. Mappa componenti coinvolti

### Componente 1 — Classe `LicenseWindow`

| Attributo | Valore |
|-----------|--------|
| **File** | `ui/license_window.py` |
| **Simbolo** | `class LicenseWindow(tk.Toplevel)` — riga 13 |
| **Ruolo** | Finestra modale Toplevel per visualizzare il testo della licenza GNU GPLv3 |
| **Dipendenze dirette** | `tkinter`, `tkinter.ttk`, `webbrowser`, `utils.resource_utils.set_window_icon`, `utils.window_utils.center_window`, `utils.i18n_utils._` |

Metodi interni rilevanti:
- `__init__(self, parent, first_run=False)` — riga 14: costruisce la UI, modalità dual-state (`first_run=True` mostra pulsanti Accetto/Esci; `first_run=False` mostra solo Chiudi)
- `on_accept(self)` — riga 53: imposta `self.accepted = True` e fa `destroy()`
- `on_exit(self)` — riga 56: imposta `self.accepted = False` e fa `destroy()`
- `_populate_content(self)` — riga 59: inietta il testo della licenza in un widget `tk.Text` attraverso tag formattati e un link cliccabile a LinkedIn via `webbrowser.open`

---

### Componente 2 — Import in `dataflow.py`

| Attributo | Valore |
|-----------|--------|
| **File** | `dataflow.py` riga 116 |
| **Simbolo** | `from ui.license_window import LicenseWindow` |
| **Ruolo** | Rende la classe disponibile a tutto il modulo principale |

---

### Componente 3 — Metodo `open_license_window` in `MainWindow`

| Attributo | Valore |
|-----------|--------|
| **File** | `dataflow.py` righe 1080–1088 |
| **Simbolo** | `MainWindow.open_license_window()` |
| **Ruolo** | Apre `LicenseWindow(self.root, first_run=False)` in sola lettura; rimuove temporaneamente l'attributo `-topmost` della finestra principale |
| **Chiamato da** | `app.btn_license` via `command=app.open_license_window` |

---

### Componente 4 — Metodo `show_first_run_license` in `MainWindow` ⚠️ DEAD CODE

| Attributo | Valore |
|-----------|--------|
| **File** | `dataflow.py` righe 1090–1129 |
| **Simbolo** | `MainWindow.show_first_run_license()` |
| **Ruolo** | Istanzia `LicenseWindow(self.root, first_run=True)`, attende risposta, salva configurazione |
| **Chiamato da** | **Nessuno** — non invocato in alcun punto del codebase. È dead code. |

---

### Componente 5 — Flusso first-run in `main_task()`

| Attributo | Valore |
|-----------|--------|
| **File** | `dataflow.py` righe 4396–4428 |
| **Simbolo** | `main_task()` |
| **Ruolo** | Verifica se `config['Settings']['license_accepted']` è `True`; se no, istanzia direttamente `LicenseWindow(root, first_run=True)`, attende risposta utente, salva configurazione |
| **Dipendenza critica** | Questo flusso è **indipendente** da `open_license_window()` e da `show_first_run_license()`. Usa `LicenseWindow` direttamente. |

---

### Componente 6 — Pulsante License nella toolbar

| Attributo | Valore |
|-----------|--------|
| **File** | `ui/main_dashboard_builder.py` riga 92 |
| **Simbolo** | `app.btn_license` |
| **Ruolo** | `ttk.Button(frame_top, text=_("≡ License"), command=app.open_license_window)` |
| **Pack** | `side="right", padx=(0, 10)` — riga 93 |

---

### Componente 7 — Pattern Help (riferimento per la sostituzione)

| Attributo | Valore |
|-----------|--------|
| **File** | `ui/window_launchers.py` righe 13–16 |
| **Simbolo** | `open_help_window(app)` |
| **Ruolo** | Rileva lingua corrente con `get_current_language()`, seleziona URL, chiama `webbrowser.open(url)` |
| **Delegante** | `MainWindow.open_help_window(self)` — `dataflow.py` riga 1573 — thin wrapper di una riga |

---

### Componente 8 — Specifiche build PyInstaller

| Attributo | Valore |
|-----------|--------|
| **File** | `dataflow.spec` riga 105 |
| **Simbolo** | `'ui.license_window'` nella lista `hiddenimports` |
| **Secondo file** | `Tools per build WIN/Creare EXE/DataFlow.spec` riga 105 — identica entry |

---

### Componente 9 — Stringhe i18n legate alla finestra License

Le seguenti chiavi nei file `.po` appartengono **esclusivamente** a `LicenseWindow`:

- `locale/it/LC_MESSAGES/dataflow.po` righe 904–905: `"Licenza d'Uso - DataFlow Procurement Software"` (titolo finestra)
- `locale/it/LC_MESSAGES/dataflow.po` righe 1695–1770: blocco multi-riga del testo della licenza
- `locale/en/LC_MESSAGES/dataflow.po` righe 1092–2889: corrispondenti chiavi EN
- Entrambi i file `.po` contengono `"≡ License"` — questa chiave **rimane necessaria** perché usata dal pulsante `btn_license`, che sopravvive alla rimozione della finestra.

---

## 2. Catena attuale di apertura della License

**Flusso click utente (uso ordinario):**

```
Utente clicca "≡ License"
  → ui/main_dashboard_builder.py:92
    app.btn_license (command=app.open_license_window)
      → dataflow.py:1080
        MainWindow.open_license_window()
          → self.root.attributes('-topmost', False)
          → LicenseWindow(self.root, first_run=False)
             → ui/license_window.py:13
               LicenseWindow.__init__(parent=root, first_run=False)
                 → mostra finestra Toplevel modale con testo licenza
                 → pulsante "❌ Chiudi" → self.destroy()
```

**Flusso first-run al lancio dell'applicazione (SEPARATO):**

```
dataflow.py:main_task()
  → legge config['Settings']['license_accepted']
  → se False: LicenseWindow(root, first_run=True)  [dataflow.py:4407]
    → attende wait_window
    → se accepted=True: salva config, prosegue
    → se accepted=False: root.destroy(), esce
```

---

## 3. Elementi che diventerebbero rimovibili

### File eliminabili

| File | Condizione |
|------|------------|
| `ui/license_window.py` | Eliminabile **solo se** il flusso first-run viene rimosso o sostituito — vedi ambiguità critica §5 |

### Import eliminabili

| File | Riga | Import |
|------|------|--------|
| `dataflow.py` | 116 | `from ui.license_window import LicenseWindow` |

### Metodi eliminabili

| File | Simbolo | Note |
|------|---------|-------|
| `dataflow.py` | `MainWindow.open_license_window()` righe 1080–1088 | Da rimuovere — sarà sostituito con delegante a `open_license_window` in `window_launchers.py` |
| `dataflow.py` | `MainWindow.show_first_run_license()` righe 1090–1129 | Dead code — eliminabile incondizionatamente |

### Attributi eliminabili

Nessun attributo di istanza `self.*` di `MainWindow` è legato esclusivamente alla finestra License.

### Stringhe/risorse eliminabili nei file `.po`

| File | Righe | Chiave | Eliminabile? |
|------|-------|--------|--------------|
| `locale/it/LC_MESSAGES/dataflow.po` | 904–905 | `"Licenza d'Uso - DataFlow Procurement Software"` (titolo finestra) | Sì |
| `locale/it/LC_MESSAGES/dataflow.po` | 1695–1770 | Blocco testo licenza (multi-riga) | Sì |
| `locale/en/LC_MESSAGES/dataflow.po` | corrispondenti | Stesse chiavi in EN | Sì |
| Entrambi i `.po` | — | `"≡ License"` | **No — rimane**, è il testo del pulsante |

### Stringhe del flusso first-run (PUNTO DUBBIO)

- `dataflow.py:4401` — `'license_accepted'` in `config.getboolean()`
- `dataflow.py:4420` — `config['Settings']['license_accepted'] = 'True'`
- Le entry i18n legate ai messaggi di errore di salvataggio config (`"Impossibile salvare l'impostazione della licenza: {}"`) — **eliminabili solo se** il flusso first-run viene rimosso con la finestra.

### Specifiche build eliminabili

| File | Riga | Entry |
|------|------|-------|
| `dataflow.spec` | 105 | `'ui.license_window'` in `hiddenimports` |
| `Tools per build WIN/Creare EXE/DataFlow.spec` | 105 | idem |

---

## 4. Implicazioni della sostituzione con link esterno

### Dove interviene logicamente il cambio

La sostituzione riguarda il solo metodo `MainWindow.open_license_window()` in `dataflow.py` righe 1080–1088. Il pulsante `btn_license` **rimane identico** (stessa posizione, stesso testo, stesso pack) e continua a chiamare `app.open_license_window`. Cambia solo il corpo del metodo.

### Pattern Help riutilizzabile

Esiste già in `ui/window_launchers.py` il pattern esatto da replicare:
- `open_help_window(app)` usa `webbrowser.open(url)` — 3 righe di corpo
- Il metodo delegante `MainWindow.open_help_window(self)` in `dataflow.py` riga 1573 è un thin wrapper di una riga: `def open_help_window(self): open_help_window(self)`

La stessa architettura a due livelli (delegante in `MainWindow` + funzione in `window_launchers.py`) può essere replicata identicamente per License.

### Differenze strutturali tra i due casi

| Aspetto | Help | License |
|---------|------|---------|
| URL variabile per lingua | Sì — 2 URL (IT/EN) + fallback | No — URL fisso unico |
| `get_current_language()` necessaria | Sì | No |
| `app` parameter usato nel body | Sì (per lingua) | No (URL statico) |
| Flusso first-run separato | Nessuno | Sì, in `main_task()` — **caso aggiuntivo** |

### Il pulsante può restare identico lato UI?

Sì, confermato. Riga 92 di `ui/main_dashboard_builder.py`:
- `text=_("≡ License")` — invariato
- `command=app.open_license_window` — invariato (cambia solo il corpo del metodo)
- `pack(side="right", padx=(0, 10))` — invariato

### Modifiche necessarie ad altri punti del codice

Oltre al replacement del corpo di `open_license_window()`, sono necessarie:
1. `dataflow.py` riga 116 — rimozione dell'import `LicenseWindow`
2. `dataflow.py` righe 4407–4428 — il flusso first-run usa ancora `LicenseWindow` — **Ambiguità critica, vedi §5**
3. `dataflow.spec` e corrispondente in `Tools per build WIN/` — rimozione di `'ui.license_window'` da `hiddenimports`

---

## 5. Analisi rischi

### RISCHIO 1 — Flusso first-run diventa broken ⚠️ ALTO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | `main_task()` righe 4407–4428 istanzia direttamente `LicenseWindow(root, first_run=True)`. Se il file viene eliminato e l'import rimosso, questa porzione di codice solleva `NameError` al lancio. |
| **Impatto** | L'applicazione non si avvia per gli utenti che non hanno mai accettato la licenza (nessun `license_accepted=True` in `config.ini`). Crash immediato. |
| **Come verificarlo** | Dopo l'implementazione: cancellare manualmente la riga `license_accepted = True` dal `config.ini` e riavviare l'app. |
| **Nota** | Il metodo `show_first_run_license()` è dead code e non è coinvolto in questo rischio. Il flusso attivo è esclusivamente in `main_task()`. |

Questo rischio richiede una decisione esplicita: il flusso first-run deve essere mantenuto, rimosso, o sostituito? Il piano di implementazione (§6) incorpora questa decisione come prerequisito.

---

### RISCHIO 2 — `webbrowser` non importato in `dataflow.py` — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Il nuovo metodo di `open_license_window` (o la funzione in `window_launchers.py`) userà `webbrowser`. `webbrowser` è già importato in `ui/window_launchers.py` riga 1. **`webbrowser` NON è importato in `dataflow.py`** (confermato da grep). |
| **Impatto** | Se il body viene scritto direttamente in `dataflow.py` senza delegare a `window_launchers.py`, bisogna aggiungere `import webbrowser`. Se si delega a `window_launchers.py`, nessun nuovo import necessario. |
| **Come verificarlo** | Grep `import webbrowser` in `dataflow.py` prima di scrivere. |

---

### RISCHIO 3 — Stringhe i18n orfane — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Le chiavi `.po` relative al testo della licenza e al titolo della finestra rimarrebbero nei file di traduzione anche dopo la rimozione della finestra. |
| **Impatto** | Nessun errore a runtime — le chiavi non usate producono solo rumore nei file `.po`. L'app funziona normalmente. |
| **Come verificarlo** | Non richiede verifica funzionale. È pulizia opzionale. |

---

### RISCHIO 4 — `dataflow.spec` non aggiornato — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Se `ui.license_window` rimane in `hiddenimports` anche dopo l'eliminazione del file, PyInstaller solleva un warning o errore durante la build dell'eseguibile. |
| **Impatto** | Nessun impatto a runtime in sviluppo; impatto sulla build distribuibile (EXE/MSIX). |
| **Come verificarlo** | Rieseguire `pyinstaller dataflow.spec` dopo l'implementazione e verificare assenza di errori relativi a `ui.license_window`. |

---

### RISCHIO 5 — `show_first_run_license` è dead code non riconosciuto — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Il metodo esiste in `MainWindow` (righe 1090–1129) ma non è chiamato da nessun punto del codebase. Potrebbe generare confusione in un eventuale futuro intervento. |
| **Impatto** | Nessuno a runtime, ma crea errore di comprensione sul flusso first-run se non eliminato insieme a `open_license_window`. |
| **Come verificarlo** | Grep `show_first_run_license` dopo la rimozione — deve restituire zero risultati. |

---

### RISCHIO 6 — Test di regressione UI manuale — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Non esistono test automatizzati per il pulsante License o per la finestra License. Il funzionamento è verificabile solo manualmente. |
| **Impatto** | Una regressione (pulsante che non risponde, URL sbagliato, crash) non verrebbe rilevata da suite automatica. |
| **Come verificarlo** | Verifica manuale: click sul pulsante → si apre il browser al corretto URL GitHub. |

---

## 6. Piano dettagliato di implementazione

> **PREREQUISITO OBBLIGATORIO prima di iniziare**: Decisione sul flusso first-run.
>
> Il piano assume che **il flusso first-run venga rimosso insieme alla finestra** (coerente con l'obiettivo dichiarato di rimozione totale). Se invece deve essere preservato in forma alternativa, STEP D deve essere modificato di conseguenza.

---

### STEP A — Aggiunta di `open_license_window` in `ui/window_launchers.py`

**Scopo:** Creare la funzione che apre il link GitHub License nel browser, seguendo il pattern già usato da `open_help_window`.

**File coinvolti:** `ui/window_launchers.py`

**Cosa fare:**
- Aggiungere una costante con l'URL `https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE`
- Aggiungere la funzione `open_license_window(app)` con corpo: `webbrowser.open(_LICENSE_URL)` — nota: nessun rilevamento lingua necessario, URL unico
- `webbrowser` è già importato in riga 1 del file: nessun nuovo import necessario

**Cosa verificare prima di STEP B:** Il file si importa senza errori; la funzione esiste e ha la firma attesa.

---

### STEP B — Sostituzione del corpo di `MainWindow.open_license_window()` in `dataflow.py`

**Scopo:** Il metodo delegante deve smettere di istanziare `LicenseWindow` e deve delegare alla nuova funzione in `window_launchers.py`.

**File coinvolti:** `dataflow.py` righe 1080–1088

**Cosa fare:**
- Sostituire l'intero corpo di `open_license_window()` con una singola riga di delega a `open_license_window(self)` importata da `window_launchers.py` — esattamente come fatto per Help a riga 1573
- Mantenere il nome del metodo identico: `def open_license_window(self):`
- Rimuovere le linee con `self.root.attributes('-topmost', False)` (non più necessarie)

**Cosa verificare prima di STEP C:** Click sul pulsante License apre il browser all'URL corretto senza errori.

---

### STEP C — Aggiornamento dell'import di `window_launchers` in `dataflow.py`

**Scopo:** Il nuovo `open_license_window` di `window_launchers.py` deve essere importato.

**File coinvolti:** `dataflow.py` riga 95

**Cosa fare:**
- Riga 95: aggiungere `open_license_window` alla riga di import esistente: `from ui.window_launchers import open_help_window, on_kpi_click, open_license_window`

**Cosa verificare prima di STEP D:** Nessun `ImportError` all'avvio.

---

### STEP D — Gestione del flusso first-run in `main_task()` ⚠️ STEP CRITICO

**Scopo:** Rimuovere l'unico uso attivo di `LicenseWindow(root, first_run=True)`.

**File coinvolti:** `dataflow.py` righe 4396–4428

**Cosa fare (ipotesi: rimozione del flusso first-run):**
- Righe 4396–4428: eliminare la lettura di `license_was_accepted` dal config, eliminare il blocco `if not license_was_accepted:` con tutta la sua logica
- Verificare che le righe successive (identità utente, show MainWindow ecc.) non dipendano dal valore di ritorno o da variabili introdotte in quel blocco

**Nota:** `license_accepted` è usato **solo** in `dataflow.py` (righe 1118, 4401, 4420) e in un file di documentazione `.md`. Riga 1118 è dentro `show_first_run_license()` che è dead code. I soli riferimenti operativi sono nelle righe 4401 e 4420 dentro `main_task()`.

**Cosa verificare prima di STEP E:** L'app si avvia correttamente su profilo pulito (senza `config.ini`) senza crash.

---

### STEP E — Rimozione di `MainWindow.show_first_run_license()` in `dataflow.py`

**Scopo:** Eliminare il metodo dead code che non è mai chiamato.

**File coinvolti:** `dataflow.py` righe 1090–1129

**Cosa fare:**
- Eliminare l'intero metodo `show_first_run_license()` inclusi i commenti delimitatori `# --- INIZIO/FINE NUOVI METODI LICENZA ---`

**Cosa verificare prima di STEP F:** Grep `show_first_run_license` restituisce zero risultati.

---

### STEP F — Rimozione dell'import `LicenseWindow` in `dataflow.py`

**Scopo:** Eliminare l'import non più necessario.

**File coinvolti:** `dataflow.py` riga 116

**Cosa fare:**
- Riga 116: eliminare `from ui.license_window import LicenseWindow`

**Cosa verificare prima di STEP G:** App si avvia senza `ImportError` o `NameError`.

---

### STEP G — Eliminazione del file `ui/license_window.py`

**Scopo:** Rimuovere fisicamente il file della finestra.

**File coinvolti:** `ui/license_window.py`

**Cosa fare:**
- Eliminare il file

**Cosa verificare prima di STEP H:** Grep `license_window` nel codebase Python restituisce zero risultati (esclusi `.spec` e `.md`).

---

### STEP H — Aggiornamento dei file `.spec` PyInstaller

**Scopo:** Rimuovere il riferimento a `ui.license_window` dagli `hiddenimports`.

**File coinvolti:** `dataflow.spec` riga 105 e `Tools per build WIN/Creare EXE/DataFlow.spec` riga 105

**Cosa fare:**
- In entrambi i file: rimuovere la riga `'ui.license_window',` dalla lista `hiddenimports`

**Cosa verificare prima di STEP I:** Rieseguire `pyinstaller dataflow.spec` e verificare che non ci siano warning/errori relativi a `ui.license_window`.

---

### STEP I — Verifica finale anti-regressione

**Scopo:** Confermare che nessun percorso di codice sia rimasto rotto.

**Cosa fare:**
1. Avviare l'app con `config.ini` modificato per verificare il comportamento senza `license_accepted` (o con flusso first-run rimosso)
2. Click sul pulsante `≡ License` → verifica apertura browser all'URL `https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE`
3. Click `❓ Guida` → verifica che Help funzioni ancora normalmente (regressione incrociata)
4. Verificare che il pulsante License mantenga posizione e aspetto identici
5. Grep finale `LicenseWindow|license_window` nel codebase Python → zero risultati attesi

---

## 7. Verdetto finale

**Fattibilità:** Alta — la feature è ben definita e il pattern di implementazione è già consolidato nel codebase (vedi Help).

**Complessità stimata:** **Bassa** per la sostituzione del pulsante; **Media** per la rimozione totale a causa dell'implicazione sul flusso first-run.

**Livello di rischio complessivo:** **Medio** — concentrato interamente in STEP D (flusso first-run), che richiede una decisione esplicita prima di procedere.

**Prerequisiti e attenzioni particolari:**

1. **Decisione esplicita sul flusso first-run** — L'obiettivo "rimozione totale della finestra" si scontra con il fatto che `LicenseWindow(root, first_run=True)` è un gate di avvio dell'applicazione. Il suo comportamento deve essere definito **prima** dell'implementazione: rimosso completamente, o sostituito con un meccanismo diverso (es. semplice `messagebox` con accettazione)?

2. **`webbrowser` NON è importato in `dataflow.py`** — Se per qualsiasi ragione si scegliesse di non delegare a `window_launchers.py`, l'import andrebbe aggiunto a `dataflow.py`. La scelta raccomandata è delegare a `window_launchers.py`, coerentemente con il pattern Help, evitando questa necessità.

3. **Due file `.spec` da aggiornare** — Il file nella cartella `Tools per build WIN/` è un duplicato non sincronizzato automaticamente: deve essere aggiornato manualmente in modo indipendente.
