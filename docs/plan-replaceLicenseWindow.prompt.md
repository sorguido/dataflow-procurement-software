# PIANO OPERATIVO — SOSTITUZIONE FINESTRA LICENSE

---

## 1. Stato attuale rilevato

### Dashboard License

Il pulsante `≡ License` in `ui/main_dashboard_builder.py` riga 92 è definito come `ttk.Button(..., command=app.open_license_window)`. Il metodo `MainWindow.open_license_window()` (`dataflow.py` righe 1080–1088) toglie temporaneamente il flag `-topmost` e istanzia `LicenseWindow(self.root, first_run=False)` da `ui/license_window.py`.

### First-run

`main_task()` (`dataflow.py` riga 4396) legge `config['Settings']['license_accepted']` dal file `config.ini`. Se assente o `False`, istanzia `LicenseWindow(root, first_run=True)` direttamente (riga 4407), fa `root.wait_window(license_prompt)`, e salva l'accettazione nel config prima di proseguire.

### Persistenza accettazione

Scritta in `config.ini` sezione `[Settings]`, chiave `license_accepted = True`. Letta all'avvio da `main_task()`. Usata esclusivamente in tre punti di `dataflow.py` (righe 1118, 4401, 4420). Riga 1118 è dead code (dentro `show_first_run_license()`). I soli punti operativi sono righe 4401 e 4420.

---

## 2. Obiettivo finale da implementare

### A) Pulsante License dashboard → link esterno

- `app.btn_license`: **invariato** (testo, posizione, pack)
- `command=app.open_license_window`: **invariato** (nome metodo)
- Il corpo di `MainWindow.open_license_window()` delega a una nuova funzione in `ui/window_launchers.py` che chiama `webbrowser.open(_LICENSE_URL)`
- Nessuna finestra interna viene più aperta

### B) First-run → dialog custom minimale

La vecchia `LicenseWindow(root, first_run=True)` è sostituita con una nuova classe `LicenseAcceptanceDialog(tk.Toplevel)` da aggiungere a `ui/dialogs/common_dialogs.py`. Il dialog:
- Mostra un messaggio breve localizzato
- Ha tre pulsanti: **License** (apre browser, dialog rimane aperto), **Accept** (accetta, chiude), **Exit** (chiude senza accettare)
- La X è trattata come **Exit** (rifiuto implicito — comportamento più sicuro, coerente con `UserIdentityDialog._prevent_close` che blocca la X quando la conferma è obbligatoria)
- L'attributo `self.accepted` su di esso funziona identicamente a quello della vecchia `LicenseWindow`

### Equivalenza logica

La logica di `main_task()` (verifica config → mostra dialog → `wait_window` → salva config → prosegue o chiude) rimane **identica strutturalmente**. Cambia solo le classi istanziate e il nome del dialog.

---

## 3. File da modificare

### 3.1 — `ui/window_launchers.py`

| Attributo | Dettaglio |
|-----------|-----------|
| **Percorso** | `ui/window_launchers.py` |
| **Motivo** | Aggiungere funzione `open_license_window(app)` che apre il link GitHub LICENSE nel browser |
| **Tipo di modifica** | Aggiunta di costante `_LICENSE_URL` e funzione `open_license_window(app)` — `webbrowser` già importato alla riga 1 |
| **Rischio** | Minimo — pattern identico a `open_help_window`, nessun nuovo import necessario |

---

### 3.2 — `ui/dialogs/common_dialogs.py`

| Attributo | Dettaglio |
|-----------|-----------|
| **Percorso** | `ui/dialogs/common_dialogs.py` |
| **Motivo** | Aggiungere la nuova classe `LicenseAcceptanceDialog(tk.Toplevel)` minimale |
| **Tipo di modifica** | Aggiunta di nuova classe in fondo al file, seguendo il pattern di `SimpleYesNoDialog` / `UserIdentityDialog` |
| **Rischio** | Basso — il file agisce come contenitore di dialog custom, è il posto architetturalmente corretto |

**Struttura del nuovo dialog (descrittiva, non codice):**
- `__init__(self, parent, url)`: chiama `super().__init__(parent)`, `withdraw()`, `set_window_icon()`, `title(...)`, `transient(parent)`, `grab_set()`, `self.accepted = False`
- `WM_DELETE_WINDOW` → gestito come Exit (rifiuto implicito)
- Layout: label con messaggio breve, frame pulsanti con tre `ttk.Button`
- Pulsante **License** → `webbrowser.open(url)` — non chiude il dialog
- Pulsante **Accept** → `self.accepted = True`, `grab_release()`, `destroy()`
- Pulsante **Exit** → `self.accepted = False`, `grab_release()`, `destroy()`
- `center_window(self)`, `deiconify()`, `wait_visibility()` — identici a `SimpleMessageDialog`

**Note su dipendenze:**
- `webbrowser` è già importato in `common_dialogs.py` riga 6 — nessun nuovo import necessario
- L'URL viene passato come parametro `url` al costruttore — zero dipendenze aggiuntive, massima flessibilità

---

### 3.3 — `dataflow.py`

| Modifica | Riga/e | Cosa fare |
|----------|--------|-----------|
| **1** | 116 | Rimuovere `from ui.license_window import LicenseWindow` |
| **2** | 95 | Aggiungere `open_license_window` all'import da `ui.window_launchers` |
| **3** | 117–125 | Aggiungere `LicenseAcceptanceDialog` all'import da `ui.dialogs.common_dialogs` |
| **4** | 1080–1088 | Sostituire corpo di `open_license_window()` con delega a `open_license_window(self)` di `window_launchers` — identico al pattern di `open_help_window` riga 1573 |
| **5** | 1090–1129 | Eliminare metodo dead code `show_first_run_license()` con i suoi commenti delimitatori |
| **6** | 4407 | Sostituire `LicenseWindow(root, first_run=True)` con `LicenseAcceptanceDialog(root, url=_LICENSE_URL)` |

**Nota su `_LICENSE_URL` in `dataflow.py`:** viene importata da `window_launchers.py` (già importato) nella stessa riga dell'import di `open_license_window`. Nessuna duplicazione.

**Rischio specifico modifica 6:** il flusso `wait_window` e `getattr(license_prompt, 'accepted', False)` rimane invariato; la compatibilità dipende da `LicenseAcceptanceDialog.accepted` che funzioni esattamente come `LicenseWindow.accepted`.

---

## 4. File potenzialmente eliminabili

### `ui/license_window.py`

| Attributo | Dettaglio |
|-----------|-----------|
| **Percorso** | `ui/license_window.py` |
| **Perché eliminabile** | Tutta la classe `LicenseWindow` cessa di essere utilizzata dopo le modifiche ai punti 3.2 e 3.3 |
| **Dipendenze da rimuovere prima** | Import `from ui.license_window import LicenseWindow` da `dataflow.py` riga 116; tutte le istanze nel codice Python; entry `'ui.license_window'` nei file `.spec` |
| **Condizione** | Eliminabile **solo dopo** che `LicenseAcceptanceDialog` è operativa e testata, e solo dopo la pulizia di tutti i riferimenti |

---

## 5. Traduzioni e localizzazione

### Nuove stringhe necessarie

Per il nuovo `LicenseAcceptanceDialog` servono le seguenti nuove chiavi msgid, **non presenti** nei `.po` attuali:

| msgid | IT msgstr | EN msgstr |
|-------|-----------|-----------|
| Corpo messaggio IT | `"Per utilizzare DataFlow Procurement Software è necessario accettare i termini e le condizioni d'uso."` | `"To use DataFlow Procurement Software, you must accept the terms and conditions of use."` |
| Titolo finestra IT | `"Accettazione Licenza"` | `"License Agreement"` |
| Pulsante License IT | `"📄 Leggi la Licenza"` | `"📄 Read License"` |

### Stringhe riutilizzabili (già nei `.po`)

Le seguenti stringhe già esistono e possono essere riutilizzate nel nuovo dialog senza aggiungere entry:

| msgid | EN msgstr | Fonte attuale |
|-------|-----------|---------------|
| `"✅ Accetto"` | `"✅ Accept"` | Vecchia `LicenseWindow` — già tradotta in entrambi i `.po` |
| `"❌ Esci"` | `"❌ Exit"` | Vecchia `LicenseWindow` — già tradotta in entrambi i `.po` |

### Stringhe obsolete (da NON eliminare in questo sprint)

Le seguenti chiavi diventano orfane dopo l'eliminazione di `LicenseWindow`, ma non causano errori se lasciate nei `.po`. La loro rimozione è pulizia opzionale per uno sprint dedicato separato:

- `"Licenza d'Uso - DataFlow Procurement Software"` (titolo finestra)
- `"Contratto di Licenza per l'Utente Finale (GNU GPLv3) - DataFlow Procurement Software\n\n"` (h1 del testo)
- Tutto il blocco multi-riga del testo legale della licenza (~20 chiavi, righe 897–1770 in IT `.po`, righe corrispondenti in EN `.po`)
- `"Impossibile salvare l'impostazione della licenza: {}\n\n..."` (da `show_first_run_license()` dead code)
- Stringhe dei tag interni: `"Sviluppatore: "`, `"E-mail: "`, `"Copyright © 2025 Guido Sorarù.\n\n"`, ecc.

### File `.po` da modificare

| File | Tipo di modifica |
|------|-----------------|
| `locale/it/LC_MESSAGES/dataflow.po` | Aggiunta delle 3 nuove coppie msgid/msgstr in IT |
| `locale/en/LC_MESSAGES/dataflow.po` | Aggiunta delle 3 nuove coppie msgid/msgstr in EN |

### File `.mo` da rigenerare

| File | Trigger |
|------|---------|
| `locale/it/LC_MESSAGES/dataflow.mo` | Dopo modifica al `.po` IT |
| `locale/en/LC_MESSAGES/dataflow.mo` | Dopo modifica al `.po` EN |

### Conferma metodo di compilazione

> ✅ La compilazione dei `.mo` deve avvenire **esclusivamente tramite `compile_translations.py`** che usa `polib.pofile(po_path).save_as_mofile(mo_path)`. Nessun altro strumento (`msgfmt`, `pybabel`, ecc.) è ammesso.

---

## 6. Piano sequenziale di implementazione

> **Prerequisito**: Per testare il first-run, rimuovere temporaneamente `license_accepted = True` dal `config.ini` locale per simulare un profilo pulito.

> **Vincolo d'ordine**: L'ordine degli STEP è vincolante. STEP C (rimozione import `LicenseWindow`) non deve precedere STEP F (sostituzione in `main_task()`). L'eliminazione fisica del file (STEP I) è l'ultimo atto.

---

### STEP A — Aggiunta di `open_license_window(app)` in `ui/window_launchers.py`

**Scopo:** Creare la funzione che apre il link GitHub LICENSE nel browser, seguendo il pattern `open_help_window`.

**File:** `ui/window_launchers.py`

**Cosa fare:**
- Aggiungere costante `_LICENSE_URL = "https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE"`
- Aggiungere funzione `open_license_window(app):` con corpo `webbrowser.open(_LICENSE_URL)`
- `webbrowser` già importato a riga 1 — nessun nuovo import

**Verificare prima di STEP B:** Modulo importabile senza errori; funzione presente e richiamabile.

---

### STEP B — Aggiunta della classe `LicenseAcceptanceDialog` in `ui/dialogs/common_dialogs.py`

**Scopo:** Creare il dialog custom minimale per il first-run, con tre pulsanti (License, Accept, Exit), modale, con `self.accepted` come interfaccia pubblica.

**File:** `ui/dialogs/common_dialogs.py`

**Cosa fare:**
- Aggiungere classe `LicenseAcceptanceDialog(tk.Toplevel)` in fondo al file
- Costruttore accetta `(self, parent, url)` per disaccoppiare l'URL dalla classe
- Pulsante **License** → `webbrowser.open(url)` — non chiude il dialog
- Pulsante **Accept** → `self.accepted = True`, `grab_release()`, `destroy()`
- Pulsante **Exit** → `self.accepted = False`, `grab_release()`, `destroy()`
- `WM_DELETE_WINDOW` → chiama la stessa logica di Exit (rifiuto implicito)
- Testo del messaggio usa `_("Per utilizzare DataFlow Procurement Software...")` localizzato
- `webbrowser` già importato a riga 6 — nessun nuovo import

**Verificare prima di STEP C:** La classe è importabile; un'istanza di test mostra il dialog con i tre pulsanti; `self.accepted` vale `False` di default; il pulsante License apre il browser senza chiudere il dialog.

---

### STEP C — Aggiornamento degli import in `dataflow.py`

**Scopo:** Rendere disponibili la nuova funzione e la nuova classe; rimuovere l'import obsoleto.

**File:** `dataflow.py`

**Cosa fare:**
- Riga 95: estendere l'import da `ui.window_launchers` aggiungendo `open_license_window` e `_LICENSE_URL`
- Righe 117–125: aggiungere `LicenseAcceptanceDialog` alla lista dell'import da `ui.dialogs.common_dialogs`
- Riga 116: rimuovere `from ui.license_window import LicenseWindow`

**Verificare prima di STEP D:** Avvio senza `ImportError` o `NameError`.

---

### STEP D — Sostituzione del corpo di `MainWindow.open_license_window()` in `dataflow.py`

**Scopo:** Il metodo delegante smette di aprire la finestra interna e delega all'apertura del browser.

**File:** `dataflow.py` righe 1080–1088

**Cosa fare:**
- Sostituire l'intero corpo con la singola riga `open_license_window(self)` — identico al pattern di `open_help_window` a riga 1573
- Rimuovere le righe con `self.root.attributes('-topmost', False)` (non più necessarie)
- Mantenere invariata la firma: `def open_license_window(self):`

**Verificare prima di STEP E:** Click sul pulsante `≡ License` nella dashboard apre il browser correttamente; nessun crash; la dashboard rimane aperta e funzionante.

---

### STEP E — Eliminazione del metodo dead code `show_first_run_license()` in `dataflow.py`

**Scopo:** Rimuovere il metodo che non è mai chiamato da nessun punto del codebase.

**File:** `dataflow.py` righe 1090–1129

**Cosa fare:**
- Eliminare l'intero metodo `show_first_run_license()` inclusi i commenti `# --- INIZIO/FINE NUOVI METODI LICENZA ---`

**Verificare prima di STEP F:** Grep `show_first_run_license` → zero risultati nel codebase Python.

---

### STEP F — Sostituzione del flusso first-run in `main_task()` in `dataflow.py`

**Scopo:** Sostituire `LicenseWindow(root, first_run=True)` con il nuovo `LicenseAcceptanceDialog`.

**File:** `dataflow.py` riga 4407

**Cosa fare:**
- Riga 4407: sostituire `LicenseWindow(root, first_run=True)` con `LicenseAcceptanceDialog(root, url=_LICENSE_URL)`
- La riga `root.wait_window(license_prompt)` rimane **invariata**
- La riga `if not getattr(license_prompt, 'accepted', False):` rimane **invariata**
- Il blocco di salvataggio config rimane **invariato**
- Il messaggio `logger.warning(f"Impossibile salvare stato licenza: {e}")` rimane **invariato** (è un log non-i18n)

**Verificare prima di STEP G:** Su profilo pulito: il dialog si apre all'avvio; i tre pulsanti funzionano; Accept permette di proseguire; Exit chiude l'app; la X è trattata come Exit.

---

### STEP G — Aggiornamento dei file `.po` con le nuove stringhe

**Scopo:** Aggiungere le nuove chiavi i18n necessarie per il dialog custom.

**File:** `locale/it/LC_MESSAGES/dataflow.po` e `locale/en/LC_MESSAGES/dataflow.po`

**Cosa fare:**
- Aggiungere le 3 nuove entry in ciascun `.po` (titolo finestra, messaggio corpo, pulsante License)
- Seguire la formattazione esistente nel file (msgid/msgstr, nessuna riga vuota extra)
- IT: `msgstr` = stesso testo del `msgid` per i tre testi in italiano
- EN: `msgstr` = testo in inglese per i tre testi

**Verificare prima di STEP H:** I file `.po` sono sintatticamente validi. Le chiavi compaiono in `po.translated_entries()` di polib.

---

### STEP H — Compilazione `.mo` con `compile_translations.py`

**Scopo:** Rigenerare i file `.mo` con le nuove stringhe.

**File:** `locale/it/LC_MESSAGES/dataflow.mo` e `locale/en/LC_MESSAGES/dataflow.mo`

**Cosa fare:**
- Eseguire `python compile_translations.py` dalla root del progetto con il virtualenv attivo

**Verificare prima di STEP I:** Output mostra `✅ Compilato: it/LC_MESSAGES/dataflow.mo` e `✅ Compilato: en/LC_MESSAGES/dataflow.mo` senza errori.

---

### STEP I — Eliminazione del file `ui/license_window.py`

**Scopo:** Rimuovere fisicamente il file della vecchia finestra.

**File:** `ui/license_window.py`

**Cosa fare:**
- Verificare tramite grep che nessun file Python importi ancora `LicenseWindow` o `license_window`
- Eliminare il file fisicamente

**Verificare prima di STEP L:** Grep `license_window|LicenseWindow` nel codebase Python → zero risultati.

---

### STEP L — Aggiornamento dei file `.spec` PyInstaller

**Scopo:** Rimuovere `'ui.license_window'` dagli `hiddenimports`.

**File:** `dataflow.spec` riga 105 e `Tools per build WIN/Creare EXE/DataFlow.spec` riga 105

**Cosa fare:**
- In entrambi i file rimuovere la riga `'ui.license_window',` dalla lista `hiddenimports`
- I due file sono indipendenti e vanno aggiornati separatamente

**Verificare dopo:** Rieseguire `pyinstaller dataflow.spec` e verificare assenza di warning/errori relativi a `ui.license_window`.

---

## 7. Analisi rischi

### RISCHIO 1 — Comportamento della X sul dialog first-run — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Se la X non è intercettata, Tkinter esegue `destroy()` predefinito senza impostare `self.accepted` |
| **Impatto** | `getattr(license_prompt, 'accepted', False)` restituisce `False` → `root.destroy()` → comportamento corretto (uscita sicura). Il `fallback=False` del getattr garantisce sicurezza anche senza intercettazione esplicita |
| **Verifica** | Chiudere il dialog con la X su profilo pulito → app deve terminare |

---

### RISCHIO 2 — `grab_release()` prima di `destroy()` — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Se il dialog usa `grab_set()` e viene distrutto senza `grab_release()`, il parent potrebbe restare bloccato su alcuni sistemi |
| **Impatto** | Root window che non risponde all'interazione dopo il dialog |
| **Verifica** | Verificare che dopo Accept/Exit la dashboard sia pienamente interattiva |

---

### RISCHIO 3 — Pulsante License nel first-run non chiude il dialog — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Il requisito richiede che License apra il browser **senza chiudere** il dialog. Su Linux/Wayland alcuni browser aprono in finestra separata che può rubare il focus |
| **Impatto** | L'utente potrebbe avere difficoltà a tornare al dialog se il browser prende tutto il focus |
| **Verifica** | Testare su sistema target: aprire browser → tornare al dialog → Accept funziona ancora |

---

### RISCHIO 4 — Compatibilità `LicenseAcceptanceDialog.accepted` con `main_task()` — BASSO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | `main_task()` usa `getattr(license_prompt, 'accepted', False)` — se il dialog non inizializza `self.accepted = False` nell'`__init__`, un `destroy()` inatteso potrebbe non trovare l'attributo |
| **Impatto** | Il `fallback=False` del `getattr` garantisce comunque uscita sicura |
| **Verifica** | Confermato a basso rischio per via del `fallback` esistente |

---

### RISCHIO 5 — Stringhe non tradotte nel dialog first-run — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Il dialog viene istanziato all'avvio. Se `_()` non è ancora inizializzato prima del dialog, le stringhe vengono mostrate non tradotte |
| **Impatto** | Dialog first-run mostrato nella lingua sbagliata |
| **Verifica** | Verificare l'ordine di inizializzazione i18n in `main_task()` — se `_()` è già attivo prima del dialog, nessun problema |

---

### RISCHIO 6 — Due file `.spec` non sincronizzati — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | `dataflow.spec` e `Tools per build WIN/Creare EXE/DataFlow.spec` sono duplicati indipendenti. Aggiornarne uno non aggiorna l'altro automaticamente |
| **Impatto** | Build Windows fallisce se il secondo `.spec` viene usato senza aggiornamento |
| **Verifica** | Verificare entrambi i file dopo STEP L |

---

### RISCHIO 7 — File `.mo` non rigenerati — MEDIO

| Attributo | Dettaglio |
|-----------|-----------|
| **Causa** | Se i `.mo` non vengono rigenerati dopo l'aggiornamento dei `.po`, le nuove stringhe del dialog appaiono come msgid non tradotte |
| **Impatto** | Dialog first-run mostra il testo italiano grezzo agli utenti EN (o viceversa) |
| **Verifica** | STEP H verifica la compilazione; testare poi l'app in entrambe le lingue |

---

## 8. Checklist finale anti-regressione

- [ ] **First-run profilo pulito (IT):** Cancellare `license_accepted` da `config.ini` → avviare app → il dialog custom appare → click Accept → app prosegue → `config.ini` contiene `license_accepted = True`
- [ ] **First-run profilo pulito (EN):** Stesso test con lingua inglese → testo del dialog in inglese
- [ ] **First-run → Exit:** Dialog aperto → click Exit → app chiusa → `config.ini` non contiene `license_accepted = True`
- [ ] **First-run → chiusura con X:** Dialog aperto → click X → app chiusa (stessa conseguenza di Exit)
- [ ] **First-run → click License:** Dialog aperto → click License → browser si apre su `https://github.com/sorguido/dataflow-procurement-software/blob/main/LICENSE` → dialog **rimane aperto e modale** → successivo click Accept funziona
- [ ] **Dashboard → pulsante License:** App avviata normalmente → click `≡ License` → browser si apre sull'URL corretto → nessuna finestra interna aperta → dashboard rimane operativa
- [ ] **Dashboard → pulsante Guida:** Click `❓ Guida` → wiki GitHub si apre normalmente (regressione incrociata)
- [ ] **Posizione pulsante License:** Layout della toolbar invariato — License a destra di Settings, a sinistra di Guida
- [ ] **Secondo avvio (licenza già accettata):** `config.ini` con `license_accepted = True` → app si avvia direttamente alla dashboard senza mostrare il dialog first-run
- [ ] **Grep residui:** `grep -r "LicenseWindow\|license_window" --include="*.py" .` → zero risultati
- [ ] **Grep import obsoleto:** `grep "from ui.license_window" dataflow.py` → zero risultati
- [ ] **File eliminato:** `ls ui/license_window.py` → file not found
- [ ] **File `.spec` aggiornati:** Grep `ui.license_window` in entrambi i file `.spec` → zero risultati
- [ ] **i18n nuove stringhe:** Avviare app in EN → dialog first-run mostra testo inglese; avviare in IT → testo italiano

---

## 9. Verdetto finale

**Fattibilità:** Alta. L'intervento è ben delimitato, il pattern di riferimento (Help) è già consolidato in codebase, le dipendenze sono chiare.

**Complessità:** **Media.** La sostituzione del pulsante dashboard è di bassa complessità. La creazione del dialog custom first-run e la sua integrazione nel flusso di `main_task()` richiede attenzione ma non è architetturalmente complessa.

**Rischio complessivo:** **Basso-Medio.** Il rischio principale (flusso first-run) è gestito con un approccio conservativo che replica esattamente la struttura esistente. Il `fallback=False` nel `getattr` di `main_task()` garantisce fail-safe. Il rischio residuo più concreto è la mancata sincronizzazione dei due file `.spec`.

**Note operative importanti:**

1. **L'ordine degli STEP è vincolante.** STEP C (rimozione import `LicenseWindow`) non deve precedere STEP F (sostituzione in `main_task()`). L'eliminazione del file fisico (STEP I) è l'ultimo atto.

2. **La costante `_LICENSE_URL` deve essere accessibile sia in `window_launchers.py` (per STEP A) sia in `dataflow.py` (per STEP F).** L'approccio raccomandato è definirla in `window_launchers.py` e importarla in `dataflow.py` nella stessa riga dell'import di `open_license_window`.

3. **Le stringhe obsolete della vecchia `LicenseWindow` nei `.po` non devono essere rimosse in questo sprint.** Sono innocue se lasciate e la loro rimozione richiederebbe una fase separata di audit i18n più ampia.
