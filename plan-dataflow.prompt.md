# Plan: DataFlow - Preparazione Pubblicazione Open Source

## TL;DR
Preparare DataFlow per la pubblicazione open source su GitHub (licenza GPLv3). Il progetto è principalmente compatibile con Linux, ma ha **8 problemi critici di dipendenze Windows** (import windll non protetti, hardcoded AppData paths, backslash paths) che causano crash/errori runtime. Il piano prevede: fix minimi e non invasivi per i problemi Windows, generazione di file di progetto standard (README, LICENSE, .gitignore), e controlli pre-pubblicazione.

---

## Fase 1: Fix Minime per Compatibilità Windows

**Obiettivo**: Correzioni minime non invasive per rendere il codice eseguibile su Linux senza crash.

### 1.1 Fix Import windll (CRITICO)
- [DataFlow 2.0.0.py](DataFlow%202.0.0.py#L3) riga 3: Wrappare import `from ctypes import windll` con `if sys.platform == 'win32':`
- [Database_Migration_Tool/Database_Migration_Tool.py](Database_Migration_Tool/Database_Migration_Tool.py#L26) riga 26-27: Proteggere importazione e funzioni `windll.shcore.SetProcessDpiAwareness()` e `windll.user32.SetProcessDPIAware()`
- **Pattern Fix**: Wrappare con `if sys.platform == 'win32':` e definire variabili di fallback vuote su Linux
- **Impatto**: Elimina crash immediato su Linux

### 1.2 Fix Hardcoded AppData Paths (ALTO)
- [DataFlow 2.0.0.py](DataFlow%202.0.0.py#L97) riga 97: Path config directory da `~/AppData/Local/DataFlow` a `~/.local/share/dataflow` su Linux
- [DataFlow 2.0.0.py](DataFlow%202.0.0.py#L151) riga 151: Stesso pattern
- [Database_Migration_Tool/ui_dialogs.py](Database_Migration_Tool/ui_dialogs.py#L52) riga 52: Stessa correzione
- **Pattern Fix**: Creare helper function `get_config_dir()` cross-platform che usa `~/.local/share/dataflow` su Linux e `~/AppData/Local/DataFlow` su Windows oppure usare `platformdirs` library (ma non aggiungere dipendenza esterna per mantenere semplicità)
- **Alternativa**: Usare pattern `os.path.join(os.path.expanduser('~'), '.local', 'share', 'dataflow')` per Linux
- **Impatto**: Log e config files si salvano in directory valide

### 1.3 Fix Backslash Paths (MEDIO-ALTO)
- [DataFlow 2.0.0.py](DataFlow%202.0.0.py#L5063) riga 5063: `os.path.expanduser("~\\Documents")` → `os.path.join(os.path.expanduser('~'), 'Documents')`
- [Database_Migration_Tool/ui_dialogs.py](Database_Migration_Tool/ui_dialogs.py#L77) riga 77: Stesso pattern per file picker initialdir
- [Database_Migration_Tool/logger_setup.py](Database_Migration_Tool/logger_setup.py#L28) riga 28: Stesso pattern per log directory
- **Pattern Fix**: Sostituire backslash paths con `os.path.join()` che è cross-platform
- **Impatto**: File picker e logging funzionano su Linux

### 1.4 Fix Script non Portatile (BASSO)
- [add_missing_translations.py](add_missing_translations.py#L97): Verificare se ha hardcoded path tipo `c:\Users\sorgu\Il mio Drive\...` - se sì, commentare o rendere configurabile
- **Impatto**: Script rimane non eseguibile (ma non è core app, è utility)

### 1.5 Gestione setup_venv.bat/.ps1 (INFORMATIVO)
- [Database_Migration_Tool/setup_venv.bat](Database_Migration_Tool/setup_venv.bat) e [.ps1](Database_Migration_Tool/setup_venv.ps1) rimangono solo Windows
- Su Linux: utenti useranno `python3 -m venv venv` manualmente (documentato in README)
- **Non serve fix**: Questi script sono helper, non core application

---

## Fase 2: Generazione File Progetto Standard

**Obiettivo**: Creare file fondamentali per pubblicazione open source (non invasivo con codebase).

### 2.1 README.md (CRITICO)
**Ubicazione**: `/README.md` (root)
**Contenuto**:
- Descrizione breve (DataFlow è app desktop Python Tkinter per gestione dati con GUI, Excel, traduzioni, SQLite)
- Badge: Python version, License (GPLv3), GitHub stars/issues (quando published)
- Funzionalità principali (lista 5-6 punti)
- Prerequisiti (Python 3.12+, venv)
- Guida installazione (clone, venv, pip install)
- Guida uso (esecuzione principale `python3 "DataFlow 2.0.0.py"`)
- Lingue supportate (IT, EN)
- Struttura cartelle
- Database migration tool (breve descrizione)
- Licenza (GPLv3)
- Contatti/contributi (se applicabile)

### 2.2 LICENSE (CRITICO)
**Ubicazione**: `/LICENSE` (root)
**Contenuto**: Testo completo MIT GPLv3 v3.0 (template standard da github.com/licenses)

### 2.3 .gitignore (CRITICO)
**Ubicazione**: `/.gitignore` (root)
**Contenuto**: Standard Python .gitignore che esclude:
- `__pycache__/`, `*.pyc`, `.Python`
- `.venv/`, `venv/`, `env/`
- `.env`
- `*.xlsx.tmp` (temp Excel files)
- `*.egg-info/`
- `dist/`, `build/`
- `.DS_Store` (macOS)
- `.idea/`, `.vscode/` (IDE files - opzionale)
- `*.spec` (PyInstaller output)
- Database temp files (`*.wal`, `*.shm`)

---

## Fase 3: Controlli Pre-Pubblicazione

**Obiettivo**: Validare che tutto sia pronto prima di pubblicare su GitHub.

### 3.1 Verifiche Code
- [ ] Eseguire `python3 "DataFlow 2.0.0.py"` su Linux → GUI si apre senza errori
- [ ] Verificare che database SQLite si crea in `~/.local/share/dataflow/` (o config dir corretta)
- [ ] Verificare traduzioni caricate (IT e EN)
- [ ] Verificare template Excel accessibili
- [ ] Test Database Migration Tool su Linux (se applicabile)

### 3.2 Verifiche Dipendenze
- [ ] Eseguire `pip install -r requirements.txt` in venv pulito → installa successo
- [ ] Nessun warning da `pip check`
- [ ] Verificare versioni dipendenze sono specifiche (non latest/wildcards)

### 3.3 Verifiche Struttura Git
- [ ] `/` contiene README.md, LICENSE, .gitignore, requirements.txt
- [ ] `/add_data/` NON contiene file sensibili
- [ ] `.gitignore` esclude correttamente __pycache__, .venv, ecc.
- [ ] Nessun file per-user salvato in repo (config locale)

### 3.4 Verifiche Documentazione
- [ ] README.md è completo e formattato
- [ ] LICENSE è presente integralmente
- [ ] Istruzioni setup/run sono chiare e testate mentalmente per utente nuovo

### 3.5 Controlli di Sicurezza
- [ ] Nessum hardcoded password/token nel codice
- [ ] Nessun path assoluto Windows (_già verificato sopra_)
- [ ] Database SQLite non contiene dati sensibili di default

---

## Relevant Files (da creare/modificare)

**Modifiche codice** (fix Windows compatibility):
- [DataFlow 2.0.0.py](DataFlow%202.0.0.py) — righe 3, 97, 151, 5063
- [Database_Migration_Tool/Database_Migration_Tool.py](Database_Migration_Tool/Database_Migration_Tool.py) — righe 26-27  
- [Database_Migration_Tool/ui_dialogs.py](Database_Migration_Tool/ui_dialogs.py) — righe 52, 77
- [Database_Migration_Tool/logger_setup.py](Database_Migration_Tool/logger_setup.py) — riga 28

**File da creare** (zero codice modifications):
- `/README.md` — nuovo
- `/LICENSE` — nuovo (GPLv3 standard)
- `/.gitignore` — nuovo (Python standard)

**File da aggiornare** (no changes needed, solo verifica):
- `/requirements.txt` — ✓ OK già presente e cross-platform
- `/Database_Migration_Tool/requirements.txt` — ✓ OK (solo stdlib, nessuna dipendenza)

---

## Verification

### Per Phase 1 (Fix Windows)
1. Aprire DataFlow su Ubuntu/Linux VM → app si avvia senza ImportError
2. Verificare che log file si crea in `~/.local/share/dataflow/` (non AppData)
3. Testare file picker → inizia da `~/Documents` (non `~\\Documents`)
4. Nessun crash relativi a `windll` o percorsi

### Per Phase 2 (Generazione file)
1. README.md aperto in GitHub preview → formattato correttamente
2. LICENSE file riconosciuto automaticamente come GPLv3 da GitHub
3. `.gitignore` esclude correttamente file build/cache

### Per Phase 3 (Pre-pub checks)
1. Eseguire checklist sopra
2. Simulare clone da GitHub → segui README, app si avvia
3. Nessun file system-specific nel repo

---

## Decisions & Scope

**INCLUSO**:
- Fix minime per importazioni windll (wrap con `if sys.platform == 'win32':`)
- Fix percorsi hardcoded con `os.path.join()` cross-platform
- Generazione README, LICENSE, .gitignore standard
- Checklist controlli pre-pubblicazione
- Nessun refactoring, nessuna modifica struttura progetto

**ESCLUSO** (intentionally):
- Refactoring codice
- Suddivisione file
- Modifica struttura cartelle
- Upgrade dipendenze
- Aggiunta nuove dipendenze (come `platformdirs`)
- Riscrittura componenti

**Assunzioni**:
- Applicazione è testata e funziona su Windows
- Python 3.12+ disponibile target user
- Utenti saranno su Linux, Windows, macOS
- Nessun database pre-populated rilasciato

---

## Further Considerations

1. **Gestione path cross-platform**: Considerare se usare helper function centralizzata per get_config_dir() o mantenere os.path.join() inline in ogni file?
   - **Raccomandazione**: Inline `os.path.join()` - meno invasivo, non richiede refactoring, legge facile in contesto

2. **CI/CD testing**: GitHub Actions workflow per testare su Linux (ubuntu-latest)?
   - **Raccomandazione**: Nice-to-have ma NON richiesto per primo release - aggiungere dopo

3. **Documentation di migrazione database**: Database_Migration_Tool ha README interno - va incluso nel main README?
   - **Raccomandazione**: Sì, breve sezione nel README principale che rimanda a Database_Migration_Tool/README.md
