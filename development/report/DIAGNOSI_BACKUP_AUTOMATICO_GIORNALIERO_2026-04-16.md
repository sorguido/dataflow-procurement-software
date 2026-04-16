# Diagnosi tecnica esecutiva — Backup Automatico Giornaliero

Data analisi: 2026-04-16  
Ambito: solo diagnosi (nessuna patch applicata)

## 1) Stato attuale rilevato

Il comportamento osservato è coerente con il codice attuale.

- Trigger backup automatico:
  - viene eseguito subito in startup tramite `MainWindow.__init__` -> `self.check_for_autobackup()` ([dataflow.py:1080](../../dataflow.py#L1080)).
  - poi viene rieseguito ogni 60 secondi via timer `after` ([dataflow.py:1343](../../dataflow.py#L1343)).
- Guard-rail giornaliero:
  - esiste solo in memoria runtime (`self.last_backup_date`) e viene inizializzato a `None` a ogni avvio app ([dataflow.py:1042](../../dataflow.py#L1042)).
  - non è persistito su file/config/database.
- Retention “max 3 copie”:
  - la routine di pruning esiste ma il pattern di ricerca file è errato, quindi non intercetta i backup esistenti e non elimina nulla ([services/settings_maintenance_service.py:82](../../services/settings_maintenance_service.py#L82)).
- Gestione `.db/.db-wal/.db-shm`:
  - prevista esplicitamente sia per backup manuale sia automatico ([services/settings_maintenance_service.py:43](../../services/settings_maintenance_service.py#L43), [services/settings_maintenance_service.py:120](../../services/settings_maintenance_service.py#L120)).

## 2) File coinvolti

File direttamente coinvolti nella catena del backup automatico:

- [dataflow.py](../../dataflow.py)
- [services/settings_preferences_service.py](../../services/settings_preferences_service.py)
- [services/settings_maintenance_service.py](../../services/settings_maintenance_service.py)

File di contesto tecnico WAL/SQLite:

- [database_manager.py](../../database_manager.py)

## 3) Funzioni/metodi coinvolti

Caricamento/salvataggio impostazioni backup:

- `SettingsWindow.load_settings` ([dataflow.py:385](../../dataflow.py#L385))
- `load_settings_snapshot` ([services/settings_preferences_service.py:12](../../services/settings_preferences_service.py#L12))
- `SettingsWindow.save_autobackup_settings` ([dataflow.py:506](../../dataflow.py#L506))
- `save_autobackup_preferences` ([services/settings_preferences_service.py:91](../../services/settings_preferences_service.py#L91))
- `read_autobackup_config` ([services/settings_maintenance_service.py:13](../../services/settings_maintenance_service.py#L13))

Scheduling/trigger:

- `MainWindow.__init__` invoca `check_for_autobackup` in startup ([dataflow.py:1080](../../dataflow.py#L1080))
- `check_for_autobackup` ([dataflow.py:1324](../../dataflow.py#L1324))
- `root.after(60000, self.check_for_autobackup)` ([dataflow.py:1343](../../dataflow.py#L1343))

Esecuzione backup automatico e retention:

- `perform_autobackup` ([dataflow.py:1345](../../dataflow.py#L1345))
- `perform_autobackup_copy` ([services/settings_maintenance_service.py:72](../../services/settings_maintenance_service.py#L72))
- blocco pruning/retention (`backup_sets`, `glob`, `while len(...) >= 3`) ([services/settings_maintenance_service.py:80](../../services/settings_maintenance_service.py#L80))

Gestione file sidecar SQLite:

- copia WAL/SHM manuale: `copy_manual_backup_bundle` ([services/settings_maintenance_service.py:24](../../services/settings_maintenance_service.py#L24))
- copia WAL/SHM autobackup: `perform_autobackup_copy` ([services/settings_maintenance_service.py:120](../../services/settings_maintenance_service.py#L120))
- WAL mode attivo: `PRAGMA journal_mode=WAL` ([database_manager.py:58](../../database_manager.py#L58))

## 4) Flusso reale del backup automatico

A. Configurazione

1. L’utente salva `enabled/hour/path` in `[AutoBackup]` via `save_autobackup_preferences` ([services/settings_preferences_service.py:100](../../services/settings_preferences_service.py#L100)).
2. Non viene salvato nessun campo “last automatic backup date/time”.

B. Startup app

1. `main_task()` crea `MainWindow` ([dataflow.py:3256](../../dataflow.py#L3256)).
2. In `MainWindow.__init__`, `self.last_backup_date = None` ([dataflow.py:1042](../../dataflow.py#L1042)).
3. Sempre in init, viene chiamato immediatamente `self.check_for_autobackup()` ([dataflow.py:1080](../../dataflow.py#L1080)).

C. Decisione “devo fare backup ora?”

In `check_for_autobackup`:

- legge config corrente (`enabled/path/hour`) ([dataflow.py:1325](../../dataflow.py#L1325));
- condizione trigger: `now.hour == int(hour) and now.date() != self.last_backup_date` ([dataflow.py:1329](../../dataflow.py#L1329));
- se vera: esegue backup e poi imposta `self.last_backup_date = now.date()` ([dataflow.py:1330](../../dataflow.py#L1330)).

D. Scheduling ricorrente

- a fine check, schedula nuovo controllo dopo 60 secondi con `after` ([dataflow.py:1343](../../dataflow.py#L1343)).

E. Backup materiale + retention

- `perform_autobackup` chiama `perform_autobackup_copy` ([dataflow.py:1361](../../dataflow.py#L1361)).
- `perform_autobackup_copy` tenta pruning pre-creazione ([services/settings_maintenance_service.py:80](../../services/settings_maintenance_service.py#L80)).
- crea file `.db` timestampato ([services/settings_maintenance_service.py:101](../../services/settings_maintenance_service.py#L101)).
- se presenti, copia anche `.db-wal` e `.db-shm` con lo stesso timestamp base ([services/settings_maintenance_service.py:122](../../services/settings_maintenance_service.py#L122), [services/settings_maintenance_service.py:134](../../services/settings_maintenance_service.py#L134)).

## 5) Root cause primaria

### 5.1 Trigger multiplo nello stesso giorno / ad ogni riapertura app

Root cause primaria confermata:

- lo stato “backup già fatto oggi” è solo `self.last_backup_date` in RAM ([dataflow.py:1042](../../dataflow.py#L1042), [dataflow.py:1331](../../dataflow.py#L1331));
- a ogni riavvio processo torna `None`;
- se si riapre l’app nella stessa ora configurata (`now.hour == hour`), la condizione torna vera e il backup riparte.

Evidenza oggettiva: non esiste nessun altro riferimento repository a `last_backup` oltre quei 3 punti in `dataflow.py`.

### 5.2 Retention che non limita a 3

Root cause primaria confermata nella retention:

- pattern glob costruito così:
  - `pattern = os.path.join(dest_folder, f"*_backup_auto_{ext.replace('*', '')}")` ([services/settings_maintenance_service.py:82](../../services/settings_maintenance_service.py#L82)).
- per `ext='*.db'` produce `*_backup_auto_.db` (senza wildcard tra `_backup_auto_` e `.db`).
- i file reali sono tipo `gestione_offerte_backup_auto_20260416_150101.db` ([services/settings_maintenance_service.py:102](../../services/settings_maintenance_service.py#L102)).
- risultato: i file esistenti non matchano, `backup_sets` resta vuoto, `while len(sorted_timestamps) >= 3` non elimina nulla.

## 6) Cause secondarie/contributive

- Trigger basato solo su ora (granularità per ora, non timestamp persistito):
  - `now.hour == int(hour)` ([dataflow.py:1329](../../dataflow.py#L1329)).
- Controllo invocato subito in startup:
  - `self.check_for_autobackup()` in init ([dataflow.py:1080](../../dataflow.py#L1080)); amplifica l’effetto del guard volatile.
- In ambienti multi-workstation sullo stesso DB/repository backup, ogni client con AutoBackup attivo può produrre backup propri (confermato anche dalla documentazione operativa), aggravando la percezione di “troppi backup”.

## 7) Impatto su retention max 3 copie

Distinzione richiesta:

- Trigger multiplo stesso giorno:
  - plausibile e supportato dal codice per riaperture nella stessa ora.
- Trigger ad ogni apertura app:
  - tecnicamente vero solo se l’apertura avviene nell’ora configurata; fuori da quell’ora non parte.
- Retention > 3:
  - indipendente dal trigger; è dovuta al mismatch del pattern di pruning, che di fatto disattiva la pulizia.

Conseguenza combinata:

- più trigger nella stessa giornata + pruning inefficace => crescita illimitata dei backup automatici.

## 8) Impatto dei file `.wal/.shm`

- `.db-wal` e `.db-shm` sono sidecar normali con SQLite in WAL mode:
  - WAL abilitato in connessione (`PRAGMA journal_mode=WAL`) ([database_manager.py:58](../../database_manager.py#L58)).
- Il codice backup li copia volontariamente se presenti ([services/settings_maintenance_service.py:120](../../services/settings_maintenance_service.py#L120), [services/settings_maintenance_service.py:132](../../services/settings_maintenance_service.py#L132)).
- Quindi la presenza di `.db-wal/.db-shm` non indica di per sé backup duplicato: può essere un singolo backup logico rappresentato da 1-3 file fisici.

Chiarimento requisito “max 3 copie”:

- Implementazione corrente intende chiaramente “3 set logici” (timestamp comune), non “3 file fisici totali”.
- Anche con retention corretta, 3 backup logici possono produrre fino a 9 file fisici (`.db`, `.db-wal`, `.db-shm`).

Nota tecnica ulteriore:

- il parser timestamp nel pruning (`split(...).rsplit('.', 1)[0]`) è coerente con l’idea di raggruppare `.db/.db-wal/.db-shm` per timestamp ([services/settings_maintenance_service.py:86](../../services/settings_maintenance_service.py#L86)); il problema è a monte nel `glob`.

## 9) Fix minimo consigliato (piano, non applicato)

Piano minimale, reversibile, basso rischio:

1. Persistenza stato ultimo autobackup riuscito
- Aggiungere in `[AutoBackup]` una chiave tipo `last_run_date=YYYY-MM-DD` (oppure `last_run_ts`).
- In `check_for_autobackup`, confrontare `now.date()` con valore persistito, non solo con variabile runtime.
- Aggiornare il valore persistito solo dopo backup completato con esito `copied=True`.

2. Correzione chirurgica pattern retention
- Correggere pattern `glob` per includere wildcard del timestamp (es. `*_backup_auto_*.db`, `*_backup_auto_*.db-wal`, `*_backup_auto_*.db-shm`).
- Lasciare invariata la logica di grouping/ordinamento già esistente.

3. Nessun refactor ampio
- Nessuna nuova dipendenza.
- Nessuna modifica alla UX.
- Nessuna modifica fuori dai due punti sopra.

## 10) Rischi del fix

- Se si persiste `last_run_date` senza validare il timezone/clock di sistema, eventuali cambi manuali dell’orologio possono alterare la frequenza percepita.
- Se si aggiorna `last_run_date` prima di verificare esito backup, si rischia salto del backup giornaliero in caso di errore copia.
- Correggendo il pruning, al primo backup utile verranno eliminati backup storici in eccesso: comportamento atteso ma da comunicare.

## 11) Test manuali da eseguire dopo il fix

1. Trigger in startup nella stessa ora
- Configurare ora backup = 15.
- Aprire app alle 15:01: deve creare 1 backup.
- Chiudere/riaprire alle 15:07, 15:11, 15:15: non deve creare altri backup nello stesso giorno.

2. Trigger timer entro stessa sessione
- Lasciare app aperta per tutta l’ora 15 con timer ogni 60s.
- Deve restare 1 backup per la data corrente.

3. Giorno successivo
- Avviare app il giorno dopo alle 15:xx.
- Deve creare esattamente 1 nuovo backup.

4. Retention rolling
- Forzare 5 esecuzioni giornaliere su date simulate o ambiente test.
- Verificare che restino solo 3 set logici più recenti.

5. Integrità set con sidecar
- In condizioni con WAL/SHM presenti, verificare che ogni set logico abbia nome timestamp coerente tra `.db/.db-wal/.db-shm`.
- Verificare che pruning elimini interamente i set più vecchi.

6. Caso senza sidecar
- Con WAL/SHM assenti (db chiuso/checkpoint), verificare che retention conti comunque correttamente i set `.db`.

---

## Mappatura richiesta (STEP A-E) sintetica

### STEP A — Mappatura codice

Entry point e callback:

- Startup app -> `MainWindow(root)` ([dataflow.py:3256](../../dataflow.py#L3256))
- Startup MainWindow -> `check_for_autobackup()` ([dataflow.py:1080](../../dataflow.py#L1080))
- Callback periodica -> `root.after(60000, self.check_for_autobackup)` ([dataflow.py:1343](../../dataflow.py#L1343))

Classi:

- `SettingsWindow` ([dataflow.py:216](../../dataflow.py#L216))
- `MainWindow` ([dataflow.py:1023](../../dataflow.py#L1023))

Funzioni servizio:

- `save_autobackup_preferences` ([services/settings_preferences_service.py:91](../../services/settings_preferences_service.py#L91))
- `read_autobackup_config` ([services/settings_maintenance_service.py:13](../../services/settings_maintenance_service.py#L13))
- `perform_autobackup_copy` ([services/settings_maintenance_service.py:72](../../services/settings_maintenance_service.py#L72))

### STEP B — Trigger giornaliero

- Decisione “ora di backup”: solo `now.hour == int(hour)` + confronto data in memoria ([dataflow.py:1329](../../dataflow.py#L1329)).
- Guard-rail persistente “ultimo backup” non esiste.
- Invocazioni verificate:
  - startup: sì
  - callback periodica: sì (ogni minuto)
  - refresh UI: no (nessuna chiamata diretta)
  - apertura finestra impostazioni: no (nessuna chiamata diretta)

### STEP C — Retention max 3

- Conta prevista per set logici (`backup_sets` con chiave timestamp) ma popolamento set fallisce per pattern glob errato.
- Pruning eseguito prima della creazione nuovo backup (scelta valida), ma inefficace per mismatch pattern.

### STEP D — SQLite WAL/SHM

- Sidecar attesi in WAL mode.
- Copia sidecar intenzionale e condizionale (`if os.path.exists(...)`).
- Influenza sulla percezione: un backup logico può apparire come 3 file fisici.

### STEP E — Riproducibilità logica dello scenario indicato

Scenario utente (ora=15) è plausibile da codice:

- Avvio 15:01 -> `last_backup_date=None` -> condizione vera -> backup.
- Chiusura/riapertura 15:07 -> nuova istanza `MainWindow`, `last_backup_date=None` -> backup.
- Chiusura/riapertura 15:11 -> stesso meccanismo -> backup.
- Chiusura/riapertura 15:15 -> stesso meccanismo -> backup.

Questa catena è direttamente giustificata da [dataflow.py:1042](../../dataflow.py#L1042), [dataflow.py:1080](../../dataflow.py#L1080), [dataflow.py:1329](../../dataflow.py#L1329).
