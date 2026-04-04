# Report Forense: Copia automatica in "Cambia Posizione DataFlow..."

---

## 1) FILE COINVOLTI

- `dataflow.py` — entry point UI, intera logica di cambio posizione, copia fisica, backup, restart
- `services/app_paths.py` — risoluzione percorsi, `get_user_documents_dataflow_dir()`, `get_db_path()`, `get_fixed_attachments_dir()`
- `services/startup_service.py` — `initialize_dataflow_directory_structure()`, log setup
- `database/db_helpers.py` — `crea_database_v4()`, apertura DB al riavvio

---

## 2) ENTRY POINT

**Funzione:** `SettingsWindow.select_standard_dataflow_location`
**File:** `dataflow.py` ~L553
**Triggerata da:** pulsante `📁 Cambia Posizione DataFlow...` definito in `SettingsWindow.__init__` ~L189
**Responsabilità dichiarata:** cambiare solo la cartella base di DataFlow e riavviare

---

## 3) FLUSSO COMPLETO

| # | Funzione / istruzione | File e riga indicativa | Cosa fa |
|---|---|---|---|
| 1 | `select_standard_dataflow_location` | `dataflow.py` ~L553 | Entry point; legge `current_dataflow_dir` dalla config corrente |
| 2 | Dialog di warning | `dataflow.py` ~L563 | Mostra avviso (paradossalmente dice "non verrà spostata automaticamente") |
| 3 | `filedialog.askdirectory` | `dataflow.py` ~L598 | L'utente sceglie la nuova cartella padre |
| 4 | `os.path.normpath / abspath` | `dataflow.py` ~L609 | Normalizza il percorso scelto |
| 5 | `os.makedirs(normalized_dir)` | `dataflow.py` ~L623 | **Crea subito** la cartella padre scelta |
| 6 | Test write-permission + path-length | `dataflow.py` ~L641 | Validazione |
| 7 | Loop username conflict | `dataflow.py` ~L718 | Gestisce conflitti `DataFlow_{username}`; imposta `source_folder` e `dest_folder` |
| 8 | `db_manager.close()` | `dataflow.py` ~L802 | Chiude il database prima della copia |
| 9 | `os.walk(source_folder)` ×n | `dataflow.py` ~L834 | Conta tutti i file della cartella sorgente (database + allegati + tutto) |
| 10 | **`copy_with_progress(source_folder, dest_folder)`** | `dataflow.py` ~L868 | **↓ PUNTO DELLA COPIA — vedi sezione 4** |
| 11 | `shutil.move(old_db, new_db)` | `dataflow.py` ~L887 | Solo se l'utente ha cambiato username: rinomina il DB copiato |
| 12 | `config['Settings']['dataflow_base_dir'] = dest_parent` | `dataflow.py` ~L914 | Salva il nuovo percorso in `config.ini` |
| 13 | Rimozione voce `custom_db_path` legacy | `dataflow.py` ~L916 | Rimuove eventuale percorso personalizzato obsoleto |
| 14 | Messaggio successo | `dataflow.py` ~L940 | Conferma "cartella copiata con successo" |
| 15 | `reset_db_cache()` + `restart_program()` | `dataflow.py` ~L966 | Invalida cache percorsi, riavvia l'app |
| 16 | **Al riavvio:** `main_task()` | `dataflow.py` ~L4397 | Esegue il bootstrap standard |
| 17 | `get_user_documents_dataflow_dir()` | `services/app_paths.py` ~L20 | Legge `dataflow_base_dir` dal config → risolve la **nuova** cartella |
| 18 | `initialize_dataflow_directory_structure` | `services/startup_service.py` ~L92 | Non chiamata: `Database/` e `Attachments/` esistono già (già copiate) |
| 19 | `crea_database_v4()` | `database/db_helpers.py` ~L14 | Apre il DB (già copiato) nella nuova posizione, crea tabelle solo se mancanti |

---

## 4) PUNTO ESATTO IN CUI AVVIENE LA COPIA

**Funzione:** `copy_with_progress` (closure locale)
**File:** `dataflow.py`
**Righe indicative:** 845–868

Struttura della copia:

```python
copy_with_progress(source_folder, dest_folder)   # riga 868
  os.makedirs(dst, exist_ok=True)                # riga 847 — crea ogni sottocartella
  for item in os.listdir(src):
      if isdir:
          copy_with_progress(s, d)               # riga 854 — ricorsione
      else:
          shutil.copy2(s, d)                     # riga 857 — copia ogni singolo file
```

`source_folder` = cartella `DataFlow_{username}` **corrente** (con `Database/` e `Attachments/` dentro).
`dest_folder` = `<nuova_cartella_padre>/DataFlow_{username}`.
Non esiste alcun filtro: tutti i file vengono copiati incondizionatamente.

---

## 5) CONDIZIONE CHE FA SCATTARE LA COPIA

**Sempre**, senza eccezioni, non appena l'utente conferma il dialogo di warning.

L'unica condizione che impedisce la copia è che la cartella sorgente sia **completamente vuota** (riga 840 lancia un'eccezione se `total_files == 0`). In qualsiasi altro scenario (anche sorgente con un solo file) la copia viene eseguita integralmente prima del riavvio.

La copia avviene **immediatamente** (scenario A), non al riavvio.

---

## 6) VERDETTO FINALE

**Perché avviene la copia automatica:**

La funzione `select_standard_dataflow_location` implementa un **"sposta con copia"**: calcola sorgente e destinazione, chiude il DB, copia ricorsivamente l'intero albero `DataFlow_{username}` (database SQLite + tutti gli allegati + qualsiasi altro file) con `shutil.copy2`, aggiorna il config, poi riavvia. Non è un effetto collaterale di bootstrap né di migrazione: è **logica esplicita e intenzionale** nelle righe 845–868.

**Incoerenza con la documentazione/UI:**

| Elemento | Testo / Comportamento |
|---|---|
| Warning iniziale (`dataflow.py` ~L568) | *"La cartella attuale non verrà spostata automaticamente"* |
| Messaggio di successo (`dataflow.py` ~L942) | *"La cartella DataFlow è stata copiata con successo"* |
| Comportamento reale | Copia completa, incondizionata, immediata |

Il warning iniziale descrive un **comportamento precedente** (cambio solo del percorso nel config, senza toccare i file), mentre in un momento successivo è stata aggiunta la logica di copia fisica (`copy_with_progress`) senza aggiornare il testo del warning. Il codice e il messaggio di successo sono coerenti tra loro; è il **testo del warning iniziale** a essere rimasto indietro rispetto all'implementazione attuale.

---

## Appendice — Tutte le operazioni shutil nel codice applicativo

| File | Riga | Operazione | Contesto |
|---|---|---|---|
| `dataflow.py` | ~L857 | `shutil.copy2(s, d)` | Inside `copy_with_progress` — copia ogni file in `DataFlow_{username}` ricorsivamente |
| `dataflow.py` | ~L887 | `shutil.move(old_db_path, new_db_path)` | Rinomina DB se username cambia |
| `dataflow.py` | ~L448 | `shutil.copy2(db_file, dest)` | Backup manuale DB |
| `dataflow.py` | ~L455 | `shutil.copy2(wal_file, wal_dest)` | Backup manuale WAL |
| `dataflow.py` | ~L464 | `shutil.copy2(shm_file, shm_dest)` | Backup manuale SHM |
| `dataflow.py` | ~L1499 | `shutil.copy2(db_file, dest_path)` | Auto-backup DB |
| `dataflow.py` | ~L1515 | `shutil.copy2(wal_file, wal_dest)` | Auto-backup WAL |
| `dataflow.py` | ~L1527 | `shutil.copy2(shm_file, shm_dest)` | Auto-backup SHM |
| `services/app_paths.py` | ~L110 | `shutil.move(old_dir, new_dir)` | Migrazione "Allegati" → "Attachments" |
| `services/startup_service.py` | ~L114 | `shutil.move(old_attachments_dir, new_attachments_dir)` | Stessa migrazione (percorso duplicato) |
| `ui/windows/attachment_window.py` | ~L380 | `shutil.copy(filepath, dest_path)` | Allega un file a una RFQ |
| `ui/windows/attachment_window.py` | ~L565 | `shutil.copy(real_full, save_path)` | Esporta un allegato |

> **Nota:** `shutil.copytree` non è usato in nessun punto del codice applicativo.
