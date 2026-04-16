# Report tecnico — Fix minimo backup automatico giornaliero

Data: 2026-04-16  
Ambito: patch minima e reversibile (senza refactor ampio)

## 1) File modificati

- `/home/guido/Repository/dataflow-procurement-software/dataflow.py`
- `/home/guido/Repository/dataflow-procurement-software/services/settings_maintenance_service.py`

Nessun altro file applicativo è stato modificato per il fix.

## 2) Diff sintetico per file

### `dataflow.py`

- Importate due nuove helper dal service manutenzione backup:
  - `read_last_autobackup_date`
  - `save_last_autobackup_date`
- In `check_for_autobackup`:
  - lettura `config_file` una sola volta;
  - lettura della data persistita `last_run_date`;
  - guard-rail giornaliero basato su data persistita invece che solo runtime;
  - salvataggio `last_run_date` solo dopo esito positivo di `perform_autobackup`.
- In `perform_autobackup`:
  - ritorno booleano esplicito (`True` solo su backup riuscito, `False` su skip/fail).

### `services/settings_maintenance_service.py`

- Aggiunte helper minime:
  - `read_last_autobackup_date(config_file)`
  - `save_last_autobackup_date(config_file, run_date)`
- Robustezza:
  - parsing data `YYYY-MM-DD` con fallback sicuro a `None` se mancante/invalida.
- Retention:
  - corretto pattern glob per intercettare realmente i backup:
    - `*_backup_auto_*.db`
    - `*_backup_auto_*.db-wal`
    - `*_backup_auto_*.db-shm`
  - mantenuta logica esistente di pruning a set logici (timestamp condiviso).

## 3) Spiegazione precisa del fix

### A) Persistenza ultimo autobackup riuscito

- È stata introdotta la chiave `[AutoBackup].last_run_date` (formato `YYYY-MM-DD`) su `config.ini`.
- Viene scritta solo dopo backup automatico completato con successo.
- In caso di errore/skip/non-copia non viene aggiornata.

### B) Trigger giornaliero corretto

- La logica oraria è stata mantenuta (`now.hour == int(hour)`).
- Il vincolo “una sola esecuzione al giorno” ora usa la data persistita (`last_run_date`) come fonte di verità.
- Riavvi multipli dell’app nella stessa ora non causano nuove esecuzioni nello stesso giorno.

### C) Retention rolling max 3 set logici

- Il pruning ora trova i file reali perché il pattern glob è corretto.
- La retention continua a ragionare per set logici basati su timestamp condiviso.
- Se oltre soglia, elimina il set più vecchio (inclusi sidecar `.db-wal`/`.db-shm` se presenti).

## 4) Rischi residui

- Se il clock sistema/data cambia manualmente, il comportamento giornaliero dipende dalla data locale risultante.
- Se in un ambiente multi-workstation più client condividono la stessa destinazione backup e stessa configurazione, possono continuare a produrre backup distinti per host/sessione (comportamento extra-scope, non modificato).
- `last_run_date` è una guardia giornaliera (non oraria/minutaria), coerente con requisito “1 backup al giorno”.

## 5) Test manuali consigliati

1. Ora backup `15`, avvio app `15:01` -> crea 1 backup.
2. Chiusura/riapertura `15:07`, `15:11`, `15:15` -> nessun nuovo backup nello stesso giorno.
3. Giorno successivo alle `15:xx` -> crea 1 nuovo backup.
4. Oltre 3 set logici -> restano solo i 3 più recenti.
5. Con `.db-wal`/`.db-shm` presenti -> pruning elimina correttamente set completo vecchio.
6. Senza sidecar -> pruning continua a funzionare sui `.db`.

## 6) Conformità vincoli richiesti

- Nessuna nuova dipendenza introdotta.
- Nessun refactor ampio.
- Nessuna modifica UX/UI.
- Nessun merge/PR eseguito.
- Modifiche limitate allo scope autorizzato (trigger persistito + retention).

