# VSM – ANALISI TRANSAZIONI PERSISTENZA (CRITICO)

## Data Analisi
23 marzo 2026

## 🔴 PROBLEMA IDENTIFICATO: OPERAZIONI NON ATOMICHE

### Sintesi Esecutiva
Le operazioni di persistenza VSM **NON sono atomiche**. Ogni operazione esegue **COMMIT multipli separati** invece di un'unica transazione atomica, creando rischio di **corruzione dati** in caso di fallimenti parziali.

---

## 1. ANALISI SITUAZIONE ATTUALE

### 1.1 Operazione: `save_event_with_impacts()`

**File:** `services/vsm_persistence.py` (linee 44-90)

**Flusso attuale:**
```python
def save_event_with_impacts(db_manager, event: VSMEvent, impacts: List[VSMImpact]) -> int:
    # COMMIT #1: Salva evento
    event_id = db_manager.insert_vsm_event(...)  # ← COMMIT alla linea 1877 di database_manager.py
    
    # COMMIT #2: Salva impacts
    db_manager.insert_vsm_impacts_batch(impacts)  # ← COMMIT alla linea 2018 di database_manager.py
    
    return event_id
```

**Problema:**
- Se `insert_vsm_event()` ha successo → **COMMIT eseguito**
- Se `insert_vsm_impacts_batch()` fallisce → **evento orfano nel database** senza impacts associati

**Scenario di fallimento:**
1. Evento VSM salvato con successo (COMMITTED)
2. Vincolo di validazione fallisce sugli impacts (es. value_type non valido)
3. **RISULTATO: Evento salvato senza impacts → DATI INCONSISTENTI**

---

### 1.2 Operazione: `update_event_with_impacts()`

**File:** `services/vsm_persistence.py` (linee 91-170)

**Flusso attuale:**
```python
def update_event_with_impacts(db_manager, event: VSMEvent, impacts: List[VSMImpact]) -> None:
    # COMMIT #1: Aggiorna evento
    db_manager.update_vsm_event(...)  # ← COMMIT alla linea 1922 di database_manager.py
    
    # COMMIT #2: Cancella impacts esistenti
    db_manager.delete_vsm_impacts_by_event_id(event.id)  # ← COMMIT alla linea 2031
    
    # COMMIT #3: Salva nuovi impacts
    db_manager.insert_vsm_impacts_batch(impacts)  # ← COMMIT alla linea 2018
```

**Problema:**
- **3 COMMIT separati** = 3 punti di fallimento potenziali
- Pattern DELETE-REGENERATE-SAVE richiesto ma non protetto da transazione

**Scenari di fallimento:**

**Scenario A:**
1. `update_vsm_event()` successo → **COMMIT #1 eseguito**
2. `delete_vsm_impacts_by_event_id()` successo → **COMMIT #2 eseguito**
3. `insert_vsm_impacts_batch()` fallisce (es. vincolo NOT NULL)
4. **RISULTATO: Evento aggiornato ma TUTTI GLI IMPACTS CANCELLATI definitivamente**

**Scenario B:**
1. `update_vsm_event()` successo → **COMMIT #1 eseguito**
2. `delete_vsm_impacts_by_event_id()` fallisce (es. errore I/O)
3. **RISULTATO: Evento modificato ma impacts rimangono nella versione precedente**

---

## 2. CODICE PROBLEMATICO NEL DETTAGLIO

### 2.1 database_manager.py - Metodi con autocommit

#### `insert_vsm_event()` (linee 1813-1880)
```python
def insert_vsm_event(self, event_name, category, process_code, date_str, 
                     description=None, responsible_person=None, status=None):
    cursor = self.conn.cursor()
    cursor.execute("""
        INSERT INTO vsm_events (...) VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (...))
    
    event_id = cursor.lastrowid
    self.conn.commit()  # ← LINEA 1877: COMMIT AUTOMATICO
    return event_id
```

#### `update_vsm_event()` (linee 1881-1950)
```python
def update_vsm_event(self, event_id, event_name, category, process_code, 
                     date_str, description=None, responsible_person=None, status=None):
    cursor = self.conn.cursor()
    cursor.execute("""
        UPDATE vsm_events SET ... WHERE id = ?
    """, (..., event_id))
    
    if cursor.rowcount == 0:
        raise ValueError(f"VSM Event with id {event_id} not found")
    
    self.conn.commit()  # ← LINEA 1922: COMMIT AUTOMATICO
```

#### `delete_vsm_impacts_by_event_id()` (linee 2020-2040)
```python
def delete_vsm_impacts_by_event_id(self, event_id: int) -> int:
    cursor = self.conn.cursor()
    cursor.execute("DELETE FROM vsm_impacts WHERE event_id = ?", (event_id,))
    deleted_count = cursor.rowcount
    self.conn.commit()  # ← LINEA 2031: COMMIT AUTOMATICO
    return deleted_count
```

#### `insert_vsm_impacts_batch()` (linee 1999-2025)
```python
def insert_vsm_impacts_batch(self, impacts: List[dict]) -> None:
    if not impacts:
        return
    
    cursor = self.conn.cursor()
    cursor.executemany("""
        INSERT INTO vsm_impacts (...) VALUES (?, ?, ?, ?)
    """, tuples_list)
    
    self.conn.commit()  # ← LINEA 2018: COMMIT AUTOMATICO
```

---

## 3. SOLUZIONE PROPOSTA

### 3.1 Strategia: Helper Methods senza Autocommit

Creare **4 metodi helper privati** in `database_manager.py` che eseguono operazioni SENZA commit, permettendo alle funzioni di persistenza di gestire manualmente le transazioni.

### 3.2 Nuovi Helper Methods da aggiungere

```python
# DatabaseManager - Aggiungi dopo la linea 2111

def _insert_vsm_event_no_commit(self, event_name, category, process_code, date_str,
                                description=None, responsible_person=None, status=None):
    """
    Inserisce un evento VSM SENZA fare commit.
    Usato per operazioni transazionali.
    """
    cursor = self.conn.cursor()
    cursor.execute("""
        INSERT INTO vsm_events (event_name, category, process_code, date, 
                               description, responsible_person, status)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (event_name, category, process_code, date_str, 
          description, responsible_person, status))
    
    return cursor.lastrowid


def _update_vsm_event_no_commit(self, event_id, event_name, category, process_code,
                                date_str, description=None, responsible_person=None, status=None):
    """
    Aggiorna un evento VSM SENZA fare commit.
    Usato per operazioni transazionali.
    """
    cursor = self.conn.cursor()
    cursor.execute("""
        UPDATE vsm_events 
        SET event_name = ?, category = ?, process_code = ?, date = ?,
            description = ?, responsible_person = ?, status = ?
        WHERE id = ?
    """, (event_name, category, process_code, date_str,
          description, responsible_person, status, event_id))
    
    if cursor.rowcount == 0:
        raise ValueError(f"VSM Event with id {event_id} not found")


def _delete_vsm_impacts_no_commit(self, event_id: int) -> int:
    """
    Cancella gli impacts di un evento SENZA fare commit.
    Usato per operazioni transazionali.
    """
    cursor = self.conn.cursor()
    cursor.execute("DELETE FROM vsm_impacts WHERE event_id = ?", (event_id,))
    return cursor.rowcount


def _insert_vsm_impacts_batch_no_commit(self, impacts: List[dict]) -> None:
    """
    Inserisce un batch di impacts SENZA fare commit.
    Usato per operazioni transazionali.
    """
    if not impacts:
        return
    
    cursor = self.conn.cursor()
    tuples_list = [
        (imp['event_id'], imp['year'], imp['month'], imp['value_type'])
        for imp in impacts
    ]
    cursor.executemany("""
        INSERT INTO vsm_impacts (event_id, year, month, value_type)
        VALUES (?, ?, ?, ?)
    """, tuples_list)
```

---

### 3.3 Modifica: `save_event_with_impacts()` con Transazione Atomica

```python
# services/vsm_persistence.py

def save_event_with_impacts(db_manager, event: VSMEvent, impacts: List[VSMImpact]) -> int:
    """
    Salva un evento VSM con i suoi impacts in UNA SINGOLA TRANSAZIONE ATOMICA.
    
    Se qualsiasi operazione fallisce, viene eseguito ROLLBACK completo.
    """
    # Validazioni preliminari
    if not event.event_name or not event.event_name.strip():
        raise ValueError("event_name non può essere vuoto")
    if not event.category or not event.category.strip():
        raise ValueError("category non può essere vuota")
    # ... altre validazioni ...
    
    # Prepara dati impacts
    impacts_data = [
        {
            'event_id': None,  # Sarà impostato dopo l'inserimento dell'evento
            'year': imp.year,
            'month': imp.month,
            'value_type': imp.value_type.value
        }
        for imp in impacts
    ]
    
    # ========================================
    # TRANSAZIONE ATOMICA
    # ========================================
    try:
        # Inizia transazione esplicita
        db_manager.conn.execute("BEGIN TRANSACTION")
        
        # Step 1: Inserisci evento (SENZA COMMIT)
        event_id = db_manager._insert_vsm_event_no_commit(
            event_name=event.event_name,
            category=event.category,
            process_code=event.process_code,
            date_str=event.date.isoformat(),
            description=event.description,
            responsible_person=event.responsible_person,
            status=event.status.value if event.status else None
        )
        
        # Step 2: Aggiorna event_id negli impacts
        for imp_data in impacts_data:
            imp_data['event_id'] = event_id
        
        # Step 3: Inserisci impacts (SENZA COMMIT)
        if impacts_data:
            db_manager._insert_vsm_impacts_batch_no_commit(impacts_data)
        
        # COMMIT UNICO: tutte le operazioni hanno successo
        db_manager.conn.commit()
        
        return event_id
        
    except Exception as e:
        # ROLLBACK: annulla TUTTE le operazioni
        db_manager.conn.rollback()
        raise RuntimeError(f"Errore durante il salvataggio atomico: {e}") from e
```

---

### 3.4 Modifica: `update_event_with_impacts()` con Transazione Atomica

```python
# services/vsm_persistence.py

def update_event_with_impacts(db_manager, event: VSMEvent, impacts: List[VSMImpact]) -> None:
    """
    Aggiorna un evento VSM e i suoi impacts in UNA SINGOLA TRANSAZIONE ATOMICA.
    
    Pattern: DELETE-REGENERATE-SAVE protetto da transazione.
    Se qualsiasi operazione fallisce, viene eseguito ROLLBACK completo.
    """
    if event.id is None:
        raise ValueError("L'evento deve avere un ID per essere aggiornato")
    
    # Validazioni preliminari
    if not event.event_name or not event.event_name.strip():
        raise ValueError("event_name non può essere vuoto")
    if not event.category or not event.category.strip():
        raise ValueError("category non può essere vuota")
    # ... altre validazioni ...
    
    # Prepara dati impacts
    impacts_data = [
        {
            'event_id': event.id,
            'year': imp.year,
            'month': imp.month,
            'value_type': imp.value_type.value
        }
        for imp in impacts
    ]
    
    # ========================================
    # TRANSAZIONE ATOMICA
    # ========================================
    try:
        # Inizia transazione esplicita
        db_manager.conn.execute("BEGIN TRANSACTION")
        
        # Step 1: Aggiorna evento (SENZA COMMIT)
        db_manager._update_vsm_event_no_commit(
            event_id=event.id,
            event_name=event.event_name,
            category=event.category,
            process_code=event.process_code,
            date_str=event.date.isoformat(),
            description=event.description,
            responsible_person=event.responsible_person,
            status=event.status.value if event.status else None
        )
        
        # Step 2: Cancella impacts esistenti (SENZA COMMIT)
        db_manager._delete_vsm_impacts_no_commit(event.id)
        
        # Step 3: Inserisci nuovi impacts (SENZA COMMIT)
        if impacts_data:
            db_manager._insert_vsm_impacts_batch_no_commit(impacts_data)
        
        # COMMIT UNICO: tutte le operazioni hanno successo
        db_manager.conn.commit()
        
    except Exception as e:
        # ROLLBACK: ripristina stato precedente
        db_manager.conn.rollback()
        raise RuntimeError(f"Errore durante l'aggiornamento atomico: {e}") from e
```

---

## 4. VANTAGGI DELLA SOLUZIONE

### 4.1 Atomicità Garantita
- **Una transazione = un commit**: tutte le operazioni riescono o falliscono insieme
- **Rollback automatico**: nessun dato parziale in caso di errore
- **Integrità referenziale**: eventi e impacts sempre sincronizzati

### 4.2 Compatibilità Retroattiva
- I metodi pubblici esistenti (`insert_vsm_event`, `update_vsm_event`, ecc.) **rimangono invariati**
- Codice esistente continua a funzionare con autocommit
- Nuovi helper privati (`_*_no_commit`) usati solo per operazioni atomiche

### 4.3 Pattern DELETE-REGENERATE-SAVE Protetto
- Cancellazione e reinserimento impacts avviene in transazione unica
- Impossibile avere evento senza impacts dopo update fallito
- Stato consistente sempre garantito

---

## 5. TEST NECESSARI

### 5.1 Test Rollback su Save
```python
def test_save_event_with_impacts_rollback_on_impact_error(db_manager_transactional):
    """
    Verifica che se l'inserimento degli impacts fallisce,
    anche l'evento venga annullato (rollback completo).
    """
    event = VSMEvent(...)
    
    # Impact con value_type NON VALIDO per forzare errore
    impacts = [VSMImpact(year=2025, month=1, value_type="INVALID_TYPE")]
    
    with pytest.raises(Exception):
        save_event_with_impacts(db_manager_transactional, event, impacts)
    
    # VERIFICA: nessun evento salvato nel database
    events = db_manager_transactional.get_all_vsm_events()
    assert len(events) == 0, "Evento non dovrebbe esistere dopo rollback"
```

### 5.2 Test Rollback su Update
```python
def test_update_event_rollback_on_impact_error(db_manager_transactional):
    """
    Verifica che se l'aggiornamento degli impacts fallisce,
    l'evento e gli impacts rimangano nel loro stato precedente.
    """
    # Setup: evento con 2 impacts esistenti
    event_id = save_event_with_impacts(...)
    original_event = db_manager_transactional.get_vsm_event_by_id(event_id)
    original_impacts = db_manager_transactional.get_vsm_impacts_by_event_id(event_id)
    
    # Tentativo update con impact invalido
    updated_event = VSMEvent(id=event_id, event_name="MODIFIED", ...)
    invalid_impacts = [VSMImpact(year=2025, month=1, value_type="INVALID")]
    
    with pytest.raises(Exception):
        update_event_with_impacts(db_manager_transactional, updated_event, invalid_impacts)
    
    # VERIFICA: evento e impacts rimangono nella versione originale
    current_event = db_manager_transactional.get_vsm_event_by_id(event_id)
    assert current_event.event_name == original_event.event_name
    
    current_impacts = db_manager_transactional.get_vsm_impacts_by_event_id(event_id)
    assert len(current_impacts) == len(original_impacts)
```

---

## 6. PRIORITÀ IMPLEMENTAZIONE

### 🔴 CRITICO - BLOCCANTE PER PRODUZIONE
**La mancanza di transazioni atomiche è un bug di severità ALTA:**
- Può causare perdita di dati (impacts cancellati ma non rigenerati)
- Può creare eventi orfani (evento senza impacts)
- Può causare inconsistenze nei report SQDC
- Viola il principio di integrità referenziale

### ⚠️ RACCOMANDAZIONE
**Implementare il fix PRIMA di qualsiasi deploy in produzione**
1. Aggiungere i 4 helper methods a `database_manager.py`
2. Modificare `save_event_with_impacts()` con gestione transazionale
3. Modificare `update_event_with_impacts()` con gestione transazionale
4. Eseguire test suite completa (12 test esistenti + 2 nuovi test rollback)
5. Validare con test di integrazione

---

## 7. CHECKLIST IMPLEMENTAZIONE

- [ ] Aggiungere `_insert_vsm_event_no_commit()` a database_manager.py
- [ ] Aggiungere `_update_vsm_event_no_commit()` a database_manager.py
- [ ] Aggiungere `_delete_vsm_impacts_no_commit()` a database_manager.py
- [ ] Aggiungere `_insert_vsm_impacts_batch_no_commit()` a database_manager.py
- [ ] Modificare `save_event_with_impacts()` con BEGIN/COMMIT/ROLLBACK
- [ ] Modificare `update_event_with_impacts()` con BEGIN/COMMIT/ROLLBACK
- [ ] Aggiungere test rollback per save operation
- [ ] Aggiungere test rollback per update operation
- [ ] Eseguire test suite completa (pytest tests/test_vsm_persistence.py -v)
- [ ] Validare che i 12 test esistenti passino ancora
- [ ] Documentare comportamento transazionale nel codice
- [ ] Code review con focus su error handling

---

## 8. RISCHI RESIDUI (POST-FIX)

### Scenario: Interruzione processo durante transazione
- **Sintomo**: Applicazione termina forzatamente durante BEGIN...COMMIT
- **Mitigazione**: SQLite garantisce rollback automatico al riavvio
- **Impatto**: Nessuna corruzione dati, operazione semplicemente non completata

### Scenario: Errore di validazione non catchato
- **Sintomo**: Exception non gestita propagata al chiamante
- **Mitigazione**: Try/except con rollback esplicito già implementato
- **Impatto**: Transazione annullata, errore loggato, utente notificato

---

## CONCLUSIONE

L'implementazione attuale **NON è production-ready** a causa della mancanza di atomicità transazionale. Il fix proposto con helper methods è **minimamente invasivo**, **retrocompatibile**, e **risolve completamente** il problema garantendo integrità dei dati.

**Stato: FIX PRONTO PER IMPLEMENTAZIONE**

---

*Documento generato: 23 marzo 2026*
*Autore analisi: GitHub Copilot (Claude Sonnet 4.5)*
*Repository: /home/guido/Repository/vsm*
