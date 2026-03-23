# Global Search Implementation - Proposta Tecnica

**Data**: 22 marzo 2026  
**Obiettivo**: Rendere la global search bar realmente globale (ricerca multi-campo)  
**Approccio**: Conservativo, minimale, reversibile

---

## 📊 ANALISI DEL COMPORTAMENTO ATTUALE

### File Coinvolti

1. **[ui/components/main_dashboard_toolbar.py](../ui/components/main_dashboard_toolbar.py)** (linee 1-200)
   - Contiene `MainDashboardToolbar` con la global search bar
   - Entry widget: `self.search_entry`
   - Placeholder: `"Search anything... (RFQ, Supplier, Code, Project...)"`
   - Callback Enter: `_on_search()`

2. **[dataflow.py](../dataflow.py)** (linee 3640-3680, 4664-4900)
   - Contiene `search_requests()` - funzione principale di ricerca
   - Inizializzazione `search_vars` con tutti i filtri

### Flusso Attuale della Ricerca

```
User digita testo → Preme Enter 
     ↓
MainDashboardToolbar._on_search()
     ↓
Se vuoto: main_window.clear_filters()
Se pieno: main_window.search_vars['num'].set(search_text)
     ↓
main_window.search_requests()
     ↓
Costruisce query SQL con AND tra tutti i campi valorizzati
```

### Campi di Ricerca Disponibili

| Key | Campo DB | Tabella | Descrizione |
|-----|----------|---------|-------------|
| `num` | `id_richiesta` | `richieste_offerta` | **Numero RfQ** (attuale target) |
| `ref` | `riferimento` | `richieste_offerta` | Riferimento/Progetto |
| `forn` | `nome_fornitore` | `richiesta_fornitori` | Fornitore |
| `cod` | `codice_materiale` | `dettagli_richiesta` | Codice Materiale |
| `desc` | `descrizione_materiale` | `dettagli_richiesta` | Descrizione Materiale |
| `ord` | `numeri_ordine` | `richieste_offerta` | Numero Ordine |
| `cod_grezzo` | `codice_grezzo` | `dettagli_richiesta` | Codice Grezzo |
| `dis_grezzo` | `disegno_grezzo` | `dettagli_richiesta` | Allegato Grezzo |
| `mat_cl` | `materiale_conto_lavoro` | `dettagli_richiesta` | Materiale c/lavoro |

### 🎯 LIMITE ATTUALE CONFERMATO

**Problema**: La global search bar scrive **SOLO** nel filtro `search_vars['num']` (numero RfQ).

**Effetto**: 
- Placeholder promette "Search anything..." ma cerca solo per numero
- Altri campi (fornitore, codice, progetto) **non vengono cercati**
- La ricerca è **limitata** nonostante il placeholder suggerisca il contrario

**Logica SQL Attuale** (AND tra criteri):
```python
# In search_requests(), ogni campo valorizzato aggiunge una clausola AND
if crit['num']: clauses.append("CAST(ro.id_richiesta AS TEXT) LIKE ?")
if crit['ref']: clauses.append("LOWER(ro.riferimento) LIKE LOWER(?)")
if crit['forn']: clauses.append("LOWER(rf.nome_fornitore) LIKE LOWER(?)")
# ... eccetera (clausole combinate con AND)
```

**Query SQL Base**:
```sql
SELECT DISTINCT ro.id_richiesta, ro.tipo_rdo, ro.data_emissione, 
                ro.data_scadenza, ro.riferimento 
FROM richieste_offerta ro 
LEFT JOIN dettagli_richiesta dr ON ro.id_richiesta=dr.id_richiesta 
LEFT JOIN richiesta_fornitori rf ON ro.id_richiesta=rf.id_richiesta
WHERE ro.stato=? 
  AND CAST(ro.id_richiesta AS TEXT) LIKE ?  -- solo se num è valorizzato
  AND LOWER(ro.riferimento) LIKE LOWER(?)   -- solo se ref è valorizzato
  -- ... (AND tra tutti i filtri)
```

**Conclusione**: Il sistema è configurato per **filtraggio combinato** (AND), non ricerca globale testuale (OR).

---

## 📋 STRATEGIA DI ESTENSIONE CONSERVATIVA

### Approccio Scelto: Global Query con OR Logic

Introdurre un nuovo campo virtuale: **`global`** che attiva una logica OR su campi selezionati, mantenendo intatta la logica AND esistente per i filtri avanzati.

### Principi della Soluzione

1. **Separazione delle Preoccupazioni**
   - Global search bar → `search_vars['global']` (ricerca testuale ampia con OR)
   - Filtri avanzati → `search_vars['num']`, `['ref']`, etc. (filtraggio preciso con AND)
   - I due sistemi possono **coesistere** o funzionare indipendentemente

2. **Campi Target per Global Search** (scelti per rilevanza e stabilità)
   - ✅ `num` - Numero RfQ (ro.id_richiesta)
   - ✅ `ref` - Riferimento/Progetto (ro.riferimento)
   - ✅ `forn` - Fornitore (rf.nome_fornitore)
   - ✅ `cod` - Codice Materiale (dr.codice_materiale)
   - ✅ `desc` - Descrizione Materiale (dr.descrizione_materiale)
   - ✅ `ord` - Numero Ordine (ro.numeri_ordine)
   - ❌ Campi specializzati esclusi: `cod_grezzo`, `dis_grezzo`, `mat_cl`

3. **Logica di Combinazione**
   ```sql
   WHERE ro.stato = ? 
   AND (
       -- Global search con OR (se presente crit['global'])
       (CAST(ro.id_richiesta AS TEXT) LIKE ? 
        OR LOWER(ro.riferimento) LIKE LOWER(?) 
        OR LOWER(rf.nome_fornitore) LIKE LOWER(?)
        OR LOWER(dr.codice_materiale) LIKE LOWER(?)
        OR LOWER(dr.descrizione_materiale) LIKE LOWER(?)
        OR LOWER(ro.numeri_ordine) LIKE LOWER(?))
   )
   AND (
       -- Filtri specifici con AND (se presenti)
       username = ? 
       AND tipo_rdo = ?
       AND data_emissione >= ? 
       -- ...
   )
   ```

### Comportamento Desiderato

| Scenario | Comportamento |
|----------|---------------|
| Global search vuota + filtri vuoti | Mostra tutti i record (stato corrente) |
| Global search piena + filtri vuoti | OR su 6 campi principali |
| Global search vuota + filtri pieni | AND tra filtri (comportamento attuale) |
| Global search piena + filtri pieni | OR sui 6 campi + AND con altri filtri |

### Esempio Pratico

**Caso 1**: Utente cerca "ACME" nella global search bar
```sql
WHERE ro.stato = 'attivo'
AND (
    CAST(ro.id_richiesta AS TEXT) LIKE '%ACME%'
    OR LOWER(ro.riferimento) LIKE LOWER('%ACME%')
    OR LOWER(rf.nome_fornitore) LIKE LOWER('%ACME%')
    OR LOWER(dr.codice_materiale) LIKE LOWER('%ACME%')
    OR LOWER(dr.descrizione_materiale) LIKE LOWER('%ACME%')
    OR LOWER(ro.numeri_ordine) LIKE LOWER('%ACME%')
)
```
**Risultati**: Tutte le RfQ che contengono "ACME" in qualsiasi dei 6 campi principali.

**Caso 2**: Utente cerca "ACME" + filtro data emissione "01/01/2026" → "31/03/2026"
```sql
WHERE ro.stato = 'attivo'
AND (
    CAST(ro.id_richiesta AS TEXT) LIKE '%ACME%'
    OR LOWER(ro.riferimento) LIKE LOWER('%ACME%')
    OR ... 
)
AND ro.data_emissione >= '2026-01-01'
AND ro.data_emissione <= '2026-03-31'
```
**Risultati**: RfQ con "ACME" nei campi principali + emesse nel Q1 2026.

---

## 🛠️ IMPLEMENTAZIONE MINIMALE

### Modifiche da Applicare

#### **MODIFICA 1**: Aggiungere campo `global` a search_vars

**File**: `dataflow.py` (linea 3643)

```python
# PRIMA:
self.search_vars = {name: tk.StringVar() for name in ['num', 'ref', 'forn', 'cod', 'desc', 'ord', 'cod_grezzo', 'dis_grezzo', 'mat_cl']}

# DOPO:
self.search_vars = {name: tk.StringVar() for name in ['global', 'num', 'ref', 'forn', 'cod', 'desc', 'ord', 'cod_grezzo', 'dis_grezzo', 'mat_cl']}
```

**Nota**: `global` è un campo "virtuale" - non ha widget UI nei filtri avanzati, è usato solo dalla global search bar.

---

#### **MODIFICA 2**: Estendere `search_requests()` per gestire ricerca globale

**File**: `dataflow.py` (circa linea 4718, dopo validazione, prima della costruzione clausole SQL)

**Logica da Implementare**:
```python
# Dopo la costruzione di clauses iniziali (stato, username, tipo, date)
# Prima dei filtri individuali (num, ref, forn, ...)

# Global Search con OR Logic
if crit['global']:
    global_query = crit['global']
    global_clauses = [
        "CAST(ro.id_richiesta AS TEXT) LIKE ?",
        "LOWER(ro.riferimento) LIKE LOWER(?)",
        "LOWER(rf.nome_fornitore) LIKE LOWER(?)",
        "LOWER(dr.codice_materiale) LIKE LOWER(?)",
        "LOWER(dr.descrizione_materiale) LIKE LOWER(?)",
        "LOWER(ro.numeri_ordine) LIKE LOWER(?)"
    ]
    clauses.append("(" + " OR ".join(global_clauses) + ")")
    # Aggiungi il parametro una volta per ogni campo OR
    for _ in global_clauses:
        params.append(f"%{global_query}%")
else:
    # Logica esistente per filtri individuali (AND)
    if crit['num']: clauses.append("CAST(ro.id_richiesta AS TEXT) LIKE ?"); params.append(f"%{crit['num']}%")
    if crit['ref']: clauses.append("LOWER(ro.riferimento) LIKE LOWER(?)"); params.append(f"%{crit['ref']}%")
    # ... (resto dei filtri come ora)
```

**Gestione Ricerca Aggregata Multi-DB**:

Per la modalità "All users", la logica in-memory deve essere estesa:
```python
# Nel blocco di filtro in-memory (dopo get_all_richieste_aggregated)
if crit['global']:
    # Verifica se la query globale matcha almeno un campo
    global_match = (
        crit['global'] in str(row[0]) or  # num
        (row[4] and crit['global'].lower() in row[4].lower())  # riferimento
        # Per forn, cod, desc, ord: interrogare DB source come già fatto
    )
    if not global_match:
        continue
```

---

#### **MODIFICA 3**: Modificare `_on_search()` per usare il campo global

**File**: `ui/components/main_dashboard_toolbar.py` (linee 175-178)

```python
# PRIMA:
if hasattr(self.main_window, 'search_vars') and 'num' in self.main_window.search_vars:
    self.main_window.search_vars['num'].set(search_text)

# DOPO:
if hasattr(self.main_window, 'search_vars') and 'global' in self.main_window.search_vars:
    # Pulisce i filtri testuali individuali per evitare conflitti
    # (mantiene filtri data/tipo/utente intatti)
    for key in ['num', 'ref', 'forn', 'cod', 'desc', 'ord']:
        if key in self.main_window.search_vars:
            self.main_window.search_vars[key].set('')
    
    # Imposta la global query
    self.main_window.search_vars['global'].set(search_text)
```

**Motivazione**: Quando l'utente usa la global search, presumiamo voglia una ricerca ampia, non una combinazione AND con filtri individuali. Puliamo i filtri testuali ma manteniamo filtri strutturati (date, tipo, utente) che sono spesso desiderati come restrizioni aggiuntive.

---

#### **MODIFICA 4**: Sincronizzare `clear_filters()` per pulire anche global

**File**: `dataflow.py` (circa linea 5065)

```python
def clear_filters(self):
    for var in self.search_vars.values(): var.set("")  # Già pulisce global incluso
    self.search_tipo.set(_("Tutte"))
    for de in self.date_entries.values(): de.delete(0, 'end')
    if self.username_filter_var:
        self.username_filter_var.set(self.current_username or self.all_users_placeholder)
    self.refresh_data()
```

**Nota**: Nessuna modifica necessaria - il loop `for var in self.search_vars.values()` pulisce già tutti i campi incluso `global`.

---

## ⚠️ CONSIDERAZIONI TECNICHE

### Validazione e Sicurezza

✅ **Lunghezza input**: Già gestita da `MAX_SEARCH_LENGTH = 100`  
✅ **SQL Injection**: Già gestita da `FORBIDDEN_CHARS` regex  
✅ **Parametrizzazione**: Uso di `?` placeholder con sqlite3 (safe)

### Performance

- **Query SQL con OR**: Potenzialmente più lenta di query su singolo campo
- **Mitigazione**: 
  - I campi `ro.id_richiesta`, `ro.riferimento` sono nella tabella principale
  - LEFT JOIN già utilizzati nell'implementazione corrente
  - Per dataset tipici (<10K RfQ), overhead accettabile
  - **Raccomandazione futura**: Aggiungere indici su `nome_fornitore`, `codice_materiale`

### Compatibilità

- ✅ **Linux/Windows**: Nessuna modifica platform-specific
- ✅ **Retrocompatibilità**: Se `global` non esiste, la logica fallback funziona
- ✅ **Database**: SQLite standard, nessuna estensione richiesta

---

## ⚡ RISCHI EVITATI

| Rischio | Come Evitato |
|---------|--------------|
| Duplicazione logica | Riuso 100% di `search_requests()` |
| Refactor massivo | Solo 3 modifiche localizzate (~50 righe totali) |
| Nuove dipendenze | Zero librerie aggiunte |
| Breaking changes | Filtri avanzati continuano a funzionare invariati |
| Regressioni | Campo vuoto + Enter continua a chiamare `clear_filters()` |
| Confusione AND/OR | Clear automatico dei filtri testuali quando si usa global |

---

## 🔍 EDGE CASES E GESTIONE

### 1. Global search + filtri avanzati contemporaneamente

**Comportamento**: 
- Global search pulisce automaticamente i filtri testuali individuali
- Mantiene filtri strutturati (data, tipo, utente)
- **Esempio**: "ACME" + Data Emissione Gen-Mar 2026 → OR su 6 campi + AND sulla data

### 2. Global search con caratteri speciali

**Gestione**:
- Sanitizzazione in `search_requests()` rimuove `';"\<>`
- Messagebox avvisa utente della sanitizzazione
- Query procede con testo pulito

### 3. Global search con risultati zero

**Comportamento**:
- Nessun risultato = albero vuoto (già gestito da `update_treeview`)
- Nessun messaggio di errore (UX standard)
- L'utente può modificare query o pulire filtri

### 4. Ricerca aggregata "All users"

**Gestione**:
- Applica filtro OR su tutti i database utente
- Filtro in-memory per dettagli (forn, cod, desc) richiede query su DB specifici
- Performance OK per dataset tipici, potenziale bottleneck se migliaia di utenti

### 5. Campo `global` nei filtri avanzati

**Nota**: Il campo `global` è invisibile nell'UI dei filtri collapsabili - è solo un container virtuale per la logica di ricerca.

---

## ✨ VANTAGGI DELLA SOLUZIONE

| Vantaggio | Dettaglio |
|-----------|-----------|
| **Minimalismo** | 3 file, ~50 righe di codice totali |
| **Reversibilità** | Rimuovere `'global'` da search_vars per rollback |
| **Chiarezza** | Separazione netta tra global search (OR) e filtri (AND) |
| **Performance** | Query SQL efficiente, nessun overhead significativo |
| **UX coerente** | Placeholder "Search anything..." diventa veritiero |
| **Manutenibilità** | Zero duplicazione, logica centralizzata |
| **Testabilità** | Modifiche isolate, facili da unit-testare |

---

## 🎯 NEXT STEPS

### Fase 1: Implementazione Base (Questo PR)
- [ ] MODIFICA 1: Aggiungere `'global'` a `search_vars`
- [ ] MODIFICA 2: Implementare logica OR in `search_requests()`
- [ ] MODIFICA 3: Modificare `_on_search()` per usare campo global
- [ ] Test manuale: ricerca per fornitore, codice, riferimento
- [ ] Test edge case: caratteri speciali, campo vuoto, risultati zero

### Fase 2: Ottimizzazioni (Opzionali, PR Successivi)
- [ ] Aggiungere indici DB su `nome_fornitore`, `codice_materiale`
- [ ] Implementare highlight dei match nella UI (visual feedback)
- [ ] Aggiungere suggerimenti autocomplete basati su history
- [ ] Metriche performance: tempo query, numero risultati

### Fase 3: Advanced Features (Futuro)
- [ ] Ricerca con operatori booleani (`AND`, `OR`, `NOT`)
- [ ] Ricerca fuzzy/tollerante a typo (Levenshtein distance)
- [ ] Salvataggio query preferite/recenti
- [ ] Export risultati ricerca (Excel, CSV)

---

## 📝 NOTE IMPLEMENTATIVE

### Testing Checklist

**Unit Tests** (manuale):
1. Global search con singola parola → verifica match multi-campo
2. Global search + filtro data → verifica AND corretto
3. Campo vuoto + Enter → verifica clear_filters()
4. Caratteri speciali (`';`) → verifica sanitizzazione
5. Query lunga (>100 char) → verifica validazione lunghezza

**Integration Tests**:
1. Ricerca locale (utente corrente) → verifica DB query
2. Ricerca aggregata ("All users") → verifica filtro in-memory
3. Ricerca con LEFT JOIN (fornitore, materiale) → verifica OR su tabelle joined
4. Passaggio global search → filtri avanzati → global search → verifica reset

**Regression Tests**:
1. Filtri avanzati senza global search → comportamento invariato
2. Pulsante "Cerca" nei filtri → funziona come prima
3. Pulsante "Pulisci Filtri" → pulisce tutto incluso global
4. Toggle filtri collapsabili → funziona invariato

---

## 📚 RIFERIMENTI

- **Issue/Task**: Global Search Enhancement
- **Documentazione esistente**: 
  - [STEP_5_COLLAPSIBLE_FILTERS.md](STEP_5_COLLAPSIBLE_FILTERS.md)
  - [ACTIONS_BUTTON_PROPOSAL.md](ACTIONS_BUTTON_PROPOSAL.md)
- **Codice rilevante**:
  - `MainDashboardToolbar` (ui/components/main_dashboard_toolbar.py)
  - `MainWindow.search_requests()` (dataflow.py:4664)
  - `DatabaseManager.search_richieste_advanced()` (database/db_helpers.py)

---

## ✅ APPROVAL CHECKLIST

Prima di implementare, verificare:

- [ ] La proposta è stata rivista e approvata
- [ ] I vincoli fondamentali sono rispettati (no refactor massivo, no nuove librerie)
- [ ] I rischi sono stati identificati e mitigati
- [ ] Il piano di testing è chiaro
- [ ] La documentazione è aggiornata
- [ ] Esiste un piano di rollback (rimuovere `'global'`)

---

**Fine Documento Tecnico**
