# Analisi approfondita: incongruenza ownership dialog VSM

Data: 2026-04-21
Progetto: DataFlow Procurement Software
Ambito: solo analisi, nessuna modifica codice

## 1. Executive summary

Il problema osservato e' reale, riproducibile dal codice ed e' circoscritto.

La causa radice non e' nei metadata aggregati, non e' nella selezione riga, non e' nel caricamento dal DB sorgente e non e' nella logica permessi/read-only.

La causa radice e' nel dialog VSM: il campo etichettato `User` viene precompilato in modo incondizionato con `current_username` e, in modalita' edit/read-only, non viene mai sovrascritto con il valore del record caricato.

In pratica:

- la tabella VSM mostra il proprietario corretto da `event.username`
- i metadata `is_mine` / `source_file` vengono valorizzati correttamente
- il dialog apre il record corretto dal DB corretto
- ma il widget `User` del dialog continua a mostrare l'utente di sessione

Verdetto finale: `Fix ora`

## 2. Comprensione del problema

Scenario osservato:

- tab VSM Saving o Cost Avoidance
- contesto aggregato / multi-DB
- selezione di un evento appartenente a un altro utente
- doppio click o azione edit

Comportamento atteso:

- apertura in sola lettura
- nessuna scrittura
- visualizzazione dell'utente reale proprietario dell'evento

Comportamento effettivo:

- apertura in sola lettura corretta
- nessuna scrittura indebita osservata
- il campo `User` nel dialog mostra `current_username` invece di `event.username`

## 3. Root cause

### Causa radice esatta

Il widget del dialog etichettato `User` e' `self.entry_buyer`, ma viene inizializzato sempre con `self.current_username` in `ui/dialogs/vsm_event_dialog.py:105-108`.

Poi, quando il dialog entra in modalita' edit e carica il record, `_load_event_data()` recupera correttamente l'evento dal DB selezionato, ma contiene un commento esplicito che evita di aggiornare quel campo:

- `ui/dialogs/vsm_event_dialog.py:555`  
  `# Campo User già popolato con current_username in __init__, non sovrascrivere`

Questa scelta rende il valore del widget indipendente dal record realmente caricato.

### Flusso completo dei dati

#### 1) Selezione riga nella UI (`tksheet`)

- Il doppio click sul foglio VSM chiama `dataflow.py:2045-2072` (`_on_vsm_sheet_double_click`).
- Prima di leggere la selezione, il codice passa sempre da `_get_selected_row_indices()`, che sincronizza i metadata con l'ordine visibile del foglio: `dataflow.py:2162-2165`.
- La sincronizzazione metadata dopo sort e' implementata in `dataflow.py:2174-2244`.

Conclusione: il mapping riga selezionata -> metadata non mostra evidenze di rottura nel caso osservato.

#### 2) Mapping riga -> metadata

Il popolamento del foglio VSM costruisce contemporaneamente:

- la riga visuale, con colonna `User = event.username`: `dataflow.py:1743-1754`
- i metadata per riga, con:
  - `event_id`
  - `username`
  - `is_mine`
  - `source_file`

Riferimenti:

- `dataflow.py:1770-1784`

In aggregazione multi-DB:

- i dati grezzi arrivano da `services/vsm_dashboard_service.py:15-29`
- `extra_meta` contiene `is_mine` e `source_file`
- il filtering per tipo e search preserva l'allineamento metadata/eventi:
  - `services/dashboard_search_service.py:76-98`
  - `dataflow.py:1499-1518`
  - `dataflow.py:2677-2732`

Conclusione: la riga e i metadata usano il valore corretto del record (`event.username`), non `current_username`.

#### 3) Caricamento record

Quando si apre il dialog da un record non owned:

- `_edit_vsm_event()` legge i metadata della riga: `dataflow.py:1858-1868`
- se `is_mine` e' falso, apre il dialog in read-only: `dataflow.py:1870-1886`
- passa anche `source_db_path` derivato da `metadata['source_file']`

Nel dialog:

- `_load_event_data()` sceglie correttamente il DB target:
  - `ui/dialogs/vsm_event_dialog.py:521-528`
- apre il database con `read_only=self.read_only`:
  - `ui/dialogs/vsm_event_dialog.py:536-537`
- `DatabaseManager` in read-only usa SQLite URI `mode=ro`:
  - `database_manager.py:48-55`
- `get_event_with_impacts()` recupera l'evento vero dal DB sorgente:
  - `services/vsm_persistence.py:216-253`
- `get_vsm_event_by_id()` legge `username` e `buyer` dal DB:
  - `database_manager.py:2185-2225`

Conclusione: il record corretto viene effettivamente caricato dal DB corretto.

#### 4) Costruzione dialog VSM

Durante `__init__` del dialog:

- il campo `User` viene creato come `self.entry_buyer`: `ui/dialogs/vsm_event_dialog.py:227-229`
- subito dopo viene valorizzato sempre con `self.current_username`: `ui/dialogs/vsm_event_dialog.py:105-108`

Questo avviene prima del caricamento del record.

#### 5) Popolamento campo `User`

Qui nasce il bug.

Nel caricamento dati:

- il codice aggiorna data, tipo, azione, descrizione, campi economici, driver, pagamenti, flag, new supplier
- ma non aggiorna il campo `User`

Riferimento chiave:

- `ui/dialogs/vsm_event_dialog.py:555`

Il valore errato quindi arriva da:

- `current_username` di sessione
- usato come default incondizionato nel dialog
- mantenuto anche in edit/read-only

### Cosa non e' la causa radice

Non ci sono evidenze che il problema derivi da:

- metadata disallineati
- `source_db_path` errato
- `is_mine` errato
- fallback su DB locale
- override specifico della modalita' read-only
- resync aggregato multi-DB

Motivo: se uno di questi pezzi fosse errato, ci aspetteremmo anche almeno uno tra questi sintomi:

- record sbagliato aperto
- permessi sbagliati
- dialog non read-only
- possibile scrittura indebita

Dal codice e dal comportamento osservato, questi sintomi non emergono.

## 4. Code path coinvolti

### Path principale

1. Dataset aggregato VSM  
   `services/vsm_dashboard_service.py:15-29`

2. Filtraggio e preservazione metadata  
   `services/dashboard_search_service.py:76-98`  
   `dataflow.py:1499-1518`  
   `dataflow.py:2677-2732`

3. Popolamento sheet con colonna `User` e metadata ownership  
   `dataflow.py:1694-1795`

4. Selezione riga e resync metadata dopo sort  
   `dataflow.py:2162-2244`

5. Apertura dialog con `read_only` e `source_db_path`  
   `dataflow.py:1837-1904`

6. Load record dal DB corretto  
   `ui/dialogs/vsm_event_dialog.py:518-607`  
   `services/vsm_persistence.py:216-253`  
   `database_manager.py:2174-2225`

7. Punto esatto del bug di visualizzazione  
   `ui/dialogs/vsm_event_dialog.py:105-108`  
   `ui/dialogs/vsm_event_dialog.py:555`

### Evidenza comparativa utile

Il dialog Derisking separato aggiorna correttamente il campo `User` con il valore del record caricato:

- `ui/dialogs/potential_supplier_dialog.py:412-416`

Questo rafforza la conclusione che il difetto sia locale al dialog VSM, non al framework multi-DB o alla UI in generale.

## 5. UI-only o problema strutturale?

### Risposta breve

Il problema osservato e' `puramente visivo / di binding nel dialog VSM`.

### Perche'

- la tabella usa `event.username` corretto
- i metadata ownership sono coerenti con la tabella
- il record viene caricato dal DB sorgente corretto
- la modalita' read-only si attiva sul criterio corretto
- il dialog non usa il record caricato per aggiornare il widget `User`

### Impatta solo il campo `User`?

Per l'evidenza disponibile: `si', riguarda solo il campo User del dialog VSM`.

Non ci sono evidenze nel codice corrente che lo stesso problema impatti:

- i permessi
- la selezione record
- `source_db_path`
- la read-only mode
- Saving / Cost Avoidance / Derisking come logica di calcolo
- altri campi del dialog VSM, che vengono popolati dal record caricato

### Esiste una incoerenza architetturale piu' ampia?

Esiste una piccola incoerenza locale di naming/intent:

- il widget mostrato come `User` si chiama `entry_buyer`
- il modello VSM ha sia `username` sia `buyer`
- il dialog visualizza di fatto il valore di sessione, non il record

Pero', allo stato delle evidenze, questa incoerenza non dimostra un problema piu' ampio di ownership o aggregazione.

Dichiarazione esplicita richiesta: `non ci sono evidenze di un problema piu' ampio tipo metadata resync o resync multi-DB`.

## 6. Tabella rischi

| Rischio | Descrizione | Prob. | Impatto | Mitigazione | Come testarlo |
|---|---|---:|---:|---|---|
| R1 | Correggere la visualizzazione ma rompere la modalita' read-only | Bassa | Alta | Non toccare `is_mine`, `source_db_path`, `_apply_read_only`, `_edit_vsm_event` | Aprire evento altrui e verificare assenza pulsante Save e campi disabilitati |
| R2 | Visualizzare utente corretto ma introdurre un salvataggio con utente errato | Bassa | Alta | Limitare il fix al popolamento del widget in edit/read-only; non toccare `_validate_and_save()` se non strettamente necessario | Aprire/modificare/salvare evento proprio e verificare che `username` resti corretto in tabella |
| R3 | Regressione su record locali owned | Bassa | Media | In create continuare a mostrare `current_username`; in edit usare il dato del record | Creare nuovo evento locale, riaprirlo, verificare `User` coerente |
| R4 | Regressione sui dialog Saving / Cost Avoidance per ordine di inizializzazione | Media | Media | Aggiornare il campo `User` dopo il load record, senza cambiare il lifecycle del form | Aprire eventi Saving e Cost Avoidance sia in create sia in edit |
| R5 | Impatto involontario su `source_db_path` e DB aggregati | Bassa | Alta | Non cambiare la scelta del DB target in `_load_event_data()` | Aprire evento remoto e verificare che i campi descrittivi/economici coincidano con la riga selezionata |
| R6 | Toccare permessi/ownership invece del solo binding UI | Bassa | Alta | Evitare modifiche in `dataflow.py`, `services/vsm_dashboard_service.py`, `dashboard_selection_policy.py` | Verificare che le azioni restino disabilitate per eventi non propri |
| R7 | Regressione su Cost Avoidance / Saving / Derisking per uso improprio del campo `buyer` | Media | Medio | Non rifattorizzare il modello; correggere solo il valore mostrato nella UI | Aprire un evento per ciascun tipo e confrontare tabella vs dialog |
| R8 | Dipendenze nascoste dall'utente corrente nel dialog | Media | Medio | Cercare e lasciare intatti i punti che dipendono da `current_username` per create/save; limitarsi al display in edit | Aprire evento proprio e altrui, controllare che solo il display cambi dove atteso |
| R9 | Sort/refresh mascherino il problema o introducano false diagnosi | Bassa | Basso | Non usare il resync come leva di fix; trattarlo solo come sanity check | Ordinare per data/user, aprire il record e verificare stesso comportamento |

## 7. Strategie di fix (max 3)

### Strategia 1 - Fix UI mirato nel solo dialog VSM

Descrizione:

- lasciare invariata tutta la pipeline dati e ownership
- in `ui/dialogs/vsm_event_dialog.py`, usare il record caricato per valorizzare il campo mostrato come `User` in modalita' edit
- in create continuare a usare `current_username`

File coinvolti:

- `ui/dialogs/vsm_event_dialog.py`

Perche' risolve:

- elimina la sorgente del valore errato
- il campo mostrera' il valore del record caricato, non quello di sessione

Rischio regressione:

- basso

Facilita' rollback:

- alta, modifica puntuale in un solo file

Impatto:

- solo UI / dialog

### Strategia 2 - Introdurre helper esplicito per il campo User nel dialog

Descrizione:

- aggiungere un helper dedicato, ad esempio concettualmente `set_displayed_user(...)`
- usarlo in create e in edit, mantenendo separati i casi:
  - create -> `current_username`
  - edit -> `event.username`

File coinvolti:

- `ui/dialogs/vsm_event_dialog.py`

Perche' risolve:

- rende esplicito il contratto del widget
- riduce la possibilita' di future regressioni nello stesso dialog

Rischio regressione:

- basso/medio

Facilita' rollback:

- alta

Impatto:

- UI, con minima pulizia locale della logica

### Strategia 3 - Riallineare semanticamente `User`, `username` e `buyer`

Descrizione:

- rivedere il ruolo del campo visuale, il naming `entry_buyer` e il rapporto tra `username` e `buyer`
- eventualmente separare display username da eventuale buyer reale

File coinvolti:

- `ui/dialogs/vsm_event_dialog.py`
- potenzialmente `models/vsm_event.py`
- potenzialmente `database_manager.py`

Perche' risolve:

- affronta la confusione semantica alla radice

Rischio regressione:

- medio/alto

Facilita' rollback:

- bassa rispetto alle altre

Impatto:

- UI + logica/modello

Nota:

- non raccomandata per questo bug, dato il vincolo di intervento minimo e low-risk

## 8. Fix raccomandato

### Soluzione raccomandata

Raccomando la `Strategia 1`.

### File esatto da modificare

- `ui/dialogs/vsm_event_dialog.py`

### Punto di intervento preciso

1. Nel costruttore del dialog, mantenere la valorizzazione con `current_username` solo per la modalita' create.
2. In `_load_event_data()`, dopo aver caricato `event`, aggiornare esplicitamente il widget `User` con il valore del record.

### Valore da mostrare

Per il bug osservato, il valore corretto da mostrare e' `event.username`, perche':

- la tabella VSM mostra `event.username`: `dataflow.py:1753`
- i metadata ownership usano `event.username`: `dataflow.py:1781`
- i permessi di ownership sono basati su questo concetto, non sul widget del dialog

### Cosa NON deve essere toccato

Non deve essere toccato:

- `dataflow.py` nella logica di selezione/open/read-only
- `services/vsm_dashboard_service.py`
- `services/dashboard_search_service.py`
- `services/dashboard_selection_policy.py`
- `source_db_path`
- `is_mine`
- `_apply_read_only()`
- `DatabaseManager`
- `get_event_with_impacts()`
- la logica di save/update, salvo necessita' strettamente dimostrata

### Perche' e' la soluzione meno rischiosa

- corregge il punto esatto che genera il bug
- non tocca ownership, permessi, aggregazione o persistenza
- e' coerente con il comportamento gia' corretto della tabella
- e' facilmente rollbackabile

## 9. Piano test manuale (max 10)

1. Evento locale owned, tab Saving  
   Aprire un evento proprio. Verificare che il dialog mostri lo stesso `User` della riga.

2. Evento locale owned, tab Cost Avoidance  
   Aprire un evento proprio. Verificare coerenza tra colonna `User` e campo `User` nel dialog.

3. Evento aggregato non owned, tab Saving  
   Aprire un evento di altro utente. Verificare che il dialog sia read-only e che `User` mostri il proprietario reale.

4. Evento aggregato non owned, tab Cost Avoidance  
   Stessa verifica del punto 3.

5. Verifica assenza scritture indebite  
   Su evento non owned, confermare che il pulsante Save non sia presente e che non sia possibile modificare campi.

6. Verifica DB sorgente corretto  
   Su evento non owned, confrontare almeno descrizione, data, azione e valori economici tra riga e dialog per escludere load dal DB sbagliato.

7. Creazione nuovo evento locale  
   Aprire nuovo evento da tab Saving o Cost Avoidance. Verificare che in create il campo `User` mostri l'utente corrente.

8. Edit + save evento proprio  
   Modificare un evento proprio, salvare, riaprire e verificare che il `User` resti coerente e che non cambi ownership.

9. Sanity check dopo sort  
   Ordinare la tabella VSM per `User` o `Date`, poi aprire un record non owned e verificare stesso comportamento corretto.

10. Sanity check dopo refresh/search  
   Applicare filtro utente o ricerca globale, aprire un record non owned e verificare che `User` nel dialog resti corretto.

## 10. Verdetto finale

### Sintesi finale

- Root cause confermata: dialog VSM, non metadata
- Problema isolato: visualizzazione campo `User`
- Nessuna evidenza di bug piu' ampio su ownership, resync metadata o multi-DB
- Fix consigliato: puntuale, locale, low-risk, in un solo file

### Decisione

`Fix ora`

Motivazione:

- il bug e' reale e con causa chiara
- la correzione puo' essere minima e circoscritta
- il rischio di regressione e' basso se il fix resta confinato al dialog

