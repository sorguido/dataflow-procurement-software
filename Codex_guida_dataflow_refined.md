# Guida Utente Operativa DataFlow

## 1. Introduzione
Questa guida descrive l’uso operativo di DataFlow partendo dal comportamento implementato nel codice applicativo.

Cosa copre:
- avvio applicazione e configurazione iniziale
- dashboard principale, tab, ricerca, filtri
- gestione completa delle RdO (RFQ)
- gestione Saving, Cost Avoidance, Derisking
- analisi SQDC
- export Excel (RdO, dashboard, KPI)
- impostazioni, backup, percorso dati, limiti multiutente

Approccio usato in questa guida:
- orientato ai passaggi reali in interfaccia
- centrato su pulsanti, campi, validazioni e salvataggi
- distinzione tra azioni reversibili e irreversibili
- nessuna funzione inventata: quando un punto non è confermabile al 100% è indicato nelle sezioni finali

Come leggere i capitoli operativi:
- `Clicca`: indica il comando preciso in toolbar, menu o finestra
- `Compila`: indica i campi da valorizzare e il formato atteso
- `Conferma`: indica cosa salvare e quale effetto vedi subito in schermata
- `Attenzione`: indica blocchi, limiti o azioni irreversibili

## Percorsi rapidi
Se vuoi arrivare subito alla parte che ti serve, usa questi percorsi:
- Voglio creare una nuova RdO: vai a [5.1 Creazione nuova RdO](#51-creazione-nuova-rdo)
- Voglio inserire articoli e prezzi: vai a [7.3 Inserimento/modifica celle articolo](#73-inserimentomodifica-celle-articolo) e [7.5 Inserimento/modifica prezzi fornitore](#75-inserimentomodifica-prezzi-fornitore)
- Voglio confrontare i fornitori con SQDC: vai a [9.5 Calcolo Cost automatico](#95-calcolo-cost-automatico) e [9.10 Micro-flow SQDC senza blocchi di validazione](#910-micro-flow-sqdc-senza-blocchi-di-validazione)
- Voglio registrare Saving o Cost Avoidance: vai a [10.2 Creazione/modifica evento Saving](#102-creazionemodifica-evento-saving) e [10.3 Creazione/modifica evento Cost Avoidance](#103-creazionemodifica-evento-cost-avoidance)
- Voglio controllare i KPI: vai a [11.1 Apertura](#111-apertura) e [11.10 Micro-flow controllo KPI di fine mese](#1110-micro-flow-controllo-kpi-di-fine-mese)
- Voglio fare backup o cambiare lingua: vai a [12.2 Lingua](#122-lingua), [12.3 Backup manuale](#123-backup-manuale), [12.4 Backup automatico](#124-backup-automatico)

Una volta scelto il percorso rapido, segui i micro-flow nel capitolo indicato: trovi sempre sequenza clic -> effetto -> passo successivo.

## 2. Primi passi

### 2.1 Primo avvio: sequenza completa
Al primo avvio DataFlow esegue questa sequenza:
1. inizializza la lingua da `config.ini` (default inglese se non configurata)
2. verifica accettazione licenza (`Settings.license_accepted`)
3. se non accettata, apre la finestra licenza in modalità bloccante
4. verifica identità utente (`User.first_name`, `User.last_name`, `User.username`)
5. se mancante, apre la finestra identità utente
6. crea la struttura cartelle utente
7. crea/inizializza il database utente
8. mostra splash screen e poi dashboard principale

Se non accetti la licenza, l’app si chiude.

### 2.2 Accettazione licenza
Nella finestra licenza di primo avvio trovi:
- `✅ Accetto`
- `❌ Esci`

Quando clicchi `✅ Accetto`, il flag viene salvato in `config.ini` e non viene richiesto di nuovo nei successivi avvii, salvo problemi di scrittura config.

### 2.3 Identificazione utente e generazione username
Se i dati utente non esistono, si apre un dialog obbligatorio con:
- `Nome`
- `Cognome`
- anteprima `Username generato`

Regola di generazione username:
- prima lettera del nome + cognome
- rimozione accenti
- rimozione spazi e caratteri non validi
- minuscolo

Esempio pratico:
- Nome: `Luca`
- Cognome: `Bianchi`
- Username: `lbianchi`

### 2.4 Creazione cartelle e database
Dopo identità valida DataFlow usa una cartella utente dedicata:
- Windows: `~/Documents/DataFlow_<username>`
- Linux/macOS: `~/DataFlow_<username>`
- oppure percorso personalizzato se configurato in `Settings.dataflow_base_dir`

Struttura principale:
- `DataFlow_<username>/Database`
- `DataFlow_<username>/Attachments`

Database utente:
- `Database/dataflow_db_<username>.db`

### 2.5 Avvii successivi
Negli avvii successivi:
- licenza non viene riproposta se già accettata
- identità utente viene letta da config
- si apre direttamente dashboard
- viene avviato il controllo periodico per backup automatico

### 2.6 Checklist operativa primo accesso
Quando entri per la prima volta, usa questa sequenza minima per evitare errori nelle fasi successive:
1. conferma licenza (`✅ Accetto`)
2. inserisci nome e cognome reali, poi verifica username generato
3. attendi la creazione automatica cartelle e database
4. apri `⚙️ Impostazioni` e controlla lingua e backup
5. torna in dashboard e verifica che i tab (`RdO`, `Saving`, `Cost Avoidance`, `Derisking`) siano navigabili

Perché serve:
- se salti il controllo iniziale di lingua/backup rischi esportazioni in lingua non desiderata o backup non configurati

Attenzione:
- se chiudi la finestra licenza senza accettare, DataFlow termina e non prosegue al dashboard

Passo successivo consigliato:
- dopo il primo accesso vai a `Panoramica interfaccia` per orientarti tra tab, barra ricerca e menu `Actions` prima di creare dati reali.

## 3. Panoramica interfaccia

### 3.1 Toolbar principale
In alto trovi:
- `➕ Nuovo Evento`
- `⚡ Actions`
- `📊 Export Excel`
- `≋ KPI`
- `⚙️ Impostazioni`
- `≡ License`
- `❓ Guida`

### 3.2 Tab principali
La dashboard include 5 tab:
- `RdO Attive`
- `RdO Archiviate`
- `Saving`
- `Cost Avoidance`
- `Derisking`

### 3.3 Ricerca globale
Sotto la toolbar principale trovi una barra unica con placeholder:
- `Search anything... (RFQ, Supplier, Code, Project...)`

A destra della barra:
- toggle `Advanced Filters` (espande/chiude pannello filtri)

### 3.4 Filtri avanzati collassabili
Il pannello filtri è unico e cambia contenuto in base al tab attivo:
- su tab RdO mostra filtri RFQ
- su tab Saving/Cost Avoidance mostra filtri VSM
- su tab Derisking il toggle Advanced Filters è disabilitato

### 3.5 Differenza operativa area RFQ vs area VSM
RFQ (`RdO Attive`, `RdO Archiviate`):
- gestione richieste, articoli, fornitori, offerte, allegati
- doppio click apre dettaglio RdO
- azioni: elimina, duplica, archivia/riattiva

VSM (`Saving`, `Cost Avoidance`, `Derisking`):
- Saving/Cost Avoidance: eventi economici con calcolo impatti
- Derisking: anagrafica fornitori potenziali (non eventi economici)
- azioni diverse per tab

### 3.6 Micro-flow orientamento dashboard
Percorso rapido per capire dove lavorare senza ambiguità:
1. clicca tab `RdO Attive`: qui lavori su richieste e offerte
2. clicca tab `Saving`: qui registri risparmi economici già negoziati
3. clicca tab `Cost Avoidance`: qui registri costo evitato rispetto alla richiesta iniziale
4. clicca tab `Derisking`: qui inserisci/aggiorni fornitori potenziali e stato qualificazione
5. usa `⚡ Actions` dopo aver selezionato una riga: il menu cambia in base al tab

Effetto pratico:
- se ti trovi nel tab sbagliato, il pulsante `➕ Nuovo Evento` apre un dialog diverso; verificare il tab prima di cliccare evita inserimenti nel modulo errato

Suggerimento:
- nella toolbar in alto, controlla sempre il tab attivo prima di usare `➕ Nuovo Evento` o `⚡ Actions`.

Una volta chiarita la differenza RFQ/VSM, passa a `Ricerca e filtri` per trovare velocemente i record su cui lavorare.

## 4. Ricerca e filtri

### 4.1 Global Search in tab RdO
La ricerca globale su RdO applica OR su questi campi:
- numero RdO
- riferimento
- nome fornitore
- codice materiale
- descrizione materiale
- numeri ordine

### 4.2 Global Search + filtri avanzati: logica combinata
La logica è:
- blocco Global Search in OR
- filtri avanzati in AND
- combinazione complessiva: `(global OR) AND (filtri impostati)`

### 4.3 Validazioni input ricerca
Nei campi ricerca testuali RdO:
- lunghezza massima: 100 caratteri per campo
- caratteri `';"\`<>` vengono rimossi automaticamente
- il sistema mostra un avviso di input sanitizzato

### 4.4 Filtri avanzati tab RdO
Filtri disponibili:
- Numero RdO
- Tipo RdO (`Tutte`, `Fornitura piena`, `Conto lavoro`)
- Riferimento
- Fornitore
- Cod. Materiale
- Desc. Materiale
- Num. Ordine
- Cod. Grezzo
- Allegato Grezzo
- Mat. c/lavoro
- Utente
- Data Emissione `Da` / `A`
- Data Scadenza `Da` / `A`

### 4.5 Logica utente e ricerca multi-database in RFQ
Comportamento:
- filtro utente = utente corrente: ricerca ottimizzata sul DB locale
- filtro utente = altro utente: ricerca aggregata su database disponibili
- filtro utente = `Tutti gli utenti`: ricerca aggregata su database disponibili

### 4.6 Ricerca in tab Saving/Cost Avoidance
La ricerca VSM usa:
- global search testuale su campi evento (descrizione, riferimento, buyer, driver, azione, tipo, nuovo fornitore, note)
- filtri avanzati VSM (data, azione, ripetitivo, range teorico/effettivo)

### 4.7 Ricerca in tab Derisking
Nel tab Derisking la ricerca globale filtra per sottostringa su:
- fornitore
- categoria
- stato
- contatto
- email
- telefono
- sito web
- note
- utente

Se la query è vuota, il dataset viene ricaricato completo.

### 4.8 Esempi pratici
Esempio A (RFQ):
1. scrivi `motore` in Global Search
2. apri Advanced Filters
3. imposta `Tipo RdO = Conto lavoro`
4. clicca `🔍 Cerca`
5. ottieni solo RdO Conto lavoro in cui almeno un campo globale contiene `motore`

Esempio B (Saving):
1. vai su tab `Saving`
2. in filtri avanzati imposta `Ripetitivo = Sì`
3. imposta `Teorico Da = 10000`
4. usa Global Search con `acciaio`
5. vedi eventi Saving ripetitivi con valore teorico >= 10.000 e testo coerente

### 4.9 Micro-flow ricerca progressiva (metodo consigliato)
Quando non trovi record al primo tentativo:
1. inserisci solo una parola chiave in Global Search
2. clicca `🔍 Cerca` e verifica quanti risultati ottieni
3. apri `Advanced Filters` e aggiungi un solo filtro (es. Tipo RdO)
4. ripeti `🔍 Cerca`
5. continua aggiungendo un filtro per volta finché restringi il dataset al punto giusto

Perché serve:
- con logica `(OR globale) AND (filtri)` aggiungere troppi filtri insieme può azzerare i risultati

Nota operativa:
- dopo ogni filtro aggiunto, controlla subito la griglia; se diventa vuota, rimuovi l’ultimo filtro prima di proseguire

Passo successivo consigliato:
- quando hai trovato le righe giuste, entra in `Gestione RFQ / RdO` o nel tab VSM interessato e lavora solo sul subset filtrato.

## 5. Gestione RFQ / RdO

### 5.1 Creazione nuova RdO
Procedura:
1. clicca `➕ Nuovo Evento` mentre sei in `RdO Attive` o `RdO Archiviate`
2. scegli tipo:
3. `📦 Fornitura piena` oppure `🔧 Conto lavoro`
4. DataFlow crea una RdO “guscio” in stato `attiva`
5. si apre automaticamente la finestra dettaglio RdO

### 5.2 Tipi RdO
`Fornitura piena`:
- griglia articoli con colonne base + colonne fornitori

`Conto lavoro`:
- griglia articoli con colonne aggiuntive: Cod. Grezzo, Allegato Grezzo, Mat. C/L

### 5.3 Apertura RdO esistente
Con doppio click su riga tab RdO:
- apre `Control Panel - Request N°...`
- se la RdO appartiene ad altro utente aggregato: apertura in sola lettura

### 5.4 Menu `Actions` in area RdO
Tab `RdO Attive`:
- `🗑 Elimina`
- `🔁 Duplica`
- `📦 Archivia`

Tab `RdO Archiviate`:
- `🗑 Elimina`
- `🔁 Duplica`
- `↩️ Riattiva`

### 5.5 Archiviazione / riattivazione
`Archivia` e `Riattiva`:
- modificano solo il campo stato della RdO
- sono operazioni reversibili
- richiedono selezione di sole RdO di proprietà utente corrente

### 5.6 Eliminazione
`Elimina`:
- operazione permanente
- elimina testata RdO e dati correlati in DB
- elimina allegati associati in DB
- tenta anche rimozione file fisici in `Attachments`
- disponibile solo per RdO proprie

**Attenzione**
- `🗑 Elimina` è irreversibile: dopo conferma non puoi ripristinare la RdO dalla UI.

### 5.7 Duplicazione RdO
`Duplica` crea nuova RdO con nuovo ID.

Comportamento implementato:
- copia testata RdO
- imposta stato `attiva`
- data emissione = oggi
- data scadenza = vuota
- note formattate azzerate
- copia dettagli articoli
- non copia fornitori associati
- non copia offerte prezzi
- non copia allegati

Differenza operativa importante:
- la duplicazione è persistente e crea un nuovo record
- è consentita solo su RdO proprie

### 5.8 Azioni temporanee vs permanenti (RFQ)
Temporanee/reversibili:
- filtri e ricerca
- archiviazione/riattivazione

Persistenti ma modificabili:
- modifica campi RdO
- aggiunta/rimozione articoli
- modifica prezzi

Permanenti/irreversibili:
- eliminazione RdO
- eliminazione articoli
- eliminazione allegati

**Importante**
- prima di usare azioni permanenti, esegui `📊 Esporta Excel` dalla finestra dettaglio RdO, così hai una copia operativa.

### 5.9 Micro-flow operativo di chiusura RdO
Esempio completo quando hai finito una trattativa:
1. apri la RdO in `RdO Attive`
2. verifica che articoli, prezzi, PO e allegati siano completi
3. esegui `📊 Esporta Excel` dalla finestra dettaglio per archiviazione esterna
4. chiudi finestra dettaglio e torna alla dashboard
5. seleziona la riga RdO
6. clicca `⚡ Actions` → `📦 Archivia`

Cosa succede dopo conferma:
- la RdO scompare da `RdO Attive` e la trovi in `RdO Archiviate` senza perdita dati

Attenzione:
- `📦 Archivia` è reversibile (`↩️ Riattiva`), ma `🗑 Elimina` è irreversibile

Una volta completata la RdO, puoi passare alla finestra dettaglio per rifinire fornitori, PO, allegati e preparare la valutazione SQDC.

## 6. Finestra dettaglio RdO

### 6.1 Struttura generale
La finestra include:
- comandi rapidi allegati/documenti/fornitori/note/export/SQDC
- sezione date e riferimento
- pulsante gestione OdA (PO)
- griglia articoli e prezzi
- pulsanti fondo: aggiungi articolo, rimuovi articolo, importa Excel

### 6.2 Modifica riferimento
Clicca sul valore in `Riferimento`:
1. si apre finestra `Modifica Riferimento`
2. inserisci nuovo testo
3. clicca `💾 Salva`
4. il dato viene scritto in `richieste_offerta.riferimento`

### 6.3 Date emissione/scadenza
Le date sono `DateEntry` con formato `dd/mm/yyyy`.

Autosalvataggio:
- alla selezione da calendario
- su uscita dal campo
- su Invio

Effetto:
- quando confermi una data, il valore viene salvato e lo vedi subito aggiornato nella dashboard
- i campi tecnici aggiornati sono `data_emissione` e `data_scadenza`

### 6.4 Aggiungere e modificare fornitori
Pulsante dinamico:
- `➕ Aggiungi Fornitori` se assenti
- `✏️ Modifica Fornitori` se presenti

Finestra fornitori:
- campo unico con nomi separati da virgola
- blocco duplicati case-insensitive
- salvataggio transazionale
- rimozione automatica offerte dei fornitori eliminati

Nota operativa:
- inserisci i fornitori prima dei prezzi: le colonne prezzo vengono create in base all’elenco fornitori salvato.

### 6.5 Inserire e gestire i numeri ordine (PO)
Pulsante `📋 Inserisci OdA` apre gestione numeri ordine.

Operazioni:
- aggiunta riga con `Numero Ordine` + `Fornitore`
- eliminazione riga selezionata
- salvataggio automatico in JSON nel campo `numeri_ordine`

Validazioni:
- PO obbligatorio
- fornitore obbligatorio
- sanitizzazione caratteri pericolosi

**Attenzione**
- se lasci vuoto `Numero Ordine` o `Fornitore`, la riga non viene salvata.

### 6.6 Note
Pulsante note dinamico:
- `📝 Aggiungi nota`
- `📝 Visualizza nota`

Editor note:
- formattazione `grassetto`, `corsivo`, `sottolineato`
- salvataggio di struttura formattata nel DB (`note_formattate`)

### 6.7 Allegati offerte/documenti interni
Pulsanti:
- `📄 Gestisci Offerte Fornitori`
- `📁 Gestisci Documenti Interni`

Funzioni disponibili:
- aggiungi file
- apri file
- download file
- elimina file

Comportamento file:
- salvataggio fisico in `Attachments`
- in DB viene memorizzato link esterno (`percorso_esterno`)
- supporto fallback per allegati legacy salvati come BLOB

Sicurezza percorso:
- prima di aprire/scaricare file esterno viene validato che il path reale resti sotto la cartella allegati

### 6.8 Export dedicato RdO
Pulsante `📊 Esporta Excel` in dettaglio RdO:
1. scegli lingua (`Italiano`/`English`)
2. DataFlow seleziona template corretto in base a tipo RdO + lingua
3. compila testata, fornitori, articoli, prezzi
4. scegli percorso salvataggio

Template usati:
- `template_rdo.xlsx`
- `template_rdo_eng.xlsx`
- `template_rdo_cl.xlsx`
- `template_rdo_eng_cl.xlsx`

### 6.9 Modalità sola lettura
Quando la RdO non è tua:
- titolo include `[SOLA LETTURA]`
- disabilitati: modifica date, riferimento, fornitori, note, SQDC, PO, modifica griglia, import/export modifica
- restano disponibili operazioni di consultazione come apertura/download allegati ed export RdO

Importante:
- se vedi `[SOLA LETTURA]`, puoi consultare ed esportare ma non salvare modifiche. Torna in dashboard e lavora su una RdO di tua proprietà.

### 6.10 Micro-flow compilazione dettaglio RdO senza salti
Sequenza consigliata nella finestra dettaglio:
1. aggiorna `Riferimento`
2. imposta `Data Emissione` e `Data Scadenza`
3. inserisci fornitori (`➕/✏️ Fornitori`)
4. popola articoli (manuale o import Excel)
5. compila prezzi fornitore nelle colonne dedicate
6. inserisci eventuali PO (`📋 Inserisci OdA`)
7. aggiungi note operative (`📝`)
8. allega offerte/documenti (`📄` / `📁`)
9. esegui export RdO (`📊 Esporta Excel`)

Perché serve:
- seguendo questo ordine eviti di inserire prezzi prima di avere i fornitori, condizione che porta a rilavorazioni

Passo successivo consigliato:
- dopo questa compilazione puoi passare direttamente a `SQDC` per confrontare i fornitori in modo strutturato.

## 7. Griglia articoli e prezzi

### 7.1 Colonne standard
Sempre presenti:
- `Codice`
- `Allegato`
- `Descrizione`
- `Q.tà` / `Q.ty`

### 7.2 Colonne aggiuntive in Conto lavoro
Solo su `Conto lavoro`:
- `Cod. Grezzo`
- `Allegato Grezzo`
- `Mat. C/L`

Dopo le colonne articolo compaiono le colonne fornitore (una per fornitore associato alla RdO).

### 7.3 Inserimento/modifica celle articolo
Le celle articolo sono salvate in `dettagli_richiesta`.

Mappatura principale:
- colonna codice -> `codice_materiale`
- allegato -> `disegno`
- descrizione -> `descrizione_materiale`
- quantità -> `quantita`
- campi conto lavoro -> rispettive colonne dedicate

### 7.4 Validazione quantità
Per `quantita`:
- se usi `.` senza `,` compare warning
- conversione con regola virgola decimale
- normalizzazione fino a 4 decimali
- formattazione pulita in cella (senza zeri finali inutili)

**Nota operativa**
- usa sempre la virgola decimale (`12,5`): riduci errori di inserimento e blocchi in fase di controllo campo.

### 7.5 Inserimento/modifica prezzi fornitore
Valori ammessi:
- numero (`123,45`)
- `X`
- `ND`

Regole:
- numeri formattati a 4 decimali con virgola
- campo vuoto = elimina offerta per quel dettaglio/fornitore
- salvataggio in `offerte_ricevute` con upsert

**Attenzione**
- se svuoti una cella prezzo, l’offerta viene rimossa per quella coppia articolo/fornitore.

### 7.6 Aggiunta e rimozione articoli
`➕ Aggiungi Articolo`:
- inserisce riga vuota in DB
- refresh griglia
- seleziona ultima riga

`🗑 Rimuovi Articolo Selezionato`:
- supporta selezione multipla
- chiede conferma
- elimina articoli e prezzi associati
- operazione permanente

### 7.7 Impatto della gestione fornitori sulla griglia
Quando modifichi elenco fornitori:
- cambiano le colonne fornitore visualizzate
- vengono eliminati i prezzi dei fornitori rimossi
- dopo salvataggio il sistema esegue refresh della griglia

### 7.8 Menu contestuale: visuale vs persistente
Nella griglia è abilitato il popup contestuale tksheet (copy/cut/paste/delete/undo/edit).

Distinzione operativa:
- azioni di selezione, ordinamento e formattazione visiva non salvano dati
- il salvataggio avviene quando termini la modifica cella e il controllo campo è superato (`end_edit_cell`)

### 7.9 Micro-flow inserimento preventivo completo in griglia
Scenario pratico (1 articolo, 2 fornitori):
1. clicca `➕ Aggiungi Articolo`
2. compila `Codice`, `Descrizione`, `Q.tà`
3. premi Invio per chiudere l’editing cella
4. inserisci prezzo del primo fornitore (es. `12,50`)
5. inserisci prezzo del secondo fornitore (es. `13,10`)
6. verifica che i valori restino in cella dopo cambio riga

Cosa succede:
- quando lasci la cella con un valore valido, DataFlow salva e al refresh ritrovi lo stesso dato

Nota:
- se il valore è non valido, la modifica non viene confermata e devi correggerla prima di passare oltre

Dopo aver completato articoli e prezzi, passa a `Import / Export Excel` per import massivi o per produrre il file da condividere.

## 8. Import / Export Excel

### 8.1 Import articoli da Excel
Pulsante: `📊 Importa da Excel` nella finestra RdO.

Flusso:
1. DataFlow mostra istruzioni struttura file
2. selezioni un `.xlsx`
3. verifica numero colonne minimo
4. importa righe valide
5. aggiorna griglia

### 8.2 Vincoli struttura file
`Fornitura piena`:
- colonna A: Codice
- colonna B: Allegato
- colonna C: Descrizione
- colonna D: Quantità

`Conto lavoro`:
- A-D come sopra
- E: Codice Grezzo
- F: Allegato Grezzo
- G: Materiale C/L

Regole di import:
- righe senza codice o quantità vengono ignorate
- se il file ha meno colonne del minimo richiesto, import bloccato

**Importante**
- prima dell’import controlla il tipo RdO nel tab attivo: un file `Conto lavoro` su una RdO `Fornitura piena` (o viceversa) genera risultati incompleti.

### 8.3 Export singola RdO
Dal dettaglio RdO (`📊 Esporta Excel`):
- export su template RdO
- lingua selezionabile
- layout diverso per Fornitura piena vs Conto lavoro

Suggerimento:
- usa questo export come snapshot operativo prima di archiviare o eliminare dati.

### 8.4 Export dashboard (RFQ)
Dal cruscotto (`📊 Export Excel`) su tab RFQ:
- export di tutte le RdO coerenti con stato tab e filtri attivi
- non solo righe visibili: ricalcola dataset completo con gli stessi criteri
- supporta dataset aggregato multi-database

### 8.5 Export VSM
Su tab `Saving` o `Cost Avoidance`:
- export dedicato VSM con colonne teorico/effettivo/realizzo/variance/ripetitivo/utente
- lingua selezionabile

### 8.6 Export Derisking
Su tab `Derisking`:
- export anagrafica fornitori potenziali
- colonne fornitore, categoria, stato, contatti, note, utente

### 8.7 Micro-flow import Excel sicuro (controllo pre-import)
Prima di cliccare `📊 Importa da Excel`:
1. verifica tipo RdO (`Fornitura piena` o `Conto lavoro`)
2. apri il file e controlla che le colonne minime siano presenti (A-D o A-G)
3. controlla che ogni riga utile abbia almeno `Codice` e `Quantità`
4. salva e chiudi il file Excel (evita lock del file)
5. esegui import

Dopo import:
1. scorri la griglia fino in fondo
2. controlla 2-3 righe campione
3. se mancano righe, controlla se erano prive di codice/quantità

### 8.8 Micro-flow export dashboard per report periodico
1. applica filtri in dashboard (utente, date, tipo)
2. clicca `📊 Export Excel`
3. scegli lingua e destinazione file
4. apri il file esportato e verifica coerenza con i filtri impostati

Attenzione:
- l’export dashboard ricostruisce l’intero dataset coerente con i filtri, non solo le righe momentaneamente visibili in viewport

Passo successivo consigliato:
- quando il dataset è completo, passa a `SQDC` se devi decidere un fornitore, oppure a `KPI` se devi consolidare il periodo.

## 9. SQDC

### 9.1 Apertura finestra SQDC
Dal dettaglio RdO con pulsante SQDC:
- se esiste file SQDC interno: `📈 Apri analisi SQDC`
- altrimenti: `📊 Crea analisi SQDC`

### 9.2 Struttura tab
Tab presenti:
- `Pesi (%)`
- `Voti (1-10)`

### 9.3 Regole pesi
Controlli:
- ogni peso deve essere numero valido
- range 0..100
- somma totale obbligatoria = 100%

Se la somma non è 100, il sistema blocca passaggio/salvataggio.

### 9.4 Regole voti
Per ogni fornitore e criterio:
- voto obbligatorio
- intero tra 1 e 10

Colonne non editabili:
- `Fornitore`
- `TOTALE`

### 9.5 Calcolo Cost automatico
Pulsante: `🔄 Calcola Cost Automaticamente`

Logica:
- usa prezzi fornitore * quantità articoli RdO
- se fornitore ha prezzi mancanti, `X`, `ND` o incompleti rispetto al numero articoli, `Cost = 0`
- mostra avviso rosso con elenco fornitori incompleti
- assegna score più alto al totale prezzo più basso

### 9.6 Calcolo totale SQDC
Totale fornitore:
- combinazione pesata Safety/Quality/Delivery/Cost
- visualizzato con 2 decimali
- in griglia vengono evidenziati in verde i vincitori (gestione parità con tolleranza)

### 9.7 Salvataggio SQDC
Pulsante: `💾 Salva SQDC`

Effetti:
- crea file Excel fisico in `Attachments`
- registra/aggiorna un allegato di tipo `Documento Interno` collegato alla RdO
- nome logico allegato: `SQDC_Analysis_RfQ_<id>.xlsx`

### 9.8 Export SQDC
Pulsante: `📊 Esporta Excel`

Comportamento:
- usa template SQDC ITA/ENG
- compila pesi e tabella fornitori
- evidenzia vincitore nel file export
- salva dove scegli tu

### 9.9 Esempio pratico
Scenario rapido:
1. apri RdO con 3 fornitori e articoli completi
2. apri SQDC
3. pesi: 25/25/25/25
4. clicca `Calcola Cost Automaticamente`
5. inserisci voti Safety/Quality/Delivery
6. verifica totale e fornitore evidenziato
7. salva SQDC nei documenti interni
8. opzionale: export file esterno da allegare a report

### 9.10 Micro-flow SQDC senza blocchi di validazione
Procedura consigliata:
1. apri SQDC
2. compila prima il tab `Pesi (%)` e porta la somma esattamente a `100`
3. passa al tab `Voti (1-10)` e inserisci solo interi
4. clicca `🔄 Calcola Cost Automaticamente`
5. se compare avviso rosso, torna in RdO e completa prezzi mancanti
6. rientra in SQDC e ricalcola
7. clicca `💾 Salva SQDC`

Perché serve:
- l’ordine pesi→voti→cost riduce i salvataggi falliti e ti porta al file SQDC finale al primo tentativo

Una volta salvato SQDC, torna nella finestra dettaglio RdO: trovi il file nei `Documenti Interni` e puoi esportarlo per il report.

## 10. Value Stream Mapping

### 10.1 Struttura modulo
Tab VSM:
- `Saving`
- `Cost Avoidance`
- `Derisking`

Distinzione chiave:
- Saving/Cost Avoidance gestiscono eventi economici (`vsm_events`, `vsm_impacts`)
- Derisking gestisce anagrafica fornitori potenziali (`potential_suppliers`)

### 10.2 Creazione/modifica evento Saving
Dal tab `Saving`:
1. clicca `➕ Nuovo Evento`
2. si apre dialog evento VSM con `Tipo Evento` bloccato su Saving
3. compila campi generali (data, azione, descrizione)
4. scegli `Driver`: `Prezzo` o `Pagamenti`
5. compila campi economici coerenti col driver
6. imposta `OPEX Ripetitivo` se necessario
7. salva

### 10.3 Creazione/modifica evento Cost Avoidance
Flusso analogo al Saving, con tipo bloccato su `Cost Avoidance`.

Campi principali:
- importo richiesto iniziale
- importo negoziato
- quantità annua (se driver Prezzo)
- % realizzo

### 10.4 Regole driver economici
Driver `Prezzo`:
- Saving teorico = `(importo_bdg - importo_negoziato) * qty`
- Cost Avoidance teorico = `(importo_richiesto_iniziale - importo_negoziato) * qty`
- valore effettivo = teorico * `%realizzo`

Driver `Pagamenti`:
- formula teorica = `spending_annuo * (delta_giorni/30) * coefficiente`
- delta giorni = giorni negoziati - giorni attuali
- per Pagamenti il valore effettivo coincide col teorico

### 10.5 Ripetitivo vs non ripetitivo
Se `OPEX Ripetitivo` attivo:
- distribuzione impatti fino a 24 mesi
- primo mese pro-rata

Se non ripetitivo:
- impatto one-shot nel mese evento

### 10.6 Salvataggio e ricalcolo impatti economici
Su salvataggio evento economico:
- aggiornamento evento
- cancellazione impatti precedenti
- ricalcolo impatti
- salvataggio batch impatti

Pattern: `DELETE -> REGENERATE -> SAVE` in transazione atomica.

Traduzione operativa:
- quando clicchi `Salva`, DataFlow ricalcola automaticamente gli impatti economici mese per mese e aggiorna la riga evento con i nuovi valori teorico/effettivo che vedi in tab
- se modifichi un evento già esistente, gli impatti vecchi vengono sostituiti da quelli ricalcolati

Nota operativa:
- dopo il salvataggio controlla subito la riga nel tab attivo; se i valori non sono attesi, riapri l’evento e correggi i campi economici.

### 10.7 Derisking (anagrafica fornitori potenziali)
Dal tab `Derisking` clicca `➕ Nuovo Evento`.

In questo tab si apre il dialog fornitore potenziale con campi:
- Fornitore (obbligatorio)
- Categoria (lista)
- Nuova categoria (testo)
- Stato (`Nuovo`, `In valutazione`, `Qualificato`, `Scartato`)
- Contatto
- E-mail
- Telefono
- Web
- Note
- Utente (auto)

### 10.8 Gestione categorie Derisking
Dal dialog fornitore puoi aprire `Gestisci Categorie`.

Funzioni:
- rinomina categoria
- unisci categoria in altra categoria
- elimina categoria solo se non usata

Dettaglio importante:
- operazioni preparate in memoria
- applicazione definitiva solo con `💾 Salva`
- commit atomico delle operazioni

### 10.9 Azioni tab VSM
Su Saving/Cost Avoidance:
- `🗑 Elimina`
- `🔁 Duplica`

Su Derisking:
- `🗑 Elimina`
- niente duplicazione fornitore dal menu Actions

### 10.10 Sola lettura su record non propri
Se selezioni evento/fornitore di altro utente:
- azioni di modifica bloccate
- doppio click apre dialog in sola lettura

### 10.11 Esempio pratico Saving
1. vai su tab `Saving`
2. crea evento con driver Prezzo
3. importo budget 120,00
4. importo negoziato 100,00
5. qty annua 100
6. realizzo 80%
7. salva

Effetto economico:
- teorico: 2.000
- effettivo: 1.600
- impatti distribuiti secondo flag ripetitivo

### 10.12 Micro-flow verifica evento dopo salvataggio
Subito dopo `Salva` in Saving o Cost Avoidance:
1. torna alla tab di origine
2. controlla che la riga sia presente con descrizione corretta
3. verifica colonne teorico/effettivo/variance
4. applica filtro testo sul riferimento appena creato
5. se i numeri non tornano, riapri con doppio click e correggi i campi economici

Attenzione:
- eliminare un evento dal menu `⚡ Actions` è irreversibile

Quando il valore evento è corretto, puoi passare al tab KPI per verificare l’effetto sul periodo di analisi.

## 11. KPI Dashboard

### 11.1 Apertura
Clicca `≋ KPI` dalla toolbar principale.

Si apre finestra con tab:
- RFQ
- Saving
- Cost Avoidance
- Derisking

### 11.2 Filtri temporali
In header KPI:
- preset rolling: `1M`, `3M`, `12M`, `3Y`, `5Y`, `10Y`, `All`
- filtro `Year` (combobox)

Regola:
- preset e Year sono mutuamente esclusivi

### 11.3 Year vs preset rolling
Se selezioni `Year`:
- KPI su anno fisso

Se selezioni preset:
- KPI su intervallo rolling fino a oggi

Se `All`:
- nessun filtro temporale

### 11.4 Contenuto tab RFQ
Card principali:
- RFQ Active
- RFQ Archived
- RFQ Total
- RFQ Not Expired
- RFQ Expired
- Work Order
- Full Supply

Chart:
- RFQ emesse per periodo

### 11.5 Contenuto tab Saving
Card principali:
- Theoretical Saving
- Actual Saving
- Average/Best/Worst/Median %
- Recurring Impact
- Non-Recurring Impact
- Carry-over anno successivo (quando filtro Year è attivo)

Chart:
- confronto teorico vs effettivo per periodo

### 11.6 Contenuto tab Cost Avoidance
Card analoghe al Saving ma su Cost Avoidance.

Include:
- carry-over anno successivo in modalità Year

### 11.7 Contenuto tab Derisking
Card:
- totale fornitori potenziali
- categorie uniche
- card dinamiche per stato fornitore

Chart:
- fornitori per categoria

Nota filtri Year:
- nel tab Derisking il combobox Year viene popolato dagli anni presenti in `potential_suppliers.created_at`

### 11.8 Esportare i KPI in Excel
Pulsante `📥 Export Excel` in finestra KPI:
1. scegli scope: `Sezione corrente` o `Tutte le sezioni`
2. scegli lingua export
3. scegli percorso file
4. DataFlow genera workbook con Summary + fogli sezione

### 11.9 Esempio pratico di lettura KPI
Scenario:
1. apri KPI
2. seleziona tab `Saving`
3. imposta `Year = 2026`
4. controlla `Actual Saving` e `Recurring Impact`
5. passa a tab `Cost Avoidance` e confronta `Actual`
6. export `Tutte le sezioni` per report mensile

### 11.10 Micro-flow controllo KPI di fine mese
Sequenza consigliata:
1. apri KPI e imposta `Year` dell’anno corrente
2. verifica tab `RFQ` (attive/scadute) per stato operativo
3. verifica tab `Saving` e annota `Actual Saving`
4. verifica tab `Cost Avoidance` e annota `Actual`
5. verifica tab `Derisking` per distribuzione stati
6. esporta `Tutte le sezioni` e allega al report mensile

Perché serve:
- usi un’unica fotografia temporale coerente, evitando confronti su periodi diversi

Passo successivo consigliato:
- dopo l’export KPI, torna in dashboard e chiudi il ciclo operativo archiviando le RdO concluse.

## 12. Impostazioni e manutenzione

### 12.1 Apertura
Clicca `⚙️ Impostazioni`.

Sezioni principali:
- Posizione DataFlow Standard
- Backup Manuale
- Backup Automatico Giornaliero
- Lingua

### 12.2 Lingua
Sezione `Lingua`:
- scelta `English` / `Italiano`
- pulsante `💾 Salva Lingua`
- dopo salvataggio viene proposto riavvio applicazione

### 12.3 Backup manuale
Sezione `Backup Manuale`:
- pulsante `💾 Backup Manuale...`

Comportamento:
- chiude temporaneamente connessione principale
- copia file `.db`
- copia eventuali `.db-wal` e `.db-shm`
- esegue controllo dimensione backup
- in caso di mismatch grave può chiedere conferma mantenimento file

**Importante**
- esegui backup manuale prima di operazioni delicate (es. spostamento cartella DataFlow o pulizie massive record).

### 12.4 Backup automatico
Sezione `Backup Automatico Giornaliero`:
- checkbox abilitazione
- scelta ora (00-23)
- scelta cartella target
- salvataggio impostazioni

Regole operative:
- se attivo senza percorso, il sistema blocca salvataggio impostazione
- controllo eseguito ogni minuto
- mantiene massimo 3 set completi di backup automatico

Nota operativa:
- imposta un percorso locale o di rete stabile; se il percorso non è accessibile all’orario impostato, il backup non parte.

### 12.5 Cambio cartella DataFlow
Pulsante `📁 Cambia Posizione DataFlow...`:

Flusso:
1. avviso iniziale con conferma
2. selezione cartella parent
3. test permessi scrittura
4. verifica lunghezza path
5. controllo conflitto username nella destinazione
6. eventuale richiesta cambio identità/username
7. copia completa cartella DataFlow utente
8. aggiornamento config
9. riavvio app

Dettagli importanti:
- la cartella originale non viene eliminata automaticamente
- il sistema suggerisce di testare prima di cancellare manualmente la sorgente

**Attenzione**
- questa operazione incide su tutto l’ambiente dati utente: dopo il riavvio verifica subito apertura dashboard, allegati e export.

### 12.6 Struttura file e configurazione
File config:
- percorso area app user-specific (`config.ini`)

Sezioni usate:
- `Settings`: lingua, licenza, path dati
- `User`: nome, cognome, username
- `AutoBackup`: enabled/hour/path

Cartelle dati:
- `Database`
- `Attachments`

### 12.7 Log e diagnostica
DataFlow usa logging rotante:
- file `dataflow.log`
- rotazione 5MB, fino a 3 file
- posizione in area dati utente locale

Sono loggati errori di:
- DB
- backup
- import/export
- caricamento finestre
- validazioni critiche

### 12.8 Routine operativa di manutenzione (settimanale)
Checklist pratica:
1. apri `⚙️ Impostazioni`
2. esegui `💾 Backup Manuale...` e verifica creazione file
3. controlla configurazione backup automatico (checkbox, ora, path)
4. verifica lingua attiva
5. in caso di anomalie, controlla `dataflow.log`

Nota:
- questa routine non modifica i dati applicativi; serve a prevenire problemi di continuità operativa

Una volta verificata la manutenzione, prosegui con il lavoro multiutente sapendo quali limiti operativi possono bloccare le modifiche.

## 13. Lavoro multiutente e limiti operativi

### 13.1 Visibilità dati RFQ multiutente
Per RdO DataFlow può aggregare dati da più database utente (`dataflow_db_*.db`) trovati nella root condivisa.

In dashboard RFQ:
- puoi vedere record altri utenti
- i metadati di ownership (`is_mine`) controllano le azioni consentite

### 13.2 Proprietà record e permessi
Su record non tuoi:
- no delete
- no duplicate
- no archive/reactivate
- doppio click apre dettaglio sola lettura

Importante:
- la proprietà del record è il primo controllo da fare quando un pulsante sembra “non funzionare”.

### 13.3 Multiutente su VSM
Saving/Cost Avoidance:
- lettura aggregata multi-database disponibile
- modifica/eliminazione consentita solo su eventi propri

Derisking:
- backend separato fornitori potenziali
- lettura da database corrente
- non è implementata un’aggregazione multi-database equivalente a RFQ/VSM-event per questa area

### 13.4 Sola lettura: cosa puoi fare
In modalità sola lettura RdO:
- consultare dati testata e griglia
- aprire/scaricare allegati
- esportare Excel RdO

Non puoi:
- modificare campi
- gestire fornitori
- gestire PO
- salvare note/SQDC
- modificare articoli/prezzi

### 13.5 Limiti operativi da tenere presenti
- concorrenza in scrittura sullo stesso DB non è il modello operativo previsto
- alcuni comportamenti visuali tksheet dipendono dal runtime GUI
- Derisking e VSM-event hanno backend diversi: non confondere i due moduli

### 13.6 Scenari multiutente tipici (cosa fare)
Scenario A: vedi la RdO di un collega ma devi lavorarla tu.
1. apri record in sola lettura per consultazione
2. torna in dashboard
3. crea una nuova RdO tua oppure usa duplicazione solo su record di tua proprietà
4. compila i campi necessari nel nuovo record

Scenario B: non riesci a usare `Actions` su una selezione multipla.
1. deseleziona tutto
2. seleziona solo righe con tuo utente
3. riapri `⚡ Actions`

Perché serve:
- evita blocchi da ownership mista e mantiene tracciabilità corretta per utente

Dopo aver chiarito i limiti multiutente, usa i workflow completi per seguire una sequenza reale dall’inizio alla chiusura.

## 14. Workflow pratici completi

### 14.1 Workflow completo RdO
1. apri tab `RdO Attive`
2. clicca `➕ Nuovo Evento`
3. scegli `Fornitura piena`
4. in dettaglio RdO imposta date e riferimento
5. apri `Aggiungi Fornitori` e inserisci elenco (es. `Alfa Srl, Beta Spa, Gamma Srl`)
6. aggiungi articoli manualmente o importa Excel
7. inserisci prezzi nelle colonne fornitore (es. articolo A: Alfa `12,40`, Beta `12,10`, Gamma `ND`)
8. allega documenti offerta e interni
9. inserisci PO associando ordine e fornitore
10. salva nota tecnica
11. esegui export singola RdO
12. opzionale: archivia da dashboard quando chiusa

Risultato operativo:
- hai una RdO completa con confronto fornitori, documenti tracciati e file export pronto per condivisione.

### 14.2 Workflow evento Cost Avoidance
1. vai su tab `Cost Avoidance`
2. clicca `➕ Nuovo Evento`
3. compila data, azione, descrizione
4. imposta importo richiesto iniziale e importo negoziato (es. `150,00` -> `132,00`)
5. imposta % realizzo (es. `75%`)
6. scegli se ripetitivo
7. salva
8. verifica riga in tab con teorico/effettivo/variance
9. export tab Cost Avoidance per report

Controllo rapido:
- dopo il salvataggio verifica subito teorico/effettivo nella tabella; se non tornano, riapri la riga dal tab attivo e correggi i campi.

### 14.3 Workflow Derisking fornitore potenziale
1. vai su tab `Derisking`
2. clicca `➕ Nuovo Evento`
3. inserisci nome fornitore (es. `Delta Components`)
4. seleziona categoria esistente o nuova categoria
5. imposta stato (es. `Nuovo`)
6. compila contatti e note
7. salva
8. se necessario, apri `Gestisci Categorie` e consolida categorie

Aggiornamento tipico:
1. dopo una settimana, riapri lo stesso fornitore con doppio click
2. aggiorna stato a `In valutazione` o `Qualificato`
3. salva e verifica aggiornamento immediato in tabella

### 14.4 Workflow SQDC + export
1. apri RdO con fornitori e prezzi completi
2. clicca pulsante SQDC
3. verifica pesi al 100%
4. calcola Cost automatico
5. completa voti mancanti
6. salva SQDC nei documenti interni
7. esegui export SQDC esterno

Scenario confronto fornitori (esempio):
- fornitore A migliore su Cost, fornitore B migliore su Delivery, fornitore C migliore su Quality.
- con pesi bilanciati (25/25/25/25) il totale evidenzia il vincitore complessivo; se cambi pesi, il ranking può cambiare.

### 14.5 Workflow KPI verifica finale
1. apri finestra KPI
2. seleziona periodo o anno
3. controlla RFQ volume e scadute
4. controlla Saving e Cost Avoidance effettivi
5. controlla Derisking per stato/categoria
6. esporta workbook KPI come allegato report management

Nota operativa:
- usa lo stesso periodo per tutte le tab KPI prima dell’export, così il report finale resta coerente.

### 14.6 Workflow end-to-end: da nuova RdO a verifica KPI
Flusso continuo consigliato:
1. crea RdO in `RdO Attive`
2. compila testata, fornitori, articoli, prezzi, PO, allegati
3. salva SQDC e allega output
4. esporta RdO per condivisione operativa
5. archivia la RdO quando chiusa
6. registra eventuale impatto economico in `Saving` o `Cost Avoidance`
7. aggiorna o inserisci fornitore potenziale in `Derisking` se necessario
8. apri KPI e verifica riflesso dei dati nel periodo corretto
9. esporta KPI e chiudi il ciclo di reporting

Risultato:
- hai una catena completa tracciabile dalla gestione operativa della richiesta fino alla sintesi manageriale

Scenario finale consigliato (operatività -> KPI):
1. chiudi e archivia le RdO concluse
2. aggiorna eventuali eventi Saving/Cost Avoidance legati alle stesse trattative
3. aggiorna Derisking sui fornitori nuovi o in avanzamento
4. apri KPI, applica il periodo del mese corrente e verifica i valori consolidati
5. esporta KPI e allega al report di chiusura periodo

## 15. Troubleshooting operativo

### 15.1 Non riesci a creare o modificare una RdO
Cause tipiche:
- RdO di altro utente aperta in sola lettura
- nessuna riga selezionata per azione
- selezione mista con record non tuoi

Cosa fare:
1. verifica colonna `Utente`
2. lavora su RdO tue
3. usa filtro utente per isolare i record

### 15.2 Ricerca non restituisce risultati attesi
Cause tipiche:
- combinazione AND con filtri avanzati troppo restrittiva
- caratteri speciali rimossi in sanitizzazione
- date non nel formato corretto

Cosa fare:
1. prova solo Global Search
2. poi aggiungi filtri uno per volta
3. usa `🔎 Pulisci Filtri`

### 15.3 Errore import Excel articoli
Cause tipiche:
- colonne insufficienti rispetto al tipo RdO
- righe senza codice o quantità
- file non `.xlsx` valido

Cosa fare:
1. conferma struttura A-D o A-G
2. verifica che codice e quantità siano valorizzati
3. riprova con file semplificato

Suggerimento:
- prova prima con 2-3 righe campione; se l’import è corretto, ripeti con il file completo.

### 15.4 Prezzo non accettato in griglia
Cause tipiche:
- formato numerico non valido
- uso punto al posto di virgola
- valore non ammesso diverso da numero/X/ND

Cosa fare:
1. usa `123,45` e non `123.45`
2. usa solo `X` o `ND` per non quotabile
3. lascia vuoto solo se vuoi cancellare l’offerta

### 15.5 Problemi allegati
Cause tipiche:
- file sorgente non più presente in `Attachments`
- percorso esterno non valido
- tentativo di accesso fuori cartella allegati bloccato

Cosa fare:
1. verifica presenza file nella cartella allegati utente
2. usa download per testare integrità
3. riallega il file se manca

### 15.6 SQDC non calcola Cost correttamente
Cause tipiche:
- prezzi mancanti per uno o più articoli
- presenza di `X` o `ND`
- quantità non valide nella RdO

Cosa fare:
1. completa prezzi numerici per tutti gli articoli
2. verifica quantità articolo
3. rilancia `Calcola Cost Automaticamente`

Importante:
- valori `X` e `ND` servono per non quotabile, ma nel calcolo Cost portano il fornitore a `Cost = 0`.

### 15.7 KPI incoerenti con attese
Cause tipiche:
- filtro anno/preset attivo senza accorgersene
- differenza tra date evento e periodo impatti economici
- nel Derisking record legacy con `created_at` nullo esclusi dai KPI

Cosa fare:
1. controlla filtro in header KPI
2. prova `All`
3. confronta con export KPI per audit

### 15.8 Backup automatico non parte
Cause tipiche:
- opzione attiva ma path vuoto
- ora non corrisponde alla fascia corrente
- percorso non accessibile

Cosa fare:
1. riapri impostazioni backup
2. verifica path, ora, checkbox
3. controlla log applicativo

### 15.9 Multiutente: record non modificabili
Cause tipica:
- record visibile ma non di proprietà utente corrente

Cosa fare:
1. filtra per il tuo utente
2. crea una copia/nuovo record tuo quando serve operare

### 15.10 Sequenza diagnostica rapida (quando non sai dove intervenire)
1. controlla in quale tab sei (`RdO`, `Saving`, `Cost Avoidance`, `Derisking`)
2. verifica ownership del record (`Utente`)
3. controlla se ci sono filtri attivi che restringono i dati
4. ripeti l’azione con un record sicuramente tuo
5. se il problema resta, apri log applicativo e cerca errori nel timestamp dell’operazione

Perché serve:
- in pochi passaggi separi errori di permessi, filtri e dati da errori tecnici reali

Dopo il troubleshooting, riprendi il flusso dal capitolo operativo corrispondente (RdO, SQDC, VSM o KPI) senza ricominciare da capo.

## 16. Glossario
- `RdO` / `RFQ`: richiesta di offerta gestita nei tab RdO.
- `Fornitura piena`: tipo RdO senza campi conto lavoro.
- `Conto lavoro`: tipo RdO con campi grezzo/materiale C/L.
- `Stato RdO`: `attiva` o `archiviata`.
- `Global Search`: ricerca testuale trasversale con logica OR su campi principali.
- `Advanced Filters`: filtri strutturati aggiuntivi con logica AND.
- `Saving`: evento di riduzione costo rispetto a budget o condizioni correnti.
- `Cost Avoidance`: costo evitato rispetto a richiesta iniziale.
- `Derisking`: area anagrafica fornitori potenziali per riduzione rischio supply.
- `Driver Prezzo`: calcolo economico basato su differenza importi e quantità.
- `Driver Pagamenti`: calcolo economico basato su delta giorni pagamento e coefficiente finanziario.
- `OPEX Ripetitivo`: evento con distribuzione impatti su più mesi (fino a 24).
- `One-shot`: evento non ripetitivo con impatto in un solo mese.
- `SQDC`: matrice di valutazione fornitore su Safety, Quality, Delivery, Cost.
- `Documento Interno`: allegato documentale non legato a un fornitore esterno.
- `Offerta Fornitore`: allegato specifico del fornitore selezionato.
- `Carry-over`: quota di impatto economico che ricade sull’anno successivo.
- `Ownership (is_mine)`: indicatore tecnico che abilita/blocca azioni di modifica.

## Note di copertura e limiti dell’analisi
1. Funzioni documentate con alta confidenza:
- avvio/licenza/identità utente
- gestione RdO, articoli, fornitori, prezzi, allegati, PO, note
- import/export RdO e dashboard
- SQDC (validazioni, calcolo cost, salvataggio, export)
- VSM Saving/Cost Avoidance (campi, validazioni, formule, persistenza impatti)
- KPI dashboard (filtri, tab, export)
- impostazioni backup/lingua/spostamento cartella DataFlow

2. Funzioni dedotte solo parzialmente:
- comportamento esatto di alcune azioni popup contestuali tksheet (copy/cut/paste) rispetto al trigger di persistenza in tutti i casi GUI
- ordinamento colonne tksheet in combinazione con dataset molto grandi e refresh frequenti

3. Aree poco chiare dal solo codice statico:
- UX finale di alcune finestre con geometry manager su diversi monitor/DPI
- messaggistica finale utente in presenza di race condition I/O rare su filesystem di rete

4. Aspetti che richiederebbero test runtime per conferma totale:
- interoperabilità completa multiutente in ambienti reali con molti database condivisi
- performance e comportamento di apertura file allegati con applicazioni esterne installate sul client
- comportamento visuale delle card/grafici KPI su resize estremo

5. Componenti legacy o rami non centrali ma presenti:
- tracce legacy `.duckdb` in utility username (mentre il runtime corrente usa `.db`)
- rami commentati/dead code relativi a vecchio Derisking event-based
- fallback BLOB allegati ancora supportato per compatibilità
- nel modello `PotentialSupplier.from_row` è presente un fallback a costante non definita (`SUPPLIER_STATUS_PROSPECT`), potenziale criticità se si attiva quel ramo

## Report finale sintetico
Aree/file codice analizzati:
- bootstrap e shell applicazione: `dataflow.py`, `services/startup_service.py`, `services/app_paths.py`, `database/db_helpers.py`
- dashboard UI e filtri: `ui/main_dashboard_builder.py`, `ui/components/main_dashboard_toolbar.py`, `ui/components/collapsible_filters.py`, `services/dashboard_controller.py`
- persistenza core RFQ/VSM/Derisking: `database_manager.py`
- finestre operative RdO: `ui/windows/view_request_window.py`, `edit_suppliers_window.py`, `edit_reference_window.py`, `purchase_order_window.py`, `notes_window.py`, `attachment_window.py`, `sqdc_analysis_window.py`
- dialog VSM/Derisking: `ui/dialogs/vsm_event_dialog.py`, `potential_supplier_dialog.py`, `manage_supplier_categories_dialog.py`
- modelli e servizi dominio: `models/vsm_event.py`, `models/vsm_impact.py`, `models/potential_supplier.py`, `services/vsm_engine.py`, `services/vsm_persistence.py`, `services/supplier_persistence.py`, `services/supplier_category_persistence.py`
- KPI: `ui/kpi_window.py`, `services/kpi_engine.py`, `services/kpi_chart_data.py`, `services/kpi_excel_export.py`
- supporto identità/validazioni/i18n/format: `utils/user_utils.py`, `utils/string_utils.py`, `utils/validation_utils.py`, `utils/format_utils.py`, `utils/i18n_utils.py`

Parti non documentabili con piena certezza senza runtime:
- comportamento visuale completo di alcuni controlli tksheet e resize multi-monitor
- casistiche estreme I/O su filesystem remoti durante backup/copia/allegati
- UX di interazione con applicazioni esterne all’apertura file

Suggerimenti per futura suddivisione in pagine Wiki/Pages:
1. `Primi Passi e Configurazione` (sezioni 1-3 + 12 base)
2. `Guida Operativa RdO` (sezioni 4-8)
3. `Guida SQDC` (sezione 9)
4. `Guida Value Stream Mapping` (sezione 10 + parte multiutente VSM)
5. `Guida KPI e Reporting` (sezione 11)
6. `Amministrazione, Backup e Troubleshooting` (sezioni 12-15)
7. `Glossario` (sezione 16)
