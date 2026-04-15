# DataFlow RFQ - Analisi tecnica e piano implementazione drag-and-drop allegati

# Obiettivo
Analizzare l'attuale gestione allegati nel Cruscotto RfQ e definire un piano tecnico low-risk, reversibile e locale per aggiungere in futuro il drag-and-drop file come canale alternativo al file picker esistente, senza alterare il comportamento corrente.

# Stato attuale individuato

## Punto di ingresso 1: documenti interni
- [FATTO] Il pulsante `📁 Manage Internal Documents` nel pannello RfQ invoca `open_attachment_window("Documento Interno")` in `ui/windows/view_request_window.py` (linee 124, 1479-1481).
- [FATTO] La finestra usata è sempre `AttachmentWindow`; il tipo è discriminato dal parametro `attachment_type` (`"Documento Interno"`) in `ui/windows/attachment_window.py` (linee 29-37, 64-67).
- [FATTO] In `add_attachment()`, per i documenti interni il fornitore è fissato a `"Interno"` e il nome file archivio segue `RfQ{request_id}_ID{next_id}{ext}` (linee 353, 375-376).

## Punto di ingresso 2: offerte fornitori
- [FATTO] Il pulsante `📄 Manage Supplier Offers` invoca `open_attachment_window("Offerta Fornitore")` in `ui/windows/view_request_window.py` (linee 123, 1479-1481).
- [FATTO] La stessa `AttachmentWindow` abilita UI specifica offerta fornitore: warning dedicato + combobox fornitori (linee 84-91, 107-110).
- [FATTO] In `add_attachment()`, per `"Offerta Fornitore"` è obbligatoria la selezione fornitore (`Select a supplier`) e il nome file archivio include fornitore sanitizzato: `RfQ{request_id}_{supplier_sanitized}_ID{next_id}{ext}` (linee 353-356, 365, 377-378).

# Flusso tecnico attuale
- [FATTO] Apertura dialog selezione file: `filedialog.askopenfilename(...)` in `AttachmentWindow.add_attachment()` (linee 343-348).
- [FATTO] Ricezione path: variabile `filepath`; se vuota il flusso termina (linea 352).
- [FATTO] Validazioni presenti:
  - blocco in read-only (linee 339-341)
  - fornitore obbligatorio per offerte fornitore (linee 353-356)
  - disponibilità cartella allegati (`attachments_base`) (linee 358-361)
- [FATTO] Persistenza file:
  - calcolo `next_id` con `DatabaseManager.get_max_allegato_id()` (linee 366-369; metodo DB linee 1019-1024)
  - copia fisica `shutil.copy(filepath, dest_path)` (linea 380)
  - cartella base allegati da `get_fixed_attachments_dir()` o calcolo da DB remoto in apertura finestra (linee 39-62; `services/app_paths.py` linee 98-116)
- [FATTO] Persistenza DB:
  - inserimento riga in `allegati_richiesta` via `insert_allegato_richiesta_link(...)` (linea 383; `database_manager.py` linee 349-358)
  - tabella `allegati_richiesta` con campi `id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore, percorso_esterno, data_inserimento` (`database_manager.py` linea 179)
- [FATTO] Refresh UI / visualizzazione nome file:
  - chiamata `self.load_attachments()` dopo insert (linea 387)
  - query `get_allegati_by_richiesta(...)` (linee 276-278; metodo DB linee 989-1002)
  - popolamento tabella con colonne `Supplier`, `File Name` (e `Insert Date` se presente) (linee 307-323)

# Mappa file / classi / metodi coinvolti
- `ui/windows/view_request_window.py`
  - `ViewRequestWindow.__init__`: crea i due pulsanti entrypoint allegati (linee 123-124)
  - `ViewRequestWindow.open_attachment_window`: apre `AttachmentWindow` (linee 1479-1481)
- `ui/windows/attachment_window.py`
  - `AttachmentWindow.__init__`: setup finestra, distinzione tipo allegato, init path allegati locale/remoto, combobox fornitori (linee 29-154)
  - `load_suppliers_for_request`: carica fornitori per richiesta (linee 329-336)
  - `add_attachment`: file picker + validazioni + copia + insert DB + refresh (linee 338-387)
  - `load_attachments`: query e refresh tabella allegati (linee 274-327)
  - `delete_attachment`, `open_attachment`, `download_attachment`: ciclo vita file allegati (linee 213-577)
- `database_manager.py`
  - schema `allegati_richiesta` (linea 179)
  - `insert_allegato_richiesta_link` (linee 349-358)
  - `get_allegati_by_richiesta` (linee 989-1002)
  - `get_allegato_file_data` (linee 1007-1014)
  - `get_max_allegato_id` (linee 1019-1024)
  - `delete_allegato` (linee 663-667)
- `services/app_paths.py`
  - `get_fixed_attachments_dir` gestione cartella `Attachments` (linee 98-116)

# Logica condivisa e logica duplicata
- [FATTO] I due ingressi richiesti (documenti interni / offerte fornitori) condividono già lo stesso percorso tecnico in `AttachmentWindow.add_attachment()`, con differenze condizionali su `attachment_type`.
- [FATTO] Non esiste oggi un helper esplicito separato tipo `attach_from_path(filepath, ...)`; la logica è inline in `add_attachment()`.
- [FATTO] Esiste logica parzialmente duplicata in `ui/windows/sqdc_analysis_window.py::save_sqdc` (linee ~744-783): calcolo `next_id`, salvataggio file in `Attachments`, insert/update DB come `Documento Interno` tramite metodo dedicato `insert_or_update_allegato_sqdc` (`database_manager.py` linee 765-798).

# Fattibilità drag-and-drop
- [FATTO] Nel codice attuale non risultano implementazioni DnD per allegati (`tkdnd`/`tkinterdnd2` non trovati nel repo; `requirements.txt` non include librerie DnD).
- [FATTO] L'architettura attuale consente estensione locale: i due ingressi passano già da una finestra unica (`AttachmentWindow`) e convergono su un unico flusso di persistenza (`add_attachment` + `insert_allegato_richiesta_link`).
- [VALUTAZIONE] La fattibilità tecnica è alta se il DnD viene agganciato localmente in `AttachmentWindow` e instradato verso la stessa logica usata dal file picker.

# Opzioni tecniche considerate

## Opzione A: estensione minima del flusso esistente
- Descrizione:
  - mantenere `➕ Add...` + `askopenfilename` invariati
  - aggiungere una drop area dentro `AttachmentWindow`
  - introdurre metodo interno unico (es. `_attach_from_path(filepath)`) chiamato sia da `add_attachment()` sia da handler drop
- Pro:
  - minima superficie di modifica (solo finestra allegati)
  - fallback già presente e invariato
  - rollback semplice (rimozione soli hook DnD/UI)
- Contro:
  - piccola riorganizzazione locale di `add_attachment()` necessaria per evitare duplicazione

## Opzione B: refactor leggero con punto unico “attach from path”
- Descrizione:
  - estrarre la logica path->validazione->copy->insert->refresh in un metodo più strutturato (classe o helper modulo) riusabile anche da `save_sqdc` in futuro
- Pro:
  - maggiore coerenza tra allegati manuali e flussi speciali
  - migliore testabilità unit/local
- Contro:
  - modifica più ampia rispetto ad A
  - rischio non necessario per obiettivo immediato

## Opzione C: altre opzioni eventuali
- C1 - Nessuna nuova dipendenza (best effort): usare solo ciò che Tk espone runtime, se disponibile.
  - Pro: nessun impatto dependency
  - Contro: comportamento DnD non garantito cross-platform
- C2 - Dipendenza opzionale DnD (`tkinterdnd2` o equivalente) attivata solo se presente.
  - Pro: comportamento più prevedibile su Windows/Linux
  - Contro: impatto packaging PyInstaller (hidden import + runtime data tkdnd da includere)
- [PUNTO DA VERIFICARE] Quale runtime Tk/Tcl è distribuito nelle build target e se include già `tkdnd`.

# Soluzione raccomandata
- Raccomandazione: **Opzione A** (estensione minima) con approccio progressivo:
  - 1) mantenere invariato il flusso attuale da pulsante/file dialog
  - 2) aggiungere drop area non invasiva nella stessa finestra
  - 3) convergere file picker e drop su un unico metodo locale `attach_from_path`
  - 4) introdurre DnD come capability opzionale: se non disponibile, la UI continua a funzionare identica a oggi
- Motivazione:
  - allinea i vincoli low-risk/reversibile
  - evita refactor globale
  - tocca solo i file strettamente coinvolti

# Piano di implementazione step-by-step
1. Preparazione tecnica locale (nessun cambio comportamento)
1.1 Mappare in `AttachmentWindow` i blocchi da estrarre da `add_attachment()` in metodo unico interno (validazioni, naming, copy, insert, refresh).
1.2 Definire firma proposta del metodo, ad esempio `_attach_from_path(filepath: str) -> None`, con riuso integrale della logica esistente.

2. Unificazione locale del canale di input file
2.1 Adattare `add_attachment()` perché rimanga solo il wrapper file picker (`askopenfilename`) + chiamata a `_attach_from_path`.
2.2 Garantire che messaggi utente, read-only e vincoli fornitore restino identici.

3. Introduzione drop area non distruttiva
3.1 In `AttachmentWindow.__init__` aggiungere widget/area visuale dedicata al drop (senza rimuovere pulsanti correnti).
3.2 Collegare eventi drag-enter/leave/drop solo all'area dedicata (evitare hook globali finestra per ridurre regressioni UI).
3.3 Nel drop handler normalizzare i path ricevuti e chiamare `_attach_from_path`.

4. Gestione robusta input da drop
4.1 Normalizzare path con spazi/caratteri speciali/formati OS.
4.2 Definire policy file multipli (consigliato MVP low-risk: accettare un solo file e mostrare warning chiaro se multipli).
4.3 Applicare le stesse validazioni già in uso (fornitore obbligatorio, read-only, attachments path disponibile).

5. Compatibilità dependency e packaging (solo se necessaria)
5.1 Verifica runtime: testare se DnD funziona senza dipendenze extra nel target reale.
5.2 Se non sufficiente, aggiungere dipendenza DnD come opzionale e aggiornare spec PyInstaller (`dataflow.spec`, `dataflow_appimage.spec`, eventuali spec build Windows) solo per i runtime data richiesti.
5.3 Mantenere fallback: se init DnD fallisce, nascondere/disabilitare drop area e lasciare solo `➕ Add...`.

6. Validazione manuale minima
6.1 Eseguire test manuali sui due flussi (Documento Interno / Offerta Fornitore).
6.2 Verificare persistenza DB (`allegati_richiesta`) e apertura/download/delete post-upload.
6.3 Verificare assenza regressioni in read-only e in RdO remote (`source_db_path`).

7. Rollout e rollback
7.1 Rilascio con flag/fallback implicito (feature attiva solo se capability DnD disponibile).
7.2 Rollback rapido: rimozione binding/widget DnD senza toccare backend allegati.

# Rischi
- Compatibilità Tkinter/DnD:
  - rischio che DnD non sia disponibile uniformemente sui runtime target.
- Packaging Windows:
  - rischio omissione runtime tkdnd/hidden imports nelle build PyInstaller.
- Linux Mint:
  - rischio differenze comportamento drag data payload e dipendenze Tk lato distro.
- Path edge case:
  - stringhe path da drop con quoting/braces, spazi, caratteri speciali.
- File multipli:
  - DnD può passare più file mentre il flusso corrente è single-file.
- Validazione file:
  - oggi non esistono whitelist estensione/size; DnD amplierebbe superficie input.
- Regressione UI:
  - rischio collisione eventi drag con selezione tksheet o focus/grab.
- Regressione persistenza:
  - rischio inconsistenza se il nuovo handler non riusa esattamente la logica copy+insert esistente.
- Differenze tra due ingressi:
  - offerte fornitore richiedono supplier obbligatorio; documenti interni no. Questo impone gestione condizionale anche nel drop handler.

# Rollback
- Strategia rollback consigliata:
  - rimuovere/disattivare binding DnD e relativa drop area in `AttachmentWindow`
  - mantenere intatto `add_attachment()` con file picker
  - nessuna migrazione DB necessaria (schema invariato)
- Impatto rollback:
  - ritorno immediato al comportamento preesistente senza impatto su allegati già salvati.

# Test manuali consigliati
1. Documento Interno - file picker
1.1 Aprire `Manage Internal Documents`.
1.2 Allegare file via `➕ Add...`.
1.3 Verificare comparsa riga con nome file e apertura/download.

2. Offerta Fornitore - file picker
2.1 Aprire `Manage Supplier Offers`.
2.2 Provare allegato senza selezionare fornitore (atteso warning).
2.3 Selezionare fornitore, allegare file, verificare riga e operazioni apertura/download.

3. Documento Interno - drag-and-drop
3.1 Trascinare file nella drop area.
3.2 Verificare stesso esito del file picker (copy, insert DB, refresh riga).

4. Offerta Fornitore - drag-and-drop
4.1 Drop senza fornitore selezionato (atteso stesso warning del picker).
4.2 Drop con fornitore selezionato (atteso successo).

5. Edge cases path/input
5.1 File con spazi e caratteri speciali nel nome.
5.2 Drop di più file (atteso comportamento definito, es. warning + primo file).
5.3 File non accessibile/permessi negati.

6. Read-only e database remoto
6.1 Aprire RdO altrui in read-only.
6.2 Verificare che Add/Drop non consentano upload.

7. Regressione generale
7.1 Cancellare allegato appena inserito e verificare rimozione DB/file.
7.2 Verificare che gestione SQDC e pulsante SQDC non subiscano regressioni.

# Domande aperte / punti da confermare
- [PUNTO DA VERIFICARE] Runtime Tk/Tcl effettivo nelle build distribuite (Win/Linux) e disponibilità nativa di DnD.
- [PUNTO DA VERIFICARE] Policy prodotto su drop multi-file: supporto batch o single-file MVP.
- [PUNTO DA VERIFICARE] Scope UX esatto della drop area: area esplicita dedicata vs intera finestra allegati.
- [PUNTO DA VERIFICARE] Eventuale necessità di messaggi UX aggiuntivi (hint testuale/internazionalizzazione) per spiegare che il file picker resta disponibile.
- [PUNTO DA VERIFICARE] Se includere o meno il flusso SQDC nel primo ciclo di unificazione (al momento è un canale interno separato).

