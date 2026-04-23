# 10 – Problemi comuni e soluzioni

## Database bloccato

**Sintomo:** Al salvataggio o alla creazione di un nuovo elemento, compare un messaggio di errore che indica che il database è occupato o bloccato.

**Causa:** Un altro utente sta eseguendo un'operazione di scrittura nello stesso momento, oppure un processo precedente non si è chiuso correttamente.

**Soluzioni:**

1. **Aspettare e riprovare.** DataFlow attende automaticamente fino a 10 secondi prima di segnalare l'errore. Se l'errore compare successivamente, riprovare dopo qualche secondo: il conflitto si risolve da solo nella maggior parte dei casi.

2. **Verificare che tutti gli utenti abbiano chiuso DataFlow.** Se un'istanza è rimasta aperta su un altro computer senza essere terminata correttamente (es. computer spento forzatamente), potrebbe esserci un file di lock residuo.

3. **Cercare file temporanei WAL.** Nella cartella `DataFlow/Database/` potrebbero essere presenti i file `dataflow_db.db-wal` e `dataflow_db.db-shm`. In condizioni normali questi file esistono mentre l'applicazione è aperta e vengono incorporati nel database alla chiusura. Se rimangono dopo che tutti gli utenti hanno chiuso DataFlow, significa che l'ultima chiusura è avvenuta in modo anomalo. Aprire e richiudere DataFlow su una qualsiasi workstation per completare il checkpoint WAL.

---

## La RdO si apre in sola lettura (ma è mia)

**Sintomo:** Aprendo una propria RdO, compare il banner rosso di sola lettura.

**Causa possibile:** Il nome utente dell'installazione corrente non corrisponde a quello registrato nella RdO. Può capitare se si è reinstallato DataFlow su un nuovo computer inserendo un nome leggermente diverso, o se si è cambiata la lingua e il sistema ha rigenerato il nome utente.

**Soluzione:** Verificare in **⚙️ Impostazioni** il nome utente corrente e confrontarlo con quello mostrato nella colonna "Utente" della RdO. Se non corrispondono, contattare il supporto per aggiornare il campo username nella RdO.

---

## Gli allegati non si aprono

**Sintomo:** Facendo clic su "Apri Selezionato" nella finestra allegati, non succede niente o compare un errore.

**Causa 1 – File non trovato:** Il file allegato punta a un percorso relativo che non esiste più (es. la cartella `Attachments/` è stata spostata o rinominata).

**Soluzione:** Verificare che la cartella `DataFlow/Attachments/{numero RdO}/` esista e contenga il file. Se la cartella è stata spostata, riallineare la posizione DataFlow nelle impostazioni.

**Causa 2 – Nessun programma associato all'estensione:** Il sistema operativo non ha un programma predefinito per aprire quel tipo di file.

**Soluzione:** Installare un programma adatto (es. Acrobat Reader per i PDF) e impostarlo come predefinito per quell'estensione.

**Causa 3 – File su rete non raggiungibile:** Se il database è condiviso e la cartella `Attachments/` si trova su un server non raggiungibile al momento.

**Soluzione:** Verificare la connessione di rete e riprovare.

---

## Errore durante l'importazione da Excel

**Sintomo:** Facendo clic su "Importa da Excel", compare un messaggio di errore o le righe non vengono importate correttamente.

**Cause possibili:**

- Il file Excel non ha il formato atteso (colonne nell'ordine sbagliato, intestazioni mancanti).
- Il file è in un formato vecchio (`.xls`) anziché `.xlsx`.
- Il file è aperto in Excel su un altro computer (blocco di lettura).

**Soluzioni:**

1. Usare il file Excel generato dalla funzione "Esporta Excel" di DataFlow come template per le importazioni.
2. Assicurarsi che il file sia in formato `.xlsx`.
3. Chiudere il file Excel (se aperto altrove) prima di importare.

---

## I valori KPI sembrano errati o incompleti

**Sintomo:** Il KPI Dashboard mostra valori diversi da quelli attesi, o eventi che sembrano mancare.

**Cosa verificare:**

1. **Controllare il filtro temporale attivo.** Un filtro per anno diverso da quello atteso, o un filtro rolling (es. 3M) che esclude eventi più vecchi, può ridurre i totali. Fare clic su **All** per vedere tutti i dati senza limiti.

2. **Ricordare che il filtro agisce sulla competenza economica, non sulla data evento.** Un evento creato a novembre dell'anno scorso con distribuzione a 12 mesi contribuisce al KPI di quest'anno per i mesi di distribuzione che ricadono nell'anno corrente.

3. **Verificare la % Realizzo degli eventi.** Se il Saving Effettivo è molto più basso del Teorico, controllare che il campo "% Realizzo" degli eventi non sia stato impostato a 0 o a un valore basso.

---

## L'applicazione è lenta all'avvio

**Causa probabile:** La cartella DataFlow si trova su una rete lenta o poco disponibile. Al primo accesso, DataFlow deve aprire il database e verificarne l'integrità.

**Soluzioni:**

1. Verificare la velocità della connessione di rete alla cartella condivisa.
2. Se non serve lavorare in modalità condivisa, spostare il database su un disco locale.

---

## I caratteri speciali nei nomi fornitore vengono rimossi

**Sintomo:** Digitando nel campo di ricerca alcuni caratteri (apostrofo, virgolette, backslash, parentesi angolari), questi vengono rimossi automaticamente con un avviso.

**Motivo:** DataFlow filtra caratteri che potrebbero causare problemi nelle query di ricerca. Il messaggio avvisa l'utente che l'input è stato sanitizzato.

**Soluzione:** Usare solo caratteri alfanumerici e spazi per la ricerca. Se il nome del fornitore contiene caratteri speciali, cercare una parte del nome senza di essi (es. cercare "Fornitore" invece di "Fornitore & Figli").

---

## Recuperare dati da un backup

Se si vuole ripristinare un backup precedente:

1. Chiudere DataFlow su tutte le workstation.
2. Rinominare il file corrente: `dataflow_db.db` → `dataflow_db.db.old`
3. Copiare il file di backup nella cartella `DataFlow/Database/` rinominandolo `dataflow_db.db`.
4. Riavviare DataFlow.

> Prima di sovrascrivere il database corrente con un backup, assicurarsi di non perdere dati importanti inseriti successivamente al backup. L'operazione è irreversibile se il vecchio file viene eliminato.

---
[← Pagina precedente](09-Lavoro-multiutente) | [Pagina successiva →](11-Best-practices)
