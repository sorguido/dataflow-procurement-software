# 04 – Gestire una RdO esistente

## Aprire una RdO

1. Nel tab **RdO Attive** (o **RdO Archiviate**), individuare la RdO desiderata.
2. Fare **doppio clic** sulla riga per aprire il pannello di controllo.

Se la RdO appartiene a un altro utente, si apre in **modalità sola lettura**: tutte le celle e i pulsanti di modifica sono disabilitati e compare un banner rosso in fondo alla finestra. È possibile consultare tutti i dati ma non modificarli.

---

## Visualizzare e modificare il riferimento

Il campo **Riferimento** nell'intestazione è cliccabile come etichetta:

1. Fare clic sull'etichetta del riferimento.
2. Si apre una piccola finestra con un campo di testo pre-compilato.
3. Modificare il testo e fare clic su **💾 Salva**, oppure **❌ Annulla** per non salvare.

---

## Modificare la griglia prezzi

La griglia prezzi è completamente editabile:

- **Fare clic su una cella** per selezionarla e modificarne il contenuto.
- Usare **Tab** o **Invio** per spostarsi alla cella successiva.
- I prezzi usano la virgola come separatore decimale (`12,50`).
- Le celle dei prezzi lasciate vuote indicano che il fornitore non ha fornito offerta.

Per le RdO di tipo Conto lavoro, le colonne **Cod. Grezzo**, **Allegato Grezzo** e **Mat. C/L** sono editabili allo stesso modo delle altre.

---

## Aggiungere e rimuovere fornitori

I fornitori già presenti nella RdO determinano le colonne prezzo. Per modificarli:

1. Fare clic su **Fornitori (N)** in alto.
2. Modificare l'elenco dei fornitori (nomi separati da virgola).
3. Fare clic su **💾 Salva**.

> Attenzione: **rimuovere un fornitore dall'elenco elimina anche tutti i prezzi inseriti per quel fornitore**. L'operazione non è reversibile.
>
> Durante la modifica dell'elenco, DataFlow può proporre nomi fornitore già presenti in altre RdO o nel Derisking. Se un nome è molto simile a uno esistente, prima del salvataggio può comparire un avviso non bloccante.

---

## Gestione allegati

DataFlow distingue due tipologie di allegati:

| Tipo | Quando utilizzarlo |
|------|--------------------|
| **Offerta Fornitore** | Offerte ricevute in formato PDF, Excel, ecc. |
| **Documento Interno** | Capitolati tecnici, specifiche, disegni, autorizzazioni, analisi SQDC |

### Aggiungere un allegato

1. Fare clic su **📄 Gestisci Offerte Fornitori** oppure **📁 Gestisci Documenti Interni**.
2. Per le offerte fornitore: selezionare prima il fornitore dal menu a tendina.
3. Fare clic su **➕ Aggiungi...** e selezionare il file nella finestra di dialogo.
4. In alternativa, trascinare direttamente il file nell'elenco allegati.
5. Il file viene copiato nella cartella `Attachments/{numero RdO}/` e il percorso relativo viene salvato nel database.

> Il file originale **non viene spostato né eliminato**. DataFlow conserva la propria copia.

### Aprire un allegato

1. Fare clic sulla riga dell'allegato per selezionarla.
2. Fare clic su **📂 Apri Selezionato**.
3. Il file viene aperto con il programma predefinito del sistema operativo.

### Scaricare un allegato

1. Selezionare la riga dell'allegato.
2. Fare clic su **⬇️ Download...**
3. Scegliere la cartella di destinazione.

### Eliminare un allegato

1. Selezionare la riga dell'allegato.
2. Fare clic su **❌ Elimina Selezionato**.

> Per le RdO di altri utenti, i pulsanti di aggiunta e di eliminazione sono disabilitati, ma l'apertura e il download rimangono sempre disponibili.

---

## Note formattate

Le note consentono di annotare contesti di negoziazione, comunicazioni ricevute o considerazioni tecniche:

1. Fare clic su **📝 Nota** (o **📝 Aggiungi Nota** se assente).
2. Nell'editor, digitare il testo. È possibile applicare:
   - **𝐁 Grassetto**
   - **𝑰 Corsivo**
   - **U̲ Sottolineato**
3. Fare clic su **💾 Salva Nota**.

Le note non hanno un limite di lunghezza pratico (limite tecnico: 1 MB di contenuto). Non è possibile salvare note con più di 10.000 elementi interni di formattazione.

---

## Numeri ordine di acquisto

Per registrare gli ordini emessi a seguito della negoziazione:

1. Fare clic su **📋 Inserisci OdA**.
2. Inserire il numero ordine e selezionare il fornitore.
3. Fare clic su **➕ Aggiungi**.
4. La tabella aggiorna immediatamente l'elenco degli ordini.
5. Fare clic su **Chiudi** (i dati vengono salvati automaticamente alla chiusura).

Per modificare un ordine già inserito, fare doppio clic sulla cella nella tabella. Per eliminarlo, selezionare la riga e fare clic su **❌ Elimina**.

---

## Archiviare una RdO

Le RdO chiuse o completate possono essere archiviate:

1. Tornare al tab **RdO Attive**.
2. Selezionare la RdO (clic singolo).
3. Fare clic su **⚡ Actions** → **📦 Archivia**.

La RdO non sarà più visibile nel tab RdO Attive, ma potrà essere consultata nel tab **RdO Archiviate**. Per riportarla attiva, selezionarla nel tab archiviato e usare **⚡ Actions** → **♻️ Riattiva**.

---

## Duplicare una RdO

Per creare una nuova RdO partendo da una esistente (stessi articoli, stessi fornitori):

1. Selezionare la RdO da copiare.
2. Fare clic su **⚡ Actions** → **📋 Duplica**.

Viene creata una copia con nuova data di emissione e numero progressivo. I prezzi inseriti nella griglia vengono copiati. Le note e gli allegati **non** vengono copiati.

---

## Eliminare una RdO

1. Selezionare la RdO.
2. Fare clic su **⚡ Actions** → **🗑 Elimina**.
3. Confermare nella finestra di dialogo.

> L'eliminazione è **irreversibile** e rimuove anche tutti gli articoli, i prezzi, le note e i metadati ordini. Gli allegati fisici rimangono nella cartella `Attachments/` e devono essere rimossi manualmente se non più necessari.

---

## Esportare la RdO

Nel pannello della RdO è disponibile il menu **📊 Esporta**, con due opzioni:

### Excel

Per condividere il confronto prezzi con colleghi o responsabili:

1. Nel pannello della RdO, fare clic su **📊 Esporta** → **📗 Excel**.
2. Scegliere la lingua del file (Italiano / English).
3. Selezionare la cartella e il nome del file.
4. Il file Excel viene generato con intestazioni in grassetto e sfondo grigio, prezzi formattati, e una colonna per ciascun fornitore.

### PDF

Per generare una RdO stampabile:

1. Nel pannello della RdO, fare clic su **📊 Esporta** → **📄 PDF**.
2. Nella finestra di export, configurare facoltativamente il logo aziendale oppure usare **Modifica PDF** per personalizzare il testo del PDF.
3. Fare clic su **Conferma Export PDF** e scegliere il file di destinazione.

Il pulsante **📥 Export Excel** nella barra degli strumenti principale esporta invece **tutte** le RdO del database in un unico file.

---
[← Pagina precedente](IT-03-Creare-una-nuova-RdO) | [Pagina successiva →](IT-05-Analisi-SQDC)
