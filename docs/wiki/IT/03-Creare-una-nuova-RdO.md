# 03 – Creare una nuova RdO

## Tipi di RdO

DataFlow gestisce due tipi di Richiesta di Offerta:

| Tipo | Quando si usa |
|------|---------------|
| **Fornitura piena** | Fornitura di componenti o prodotti finiti. Il fornitore fornisce tutto (materiale + lavorazione). |
| **Conto lavoro** | Il grezzo o il materiale è fornito dall'azienda; il fornitore esegue solo la lavorazione. Richiede l'inserimento di codice grezzo, disegno e materiale. |

---

## Avviare la creazione

1. Fare clic su **➕ Nuovo Evento** nella barra degli strumenti, mentre si è nel tab **RdO Attive** o **RdO Archiviate**.
2. Nella finestra di selezione tipo, scegliere:
   - **📦 Fornitura piena**
   - **🔧 Conto lavoro**
   - **❌ Annulla** per tornare indietro
3. Si apre automaticamente il **pannello di controllo** della nuova RdO.

---

## Pannello di controllo della RdO

Il titolo della finestra mostra: `Control Panel - User: [nome] - Request N° [numero] - [tipo]`.

### Intestazione

| Campo | Note |
|-------|------|
| **Data Emissione** | Selettore data. Viene salvato automaticamente alla selezione o alla perdita del focus. |
| **Data Scadenza** | Selettore data. Stesso comportamento automatico. |
| **Riferimento** | Etichetta cliccabile. Fare clic per aprire la finestra di modifica e inserire il riferimento (es. nome progetto, commessa, cliente). |

### Pulsanti della barra superiore

| Pulsante | Azione |
|----------|--------|
| **📄 Gestisci Offerte Fornitori** | Apre la gestione allegati per i documenti ricevuti dai fornitori |
| **📁 Gestisci Documenti Interni** | Apre la gestione allegati per i documenti interni (capitolati, specifiche, ecc.) |
| **Fornitori (N)** o **➕ Aggiungi Fornitori** | Apre la finestra per inserire o modificare l'elenco fornitori |
| **📝 Nota** o **📝 Aggiungi Nota** | Apre l'editor note con formattazione testo (grassetto, corsivo, sottolineato) |
| **📊 Esporta Excel** | Esporta la griglia prezzi in un file Excel |
| **📊 SQDC** o **📊 SQDC ✓** | Apre l'analisi SQDC (il ✓ indica che esiste già un'analisi salvata) |

---

## Inserire i fornitori

Prima di inserire i prezzi nella griglia è consigliabile impostare i fornitori:

1. Fare clic su **➕ Aggiungi Fornitori** (o sul pulsante **Fornitori (N)** se ci sono già).
2. Inserire i nomi dei fornitori nel campo testo, **separati da virgola** (es. `Fornitore A, Fornitore B, Fornitore C`).
3. Fare clic su **💾 Salva**.

> DataFlow non accetta nomi fornitore duplicati (case-insensitive). Se si inserisce due volte lo stesso nome, il salvataggio viene bloccato con un avviso.

Ogni fornitore aggiungerà **una colonna prezzo** nella griglia.

---

## Inserire gli articoli manualmente

Nella parte centrale del pannello di controllo si trova la **griglia prezzi**. Per aggiungere articoli:

1. Fare clic su **➕ Aggiungi Articolo** (pulsante in basso a sinistra).
2. Viene aggiunta una riga vuota.
3. Fare clic sulla cella da compilare e digitare il valore.

### Colonne per RdO Fornitura piena

| Colonna | Contenuto |
|---------|-----------|
| **Pos.** | Numero posizione (automatico) |
| **Allegato** | Riferimento al disegno tecnico o documento |
| **Descrizione** | Descrizione dell'articolo |
| **Q.tà** | Quantità richiesta (usare la virgola come separatore decimale) |
| **[Fornitore 1]** | Prezzo unitario offerto dal fornitore 1 |
| **[Fornitore 2...]** | Una colonna per ogni fornitore inserito |

### Colonne aggiuntive per RdO Conto lavoro

Dopo le colonne base, compaiono tre colonne in più:

| Colonna | Contenuto |
|---------|-----------|
| **Cod. Grezzo** | Codice del materiale grezzo fornito dall'azienda |
| **Dis. Grezzo** | Numero del disegno del grezzo |
| **Mat. C/L** | Descrizione del materiale da lavorare |

### Inserimento prezzi

I prezzi vengono inseriti direttamente nelle celle della colonna del fornitore corrispondente. Regole:

- Usare la **virgola** come separatore decimale (es. `12,50`)
- Non usare il punto come separatore decimale
- Non inserire il simbolo valuta (€)
- Il campo vuoto equivale a "offerta non ricevuta"

---

## Importare gli articoli da Excel

Se esiste già un file Excel con la lista articoli, è possibile importare le righe senza inserirle manualmente:

1. Fare clic su **📊 Importa da Excel** (pulsante in basso, accanto ad Aggiungi Articolo).
2. Selezionare il file Excel nella finestra di dialogo.
3. Il sistema legge il file e inserisce le righe nella griglia.

> Il file Excel deve avere il formato compatibile (lo stesso che si ottiene esportando da DataFlow). In caso di errore di formato, viene mostrato un messaggio con il problema riscontrato.

---

## Rimuovere un articolo

1. Fare clic sulla riga da rimuovere per selezionarla.
2. Fare clic su **🗑 Rimuovi Articolo Selezionato**.

> La rimozione è immediata e non richiede conferma. Operare con attenzione.

---

## Salvare la RdO

La RdO viene **salvata automaticamente** ogni volta che si modifica un dato (data, riferimento, prezzi nella griglia). Non esiste un pulsante "Salva" esplicito. Si può chiudere il pannello in qualsiasi momento.

---

## Aggiungere note alla RdO

Per annotazioni, contesti di negoziazione o riepilogo della trattativa:

1. Fare clic su **📝 Aggiungi Nota**.
2. Digitare il testo nell'editor. Usare i pulsanti di formattazione: **𝐁** grassetto, **𝑰** corsivo, **U̲** sottolineato.
3. Fare clic su **💾 Salva Nota**.

Le note supportano testo formattato con stili. La finestra note blocca la schermata sottostante durante la modifica; viene rilasciata alla chiusura.

---

## Inserire il numero ordine di acquisto

Quando viene emesso un ordine di acquisto a seguito della RdO:

1. Fare clic su **📋 Inserisci OdA** nell'intestazione del pannello.
2. Inserire il **Numero Ordine** nel campo testo.
3. Selezionare il **Fornitore** dal menu a tendina (solo i fornitori già presenti nella RdO).
4. Fare clic su **➕ Aggiungi**.
5. Fare clic su **Chiudi** per salvare.

È possibile inserire più ordini per la stessa RdO (uno per fornitore o per tranche).
