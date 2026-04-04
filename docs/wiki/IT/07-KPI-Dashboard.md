# 07 – KPI Dashboard

## Aprire la dashboard KPI

Fare clic su **≋ KPI** nella barra degli strumenti principale. La finestra si apre a schermo intero.

---

## Struttura della finestra

La parte superiore contiene i controlli di filtro temporale:

### Filtro per Periodo (finestra mobile)

I pulsanti **1M**, **3M**, **12M**, **3Y**, **5Y**, **10Y**, **All** filtrano i dati con una finestra mobile calcolata a ritroso da oggi:

| Pulsante | Intervallo |
|----------|-----------|
| 1M | Ultimi 30 giorni |
| 3M | Ultimi 90 giorni |
| 12M | Ultimi 365 giorni |
| 3Y | Ultimi 3 anni |
| 5Y | Ultimi 5 anni |
| 10Y | Ultimi 10 anni |
| All | Tutti i dati, senza limite temporale |

### Filtro per Anno (anno solare)

Il menu a tendina **Anno** permette di selezionare un anno solare specifico. Quando si usa il filtro per Anno, si ottengono **esattamente 12 mesi fissi** da gennaio a dicembre, indipendentemente dalla data odierna.

> I filtri Periodo e Anno sono **mutualmente esclusivi**: selezionarne uno deseleziona automaticamente l'altro.

### Come il filtro si applica ai KPI di Saving e Cost Avoidance

> **Importante:** il filtro temporale per i KPI di Saving e Cost Avoidance agisce sull'**anno e mese di competenza economica** (ossia il mese in cui l'impatto si manifesta), non sulla data in cui l'evento è stato creato.

Questo significa che:
- Un evento di gennaio con distribuzione a 12 mesi ha impatti fino a dicembre.
- Filtrando su "anno corrente", vengono inclusi anche gli impatti di eventi creati l'anno precedente purché abbiano competenza nell'anno corrente.
- Questo riflette fedelmente la realtà contabile del procurement.

---

## Tab RFQ

Mostra i KPI relativi all'attività di emissione richieste di offerta.

### Schede KPI disponibili

| KPI | Significato |
|-----|-------------|
| RFQ Attive | Numero di RdO nello stato "attiva" nel periodo |
| RFQ Archiviate | Numero di RdO nello stato "archiviata" |
| RFQ Totali | Somma attive + archiviate |
| RFQ Non Scadute | RdO con data scadenza nel futuro |
| RFQ Scadute | RdO con data scadenza passata |
| Conto Lavoro | RdO di tipo lavorazione esterna |
| Fornitura Piena | RdO di tipo fornitura completa |

### Grafico

Istogramma che mostra il numero di RdO emesse per mese. Ogni barra corrisponde a un mese del periodo selezionato. I mesi senza attività mostrano una barra a zero.

### Tabella dettagli

Sotto il grafico compare la tabella `Periodo | RFQ Emesse` con i dati numerici ordinati dal più recente al più antico.

---

## Tab Saving

### Schede KPI disponibili

| KPI | Significato |
|-----|-------------|
| **Saving Teorico** | Somma del valore teorico mensile di tutti gli eventi Saving nel periodo filtrato |
| **Saving Effettivo** | Somma del valore effettivo (Teorico × % Realizzo / 100) |
| **% Saving Media** | Media ponderata delle percentuali saving di ogni evento |
| **% Saving Migliore** | Percentuale più alta registrata tra tutti gli eventi del periodo |
| **% Saving Peggiore** | Percentuale più bassa registrata tra tutti gli eventi |
| **% Saving Mediana** | Valore mediano delle percentuali saving |
| **Impatto Ricorrente (€)** | Saving da eventi con opzione OPEX Ripetitivo attiva |
| **Impatto Non Ricorrente (€)** | Saving da eventi una tantum |

### Come viene calcolata la % Saving Media

Non è una media aritmetica semplice, ma una **media ponderata**:

$$\text{Saving\% Media} = \frac{\sum \text{Saving}_{evento}}{\sum \text{Base}_{evento}} \times 100$$

La "Base" è `Importo Budget × Quantità Annua` per il driver Prezzo, oppure `Spending Annuo` per il driver Pagamenti. Questo evita che un evento piccolo con percentuale alta distorca il risultato complessivo.

### Carry-over (solo con filtro Anno)

Quando si usa il **filtro per Anno**, compare un KPI aggiuntivo: **Carry-over verso anno successivo (€)**.

Questo valore rappresenta il totale degli impatti economici di eventi già creati nell'anno selezionato (o prima) che si manifesteranno nell'**anno successivo**. È utile per proiezioni di budget e per dare evidenza al management del valore "già in pipeline" per l'anno prossimo.

Un saving da un contratto pluriennale siglato a novembre, per esempio, avrà impatti negli undici mesi successivi: il valore dei mesi dell'anno prossimo è il carry-over.

### Grafico

Istogramma doppio: barre blu per il Saving Teorico e barre arancioni per il Saving Effettivo, affiancate per ogni mese del periodo.

---

## Tab Cost Avoidance

Struttura identica al tab Saving. I KPI corrispondenti sono:

| KPI | Significato |
|-----|-------------|
| **Cost Avoidance Teorico** | Somma del valore teorico degli eventi Cost Avoidance nel periodo |
| **Cost Avoidance Effettivo** | Valore effettivo dopo l'applicazione della % realizzo |
| **% CA Media** | Media ponderata delle % di avoidance |
| **% CA Migliore / Peggiore / Mediana** | Statistiche per evento |
| **Impatto Ricorrente / Non Ricorrente** | Distribuzione per tipo di evento |
| **Carry-over verso anno successivo** | Solo con filtro Anno attivo |

---

## Tab Derisking

### Schede KPI disponibili

| KPI | Significato |
|-----|-------------|
| **Totale Fornitori Potenziali** | Numero di fornitori registrati nel periodo |
| **Categorie Uniche** | Numero di categorie merceologiche distinte |
| **Nuovo** / **In valutazione** / **Qualificato** / **Scartato** | Conteggio per stato |

### Grafico

Istogramma con il numero di nuovi fornitori registrati per mese.

### Tabella dettagli

Riepilogo per categoria con il numero di fornitori in ciascuna.

---

## Esportare i KPI in Excel

Fare clic su **📥 Export Excel** nell'angolo in alto a destra della finestra KPI.

**Passo 1 – Scegliere la sezione da esportare:**
- **📋 Sezione corrente** – esporta solo i dati del tab attivo
- **📊 Tutte le sezioni** – esporta tutti e quattro i tab in un unico file Excel

**Passo 2 – Scegliere la lingua:** Italiano o English.

**Passo 3 – Scegliere dove salvare il file.**

Il file Excel generato contiene:
- Un foglio **Riepilogo** con i metadati (data esportazione, filtro applicato, perimetro)
- Un foglio per ogni sezione esportata con i valori numerici
- Formattazione numerica: monetario `€ 1.234,56`; percentuale `12,34%`

---

## Aggiornamento dei dati

I dati nella finestra KPI riflettono lo stato del database al momento dell'apertura. Per aggiornare i valori dopo aver aggiunto nuovi eventi, chiudere e riaprire la finestra KPI, oppure cambiare il filtro temporale per forzare un ricalcolo.
