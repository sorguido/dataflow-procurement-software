# 06 – Value Stream Mapping

## Panoramica

Il modulo **Value Stream Mapping** (d'ora in poi indicato per esteso o come "modulo VSM") permette di registrare e tracciare le attività di miglioramento economico del procurement. Le attività sono suddivise in tre tipologie:

| Tab | Tipologia | Descrizione |
|-----|-----------|-------------|
| **Saving** | Reducing cost | Riduzione di costi rispetto al budget o al costo storico |
| **Cost Avoidance** | Avoiding cost increase | Blocco di un aumento di costo richiesto dal fornitore o dal mercato |
| **Derisking** | Supply chain risk reduction | Qualifica di nuovi fornitori per ridurre la dipendenza da un singolo fornitore |

---

## Creare un nuovo evento Saving

1. Fare clic sul tab **Saving**.
2. Fare clic su **➕ Nuovo Evento**.
3. Si apre la finestra **"Nuovo Evento VSM"** con il tipo preimpostato su "Saving".

### Sezione Informazioni Generali

| Campo | Note |
|-------|------|
| **Data Evento** | Data in cui la negoziazione è stata conclusa o formalizzata (obbligatoria) |
| **Tipo Evento** | Preimpostato su "Saving"; non modificabile |
| **Azione** | Selezionare: `Negoziazione` / `Derisking` / `Altro` |
| **Utente** | Compilato automaticamente con il proprio nome; non modificabile |

### Sezione Descrizione

Campo testo libero. Descrivere brevemente l'oggetto della negoziazione, il fornitore, il contesto.

### Sezione Riferimento

Campo testo per collegare l'evento a una RdO, un OdA, un contratto o un fornitore specifico.

### Sezione Dati Economici – Driver Prezzo

Compilare quando il saving deriva da una riduzione di prezzo unitario:

| Campo | Note |
|-------|------|
| **Importo a Budget (€)** | Prezzo unitario o importo complessivo a budget (usare la virgola come separatore decimale) |
| **Importo Negoziato (€)** | Prezzo unitario o importo effettivamente negoziato |
| **Quantità Annua** | Numero di pezzi o unità annui. Default: 1 |
| **% Realizzo** | Percentuale prevista di effettivo conseguimento del saving. Default: 100 |

Il valore teorico viene calcolato automaticamente al salvataggio:

$$V_{teorico} = Q_{annua} \times (\text{Importo Budget} - \text{Importo Negoziato})$$

### Sezione Dati Economici – Driver Pagamenti

Compilare quando il saving deriva da un miglioramento dei termini di pagamento (es. da 30 a 90 giorni):

| Campo | Note |
|-------|------|
| **Spending Annuo (€)** | Spesa annua con il fornitore su cui si applica il saving di pagamento |
| **Termini Pagamento Attuali (giorni)** | Es. `30` |
| **Termini Pagamento Negoziati (giorni)** | Es. `90` |
| **Financial Impact (% per 30 giorni)** | Coefficiente finanziario per 30 giorni. Default: `0,50%` (configurato nelle impostazioni) |

$$V_{teorico} = \text{Spending Annuo} \times \frac{(\text{Giorni Neg.} - \text{Giorni Att.})}{30} \times \text{Coefficiente}$$

---

## Creare un nuovo evento Cost Avoidance

Il flusso è identico al Saving. La differenza è nella sezione Dati Economici – Driver Prezzo:

| Campo | Note |
|-------|------|
| **Importo Richiesto Iniziale (€)** | Prezzo o importo richiesto originariamente dal fornitore |
| **Importo Negoziato (€)** | Importo effettivamente concordato dopo la negoziazione |
| **Quantità Annua** | Default: 1 |
| **% Realizzo** | Default: 100 |

$$V_{teorico} = Q_{annua} \times (\text{Importo Richiesto Iniziale} - \text{Importo Negoziato})$$

> Il driver Pagamenti **non è disponibile** per gli eventi Cost Avoidance.

---

## Distribuire il valore nel tempo (OPEX Ripetitivo)

Per default, il valore economico di un evento viene registrato **interamente nel mese della data evento** (impatto una tantum).

Se la negoziazione genera un beneficio economico che si ripeterà ogni mese (tipicamente OPEX: contratti di servizio, forniture a canone, accordi di fornitura pluriennali), attivare l'opzione:

- **☑ OPEX Ripetitivo (distribuzione multi-mese)**

Con questa opzione attiva, DataFlow distribuisce il valore teorico su **fino a 24 mesi** a partire dal mese dell'evento, con un **pro-rata per il primo mese**:

$$\text{Coefficiente primo mese} = \frac{30 - \text{giorno dell'evento} + 1}{30}$$

**Esempio pratico:**  
Saving da negoziazione servizio manutenzione: 12.000 € annui = 1.000 €/mese.  
Evento in data 15 marzo → coefficiente primo mese = (30 - 15 + 1) / 30 = 0,533  
- Marzo: 1.000 × 0,533 = 533 €
- Da aprile a febbraio (23 mesi): 1.000 € / mese
- Ultimo mese: aggiustato per garantire che la somma totale sia esattamente 24.000 €

Il valore effettivo di ogni mese = valore teorico mensile × (% Realizzo / 100).

---

## Il tab Derisking – Registro Fornitori Potenziali

Il tab Derisking non registra eventi economici ma costruisce un **registro di fornitori potenziali** per la valutazione e qualifica.

### Aggiungere un nuovo fornitore potenziale

1. Fare clic sul tab **Derisking**.
2. Fare clic su **➕ Nuovo Evento**.
3. Si apre la finestra **"Nuovo Fornitore"**.

### Campi della scheda fornitore

| Sezione | Campo | Note |
|---------|-------|------|
| Informazioni Generali | **Fornitore** | Ragione sociale (obbligatoria) |
| | **Categoria** | Selezione da catalogo categorie esistenti |
| | **Nuova categoria** | Inserire se la categoria non esiste ancora (viene creata automaticamente) |
| | **Stato** | `Nuovo` / `In valutazione` / `Qualificato` / `Scartato` |
| Contatti | **Contatto** | Nome del referente commerciale |
| | **E-mail** | Cliccabile per aprire il client di posta |
| | **Telefono** | |
| | **Web** | URL del sito (cliccabile per aprire il browser) |
| Note | | Testo libero |

Fare clic su **💾 Salva** per salvare la scheda.

### Aggiornare lo stato di un fornitore

1. Fare doppio clic sulla riga del fornitore nel tab Derisking.
2. Modificare il campo **Stato**.
3. Fare clic su **💾 Salva**.

Lo stato avanza tipicamente da `Nuovo` → `In valutazione` → `Qualificato` (o `Scartato`).

---

## Gestire le categorie fornitori

Le categorie permettono di raggruppare i fornitori per famiglia merceologica.

### Accedere alla gestione categorie

Nella finestra fornitore, fare clic su **Gestisci Categorie**.

### Rinominare una categoria

1. Selezionare la categoria dall'elenco.
2. Inserire il nuovo nome nel campo **Nuovo nome**.
3. Fare clic su **Rinomina**.

La rinomina è **in sospeso** finché non si fa clic su **💾 Salva**.

### Unire due categorie

1. Selezionare la categoria da unire (sorgente) dall'elenco.
2. Nel campo **Unisci con**, scegliere la categoria di destinazione.
3. Fare clic su **Unisci**.

Tutti i fornitori della categoria sorgente vengono spostati nella categoria di destinazione. La categoria sorgente viene rimossa.

### Eliminare una categoria

Una categoria può essere eliminata solo se **non ha fornitori associati**. Il contatore **"Fornitori associati: N"** mostra quanti fornitori appartengono alla categoria selezionata. Se N > 0, il pulsante Elimina viene bloccato.

Tutte le modifiche rimangono in sospeso finché non si fa clic su **💾 Salva**. Fare clic su **❌ Annulla** per scartare tutte le modifiche.

---

## Modificare o eliminare un evento

1. Selezionare la riga dell'evento nel tab corrispondente.
2. Fare clic su **⚡ Actions**:
   - Scegliere **Modifica** per aprire la finestra di modifica.
   - Scegliere **Elimina** per eliminare l'evento.

> L'eliminazione rimuove anche tutti gli impatti mensili associati all'evento. L'operazione è irreversibile.

Quando si modifica un evento, DataFlow **ricalcola e ricrea automaticamente** tutti gli impatti mensili. Il calcolo precedente viene scartato.

---

## Visualizzare un evento di un altro utente

Gli eventi di altri utenti sono visibili nell'elenco ma si aprono in **modalità sola lettura**. Viene mostrata la finestra con tutti i dati, ma i campi sono disabilitati e l'unico pulsante disponibile è **✖ Chiudi**.
