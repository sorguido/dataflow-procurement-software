# 12 – Glossario

## Termini procurement

**Budget (Importo a Budget)**  
Costo pianificato o storicamente pagato per un articolo o servizio. Costituisce la baseline rispetto a cui si misura il Saving.

**Buyer**  
Acquirente professionale all'interno dell'ufficio acquisti. In DataFlow, ogni buyer ha un nome utente univoco e la titolarità sulle proprie RdO ed eventi.

**Category Manager**  
Responsabile di una famiglia merceologica (categoria). Può gestire un portafoglio di fornitori e negociare accordi quadro.

**Conto Lavoro**  
Modalità di fornitura in cui il materiale grezzo è fornito dall'azienda committente e il fornitore esegue solo la lavorazione. In DataFlow corrisponde al tipo di RdO "Conto lavoro".

**Cost Avoidance (Evitamento Costo)**  
Valore economico generato dal blocco o dalla riduzione di un aumento di costo richiesto dal fornitore. Differisce dal Saving perché non riduce un costo già pagato, ma impedisce un incremento futuro. Formula: `(Importo Richiesto Iniziale − Importo Negoziato) × Quantità Annua`.

**Derisking**  
Attività di riduzione del rischio nella supply chain attraverso la qualifica di fornitori alternativi. In DataFlow, il modulo Derisking gestisce il registro dei fornitori potenziali.

**Driver (di Saving o Cost Avoidance)**  
La leva negoziale che genera il valore economico:
- **Driver Prezzo**: riduzione del prezzo unitario o dell'importo complessivo.
- **Driver Pagamenti**: miglioramento dei termini di pagamento (dilazione maggiore = minore costo finanziario).

**Fornitura Piena**  
Modalità di fornitura in cui il fornitore fornisce il prodotto completo (materiale + lavorazione + eventuali componenti). In DataFlow corrisponde al tipo di RdO "Fornitura piena".

**KPI (Key Performance Indicator)**  
Indicatore chiave di prestazione. In procurement, i KPI principali sono Saving, Cost Avoidance e il numero di fornitori qualificati.

**OdA / PO (Ordine di Acquisto / Purchase Order)**  
Documento formale con cui l'azienda acquirente ordina beni o servizi a un fornitore, specificando quantità, prezzo e condizioni.

**OPEX (Operational Expenditure)**  
Spesa operativa ricorrente (contrapposta a CAPEX, spesa in conto capitale). In DataFlow, un evento OPEX Ripetitivo genera impatti economici distribuiti nel tempo.

**% Realizzo**  
Percentuale del valore teorico che si prevede venga effettivamente conseguita. Un saving da negoziazione non ancora formalizzata potrebbe avere % Realizzo = 50%, indicando incertezza sul pieno conseguimento. Il Saving Effettivo = Saving Teorico × (% Realizzo / 100).

**RdO (Richiesta di Offerta) / RFQ (Request for Quotation)**  
Documento con cui l'azienda richiede a uno o più fornitori un'offerta per la fornitura di beni o servizi. In DataFlow corrisponde a una scheda con griglia prezzi, allegati e gestione fornitori.

**Saving (Risparmio)**  
Riduzione di costo rispetto al budget o al costo storico. Formula base: `(Importo Budget − Importo Negoziato) × Quantità Annua`.

**Spending Annuo**  
Spesa complessiva annua con un fornitore, usata come base di calcolo per i saving da condizioni di pagamento.

**SQDC**  
Framework di valutazione multicriterio: Safety, Quality, Delivery, Cost. Utilizzato per confrontare i fornitori in modo strutturato.

---

## Termini applicativi DataFlow

**Carry-over**  
Valore economico (Saving o Cost Avoidance) generato da eventi dell'anno corrente (o precedenti) che si manifesta nell'anno successivo. Visibile come KPI aggiuntivo quando si usa il filtro per anno nel KPI Dashboard. Utile per le proiezioni di budget.

**Competenza economica**  
Il mese in cui un impatto economico si manifesta contabilmente. In DataFlow, i KPI di Saving e Cost Avoidance sono filtrati per competenza, non per data di inserimento dell'evento.

**Evento VSM**  
Registrazione di una negoziazione o di un'attività Derisking nel modulo Value Stream Mapping. Ogni evento genera uno o più impatti mensili (la distribuzione nel tempo del valore economico).

**Impatto mensile**  
Record derivato dall'evento VSM che rappresenta il valore economico (teorico ed effettivo) attribuito a un singolo mese. È la granularità minima usata dai KPI.

**Modalità Sola Lettura**  
Comportamento automatico di DataFlow quando si apre un elemento (RdO, evento, fornitore) di proprietà di un altro utente. Tutti i campi di modifica sono disabilitati; la visualizzazione è piena.

**Nome Utente (Username)**  
Identificativo univoco di ogni buyer, generato automaticamente al primo avvio nel formato `nome.cognome`. Non è modificabile dall'utente.

**Pro-rata primo mese**  
Coefficiente applicato al primo mese di distribuzione di un evento OPEX Ripetitivo. Tiene conto del fatto che l'evento è avvenuto in un giorno specifico del mese, non il primo giorno. Formula: `(30 − giorno + 1) / 30`.

**Value Stream Mapping (VSM)**  
Nel contesto di DataFlow, indica il modulo di tracciamento delle attività di valore del procurement (Saving, Cost Avoidance, Derisking). Il nome si ispira alla metodologia Lean del Value Stream Mapping, adattata al contesto degli acquisti.

**WAL (Write-Ahead Log)**  
Modalità operativa del database SQLite che permette a più utenti di leggere simultaneamente mentre uno solo scrive. Garantisce la coerenza dei dati in ambienti multiutente senza richiedere un server dedicato.
