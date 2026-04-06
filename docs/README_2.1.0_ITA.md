# DataFlow Procurement Software

📊 DataFlow è l'applicazione desktop indispensabile per mettere fine alla dispersione delle informazioni tipica degli uffici acquisti.

👨‍💼 Sviluppato da un buyer per i buyer, DataFlow offre una piattaforma completa per gestire ogni fase delle Richieste di Offerta (RdO) e analizzare le quotazioni dei fornitori, nonché per misurare le performance dell'ufficio acquisti attraverso un modulo KPI dedicato.

✨ Novità della versione 2.1.0: modulo Value Stream Mapping (VSM) per il tracciamento di Saving, Cost Avoidance e attività di Derisking. Nuova finestra KPI Analysis con grafici ed export Excel. Anagrafica Fornitori Potenziali integrata nel workflow di Derisking. Barra di ricerca globale e architettura dashboard modulare.

🎯 Il software DataFlow si posiziona come strumento di nicchia e di supporto alle decisioni, colmando il divario tra la gestione base (Excel) e i costosi moduli ERP.

📧 Se il tuo flusso di lavoro è frammentato tra infinite email, allegati dispersi e fogli Excel non collegati, DataFlow porta ordine.

💪 Il vero punto di forza di DataFlow risiede nella sua capacità di standardizzare i processi.
Ogni nuova richiesta, che si tratti di fornitura standard o conto lavoro, segue un percorso definito.

🔄 Flusso di lavoro tipico consigliato:

1️⃣ Inserisci una Richiesta di Offerta (RdO) con eventuali documenti interni (email, disegni, ecc.).  
2️⃣ Esporta la RdO in Excel.  
3️⃣ Copia e incolla le colonne del foglio contenenti solo le informazioni necessarie ai fornitori nell'email e invia la tua RdO.  
4️⃣ Ricevi i preventivi e inseriscili nel Pannello di Controllo RdO.  
5️⃣ Allega le quotazioni ricevute al file.  
6️⃣ Se necessario, esegui l'analisi SQDC (Safety, Quality, Delivery, Cost) per determinare il fornitore vincitore.  
7️⃣ Salva l'analisi SQDC nella RdO come riferimento futuro.  
8️⃣ Registra i risultati della negoziazione come eventi VSM (Saving / Cost Avoidance / Derisking) per costruire la tua storia di KPI acquisti.

📌 Punto di riferimento unico: ogni RdO è un record completo che include la storia della negoziazione, gli articoli richiesti, i riferimenti tecnici (disegni, codici) e le scadenze.

⚖️ Confronto obiettivo: non dovrai più estrarre manualmente i dati dai PDF delle quotazioni.
Inserisci i prezzi voce per voce per ottenere un confronto immediato e trasparente tra offerte di fornitori diversi.

📁 Archivio integrato: documenti di quotazione, comunicazioni interne, disegni e specifiche tecniche non sono più allegati "persi" in una casella di posta, ma archiviati in modo logico e consultabili direttamente dal tab della Richiesta di Offerta.

👥 Collaborazione in team: ogni utente ha la propria area dati personale, con la possibilità di visualizzare le RdO dei colleghi in modalità lettura. Ideale per uffici acquisti con più buyer e per i responsabili che necessitano di visibilità sulle negoziazioni in corso.

📈 KPI Acquisti: traccia l'impatto misurabile della tua attività negoziale nel tempo. La finestra KPI Analysis fornisce metriche aggregate sull'attività RdO, Saving, Cost Avoidance e Derisking, con grafici interattivi e export Excel.

🎯 DataFlow ti rimette in controllo. Garantisce che le decisioni di acquisto siano sempre basate su dati completi, trasparenti e facilmente accessibili, permettendoti di concentrarti sulla negoziazione strategica piuttosto che sulla ricerca delle informazioni.

🔒 Trasforma la gestione delle RdO da un'attività frammentata in un processo snello, centralizzato e sicuro, con i dati archiviati localmente e sotto il tuo pieno controllo.

---

Originariamente sviluppato per Windows e pubblicato sul [Microsoft Store](https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=en-EN&gl=EN), il progetto è ora rilasciato come edizione open-source per Linux con licenza **GNU GPLv3**.

L'applicazione è scritta in **Python** con interfaccia grafica **Tkinter** e utilizza **SQLite** come motore di database locale.  
Il port Linux include gestione dei percorsi cross-platform, supporto icone finestra Linux, supporto multilingua (italiano e inglese), import/export Excel, gestione allegati, tracciamento ordini di acquisto, note, analisi SQDC, tracciamento eventi VSM, analisi KPI e anagrafica fornitori potenziali.  
Il codebase include un database manager dedicato con supporto SQLite/WAL e logica per l'aggregazione dei dati su database di più utenti.

---

## Highlights

- Applicazione desktop per la gestione degli acquisti e delle RdO
- Interfaccia grafica Python + Tkinter
- Backend database SQLite (modalità WAL)
- Import/export Excel
- Gestione allegati
- Tracciamento ordini di acquisto (PO)
- Gestione note
- Workflow export/salvataggio analisi SQDC
- **Modulo VSM**: tracciamento eventi Saving, Cost Avoidance e Derisking
- **Finestra KPI Analysis**: metriche aggregate con grafici e export Excel
- **Anagrafica Fornitori Potenziali**: gestione e qualifica di nuovi fornitori (workflow Derisking)
- **Barra di ricerca globale**: ricerca multi-campo su tutti i tab della dashboard principale
- Supporto lingue italiano e inglese
- Port Linux con correzioni per comportamenti specifici della piattaforma
- Distribuzione Windows esistente sul Microsoft Store

---

## Aree Funzionali

### Gestione RdO
Crea, gestisci e archivia le Richieste di Offerta. Registra i preventivi dei fornitori voce per voce, allega documenti, traccia gli ordini di acquisto ed esegui analisi SQDC a supporto della selezione del fornitore.

### Value Stream Mapping (VSM) — Nuovo in 2.1.0
Traccia i risultati delle negoziazioni come eventi strutturati direttamente dalla dashboard principale:

- **Saving**: riduzione di prezzo ottenuta tramite negoziazione, con driver opzionale per i termini di pagamento
- **Cost Avoidance**: aumento di costo evitato, con percentuale di realizzo configurabile
- **Derisking**: attività di riduzione del rischio nella supply chain, con tracciamento introduzione nuovo fornitore

Ogni evento genera proiezioni di impatto economico mensile. Gli eventi OPEX-ripetitivi propagano il loro effetto fino a 24 mesi.

### KPI Analysis — Nuovo in 2.1.0
Una finestra dedicata fornisce KPI acquisti aggregati su quattro dimensioni:

- **KPI RdO**: volume, breakdown attive/archiviate, copertura fornitori e codici prodotto
- **KPI Saving**: importi teorici vs. effettivi con grafici di tendenza
- **KPI Cost Avoidance**: costi evitati nel tempo
- **KPI Derisking**: nuovi fornitori introdotti, distribuzione degli stati di qualifica

Filtri per anno o intervallo di date personalizzato. Export Excel disponibile.

### Anagrafica Fornitori Potenziali — Nuovo in 2.1.0
Gestisci il ciclo di vita dei potenziali nuovi fornitori direttamente nel tab Derisking:

- Registra ragione sociale, categoria merceologica, contatti e stato di qualifica
- Stati: Nuovo, In valutazione, Qualificato, Scartato
- Integrato con il workflow VSM Derisking

---

## Screenshot

![Finestra Principale](docs/screenshot/EN/1.png)
![Finestra Principale](docs/screenshot/EN/2.png)
![Finestra Principale](docs/screenshot/EN/3.png)
![Finestra Principale](docs/screenshot/EN/4.png)
![Finestra Principale](docs/screenshot/EN/5.png)
![Finestra Principale](docs/screenshot/EN/6.png)
![Finestra Principale](docs/screenshot/EN/7.png)

---

## Stato del Progetto

La versione 2.1.0 segna la transizione da uno strumento di pura gestione RdO a una piattaforma più ampia per le performance degli acquisti, aggiungendo il tracciamento strutturato dei risultati negoziali e le capacità di misurazione dei KPI.

---

## Stack Tecnologico

- **Linguaggio:** Python
- **GUI:** Tkinter
- **Database:** SQLite (modalità WAL)
- **File principale:** `dataflow.py`

### Dipendenze principali

- `openpyxl`
- `Pillow`
- `polib`
- `tkcalendar`
- `tksheet`

---

## Installazione

### Installazione su Linux

# [📥 DOWNLOAD 📥](https://github.com/sorguido/dataflow-procurement-software/releases)

➡ Scarica il pacchetto AppImage e fai doppio clic (per tutte le distribuzioni Linux)

oppure

➡ Scarica il pacchetto Linux (.deb)

Installa con doppio clic o da terminale:

```bash
sudo apt install ./dataflow_2.1.0_amd64.deb
```

## Installazione da sorgente

### 1. Clona il repository

```bash
git clone https://github.com/sorguido/dataflow-procurement-software.git
cd dataflow-procurement-software
```

### 2. Crea un ambiente virtuale

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Installa le dipendenze

```bash
pip install -r requirements.txt
```

### 4. Avvia l'applicazione

```bash
python3 dataflow.py
```

---

## Requisiti

```txt
openpyxl
Pillow
polib
tkcalendar
tksheet
```

---

## Licenza

Questo progetto è rilasciato con licenza **GNU General Public License v3.0**.

Il testo completo della licenza è incluso nel repository nel file `LICENSE`.

---

## Nota sulla versione Windows

Esiste anche una versione Windows di DataFlow, pubblicata sul Microsoft Store:  
https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=en-EN&gl=EN

L'edizione è open-source GNU GPLv3. Se future versioni Windows/Linux saranno allineate allo stesso modello di licenza, potranno essere distribuite tramite questo repository o un flusso di packaging correlato.

---

## Contribuire

DataFlow è disponibile come progetto open-source.

L'applicazione è stata rilasciata con licenza GNU GPLv3 e il codice sorgente è disponibile su GitHub.

Gli sviluppatori interessati a migliorare o adattare il software, incluse le future versioni Windows, sono benvenuti a contribuire.

---