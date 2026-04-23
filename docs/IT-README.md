# DataFlow Procurement Software

📊 DataFlow è l'applicazione desktop essenziale progettata per porre fine alla dispersione di informazioni tipica degli uffici acquisti.

👨‍💼 Sviluppato da un buyer per i buyer, DataFlow offre una piattaforma completa per gestire ogni fase delle Requests for Quotation (RFQ), analizzare le offerte dei fornitori e misurare le performance procurement attraverso un tracciamento KPI dedicato.

✨ Cosa c'è di nuovo nella versione 2.2.0: architettura a servizi più modulare, traduzione runtime centralizzata con `tr(...)`, ricerca dashboard e filtri contestuali migliorati, export RFQ in PDF con supporto logo/template e affinamenti nei flussi VSM, KPI, Derisking e impostazioni.

🎯 Il software 'DataFlow' si posiziona come uno strumento di nicchia e di supporto decisionale che colma il divario tra la gestione di base (Excel) e i costosi moduli ERP.

📧 Se il tuo flusso di lavoro è attualmente frammentato tra innumerevoli email, allegati sparsi e fogli Excel scollegati, DataFlow porta ordine.

💪 La vera forza di DataFlow risiede nella sua capacità di standardizzare i processi.
Ogni nuova richiesta, sia per una fornitura standard sia per una lavorazione conto terzi, segue un percorso definito.

🔄 Flusso di lavoro tipico suggerito:

1️⃣ Inserisci una Request for Quotation (RFQ) e gli eventuali documenti interni (email, disegni, ecc.).  
2️⃣ Esporta la RFQ in Excel.  
3️⃣ Copia e incolla in un'email le colonne del foglio di calcolo contenenti solo le informazioni necessarie ai fornitori e invia la tua RFQ.  
4️⃣ Ricevi le quotazioni e inseriscile nel RFQ Control Panel.  
5️⃣ Allega le quotazioni ricevute al fascicolo.  
6️⃣ Se necessario, esegui l'analisi SQDC (Safety, Quality, Delivery, Cost) per determinare il fornitore vincitore.  
7️⃣ Salva l'analisi SQDC nella RFQ per riferimento futuro.  
8️⃣ Registra i risultati della negoziazione come eventi VSM (Saving / Cost Avoidance / Derisking) per costruire la tua storia KPI procurement.

📌 Punto di riferimento unico: ogni RFQ è una registrazione completa che include la cronologia della negoziazione, gli articoli richiesti, i riferimenti tecnici (disegni, codici) e le scadenze.

⚖️ Confronto oggettivo: non dovrai più estrarre manualmente i dati dai PDF delle quotazioni.
Inserisci i prezzi articolo per articolo per ottenere un confronto immediato e trasparente tra le offerte dei diversi fornitori.

📁 Archivio integrato: documenti di quotazione, comunicazioni interne, disegni e specifiche tecniche non sono più allegati 'persi' in una casella di posta, ma sono archiviati in modo logico e possono essere consultati direttamente dalla scheda Request for Quotation.

👥 Collaborazione di team: ogni utente ha la propria area dati personale, con la possibilità di visualizzare le RFQ dei colleghi in modalità sola lettura. Ideale per uffici acquisti con più buyer e per i responsabili che hanno bisogno di visibilità sulle negoziazioni in corso.

📈 KPI Procurement: monitora nel tempo l'impatto misurabile della tua attività negoziale. La finestra KPI Analysis fornisce metriche aggregate sull'attività RFQ, Savings, Cost Avoidances e Derisking, con grafici interattivi ed export Excel.

🎯 DataFlow ti rimette al comando. Garantisce che le decisioni di acquisto siano sempre basate su dati completi, trasparenti e facilmente accessibili, permettendoti di concentrarti sulla negoziazione strategica invece che sulla ricerca delle informazioni.

🔒 Trasforma la gestione delle RFQ da attività frammentata a processo snello, centralizzato e sicuro, con dati archiviati localmente e sotto il tuo pieno controllo.

---

Originariamente sviluppato per Windows e pubblicato anche sul [Microsoft Store](https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=en-EN&gl=EN), il progetto è ora rilasciato come edizione Linux open-source sotto licenza **GNU GPLv3**.

L'applicazione è scritta in **Python** con una GUI **Tkinter** e utilizza **SQLite** come motore di database locale.  
L'attuale porting Linux include gestione cross-platform dei path, supporto alle icone finestra su Linux, supporto multilingua (italiano e inglese), export Excel/PDF, gestione allegati, tracciamento degli ordini di acquisto, note, supporto all'analisi SQDC, tracciamento eventi VSM, analisi KPI e un registro dei fornitori potenziali.  
Il codebase include un database manager dedicato con supporto SQLite/WAL e logica per aggregare dati tra più database utente.

---

## In evidenza

- Applicazione desktop per procurement e gestione RFQ
- Interfaccia grafica Python + Tkinter
- Backend database SQLite
- Import/export Excel
- Gestione allegati
- Tracciamento Purchase order (PO)
- Gestione note
- Workflow di export/salvataggio analisi SQDC
- **Modulo VSM**: traccia eventi Saving, Cost Avoidance e Derisking
- **Finestra KPI Analysis**: metriche aggregate con grafici ed export Excel
- **Registro dei fornitori potenziali**: gestisci e qualifica nuovi fornitori (workflow Derisking)
- **Export RFQ in PDF**: esporta Requests for Quotation con logo persistente e template lingua modificabile
- **Barra di ricerca globale**: ricerca multi-campo con filtri contestuali in tutte le schede principali della dashboard
- Supporto lingua inglese e italiana
- Porting compatibile Linux con fix per comportamenti platform-specific
- Distribuzione Windows esistente su Microsoft Store

---

## Aree funzionali

### Gestione RFQ
Crea, gestisci e archivia le Requests for Quotation. Registra le offerte dei fornitori articolo per articolo, allega documenti, traccia gli ordini di acquisto ed esegui analisi SQDC per supportare le decisioni di selezione del fornitore.

### Value Stream Mapping (VSM) — Versione 2.2.0
Traccia i risultati delle negoziazioni come eventi strutturati direttamente dalla dashboard principale:

- **Saving**: riduzione di prezzo ottenuta tramite negoziazione, con driver opzionale relativo ai termini di pagamento
- **Cost Avoidance**: aumento di costo evitato, con percentuale di realizzo configurabile
- **Derisking**: attività di riduzione del rischio di supply chain, con tracciamento dell'introduzione di nuovi fornitori

Ogni evento genera proiezioni mensili di impatto economico. Gli eventi OPEX-ripetitivi propagano il loro effetto fino a 24 mesi.

### Analisi KPI — Versione 2.2.0
Una finestra dedicata fornisce KPI procurement aggregati su quattro dimensioni:

- **RFQ KPIs**: volume, ripartizione attive/archiviate, copertura fornitori e codici prodotto
- **Saving KPIs**: importi di saving teorici vs. effettivi con grafici di trend
- **Cost Avoidance KPIs**: costi evitati nel tempo
- **Derisking KPIs**: nuovi fornitori introdotti, distribuzione degli stati di qualificazione

Filtri per preset di periodo, anno o intervallo date personalizzato. Export in Excel disponibile.

### Registro dei fornitori potenziali — Versione 2.2.0
Gestisci il ciclo di vita dei potenziali nuovi fornitori direttamente nella scheda Derisking:

- Registra nome fornitore, categoria, dettagli di contatto e stato di qualificazione
- Stati: Nuovo, In valutazione, Qualificato, Scartato
- Suggerimenti nome fornitore e gestione di categorie riutilizzabili
- Integrato con il workflow VSM Derisking

---

## Screenshot

![Main Window](docs/screenshot/EN/1.png)

![Main Window](docs/screenshot/EN/2.png)

![Main Window](docs/screenshot/EN/3.png)

![Main Window](docs/screenshot/EN/4.png)

![Main Window](docs/screenshot/EN/5.png)

![Main Window](docs/screenshot/EN/6.png)

![Main Window](docs/screenshot/EN/7.png)

---

## Stato del progetto

La versione 2.2.0 consolida questa transizione con un'architettura più modulare, una copertura più ampia della traduzione runtime, una migliore gestione di ricerca/filtri della dashboard, export RFQ in PDF e affinamenti nei workflow di manutenzione ed export.

---

## Stack tecnologico

- **Language:** Python
- **GUI:** Tkinter
- **Database:** SQLite (WAL mode)
- **Main file:** `dataflow.py`

### Dipendenze principali

- `openpyxl`
- `Pillow`
- `polib`
- `reportlab`
- `tkcalendar`
- `tksheet`
- `tkinterdnd2`

---

## Installazione

# [📥 DOWNLOAD 📥](https://github.com/sorguido/dataflow-procurement-software/releases)

### Installazione Linux

➡ Scarica il pacchetto AppImage e fai doppio clic su di esso (per tutte le distribuzioni Linux)

Dopo aver scaricato il pacchetto, assegna i permessi di esecuzione:

```bash
chmod +x DataFlow_2.2.0_Linux_x86_64.AppImage
```

### Installazione Windows

➡ [Scarica](https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=en-EN&gl=EN) il pacchetto Msix dal Microsoft Store

### Installazione da sorgente

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

### 4. Esegui l'applicazione

```bash
python3 dataflow.py
```

---

## Requisiti

```txt
openpyxl
Pillow
polib
reportlab
tkcalendar
tksheet
tkinterdnd2
```

---

## Licenza

Questo progetto è rilasciato sotto la **GNU General Public License v3.0**.

Il testo completo della licenza è incluso nel repository nel file `LICENSE`.

---

## Nota sulla versione Windows

Esiste anche una versione Windows di DataFlow ed è stata pubblicata sul Microsoft Store:  
https://apps.microsoft.com/detail/9nt3bbg1w0k7?hl=en-EN&gl=EN

L'edizione è una release open-source GNU GPLv3. Se future release Windows/Linux saranno allineate allo stesso modello di licenza, potrebbero anch'esse essere distribuite tramite questo repository o un workflow di packaging correlato.

---

## Contribuire

DataFlow è un progetto open-source sviluppato con una forte attenzione a stabilità, chiarezza e usabilità nel mondo reale dei workflow procurement.

I contributi non sono solo benvenuti — sono essenziali per l'evoluzione sostenibile del progetto.

Tuttavia, questo progetto segue un insieme di principi guida (“dogma”) per preservarne affidabilità e usabilità:

- Preferire **modifiche piccole, chirurgiche e reversibili**
- Evitare **refactor globali** se non strettamente necessari
- Non introdurre **nuove dipendenze** senza una forte giustificazione
- Dare priorità alla **stabilità rispetto all'innovazione**
- Mantenere **coerenza UI e comportamento prevedibile**
- Ogni modifica deve puntare a **evitare regressioni**

### Da dove iniziare

Se vuoi contribuire, i punti di ingresso migliori sono:

- Bug fix (specialmente UI, traduzioni ed export Excel)
- Piccoli miglioramenti di usabilità
- Fix minori di performance o stabilità

### Aree che richiedono cautela

Le seguenti aree sono considerate ad alto impatto e non dovrebbero essere modificate senza discussione preventiva:

- `dataflow.py` logica core
- Schema database e layer di persistenza
- Comportamento della dashboard principale e flusso di navigazione
- Packaging e build system

### Approccio

Questo progetto è mantenuto con una **mentalità quality-first**.  
I contributi devono allinearsi ai principi sopra indicati.

In caso di dubbio, apri un issue prima di iniziare il lavoro.
