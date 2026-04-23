# 09 – Lavoro multiutente

## Come funziona la condivisione del database

DataFlow supporta il lavoro contemporaneo di più utenti sullo stesso database. Il meccanismo si basa su un file SQLite condiviso in una cartella di rete (server aziendale, NAS, cartella condivisa).

**Non è necessario un server applicativo.** Ogni utente ha la propria installazione di DataFlow sul proprio computer; tutti puntano alla stessa cartella di rete dove risiede il database.

---

## Configurare la cartella condivisa

### Passo 1 – Preparare la cartella sul server

Su un percorso di rete accessibile a tutti gli utenti (es. `\\server\DataFlow\`), creare la struttura:

```
DataFlow\
├── Database\
└── Attachments\
```

Se esiste già un database di un utente, copiare il file `dataflow_db.db` in `DataFlow\Database\`.

### Passo 2 – Configurare ogni workstation

Su ogni computer:

1. Aprire DataFlow.
2. Andare in **⚙️ Impostazioni**.
3. Fare clic su **📁 Cambia Posizione DataFlow...** e selezionare la cartella di rete (es. `\\server\DataFlow\`).
4. Riavviare DataFlow.

### Passo 3 – Verificare l'identità utente

Ogni utente deve avere un'identità unica (nome utente generato al primo avvio). DataFlow usa il nome utente per distinguere i dati di ciascun buyer. Se due utenti hanno lo stesso nome (es. due "Marco Rossi"), potrebbero avere nomi utente identici; in tal caso è necessario che uno dei due venga rinominato prima della condivisione (operazione che richiede assistenza tecnica).

---

## Come DataFlow gestisce il database condiviso

- Il database usa la modalità **WAL (Write-Ahead Log)** di SQLite, che consente **più lettori simultanei** e **un solo scrittore alla volta**.
- Le scritture hanno un timeout di **10 secondi**: se il database è occupato da un'altra scrittura, DataFlow attende fino a 10 secondi prima di segnalare un errore.
- Le operazioni di lettura (visualizzazione, ricerca) usano l'accesso **sola lettura** per non interferire mai con le scritture in corso.

In condizioni normali, con 5–10 utenti, non si verificano conflitti visibili.

---

## Visibilità dei dati

Ogni utente vede **tutti i dati di tutti gli utenti**, ma con le seguenti restrizioni:

| Dato | Proprietario | Altro utente |
|------|-------------|--------------|
| RdO | Apre e modifica liberamente | Apre in sola lettura |
| Evento VSM (Saving/CA) | Modifica e cancella | Solo lettura |
| Fornitore Derisking | Modifica e cancella | Solo lettura |

La **modalità sola lettura** viene segnalata da:
- Banner rosso in fondo alla finestra RdO: *"⚠️ MODALITÀ SOLA LETTURA: Stai visualizzando una RdO di un altro utente."*
- Tutti i campi e pulsanti di modifica disabilitati.
- Solo apertura e download degli allegati rimangono attivi.

---

## Filtro per utente

Nella barra dei filtri avanzati, il campo **Utente** permette di filtrare per vedere solo le RdO di uno specifico buyer o di tutti contemporaneamente.

- Selezionare il proprio nome per vedere solo le proprie RdO (ricerca solo nel database locale).
- Selezionare **"(Tutti gli Utenti)"** per vedere le RdO di tutti i buyer (ricerca aggregata su tutti i database nella stessa cartella `Database/`).
- Selezionare il nome di un collega per vedere solo le sue RdO.

---

## Database multipli nella stessa cartella

DataFlow supporta una variante avanzata della condivisione: ogni utente può avere il **proprio file di database separato** dentro la stessa cartella `Database/`. In questo caso, la funzione di **ricerca aggregata** legge tutti i file `*.db` presenti nella cartella e li presenta all'utente come se fossero uno solo (in sola lettura per i dati degli altri).

Questa modalità si attiva automaticamente quando ogni utente punta alla stessa cartella di rete ma ha creato il proprio database localmente e poi lo ha copiato lì.

---

## Backup e database condiviso

Quando il database è condiviso su rete, il backup automatico giornaliero deve essere configurato **su una sola workstation** (es. il computer del responsabile acquisti o un server). Configurarlo su più workstation crea backup ridondanti, il che è accettabile ma non necessario.

Prima di eseguire un backup manuale di un database condiviso, verificare che nessun altro utente stia compiendo operazioni di scrittura.

---

## Comportamento in caso di rete non disponibile

Se la cartella di rete non è raggiungibile all'avvio, DataFlow non trova il database e mostra un errore all'apertura. In questo caso:

1. Verificare la connessione di rete.
2. Assicurarsi di avere i permessi di lettura/scrittura sulla cartella condivisa.
3. Se necessario, lavorare temporaneamente con una copia locale del database e riallinearla manualmente in seguito.

DataFlow non gestisce in automatico i conflitti di merge tra due database che hanno evoluto indipendentemente.

---
[← Pagina precedente](IT-08-Impostazioni-e-manutenzione) | [Pagina successiva →](IT-10-Problemi-comuni-e-soluzioni)
