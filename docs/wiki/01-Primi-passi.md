# 01 – Primi passi

## Requisiti di sistema

DataFlow funziona su **Windows 10/11** e **Linux** (distribuzione con desktop grafico). Non richiede connessione internet. Per l'uso condiviso su rete è sufficiente che tutti i computer abbiano accesso alla stessa cartella su un server o NAS.

---

## Primo avvio

Al primo avvio, DataFlow mostra una finestra di configurazione identità. È necessario inserire **Nome** e **Cognome** prima di poter utilizzare l'applicazione.

### Impostare la propria identità

1. Alla comparsa della finestra **"Identità Utente"**, inserire il proprio nome nel campo **Nome** e il cognome nel campo **Cognome**.
2. Il campo **Nome utente** si aggiorna automaticamente in tempo reale con il formato `nome.cognome` (tutto minuscolo, senza accenti).
3. Fare clic su **Conferma**.

> Il nome utente generato automaticamente è definitivo e non modificabile in seguito. Verrà usato per identificare tutte le RdO e gli eventi creati.

Non è possibile chiudere questa finestra senza inserire entrambi i campi. Il pulsante di chiusura (X) non ha effetto su questa schermata.

---

## Schermata splash

Dopo la configurazione dell'identità, compare brevemente una schermata di avvio mentre l'applicazione inizializza il database e le cartelle di lavoro.

---

## Struttura cartelle creata automaticamente

DataFlow crea automaticamente la seguente struttura nella posizione configurata (default: `Documenti/DataFlow/`):

```
DataFlow/
├── Database/
│   └── dataflow_db.db          ← database principale
└── Attachments/
    └── {numero RdO}/           ← allegati per ogni RdO
```

Non spostare manualmente questi file mentre l'applicazione è aperta.

---

## Cambiare la posizione del database

Se si vuole usare una cartella condivisa su rete (per lavoro multiutente):

1. Aprire le **Impostazioni** (pulsante ⚙️ nella barra degli strumenti).
2. Nella sezione **"Posizione DataFlow Standard"**, leggere il percorso attuale.
3. Fare clic su **📁 Cambia Posizione DataFlow...** e scegliere la nuova cartella.
4. Riavviare l'applicazione.

> Dopo aver cambiato la posizione, il database vecchio **non viene spostato automaticamente**. Copiare manualmente la cartella `DataFlow/` nella nuova posizione prima di riavviare, oppure l'applicazione creerà un nuovo database vuoto.

Vedere la sezione [Lavoro multiutente](09-Lavoro-multiutente.md) per le istruzioni complete.

---

## Prima RdO di prova

Per familiarizzare con l'applicazione, ecco come creare la prima Richiesta di Offerta in pochi passi:

1. Fare clic su **➕ Nuovo Evento** nella barra degli strumenti.
2. Selezionare **📦 Fornitura piena**.
3. Si apre il pannello di controllo della RdO. Inserire una data di scadenza e un riferimento (es. il nome del progetto).
4. Fare clic su **➕ Aggiungi Articolo** per inserire la prima riga.
5. Fare clic su **Fornitori** per indicare i fornitori da invitare.
6. Chiudere la finestra: la RdO viene salvata automaticamente ed è visibile nel tab **RdO Attive**.

---

## Avvio successivo al primo

Dai successivi avvii, DataFlow si apre direttamente sulla schermata principale senza richiedere di nuovo l'identità. L'applicazione ricorda la posizione del database e la lingua e valuta selezionate.

---
[Pagina successiva →](02-Schermata-principale)
