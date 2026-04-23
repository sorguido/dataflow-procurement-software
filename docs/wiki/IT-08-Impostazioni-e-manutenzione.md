# 08 – Impostazioni e manutenzione

## Aprire le impostazioni

Fare clic su **⚙️ Impostazioni** nella barra degli strumenti. Si apre la finestra **"Impostazioni e Manutenzione"**.

---

## Posizione DataFlow

### Visualizzare la posizione corrente

Nella sezione **"Posizione DataFlow Standard"** è indicato il percorso della cartella attiva. Questa cartella contiene il database e gli allegati.

### Cambiare la posizione

1. Fare clic su **📁 Cambia Posizione DataFlow...**
2. Selezionare la nuova cartella nella finestra di dialogo.
3. Fare clic su **OK**.
4. Un messaggio avvisa che è necessario riavviare l'applicazione per rendere effettiva la modifica.

> **Attenzione:** cambiare posizione **non copia** automaticamente il database esistente. Prima di riavviare, copiare manualmente l'intera cartella `DataFlow/` nella nuova posizione.

---

## Backup manuale

Per creare una copia di sicurezza del database in qualsiasi momento:

1. Fare clic su **💾 Backup Manuale...**
2. Scegliere la cartella e il nome del file di backup nella finestra di salvataggio.
3. Il file `dataflow_db.db` viene copiato nella posizione scelta.

Si consiglia di eseguire un backup manuale prima di operazioni critiche (es. migrazione a un nuovo server, cambio di percorso del database).

---

## Backup automatico giornaliero

DataFlow può eseguire automaticamente un backup del database ogni giorno all'ora configurata.

### Configurare il backup automatico

1. Attivare il **checkbox "Backup Automatico Giornaliero"**.
2. Impostare l'**Ora** (da 00 a 23) nel menu a tendina.
3. Nel campo **"Salva in:"**, fare clic su **📁 Scegli...** per selezionare la cartella di destinazione dei backup.
4. Fare clic su **💾 Salva Impostazioni Backup**.

DataFlow mantiene al massimo **3 backup automatici**. Quando ne viene creato uno nuovo, il più vecchio viene eliminato automaticamente.

Le impostazioni vengono salvate nel file `config.ini` nella sezione `[AutoBackup]`.

---

## Lingua e valuta

1. Nel menu a tendina **"Lingua"**, selezionare:
   - **Italiano**
   - **English**
2. Nel menu a tendina **"Valuta"**, selezionare:
   - **Nessuna**
   - **EUR**
   - **USD**
   - **GBP**
   - **CHF**
3. Fare clic su **💾 Salva Impostazioni**.
4. Un messaggio avvisa che il cambio richiede il riavvio dell'applicazione.

Lingua e valuta vengono salvate nel file `config.ini`. Al riavvio, la lingua selezionata viene applicata all'interfaccia e la valuta scelta viene usata dove applicabile per la formattazione degli importi.

---

## File di configurazione

Le impostazioni dell'applicazione sono salvate nel file `config.ini`. Questo file si trova nella stessa cartella del database (`DataFlow/`). Le sezioni principali sono:

| Sezione | Contenuto |
|---------|-----------|
| `[Settings]` | Lingua, valuta, percorso database, coefficiente finanziario pagamenti |
| `[AutoBackup]` | Attivazione, ora, cartella destinazione backup |
| `[User]` | Nome, cognome, nome utente generato |

In caso di problemi con l'applicazione, un tecnico potrà leggere questi valori per diagnosticare la configurazione.

---

## Percorsi dei file di log

I log di diagnostica vengono scritti automaticamente in:

- **Windows:** `%LOCALAPPDATA%\DataFlow\dataflow.log`
- **Linux:** `~/.local/share/DataFlow/dataflow.log`

Vedere la sezione [Log e diagnostica](IT-13-Log-e-diagnostica) per maggiori dettagli.

---
[← Pagina precedente](IT-07-KPI-Dashboard) | [Pagina successiva →](IT-09-Lavoro-multiutente)
