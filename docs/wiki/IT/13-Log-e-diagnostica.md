# 13 – Log e diagnostica

## A cosa servono i log

DataFlow registra automaticamente le operazioni significative e gli errori in un file di log. Il log è utile per:

- Diagnosticare errori che si verificano all'avvio o durante l'uso.
- Supporto tecnico: il file di log è la prima informazione richiesta in caso di segnalazione di un problema.
- Verificare se un'operazione (es. backup automatico) è avvenuta correttamente.

---

## Posizione del file di log

| Sistema operativo | Percorso |
|-------------------|----------|
| **Windows** | `C:\Users\[nome_utente]\AppData\Local\DataFlow\dataflow.log` |
| **Linux** | `~/.local/share/DataFlow/dataflow.log` |

Su Windows, la cartella `AppData\Local` è nascosta per impostazione predefinita. Per accedervi, digitare `%LOCALAPPDATA%\DataFlow\` nella barra degli indirizzi di Esplora file.

---

## Rotazione automatica del log

Il file di log ha una dimensione massima di **5 MB**. Quando raggiunge il limite, viene rinominato in `dataflow.log.1` e ne viene creato uno nuovo. DataFlow mantiene al massimo **3 file di backup**:

```
dataflow.log       ← il più recente
dataflow.log.1     ← backup 1
dataflow.log.2     ← backup 2
dataflow.log.3     ← backup più vecchio
```

I file più vecchi vengono eliminati automaticamente.

---

## Come leggere il file di log

Il file di log è un file di testo (apribile con Blocco Note o qualsiasi editor di testo). Ogni riga ha il formato:

```
2026-04-04 09:15:32,123 - INFO - [modulo] - Messaggio descrittivo
2026-04-04 09:15:33,456 - ERROR - [modulo] - Messaggio di errore
```

### Livelli di gravità

| Livello | Significato |
|---------|-------------|
| `INFO` | Operazione normale registrata a titolo informativo |
| `WARNING` | Situazione insolita che non ha impedito l'operazione |
| `ERROR` | Errore che ha impedito un'operazione; richiede attenzione |
| `CRITICAL` | Errore grave che ha causato la chiusura dell'applicazione |

---

## Operazioni registrate nel log

Il log include (non esaustivo):

- Avvio dell'applicazione e versione
- Apertura del database (percorso, esito)
- Creazione, modifica, eliminazione di RdO
- Creazione, modifica, eliminazione di eventi Value Stream Mapping
- Operazioni di backup (manuale e automatico) con esito
- Errori di validazione dei dati
- Errori di scrittura su database
- Apertura e salvataggio di allegati
- Cambio di lingua
- Cambio di percorso del database

---

## Pulizia dei file temporanei

All'avvio, DataFlow elimina automaticamente:

- I file temporanei `_MEI*` lasciati da sessioni PyInstaller precedenti non chiuse correttamente.
- I file con prefisso `tmp*` più vecchi di 24 ore nella cartella temporanea di sistema.

Questo processo avviene silenziosamente in background senza impattare l'avvio.

---

## Segnalare un problema con i log

Quando si segnala un problema al supporto tecnico:

1. Aprire la cartella dei log (vedi percorso sopra).
2. Aprire `dataflow.log` con un editor di testo.
3. Cercare le righe con `ERROR` o `CRITICAL` nelle ore in cui si è verificato il problema.
4. Allegare il file `dataflow.log` alla segnalazione (non copiare solo poche righe: l'intero file è utile per il contesto).

Vedere la sezione [Supporto](14-Supporto.md) per i recapiti di segnalazione.
