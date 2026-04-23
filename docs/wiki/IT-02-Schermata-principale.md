# 02 – Schermata principale

## Layout generale

La schermata principale è organizzata verticalmente in cinque aree:

| Area | Contenuto |
|------|-----------|
| Barra degli strumenti | Logo + pulsanti operativi |
| Barra di ricerca globale | Campo di ricerca unico + pulsante filtri avanzati |
| Filtri avanzati | Pannello a scomparsa con campi di ricerca dedicati |
| Notebook | 5 tab: RdO Attive, RdO Archiviate, Saving, Cost Avoidance, Derisking |
| Piè di pagina | Versione, autore, licenza |

---

## Barra degli strumenti

| Pulsante | Azione |
|----------|--------|
| **➕ Nuovo Evento** | Crea una nuova RdO (sui tab RdO) oppure un nuovo evento Value Stream Mapping (sui tab Saving / Cost Avoidance / Derisking) |
| **⚡ Actions** | Menu contestuale con azioni sulla riga selezionata (attivo solo se una riga è selezionata) |
| **📥 Export Excel** | Esporta l'elenco RdO in un file Excel |
| **≋ KPI** | Apre la finestra di analisi KPI |
| **⚙️ Impostazioni** | Apre le impostazioni dell'applicazione |
| **≡ License** | Visualizza la licenza software |
| **❓ Guida** | Apre la guida utente integrata |

### Il menu ⚡ Actions

Il menu Actions è **sempre disabilitato** finché non si seleziona una riga. Diventa attivo al clic su un elemento:

- **Sui tab RdO**: permette di eliminare, duplicare, archiviare o riattivare la RdO selezionata.
- **Sui tab VSM (Saving / Cost Avoidance / Derisking)**: permette di modificare o eliminare l'evento selezionato.

---

## Barra di ricerca globale

Il campo di ricerca ampio al centro della barra è la modalità più rapida per trovare qualsiasi elemento. Nei tab RdO cerca contemporaneamente su:

- Numero RdO
- Riferimento progetto
- Nome fornitore
- Codice materiale
- Disegno / allegato
- Descrizione materiale
- Numero ordine di acquisto
- Cod. Grezzo
- Allegato Grezzo
- Mat. C/L

**Come usarla:**
1. Digitare una parola chiave (es. il nome di un fornitore, un codice pezzo, un riferimento progetto).
2. Premere **Invio**.
3. I risultati vengono mostrati nel tab attivo. Tutti gli altri tab si aggiornano in parallelo.
4. Per cancellare la ricerca, svuotare il campo e premere **Invio**.

Nei tab Saving, Cost Avoidance e Derisking, la stessa barra cerca invece nei principali campi testuali visibili dell'elenco attivo.

La ricerca è **case-insensitive** e usa la logica **OR**: un risultato viene mostrato se contiene il testo cercato almeno in uno dei campi interrogati.

---

## Filtri avanzati

Per ricerche più precise, fare clic sull'etichetta **⌄ Filtri Avanzati** (a destra della barra di ricerca). Il pannello si espande mostrando campi distinti per ciascun criterio.

> Il pulsante Filtri Avanzati è disabilitato quando il tab **Derisking** è attivo, perché la ricerca fornitori avviene direttamente nell'elenco.

### Filtri per le RdO

| Campo | Descrizione |
|-------|-------------|
| Numero RdO | Cerca per numero identificativo |
| Tipo RdO | Combobox: Tutte / Fornitura piena / Conto lavoro |
| Riferimento | Testo libero sul riferimento progetto |
| Fornitore | Nome fornitore (anche parziale) |
| Cod. Materiale | Codice articolo |
| Desc. Materiale | Descrizione articolo |
| Num. Ordine | Numero ordine di acquisto |
| Cod. Grezzo | Solo RdO di tipo Conto lavoro |
| Allegato Grezzo | Solo RdO di tipo Conto lavoro |
| Mat. c/lavoro | Solo RdO di tipo Conto lavoro |
| Utente | Filtra per buyer (combobox con tutti gli utenti) |
| Data Emissione Da / A | Intervallo date di emissione |
| Data Scadenza Da / A | Intervallo date di scadenza |

### Filtri per gli eventi VSM (tab Saving / Cost Avoidance)

| Campo | Descrizione |
|-------|-------------|
| Utente | Buyer proprietario dell'evento |
| Da / A | Intervallo date evento |
| Azione | Negoziazione / Derisking / Altro |
| Ripetitivo | Sì / No (filtra eventi OPEX ricorrenti) |
| Valore Teorico Da / A | Intervallo importo teorico |
| Valore Effettivo Da / A | Intervallo importo effettivo |

Dopo aver impostato i filtri, fare clic su **🔍 Cerca**. Per riportare la vista completa, fare clic su **🔎 Pulisci Filtri**.

I filtri usano logica **AND**: ogni campo attivo aggiunge un vincolo al risultato.

---

## I cinque tab principali

### Tab RdO Attive e RdO Archiviate

Mostrano l'elenco delle richieste di offerta in una tabella con le colonne:

- **N°** – Numero progressivo assegnato automaticamente
- **Tipo** – Fornitura piena / Conto lavoro
- **Data Emissione**
- **Data Scadenza** – Le RdO scadute appaiono evidenziate in rosso
- **Riferimento** – Progetto o descrizione breve
- **Utente** – Buyer proprietario della RdO

Fare **doppio clic** su una riga per aprire il pannello di controllo della RdO. Le RdO di altri utenti si aprono in **modalità sola lettura** (banner rosso in basso: *"Stai visualizzando una RdO di un altro utente"*).

Le colonne sono **ordinabili**: fare clic sull'intestazione per ordinare in modo crescente o decrescente.

### Tab Saving e Cost Avoidance

Mostrano l'elenco degli eventi Value Stream Mapping di tipo economico, con le colonne:

- Data, Tipo, Azione, Descrizione, Riferimento, Driver, Valore Teorico, Valore Effettivo, % Realizzo, Ripetitivo, Utente

Fare **doppio clic** per aprire il dettaglio dell'evento. Vedere la sezione [Value Stream Mapping](IT-06-Value-Stream-Mapping) per le istruzioni complete.

### Tab Derisking

Mostra il registro dei fornitori potenziali, con le colonne:

- Fornitore, Categoria, Stato, Contatto, E-mail, Telefono, Utente

Fare **doppio clic** per aprire la scheda fornitore.

---

## Ordinamento delle colonne

Fare clic sull'intestazione di una colonna per ordinare la lista in modo crescente. Fare clic di nuovo per invertire l'ordine. L'ordinamento è visivo e non modifica i dati.

---

## Scorciatoie da tastiera

| Azione | Scorciatoia |
|--------|-------------|
| Avviare una ricerca | Digitare nel campo di ricerca + **Invio** |
| Aprire un elemento | **Doppio clic** sulla riga |
| Aprire la guida | Pulsante **❓ Guida** |

---
[← Pagina precedente](IT-01-Primi-passi) | [Pagina successiva →](IT-03-Creare-una-nuova-RdO)
