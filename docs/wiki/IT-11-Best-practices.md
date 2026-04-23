# 11 – Best practices

## Gestione delle RdO

### Usa il campo Riferimento in modo consistente

Il campo **Riferimento** è il principale campo di ricerca libera. Inserire sempre un valore che permetta di ritrovare la RdO in futuro:
- Il nome del progetto o della commessa
- Il codice cliente o il numero di contratto
- Una descrizione breve ma univoca

Evitare riferimenti generici come "Vari", "Test" o "Offerta aprile".

### Archivia le RdO chiuse

Quando una RdO è conclusa (ordine emesso, gara terminata, o non più rilevante), **archiviarla** invece di lasciarla nel tab RdO Attive. Questo riduce il rumore nella vista principale e migliora le prestazioni della ricerca.

### Inserisci sempre i numeri ordine

Appena viene emesso un ordine di acquisto, inserire il numero OdA nella RdO tramite il pulsante **📋 Inserisci OdA**. Questo permette di cercare per numero ordine e di tracciare la pipeline da RdO a OdA.

### Allega le offerte ricevute

Salvare le offerte dei fornitori come allegati nella RdO. In questo modo l'intera documentazione della gara è centralizzata e accessibile a tutto il team, anche a distanza di mesi.

---

## Registrazione nel Value Stream Mapping

### Registra tutti i saving, anche i piccoli

Il modulo Value Stream Mapping è tanto più utile quanto più è completo. Registrare anche i saving di importo modesto, perché contribuiscono alle statistiche aggregate e al report KPI verso il management.

### Scegli la data evento con cura

La data evento influenza la distribuzione mensile degli impatti. Usa la data effettiva in cui la negoziazione è stata formalizzata (firma del contratto, approvazione del nuovo listino, data email di conferma del fornitore).

### Usa OPEX Ripetitivo per i contratti pluriennali

Ogni volta che una negoziazione produce un beneficio che si ripeterà ogni mese (servizi, forniture a canone, accordi pluriennali), attivare **☑ OPEX Ripetitivo**. Questo distribuisce il valore sui mesi reali di competenza e rende i grafici KPI molto più accurati.

### Compila sempre la % Realizzo realisticamente

La % Realizzo non è un campo decorativo: influisce direttamente sul **Saving Effettivo** mostrato nei KPI. Se si prevede che una negoziazione si materializzerà solo parzialmente (es. perché l'ordine non è ancora emesso, o perché la controparte negoziale non ha ancora approvato), impostare una percentuale inferiore a 100 e aggiornarla quando si ha la conferma definitiva.

### Usa il Riferimento per collegare VSM e RdO

Nel campo **Riferimento** di un evento Value Stream Mapping, inserire il numero della RdO corrispondente (es. `RdO-124` o `RdO 124`). Questo permette di trovare rapidamente l'evento partendo dalla RdO e viceversa.

---

## Value Stream Mapping – Derisking

### Aggiorna lo stato dei fornitori potenziali

Il valore del registro Derisking dipende dalla qualità degli aggiornamenti. Quando un fornitore avanza nel processo di valutazione, aggiornare il campo **Stato** dalla scheda fornitore. Uno stato aggiornato permette ai KPI di Derisking di dare una fotografia reale del portafoglio fornitori.

### Usa le categorie in modo uniforme

Prima di inserire nuovi fornitori, verificare se la categoria necessaria esiste già nel catalogo. Usare la stessa nomenclatura evita duplicati (`Meccanica CNC` vs `CNC Meccanica`) e mantiene i grafici KPI per categoria accurati.

---

## Filtri e ricerca

### Usa la ricerca globale per le ricerche veloci

Per trovare rapidamente una RdO di cui si ricorda il fornitore o il codice materiale, digitare direttamente nel campo di ricerca globale. È la modalità più rapida.

### Usa i filtri avanzati per i report periodici

Per le revisioni mensili o trimestrali, usare i **filtri avanzati** con l'intervallo di date emissione. Per esempio, alla fine del mese, filtrare per "Data Emissione: dal 01/01 al 31/12" per vedere tutte le RdO dell'anno con il totale corretto.

### Pulisci sempre i filtri dopo una ricerca specifica

Dopo aver usato filtri particolari per una ricerca puntuale, fare clic su **🔎 Pulisci Filtri** per tornare alla visualizzazione completa. Lasciare filtri attivi può dare l'impressione che dei dati siano mancanti.

---

## Backup

### Esegui un backup manuale prima di ogni operazione delicata

Prima di:
- Cambiare la posizione del database
- Migrare su un nuovo server
- Aggiornare DataFlow a una nuova versione

...eseguire sempre un backup manuale da **⚙️ Impostazioni → 💾 Backup Manuale**.

### Configura il backup automatico

Attivare il backup automatico giornaliero e puntarlo a una cartella diversa da quella del database (idealmente su un'unità o server separato). In questo modo si ha sempre una copia recente in caso di guasto del server.

---

## Pulizia periodica

Una volta ogni 3–6 mesi, è utile:

1. **Archiviare le RdO obsolete** ancora nei tab Attivi.
2. **Verificare e aggiornare i fornitori potenziali** nel tab Derisking (eliminare i `Scartati` definitivi se non servono più storicamente, o tenerli per memoria storica).
3. **Riesaminare gli eventi VSM** con % Realizzo ancora bassa: o aggiornarle se la negoziazione è diventata definitiva, o documentare perché il realizzo è rimasto parziale.

---
[← Pagina precedente](IT-10-Problemi-comuni-e-soluzioni) | [Pagina successiva →](IT-12-Glossario)
