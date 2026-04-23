# 05 – Analisi SQDC

## Cos'è l'analisi SQDC

SQDC è un framework di valutazione multicriterio per la selezione del fornitore. Le quattro dimensioni valutate sono:

| Criterio | Significato in DataFlow |
|----------|------------------------|
| **S – Safety** | Conformità a standard di sicurezza, rischio regolatorio, solidità finanziaria e geopolitica |
| **Q – Quality** | Capacità di rispettare le specifiche tecniche concordate |
| **D – Delivery** | Rispetto dei tempi di consegna offerti, flessibilità e prontezza |
| **C – Cost** | Competitività del prezzo, condizioni di pagamento, costi accessori |

Il risultato è un **punteggio ponderato** per ogni fornitore. Il fornitore con il punteggio più alto è quello consigliato dall'analisi.

---

## Aprire l'analisi SQDC

1. Aprire il pannello di controllo di una RdO.
2. Fare clic su **📊 SQDC** nella barra superiore.
3. Si apre la finestra **"Analisi SQDC – RdO N° [numero]"**.

Se la RdO non ha fornitori, il pulsante SQDC non è disponibile.

---

## Tab 1 – Pesi (%)

Il primo tab permette di impostare l'importanza relativa dei quattro criteri:

1. Inserire nei quattro campi un valore tra 1 e 100 per ciascun criterio (Safety, Quality, Delivery, Cost).
2. La somma dei quattro valori **deve essere esattamente 100**. In caso contrario, non è possibile passare al tab successivo.
3. I valori predefiniti sono **25% per ciascun criterio** (distribuzione uniforme).

**Esempio pesi personalizzati:**
- Safety: 20%
- Quality: 30%
- Delivery: 20%
- Cost: 30%

---

## Tab 2 – Voti (1–10)

Il secondo tab mostra una tabella con una riga per ogni fornitore e quattro colonne di punteggio:

| Colonna | Tipo |
|---------|------|
| **Fornitore** | Solo lettura (determinato dalla RdO) |
| **Safety** | Inserire voto da 1 a 10 |
| **Quality** | Inserire voto da 1 a 10 |
| **Delivery** | Inserire voto da 1 a 10 |
| **Cost** | Inserire voto da 1 a 10 |
| **TOTALE** | Calcolato automaticamente |

I voti devono essere **interi da 1 a 10**. Valori non validi (lettere, decimali, valori fuori range) vengono segnalati e rifiutati.

La colonna **TOTALE** viene calcolata con la formula:

$$\text{TOTALE} = \frac{(S_{safety} \times W_{safety}) + (S_{quality} \times W_{quality}) + (S_{delivery} \times W_{delivery}) + (S_{cost} \times W_{cost})}{100}$$

Dove $W$ sono i pesi impostati nel tab precedente e $S$ sono i voti inseriti.

Quando tutti i voti sono completi, **la riga del fornitore con il punteggio più alto viene evidenziata in verde**. In caso di parità (differenza inferiore a 0,01), tutti i fornitori a pari merito vengono evidenziati.

---

## Calcolo automatico del costo

Il pulsante **🔄 Calcola Cost Automaticamente** recupera i prezzi dalla griglia RdO e calcola il voto "Cost" in modo proporzionale:

- Il fornitore con il costo totale più basso riceve **voto 10**.
- Gli altri fornitori ricevono un voto proporzionalmente più basso.
- I fornitori con prezzi mancanti ricevono **voto 0** e compare un avviso rosso nella schermata.

Questa funzione è utile come punto di partenza per il criterio economico; il buyer può comunque modificare manualmente il voto automatico.

---

## Salvare l'analisi

1. Completare i pesi e tutti i voti.
2. Fare clic su **💾 Salva SQDC**.
3. L'analisi viene salvata come **Documento Interno** allegato alla RdO.

Dopo il salvataggio, il pulsante nella barra della RdO mostra **📊 SQDC ✓** per indicare che esiste un'analisi salvata.

---

## Esportare in Excel

Fare clic su **📊 Esporta Excel** nella finestra SQDC.

Il file Excel generato contiene:
- La matrice pesi / punteggi / totali
- Il fornitore vincente evidenziato in verde
- L'avviso sulle offerte mancanti, se applicabile

---

## Comportamento in sola lettura

Se la RdO appartiene a un altro utente, la finestra SQDC si apre in modalità sola lettura: è possibile consultare i pesi e i punteggi salvati, esportare in Excel, ma non modificare nulla.

---
[← Pagina precedente](04-Gestire-una-RdO-esistente) | [Pagina successiva →](06-Value-Stream-Mapping)
