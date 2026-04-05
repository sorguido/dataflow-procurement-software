# DataFlow 2.1.0
@sorguido

## v2.1.0

DataFlow 2.1.0 – Feature Release

---

## 🆕 Aggiunto

- **Modulo VSM — Value Stream Mapping**: nuova area funzionale, indipendente dal workflow RdO, per il tracciamento dei risultati negoziali quantificati
  - Tre tipi di evento supportati: **Saving**, **Cost Avoidance**, **Derisking**
  - Dialog eventi VSM con layout form dinamico (campi mostrati/nascosti in base al tipo di evento)
  - Modalità visualizzazione in sola lettura per gli eventi di altri utenti (coerenza multiutenza)
  - Flag OPEX-ripetitivo: gli eventi Saving e Cost Avoidance con impatto ricorrente si propagano fino a 24 mesi
  - Driver termini di pagamento (Pagamenti) per gli eventi Saving, con coefficiente di costo opportunità mensile configurabile

- **VSM Engine**: motore di proiezione degli impatti economici mensili
  - Calcolo pro-rata per il primo mese
  - Rigenerazione deterministica di tutti gli impatti ad ogni aggiornamento evento (pattern DELETE–REGENERATE–SAVE)
  - Transazioni database atomiche: rollback completo in caso di errore di inserimento

- **Finestra KPI Analysis**: finestra dedicata con KPI acquisti aggregati
  - Quattro tab: RdO, Saving, Cost Avoidance, Derisking
  - Filtri periodo: selettore anno o intervallo date personalizzato (preset: ultimi 3, 6, 12 mesi o Tutti)
  - Card KPI con dati reali dal motore KPI
  - Grafici a barre integrati (Tkinter puro, nessuna libreria grafica esterna richiesta)
  - Export Excel del riepilogo KPI completo

- **Anagrafica Fornitori Potenziali** (tab Derisking): modulo dedicato per la gestione della qualifica dei fornitori
  - Modello `PotentialSupplier` con ragione sociale, categoria merceologica, recapiti, sito web e note
  - Ciclo di vita della qualifica: Nuovo → In valutazione → Qualificato / Scartato
  - Ownership per utente con visibilità multiutente
  - Integrato con gli eventi VSM Derisking (campo nuovo fornitore)

- **Tre nuovi tab nella dashboard principale**: Saving, Cost Avoidance, Derisking — accessibili direttamente accanto ai tab RdO esistenti

- **Barra di ricerca globale** (`MainDashboardToolbar`): campo di ricerca centrale con testo placeholder, attivo su tutti i tab della dashboard
  - Ricerca OR multi-campo: numero RdO, riferimento, fornitore, codice articolo, descrizione, numero ordine
  - Coesiste con i filtri avanzati (ricerca globale OR + filtri contestuali AND)
  - Ricerca vuota attiva il reset dei filtri

- **Gestione Categorie Fornitori**: dialog per la creazione e gestione delle etichette di categoria riutilizzabili (usate nell'anagrafica Fornitori Potenziali)

---

## ✨ Miglioramenti

- Il pannello filtri della dashboard principale include ora un sub-frame VSM dedicato (utente, dal, al) mostrato/nascosto in modo contestuale al passaggio tra tab RdO e tab VSM
- I grafici KPI si adattano al ridimensionamento del canvas; i grafici a doppia barra includono legenda e label degli assi
- Duplicazione eventi VSM: gli eventi esistenti possono essere duplicati dalla dashboard, preservando tutti i campi
- Filtro username VSM nella dashboard: popolato dall'aggregazione multi-database, coerente con il comportamento del filtro username RdO esistente
- Funzioni `get_available_years` e `get_available_years_derisking` espongono liste di anni distinte per i combo filtro
- Formattazione numerica migliorata per i valori monetari nelle card KPI e nell'export Excel (locale italiano: separatore migliaia punto, decimale virgola)
- L'export Excel KPI include supporto bilingue (italiano / inglese) coerente con gli export Excel esistenti

---

## 🛠 Correzioni

- Il controller della dashboard è stato separato da `MainWindow.__init__`, eliminando diversi casi in cui lo stato dei filtri non veniva preservato al passaggio tra tab
- Il sub-frame filtri VSM è correttamente nascosto al ritorno ai tab RdO e mostrato all'attivazione dei tab VSM
- Validazione difensiva in `populate_username_filter` previene il crash su tuple di lunghezza variabile provenienti da query multi-database aggregate
- Guard per il rendering dell'immagine logo: le immagini con dimensioni zero non causano più `ZeroDivisionError` all'avvio

---

## ♻️ Refactoring

- La costruzione UI di `MainWindow` è stata estratta in `ui/main_dashboard_builder.py` (pure widget builder, nessun caricamento dati)
- La logica di orchestrazione della dashboard è stata estratta in `services/dashboard_controller.py`
- `MainDashboardToolbar` e `CollapsibleFilters` introdotti come componenti UI standalone sotto `ui/components/`
- VSM persistence, VSM engine, KPI engine, KPI chart data e KPI Excel export risiedono ciascuno in moduli dedicati sotto `services/` e `ui/`
- Pacchetto `models/` esteso con i dataclass `VSMEvent`, `VSMImpact` e `PotentialSupplier`
- `utils/vsm_config.py` introdotto per i parametri VSM configurabili per utente (coefficiente pagamenti) archiviati in `config.ini`
- `utils/validation_utils.py` esteso con validazione formato email e sito web usata dal dialog Fornitori Potenziali
- Suite di test estesa: `test_vsm_engine.py`, `test_vsm_persistence.py`, `test_vsm_event_model.py`, `test_supplier_category_persistence.py`

---

## 📦 Distribuzione

- Pacchetti di distribuzione aggiornati per la versione 2.1.0:
  - `.deb`
  - `AppImage`
  - `.exe`
- `dataflow.spec` aggiornato per includere i nuovi moduli e asset introdotti in questa release

---

## 🔒 Altro

- Nessuna modifica incompatibile ai dati RdO esistenti o allo schema database per le tabelle esistenti
- Le nuove tabelle database (`vsm_events`, `vsm_impacts`, `potential_suppliers`, `supplier_categories`) vengono create in modo trasparente al primo avvio
- Nessuna nuova dipendenza esterna obbligatoria introdotta

---

> Questa release introduce nuove aree funzionali ed estende le capacità di DataFlow oltre la gestione delle RdO. I team acquisti possono ora tracciare l'impatto quantificato delle attività negoziali e misurare le performance dei buyer attraverso un livello KPI dedicato.
