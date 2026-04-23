# DataFlow 2.2.0
@sorguido

## v2.2.0

DataFlow 2.2.0 – Release di funzionalità

---

## 🆕 Added

- **Export RFQ in PDF**: nuovo flusso di export direttamente dalla finestra RFQ
  - Generazione documento A4 multipagina basata su ReportLab
  - Logo aziendale persistente opzionale per gli export PDF
  - Template testuali esterni specifici per lingua e modificabili con validazione del placeholder `{{TABLE}}`

- **Servizi di export dedicati**: logica di export Excel estratta da `dataflow.py` in servizi dedicati
  - Export RFQ
  - Export VSM
  - Export Derisking
  - Generazione sicura di filename di export con timestamp

- **Sistema di suggerimento nome fornitore**: flusso riutilizzabile di suggerimento e rilevamento soft dei duplicati per i campi input fornitore
  - Suggerimenti costruiti sia dalla storia fornitori RFQ sia dal registro fornitori Derisking
  - Integrato nei dialog di modifica fornitori e dei fornitori potenziali

- **Servizi di manutenzione impostazioni**: layer di supporto dedicato per i flussi operativi di manutenzione
  - Copia bundle di backup manuale includendo DB/WAL/SHM quando disponibili
  - Backup automatico giornaliero con retention timestamped
  - Helper di validazione per la migrazione della cartella DataFlow e lifecycle di restart controllato

- **Sheet factory layer**: builder standalone per le istanze `tksheet` di RFQ, VSM e Derisking

---

## ✨ Improvements

- Ricerca dashboard e filtri avanzati allineati tra le schede RFQ, Saving, Cost Avoidance e Derisking
- La ricerca globale ora coesiste in modo più consistente con i filtri contestuali e i flussi di export
- La copertura della traduzione runtime è consolidata attorno a `tr(...)` e a helper centralizzati di normalizzazione per RFQ type, VSM action e Derisking status
- I flussi principali dei dialog sono standardizzati attraverso componenti dialog riutilizzabili per messaggio, conferma, lingua export, splash, licenza e prompt identità
- I filtri di KPI Analysis sono estesi con preset rolling (`1M`, `3M`, `12M`, `3Y`, `5Y`, `10Y`, `ALL`) e gestione dedicata dell'anno Derisking
- Il flusso dashboard Derisking ora usa logica dedicata di popolamento supplier-sheet e auto-sizing delle colonne
- I flussi di export RFQ, VSM, KPI e Derisking sono allineati più strettamente al comportamento bilingue della UI

---

## 🛠 Fixes

- L'input di ricerca ora applica sanitizzazione difensiva e validazione di lunghezza prima dell'esecuzione della query
- I controlli di ownership delle righe dashboard sono applicati tramite helper dedicati di selection-policy prima di abilitare azioni distruttive
- Il caricamento delle note ora preferisce `json.loads()` con fallback per contenuti serializzati legacy
- L'attivazione del drag-and-drop degli allegati è protetta in modo più sicuro quando `tkdnd` / `tkinterdnd2` non sono disponibili nel runtime
- Il flusso di restart dopo le modifiche alle impostazioni è gestito tramite helper dedicati di lifecycle per un comportamento di riavvio più affidabile
- I flussi di migrazione cartella DataFlow e backup includono validazione più conservativa e salvaguardie di copia

---

## ♻️ Refactoring

- Ampie porzioni della logica operativa di `dataflow.py` sono state estratte in servizi dedicati sotto `services/`
- La logica dashboard è stata suddivisa tra moduli controller, search, actions policy, selection policy, RFQ dashboard, VSM dashboard e Derisking dashboard
- La logica impostazioni è stata suddivisa tra servizi preferences, maintenance, location e restart lifecycle
- L'export RFQ in PDF è stato suddiviso in moduli di servizio dedicati per export, logo e template
- Il nuovo `ui/sheet_factories.py` centralizza la costruzione delle `tksheet` per le viste dashboard
- Nuovi moduli utility introdotti per la generazione dei filename di export e la normalizzazione dei nomi fornitore
- `utils/i18n_utils.py` è stato rifattorizzato in un translation service centralizzato con inizializzazione runtime e helper di normalizzazione domain-specific

---

## 📦 Packaging

- Pacchetti di distribuzione aggiornati per la versione 2.2.0:
  - `AppImage`
  - `.exe`
- `requirements.txt` aggiornato per includere le runtime dependencies richieste dai nuovi flussi di export e drag-and-drop
- Le root PyInstaller specs sono state aggiornate per raccogliere i moduli e gli asset richiesti dall'export PDF e dalle risorse runtime multilingua

---

## 🔒 Other

- Nessuna breaking change ai workflow utente esistenti di RFQ / VSM / Derisking
- L'evoluzione del database resta conservativa e orientata alla migrazione per nuovi campi e tabelle
- L'aggregazione dashboard multi-user resta in sola lettura per i dati appartenenti ad altri utenti

---

## v2.1.0

DataFlow 2.1.0 – Release di funzionalità

---

## 🆕 Added

- **Modulo VSM — Value Stream Mapping**: nuova area funzionale, indipendente dal workflow RFQ, per tracciare risultati di negoziazione quantificati
  - Tre tipi di evento supportati: **Saving**, **Cost Avoidance**, **Derisking**
  - Dialog VSM Event con layout form dinamico (campi mostrati/nascosti in base al tipo evento)
  - Modalità di visualizzazione sola lettura per eventi appartenenti ad altri utenti (coerenza multi-user)
  - Flag OPEX-ripetitivo: gli eventi Saving e Cost Avoidance con impatto ricorrente si propagano fino a 24 mesi
  - Driver termini di pagamento (Pagamenti) per eventi Saving, con coefficiente mensile di costo opportunità configurabile

- **VSM Engine**: motore automatico di proiezione degli impatti mensili
  - Calcolo pro-rata per il primo mese
  - Rigenerazione deterministica di tutti gli impatti a ogni aggiornamento evento (pattern DELETE–REGENERATE–SAVE)
  - Transazioni database atomiche: rollback completo in caso di qualsiasi errore di insert

- **Finestra KPI Analysis**: finestra dedicata con KPI procurement aggregati
  - Quattro schede: RFQ, Saving, Cost Avoidance, Derisking
  - Filtri periodo: selettore anno o intervallo date personalizzato (preset: ultimi 3, 6, 12 mesi oppure All)
  - KPI card con dati reali dal KPI engine
  - Grafici a barre integrati (pure Tkinter, nessuna external chart library richiesta)
  - Export Excel del riepilogo KPI completo

- **Registro dei fornitori potenziali** (scheda Derisking): modulo dedicato alla gestione della qualificazione fornitore
  - Modello `PotentialSupplier` con nome, categoria, dettagli di contatto, website e note
  - Ciclo di vita della qualificazione: New → Under Evaluation → Qualified / Rejected
  - Ownership per utente con visibilità multi-user
  - Integrato con gli eventi VSM Derisking (campo new supplier)

- **Tre nuove schede nella dashboard principale**: Saving, Cost Avoidance, Derisking — direttamente accessibili accanto alle schede RFQ esistenti

- **Barra di ricerca globale** (`MainDashboardToolbar`): entry di ricerca centrale con testo placeholder, attiva in tutte le schede dashboard
  - Ricerca OR multi-campo: numero RFQ, riferimento, fornitore, part code, description, order number
  - Coesiste con i filtri avanzati (global OR + contextual filters AND)
  - La ricerca vuota attiva il reset dei filtri

- **Gestione Supplier Category**: dialog per creare e gestire label di categoria fornitore riutilizzabili (usate nel registro dei fornitori potenziali)

---

## ✨ Improvements

- Il pannello filtri della dashboard principale ora include un sotto-frame VSM dedicato (user, date from, date to) mostrato/nascosto contestualmente quando si passa tra schede RFQ e VSM
- I grafici KPI si adattano al resize del canvas; i grafici dual-bar includono legenda e label degli assi
- Duplicazione evento VSM: gli eventi esistenti possono essere duplicati dalla dashboard, preservando tutti i campi
- Filtro username VSM sulla dashboard: popolato tramite aggregazione multi-database, coerente con il comportamento del filtro username RFQ esistente
- Le funzioni `get_available_years` e `get_available_years_derisking` espongono liste anno distinte per le combo filtro
- Migliorata la formattazione numerica per i valori monetari nelle KPI card e nell'export Excel (locale italiano: separatore migliaia punto, decimale virgola)
- L'export KPI Excel include supporto bilingue (italiano / inglese) coerente con gli export Excel esistenti

---

## 🛠 Fixes

- Dashboard controller separato da `MainWindow.__init__`, eliminando diversi casi in cui lo stato dei filtri non veniva preservato tra i cambi di scheda
- Il sotto-frame filtri VSM viene correttamente nascosto tornando alle schede RFQ e mostrato attivando le schede VSM
- La validazione difensiva in `populate_username_filter` previene crash su tuple aggregate a lunghezza variabile provenienti da query multi-database
- La protezione del rendering dell'immagine logo: immagini con dimensione zero non causano più `ZeroDivisionError` all'avvio

---

## ♻️ Refactoring

- La costruzione UI di `MainWindow` è stata estratta in `ui/main_dashboard_builder.py` (pure widget builder, nessun data loading)
- La logica di orchestrazione dashboard è stata estratta in `services/dashboard_controller.py`
- `MainDashboardToolbar` e `CollapsibleFilters` introdotti come componenti UI standalone sotto `ui/components/`
- VSM persistence, VSM engine, KPI engine, KPI chart data e KPI Excel export vivono ciascuno in moduli dedicati sotto `services/` e `ui/`
- Il package `models/` è stato esteso con le dataclass `VSMEvent`, `VSMImpact` e `PotentialSupplier`
- `utils/vsm_config.py` introdotto per parametri VSM configurabili dall'utente (payment coefficient) memorizzati in `config.ini`
- `utils/validation_utils.py` esteso con validazione formato email e website usata dal dialog Potential Supplier
- Test suite estesa: `test_vsm_engine.py`, `test_vsm_persistence.py`, `test_vsm_event_model.py`, `test_supplier_category_persistence.py`

---

## 📦 Packaging

- Pacchetti di distribuzione aggiornati per la versione 2.1.0:
  - `AppImage`
  - `.exe`
- `dataflow.spec` aggiornato per includere i nuovi moduli e asset introdotti in questa release

---

## 🔒 Other

- Nessuna breaking change ai dati RFQ esistenti o allo schema database per le tabelle esistenti
- Nuove tabelle database (`vsm_events`, `vsm_impacts`, `potential_suppliers`, `supplier_categories`) create in modo trasparente al primo avvio
- Nessuna nuova mandatory external dependency introdotta

---

> Questa release introduce nuove aree funzionali ed estende le capacità di DataFlow oltre la gestione RFQ. I team procurement possono ora tracciare l'impatto quantificato delle attività di negoziazione e misurare la performance del buyer attraverso un layer KPI dedicato.
