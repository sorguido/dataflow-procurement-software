# Report Fix Minimo Global Search i18n (Soluzione A)

Data: 2026-04-20
Ambito: DataFlow Procurement Software
Modalita: STRICT / LOW RISK / NO REGRESSIONS

## A. File modificati
- `services/dashboard_search_service.py` (unico file di codice toccato)
- `development/report/global_search_i18n_solutionA_fix_2026-04-20.md` (questo report richiesto)

## B. Funzioni/blocchi toccati
In `services/dashboard_search_service.py`:
- Nuovo blocco costanti: `_GLOBAL_SEARCH_EN_TO_IT_EXACT`
- Nuovo helper locale: `_normalize_global_search_query_for_closed_domains(query)`
- Aggiornata `filter_derisking_suppliers_by_query(...)`:
  - normalizzazione query applicata una sola volta a monte
  - confronto finale invariato: substring su valori raw dei record
- Aggiornata `filter_vsm_events_by_query(...)`:
  - normalizzazione query applicata una sola volta a monte
  - confronto finale invariato: substring su valori raw dei record

## C. Strategia applicata
Implementata esclusivamente la Soluzione A:
- query utente normalizzata in ingresso EN -> canonico IT (solo token noti)
- nessun doppio match raw+translated
- nessuna traduzione runtime dei valori record
- nessun uso di `tr()` nella logica di search
- confronto finale rimasto case-insensitive e basato su valori raw DB

## D. Mapping introdotti
Mapping esplicito EN -> IT (exact match sulla query normalizzata, case-insensitive):

Domino `action` (VSM):
- `negotiation` -> `negoziazione`
- `other` -> `altro`
- `derisking` -> `derisking`

Dominio `supplier_status` (Derisking):
- `new` -> `nuovo`
- `under evaluation` -> `in valutazione`
- `qualified` -> `qualificato`
- `rejected` -> `scartato`
- `approved` -> `qualificato` (alias compatibilita)

Nota RFQ status:
- in `services/dashboard_search_service.py` non sono presenti campi RFQ status nel punto di match globale; quindi nessun mapping RFQ e stato lasciato invariato.

## E. Casi coperti
Verifica logica/manuale eseguita via script `python3` sui metodi del service:
- Caso 1 (IT canonico): `qualificato` continua a matchare record con `Qualificato`
- Caso 2 (EN mappato):
  - `qualified` matcha `Qualificato`
  - `new` matcha `Nuovo`
  - `negotiation` matcha `Negoziazione`
- Caso 3 (query non mappata): comportamento invariato (`plastic` matcha su campi raw come prima)
- Caso 4 (Advanced Filters): nessun impatto, nessun file/flow dei filtri avanzati toccato
- Caso 5 (robustezza): nessun errore con dataset vuoti / query `None` / campi stringa mancanti (`or ""` invariato)

## F. Rischi residui
- La normalizzazione e intenzionalmente minimale e per exact-label: frasi miste (es. testo libero con parole chiuse dentro una frase) non vengono riscritte integralmente.
- Il fix e confinato al service `dashboard_search_service.py` (Derisking/VSM). Se esistono mismatch i18n della Global Search RFQ in altri path, restano fuori scope per vincolo richiesto.

## G. Rollback semplice
Rollback immediato e reversibile:
1. Ripristinare `services/dashboard_search_service.py` alla revisione precedente
2. Rimuovere l'helper `_normalize_global_search_query_for_closed_domains` e il dict `_GLOBAL_SEARCH_EN_TO_IT_EXACT`
3. Ripristinare uso diretto di `query` nei due filtri

Nessuna migrazione DB, nessun cambio schema, nessuna dipendenza nuova.

## H. Conferme esplicite
Confermo che:
- non e stata implementata la soluzione B (nessun doppio match raw+translated)
- non e stato toccato `tr()`
- non e stato modificato `utils/i18n_utils.py`
- non sono stati modificati Advanced Filters
- non sono stati modificati DB/schema/persistence
- non e stata modificata la UI
- non sono stati toccati altri file di codice fuori scope (unico file codice: `services/dashboard_search_service.py`)
