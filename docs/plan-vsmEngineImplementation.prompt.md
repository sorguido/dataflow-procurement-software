## Plan: Implementazione VSM Engine per calcolo impatti mensili

Creare il motore di calcolo VSM che genera automaticamente gli impatti economici mensili (VSMImpact) a partire da eventi VSM (VSMEvent), seguendo regole di business specifiche per riverbero, pro-rata e multiutenza.

**Steps**

1. **Creare eccezione custom VSMError**
   - Definire `services/vsm_engine.py` con classe `VSMError(Exception)`
   - Per gestione errori specifica del modulo (dati mancanti, tipo evento non supportato)
   - Coerente con pattern `DatabaseError` esistente

2. **Implementare funzione calcolo mesi distribuzione**
   - Helper privato `_calculate_distribution_months(event: VSMEvent) -> list[tuple[int, int]]`
   - Logica: se `opex_ripetitivo=True` → massimo 24 mesi, altrimenti solo fino a dicembre anno evento
   - Input: anno/mese da `event_date`
   - Output: lista di tuple `(year, month)` in ordine cronologico

3. **Implementare calcolo coefficiente pro-rata primo mese**
   - Helper privato `_calculate_first_month_coefficient(event_date: datetime) -> float`
   - Convenzione scelta: **giorni residui incluso giorno evento / 30**
   - Esempio: evento 16 marzo → (30 - 16 + 1) / 30 = 15/30 = 0.5
   - Documentare nei docstring la convenzione commerciale a 30 giorni

4. **Implementare distribuzione valore sui mesi** *(depends on 2, 3)*
   - Helper `_distribute_value(total_value: float, months: list, first_month_coeff: float) -> list[float]`
   - Primo mese: `total_value / num_mesi * first_month_coeff`
   - Mesi successivi: `total_value / num_mesi * 1.0`
   - Garantire coerenza matematica: somma quote = total_value (gestire arrotondamenti nell'ultimo mese)

5. **Implementare funzione principale generate_impacts_for_event** *(depends on 1, 2, 3, 4)*
   - Firma: `generate_impacts_for_event(event: VSMEvent) -> list[VSMImpact]`
   - Validazioni iniziali: `event_date`, `username`, `event_type` presenti
   - Branch per tipo evento:
     - **Derisking**: return `[]` (nessun impatto economico)
     - **Saving / Cost Avoidance**: calcolo impatti con riverbero
   - Usare `event.calculate_theoretical_value()` e `event.calculate_effective_value()` esistenti
   - Propagare `event.id`, `event.username` in ogni VSMImpact
   - Ordinamento cronologico finale (year asc, month asc)

6. **Aggiungere funzione batch (opzionale)**
   - `generate_impacts_for_events(events: list[VSMEvent]) -> dict[int, list[VSMImpact]]`
   - Mappa `event_id -> impacts[]`
   - Gestione robusta: se un evento fallisce, loggare e continuare con gli altri

7. **Creare file di test minimali** *(parallel with step 5)*
   - `tests/test_vsm_engine.py` usando solo `unittest` (stdlib)
   - Test cases:
     1. Saving ripetitivo con 24 mesi
     2. Cost Avoidance non ripetitivo (solo anno corrente)
     3. Primo mese pro-rata (eventi a metà mese)
     4. Derisking → lista vuota
     5. Propagazione `username` e `event_id`
     6. Ordinamento cronologico
     7. Validazione errori su dati mancanti

**Relevant files**

- `models/vsm_event.py` — Riutilizzare `calculate_theoretical_value()`, `calculate_effective_value()` (no modifiche)
- `models/vsm_impact.py` — Istanziare VSMImpact con campi calcolati (no modifiche)
- `database_manager.py` — Pattern `DatabaseError` come riferimento per `VSMError`
- `services/startup_service.py` — Esempio uso logger per pattern logging coerente

**Verification**

1. **Test unitari**: `python -m unittest tests.test_vsm_engine`
   - Verificare tutti i 7 casi di test passano
2. **Test manuale interattivo**: creare script `test_vsm_manual.py` che genera sample events e stampa impacts
   - Evento Saving ripetitivo 12 mesi → verifica 12 impact generati
   - Evento con data 15/03/2026 → primo impact marzo con coeff ~0.5
3. **Import test**: `python -c "from services.vsm_engine import generate_impacts_for_event"`
4. **Verifica matematica**: somma `valore_teorico` degli impacts deve uguagliare `event.calculate_theoretical_value()` (tolleranza arrotondamento)

**Decisions**

- **Convenzione pro-rata**: giorni residui *incluso giorno evento* diviso 30 (es. evento giorno 16 → 15 giorni residui → coeff 0.5)
- **Gestione Derisking**: restituisce lista vuota `[]` invece di generare impatti fittizi
- **Durata non ripetitivo**: distribuzione fino a dicembre dell'anno evento (non solo 1 mese)
- **Arrotondamenti**: eventuale differenza centesimale accodata all'ultimo mese per conservare total
- **Logging**: usare `logging.getLogger('DataFlow.VSMEngine')` per coerenza con progetto
- **Eccezioni**: sollevare `VSMError` per errori business logic, non usare `ValueError` o `Exception` generica
- **Location modulo**: `services/vsm_engine.py` (non root) per coerenza con `startup_service.py`

**Further Considerations**

1. **Gestione event_id = None**: evento non ancora persistito → generare comunque impacts con `event_id=None` o bloccare con VSMError?
   - **Raccomandazione**: accettare `event_id=None`, sarà aggiornato alla persistenza. Documentare nel docstring.

2. **Arrotondamenti ultimo mese**: distribuire remainder nell'ultimo impact o proporzione matematica esatta?
   - **Raccomandazione**: remainder nell'ultimo mese per semplicità e conservazione total.

3. **Test infrastruttura**: creare directory `tests/` o file singolo `test_vsm_engine.py` in root?
   - **Raccomandazione**: `tests/` directory per scalabilità futura, con `__init__.py` vuoto.
