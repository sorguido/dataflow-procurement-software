# Derisking Advanced Filter - User Visibility Bug Analysis

## 1. Sintomo osservato
Nel tab Derisking, dopo la patch che ha riabilitato gli Advanced Filters:
- la combo `User` è visibile e selezionabile;
- selezionando un utente diverso da quello corrente la lista può diventare vuota;
- non compare alcun errore UI né stack trace a terminale.

## 2. File e punti di codice analizzati
- `ui/main_dashboard_builder.py:190-210`
  - creazione `vsm_username_filter_var` e combo `User` condivisa VSM.
- `dataflow.py:1156-1205`
  - `populate_vsm_username_filter()` popola utenti della combo.
- `dataflow.py:1207-1215`
  - `_on_vsm_username_filter_changed()` ricarica Saving/CA/Derisking.
- `dataflow.py:1141-1154`
  - `_get_active_username_filter()` trasforma valore combo in filtro effettivo.
- `dataflow.py:1521-1534`
  - `_load_potential_suppliers()` applica filtro utente su Derisking.
- `dataflow.py:2613-2634`
  - `_search_derisking_suppliers()` applica stesso filtro nel percorso search.
- `services/supplier_persistence.py:121-132`
  - `get_all_suppliers()` delega a DB manager.
- `database_manager.py:2662-2699`
  - `get_all_potential_suppliers(username)` query finale `WHERE username = ?`.
- `database_manager.py:2304-2399`
  - `get_all_vsm_events_aggregated(...)`: aggregazione multi-DB usata per VSM.
- `services/app_paths.py:173-179`
  - `get_db_path()` punta al DB dell'utente corrente (`dataflow_db_<username>.db`).
- `ui/dialogs/potential_supplier_dialog.py:433-451` + `database_manager.py:2531-2567`
  - persistenza `username` fornitore senza normalizzazione forzata in lowercase.

## 3. Flusso reale del filtro utente
1. Popolamento combo utenti:
   - `populate_vsm_username_filter()` legge username da `get_all_vsm_events_aggregated(get_db_path())` (`dataflow.py:1169-1176`).
   - Quindi la combo prende il dominio utenti dagli **eventi VSM aggregati multi-DB**, non dai fornitori Derisking.
   - I valori vengono normalizzati in lowercase e popolati come `[All users, user1, user2, ...]` (`dataflow.py:1194-1201`).

2. Valore salvato nella variabile Tkinter:
   - la combo è legata a `vsm_username_filter_var` (`ui/main_dashboard_builder.py:190-205`).
   - alla selezione, la variabile contiene il testo selezionato (placeholder o username).

3. Valore letto e trasformato:
   - `_get_active_username_filter(self.vsm_username_filter_var)` (`dataflow.py:1530`, `2622`) restituisce:
     - `None` se valore vuoto o `All users` (`dataflow.py:1151-1153`),
     - `value.lower()` altrimenti (`dataflow.py:1154`).

4. Applicazione a Derisking:
   - `_load_potential_suppliers()` apre **solo** `DatabaseManager(get_db_path())` (`dataflow.py:1531`).
   - `get_all_suppliers(..., username=username_filter)` (`dataflow.py:1532`) -> `get_all_potential_suppliers(username)` (`services/supplier_persistence.py:132`).

5. Filtro/query finale:
   - se `username_filter` è valorizzato: `SELECT ... FROM potential_suppliers WHERE username = ?` (`database_manager.py:2674-2685`).
   - confronto case-sensitive, senza `LOWER(...)`.

### Verifica mismatch richiesti
- Mismatch `combo domain` vs `dataset Derisking`:
  - combo utenti da VSM eventi aggregati;
  - Derisking letto da `potential_suppliers` del solo DB locale.
- Mismatch `valore visualizzato` vs `valore atteso query`:
  - combo usa lowercase; query richiede `username` esatto (`=`), potenzialmente case-sensitive.
- Vincolo implicito su utente corrente:
  - non è hardcoded nel filtro, ma è implicito nel data source: `get_db_path()` punta al DB utente corrente.
- Coerenza `All users`:
  - `All users -> None` è coerente semanticamente;
  - ma in Derisking significa “tutti gli utenti nel DB locale”, non “tutti gli utenti aggregati”.
- Filtri addizionali che svuotano i risultati:
  - nel percorso base combo-change no (solo filtro username);
  - nel percorso search può sommarsi anche `global` (`dataflow.py:2621-2654`).
- Differenze `current user / username / owner`:
  - Derisking usa il campo `username` su `potential_suppliers`;
  - nessun campo owner separato nel flusso analizzato.

## 4. Causa più probabile
**Causa primaria (alta confidenza): disallineamento tra sorgente utenti della combo e sorgente dati Derisking interrogata.**

In dettaglio:
- la combo mostra utenti trovati negli eventi VSM aggregati (anche altri DB);
- Derisking interroga `potential_suppliers` del solo DB locale;
- selezionando un utente “valido” in combo ma non presente nel `potential_suppliers` locale (o presente in DB sibling) il risultato è vuoto.

Questa dinamica spiega perfettamente sintomo + assenza errori.

## 5. Cause alternative da non escludere
1. **Case mismatch username**
   - filtro passato in lowercase (`_get_active_username_filter`), ma query Derisking usa `WHERE username = ?` senza normalizzazione (`database_manager.py:2682`).
   - se record storici hanno maiuscole/minuscole diverse, può risultare vuoto anche con utente teoricamente presente.

2. **Dominio combo non allineato al dominio supplier**
   - anche senza multi-DB: la combo deriva da `vsm_events`, non da `potential_suppliers`; può esporre utenti senza record Derisking.

3. **Global search residua nel percorso Search**
   - se viene usato il pulsante/Enter search, il filtro `global` può ulteriormente restringere (`dataflow.py:2621-2654`).

## 6. Perché non emergono errori a runtime
- `SELECT ... WHERE username = ?` con zero righe è comportamento normale, non eccezione.
- il codice tratta lista vuota come caso valido e popola la griglia con 0 righe (`dataflow.py:1533`, `2633`).
- i blocchi `try/except` mostrano dialog solo su eccezioni reali DB/logic, non su empty result set.

## 7. Fix minimo consigliato
**Non implementato in questo task (solo proposta).**

Fix minimo, sicuro, reversibile:
1. Introdurre per Derisking un helper dataset con semantica allineata a VSM (`services/vsm_dashboard_service.py:15-29` come pattern):
   - `username_filter is None` -> aggregazione multi-DB `potential_suppliers`;
   - `username_filter == current_user` -> locale;
   - `username_filter altro utente` -> aggregazione multi-DB filtrata per utente.
2. Usare tale helper sia in `_load_potential_suppliers()` sia in `_search_derisking_suppliers()`.
3. Hardening di compatibilità (consigliato ma minimale): confronto case-insensitive per Derisking (`LOWER(username)=?`) o normalizzazione in scrittura.

Perché è il più sicuro:
- non cambia UX;
- mantiene default conservativo (utente corrente);
- modifica confinata al percorso Derisking data-loading.

## 8. Verifiche manuali post-fix
1. Apri Derisking con default: deve mostrare stesso dataset pre-fix (utente corrente).
2. Seleziona `All users`: devono apparire anche supplier da DB sibling (se presenti).
3. Seleziona un altro utente con supplier presenti: lista non vuota e coerente.
4. Seleziona un utente senza supplier: lista vuota ma senza errori (atteso).
5. Verifica percorso search (`global`) con stesso utente: filtri combinati coerenti.
6. Verifica `Clear Filters`: ritorno a default utente corrente.
7. Re-test Saving/Cost Avoidance user filter: nessuna regressione.
8. Re-test edit/delete read-only behavior su record altri utenti (se previsto).

## 9. Rischi residui
- Medio-basso se fix confinato al loader Derisking.
- Rischio principale: introdurre regressione sulla semantica "All users" tra tab VSM e Derisking se non si allinea bene il dataset helper.
- Rischio secondario: metadata ownership/read-only in Derisking se si aggiunge aggregazione senza tracciare `source_file` (se necessario in workflow edit).

## 10. Rollback
Per questa attività di sola analisi, rollback = eliminazione del report.

Comando:
```bash
git restore --staged development/report/derisking_advanced_filter_user_bug_analysis.md 2>/dev/null || true
rm -f development/report/derisking_advanced_filter_user_bug_analysis.md
```

Se in un task successivo verrà applicato il fix, rollback consigliato: ripristinare esclusivamente i file del fix Derisking (senza toccare RFQ/Saving/CA).
