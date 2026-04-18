# VSM Export Global Search Alignment Fix Report

## A. Diagnosi sintetica finale
- **Causa precisa**: gli export VSM non riusavano il dataset effettivamente visualizzato nel tab corrente. In particolare:
  - `Saving/Cost Avoidance` (`_export_vsm_excel`) ricostruiva eventi da pipeline separata lato export.
  - `Derisking` (`_export_derisking_excel`) ricaricava fornitori dal DB tramite funzione dedicata.
- **Effetto**: la griglia era correttamente filtrata dalla Global Search, ma l'export prendeva un set dati diverso.
- **Conferma scope bug**: la divergenza reale riguarda la **Global Search** nei tab VSM; i filtri avanzati risultavano già gestiti correttamente dal percorso di ricerca/vista.

## B. Strategia scelta
- **Approccio adottato (minimo e conservativo)**:
  1. all'export, forzare un `search_requests()` per riallineare la vista ai filtri correnti (comportamento accettato);
  2. esportare il **dataset visualizzato** (cache dominio aggiornata in fase di populate sheet), non un dataset ricostruito separatamente.
- **Perché è la più sicura**:
  - riusa pipeline già in uso dalla dashboard per applicare Global Search;
  - evita nuova logica filtro duplicata nell'export;
  - non tocca RFQ/KPI né layout/UI;
  - modifica circoscritta a punti VSM export + cache view dataset.

## C. File modificati
- `dataflow.py`

## D. Patch applicata
- In `_populate_vsm_sheet(...)` è stato aggiunto il cache del dominio visualizzato:
  - `sheet._visible_vsm_events = list(events or [])`
- In `_populate_potential_suppliers_sheet(...)` è stato aggiunto il cache del dominio visualizzato:
  - `sheet._visible_suppliers = list(suppliers or [])`
- In `_export_vsm_excel(...)`:
  - rimosso il ricalcolo dataset separato da DB;
  - aggiunto riallineamento `self.search_requests()`;
  - export eseguito da `sheet._visible_vsm_events`.
- In `_export_derisking_excel(...)`:
  - rimosso il caricamento separato `load_derisking_suppliers_for_export(...)`;
  - aggiunto riallineamento `self.search_requests()`;
  - export eseguito da `self.sheet_derisking._visible_suppliers`.
- Pulizia import non più usato:
  - rimosso `load_derisking_suppliers_for_export` dall'import di `services.excel_export_service`.

### Evidenza tecnica (linee principali)
- `dataflow.py:1568-1571` cache dataset Derisking visualizzato
- `dataflow.py:1696-1699` cache dataset Saving/Cost Avoidance visualizzato
- `dataflow.py:3130-3151` export VSM allineato a `search_requests()` + cache visualizzata
- `dataflow.py:3158-3172` export Derisking allineato a `search_requests()` + cache visualizzata

### Verifica tecnica eseguita
- `python3 -m py_compile dataflow.py` → OK

## E. Test manuali da eseguire
1. **Saving + sola Global Search → export coerente con griglia**
   1. Aprire tab Saving.
   2. Inserire una query in Global Search che riduca chiaramente le righe.
   3. Eseguire export Excel.
   4. Verificare che il file contenga solo le righe presenti nella griglia.

2. **Cost Avoidance + sola Global Search → export coerente con griglia**
   1. Aprire tab Cost Avoidance.
   2. Inserire query Global Search discriminante.
   3. Eseguire export Excel.
   4. Verificare identità tra righe esportate e righe visualizzate.

3. **Derisking + sola Global Search → export coerente con griglia**
   1. Aprire tab Derisking.
   2. Inserire query Global Search su campo fornitore/note.
   3. Eseguire export Excel.
   4. Verificare che l'export contenga solo i supplier visibili.

4. **Saving/CA + filtri avanzati senza Global Search → conferma non regressione**
   1. Pulire Global Search.
   2. Applicare filtri avanzati (date/azione/repetitive/importi) su Saving e poi su Cost Avoidance.
   3. Eseguire export.
   4. Verificare che l'export rispetti i filtri avanzati come prima.

5. **Filtri modificati ma Search non premuto → export coerente con vista riallineata**
   1. In un tab VSM, modificare filtri (global e/o avanzati) senza premere Search.
   2. Avviare export.
   3. Verificare che la vista venga riallineata e che il file esporti il dataset risultante visibile.

6. **Smoke test RFQ export invariato**
   1. Aprire tab RFQ Attive/Archiviate.
   2. Eseguire export con e senza filtri.
   3. Verificare assenza di cambiamenti comportamentali rispetto a prima.

## F. Rischi residui / rollback
- **Rischi residui reali**:
  - dipendenza dal refresh `search_requests()` al click export: se la ricerca fallisce per errore runtime, l'export VSM non procede (comportamento fail-safe);
  - in Derisking, export con vista vuota mantiene il comportamento esistente del ramo export (nessun warning aggiuntivo introdotto).
- **Rollback semplice**:
  1. ripristinare `dataflow.py` alla revisione precedente (unico file modificato);
  2. in particolare, ripristinare `_export_vsm_excel`/`_export_derisking_excel` ai loader precedenti e rimuovere cache `_visible_*`;
  3. ripristinare import `load_derisking_suppliers_for_export` se necessario.
