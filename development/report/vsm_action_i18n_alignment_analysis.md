# VSM Action I18N Alignment — Analisi e Piano

## 1. Evidenza del problema
Con UI in inglese, nei tab `Saving` e `Cost Avoidance` la colonna `Action` può mostrare valori italiani (es. `Negoziazione`).

Evidenza tecnica nel rendering:
- `dataflow.py:1733` usa `tr(event.action)` durante il populate griglia VSM.
- Se `event.action` è canonico IT (`Negoziazione`), `tr("Negoziazione")` non trova msgid EN e ritorna input (IT).

Conferma catalogo gettext:
- in `locale/en/LC_MESSAGES/dataflow.po` esistono `msgid "Negotiation"` e `msgid "Other"`, non `msgid "Negoziazione"`.

---

## 2. Stato attuale del dominio Action
Origine e uso del valore `action`:

- Canonico dominio: prevalentemente IT.
  - `models/vsm_event.py:28` e `:52` indicano `action` come `Negoziazione / Derisking`.
- Inserimento/modifica dialog:
  - `ui/dialogs/vsm_event_dialog.py:124-134` (`_get_action_internal`) converte display EN (`Negotiation`/`Other`) in canonico IT (`Negoziazione`/`Altro`).
  - `ui/dialogs/vsm_event_dialog.py:610` salva `action = self._get_action_internal()`.
- Rendering griglia Saving/CA:
  - `dataflow.py:1733` usa `tr(event.action)` direttamente.
- Filtro avanzato Action:
  - opzioni UI in EN in `ui/main_dashboard_builder.py:235` (`tr("Negotiation")`, `tr("Other")`).
  - confronto filtro in `services/vsm_dashboard_service.py:103` con `tr(event.action)`.
- Export Excel VSM:
  - `services/excel_export_service.py:352` mapping locale `action_map_en = {"Negoziazione": "Negotiation", "Altro": "Other"}`.
  - `services/excel_export_service.py:381` usa mapping solo in export EN.

Conclusione stato attuale: il dominio `action` non è centralizzato in i18n; è gestito con mapping locali distribuiti + `tr()` diretto su canonico.

---

## 3. Confronto con Derisking
- Somiglianze:
  - valore canonico business non allineato ai msgid EN.
  - uso di `tr(canonico_IT)` in punti UI.
  - mapping presenti ma locali/distribuiti, non single source of truth.

- Differenze:
  - VSM Action è dominio condiviso tra due tab (`Saving`, `Cost Avoidance`) e tocca anche il filtro avanzato Action.
  - Derisking era concentrato soprattutto su grid/export/dialog status.

- Stessa natura: **SI**.
  - È la stessa classe di problema i18n (bypass della pipeline `normalize -> msgid EN -> tr()`), con impatto VSM più trasversale.

---

## 4. Mapping esistente
- Dove si trova:
  - `ui/dialogs/vsm_event_dialog.py:124-145` (`_get_action_internal` / `_set_action_display`).
  - `services/excel_export_service.py:352` (`action_map_en`).

- Se esiste già:
  - Sì, esiste ma solo in punti locali.

- Se è duplicato:
  - Sì, stessa semantica replicata in dialog + export.

- Se è centralizzato:
  - No, in `utils/i18n_utils.py` non esiste ad oggi un dominio centralizzato VSM Action equivalente a RFQ/Derisking.

---

## 5. Root cause
Root cause reale: combinazione di due fattori.

1. Canonico dominio `action` salvato in IT (`Negoziazione`, `Altro`) in parte dei flussi.
2. UI/service usano `tr(event.action)` direttamente (`dataflow.py:1733`, `services/vsm_dashboard_service.py:103`) invece di passare da un mapping dominio centralizzato verso msgid EN.

Effetto: in UI EN il fallback diventa italiano e la colonna `Action` resta in IT.

---

## 6. Strategia consigliata
Il problema è della stessa natura di Derisking, quindi la strategia corretta è la stessa tipologia di allineamento:

- introdurre dominio i18n centralizzato VSM Action (pattern RFQ/Derisking);
- sostituire gli usi locali/distribuiti con API centrali di dominio;
- mantenere invariati DB, `tr(...)`, business logic e layout.

Punti da riallineare nello stesso disegno:
- griglia Saving/CA (`dataflow.py`);
- filtro avanzato Action (`services/vsm_dashboard_service.py`);
- export VSM (`services/excel_export_service.py`);
- dialog VSM (`ui/dialogs/vsm_event_dialog.py`).

---

## 7. Proposta di centralizzazione
Coerente proporla: **SI**.

Single source of truth nel layer i18n (`utils/i18n_utils.py`), estendendo pattern già esistenti:
- normalizzazione canonico (`normalize_vsm_action(...)`)
- traduzione runtime (`translate_vsm_action(...)`)

Obiettivo: eliminare mapping duplicati in dialog/export e impedire nuovi `tr(canonico_IT)` su `action`.

---

## 8. Impatto
- UI:
  - colonna `Action` coerente con lingua applicazione in Saving/Cost Avoidance.
- export:
  - nessun cambio formato; sostituzione mapping locale con mapping centrale.
- KPI:
  - impatto minimo/nullo diretto (KPI engine/chart non espongono normalmente il dominio `action` come label primaria).
- dialog:
  - comportamento utente invariato, conversione interna resa centralizzata.
- DB:
  - nessuna modifica schema/valori canonici.

---

## 9. Rischi
Solo rischi concreti:

- Valori legacy fuori dominio previsto (`action` non mappata): necessario fallback pass-through per non rompere visualizzazione.
- Filtro avanzato Action: oggi è già fragile in EN (confronto su `tr(event.action)`); riallineamento può cambiare risultati e va validato con test mirato.
- Global search: ricerca testuale opera su campo raw `action` (`dataflow.py:2696-2699`), quindi la semantica ricerca EN vs canonico IT va verificata dopo allineamento UI.
- Regressione funzionale tra Saving e Cost Avoidance: dominio condiviso, quindi qualunque disallineamento impatta entrambi i tab.

---

## 10. Strategia di test
1. UI EN: aprire Saving e Cost Avoidance e verificare `Action` tutta in EN (`Negotiation`, `Other`, `Derisking`).
2. UI IT: stessi tab, verificare `Action` in IT canonico (`Negoziazione`, `Altro`, `Derisking`).
3. Filtro avanzato Action EN: selezionare `Negotiation` e verificare match corretto su record con canonico IT.
4. Filtro avanzato Action IT: selezionare `Negoziazione`/equivalente UI IT e verificare parità comportamento.
5. Export EN/IT di Saving e Cost Avoidance: verificare coerenza Action con lingua scelta export e coerenza con griglia.
6. Dialog VSM create/edit: verificare che selezione Action resti coerente e che salvataggio/modifica non introducano valori anomali.
7. Smoke test Derisking/RFQ: verificare assenza regressioni su domini già allineati.

---

## 11. Coerenza con refactor i18n
**SI**.

Note:
- segue il principio `tr(...)` runtime-only;
- evita branch lingua manuali;
- evita duplicazione mapping locale;
- applica il principio dominio + single source of truth già adottato per RFQ e Derisking.
