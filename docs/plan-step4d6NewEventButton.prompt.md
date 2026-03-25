# STEP 4D.6 — TRASFORMARE IL PULSANTE "+ Nuova RdO" IN "+ Nuovo Evento" DINAMICO

## CONTESTO

- Nel codice attuale NON esiste un pulsante globale "New Event"
- Esiste il pulsante attuale "+ Nuova RdO" nella toolbar principale
- VSM è già integrato nei tab:
  - Saving
  - Cost Avoidance
  - Derisking
- Il dialog VSM esiste già e viene usato per edit
- RFQ deve continuare a funzionare come prima

## OBIETTIVO

Trasformare il pulsante globale esistente:
- da "+ Nuova RdO"
- a "+ Nuovo Evento"

con comportamento dinamico basato sul tab attivo.

## COMPORTAMENTO ATTESO

**SE tab attivo = RFQ:**
- click su "+ Nuovo Evento" → apre la normale creazione nuova RdO
- comportamento RFQ invariato

**SE tab attivo = Saving:**
- click su "+ Nuovo Evento" → apre VSMEventDialog in modalità CREATE
- event_type = "Saving"

**SE tab attivo = Cost Avoidance:**
- click su "+ Nuovo Evento" → apre VSMEventDialog in modalità CREATE
- event_type = "Cost Avoidance"

**SE tab attivo = Derisking:**
- click su "+ Nuovo Evento" → apre VSMEventDialog in modalità CREATE
- event_type = "Derisking"

## IMPLEMENTAZIONE

### MODIFICA 1: Pulsante toolbar (dataflow.py linea ~3594)

**Cambia:**
```python
# 1. New RfQ
self.btn_new_rdo = ttk.Button(frame_top, text=_("➕ Nuova RdO"), command=self.open_new_request_window)
self.btn_new_rdo.pack(side="left", padx=(0, 10))
```

**In:**
```python
# 1. New Event (dinamico: RdO o VSM in base al tab)
self.btn_new_event = ttk.Button(frame_top, text=_("➕ Nuovo Evento"), command=self.open_new_event)
self.btn_new_event.pack(side="left", padx=(0, 10))
```

### MODIFICA 2: Nuovo handler open_new_event() (dataflow.py prima di open_new_request_window, circa linea ~5925)

**Aggiungi questo metodo PRIMA di `open_new_request_window()`:**

```python
def open_new_event(self):
    """Handler dinamico per pulsante + Nuovo Evento.
    
    Step 4D.6: Routing intelligente basato sul tab attivo:
    - RFQ (attiva/archiviata): crea nuova RdO
    - VSM (Saving/Cost Avoidance/Derisking): crea nuovo evento VSM
    """
    _, status = self.get_current_tree_and_status()
    
    # Branch 1: RFQ - usa la logica esistente
    if status in ('attiva', 'archiviata'):
        self.open_new_request_window()
    
    # Branch 2: VSM - apri dialog CREATE
    elif status.startswith('vsm_'):
        # Mappa status → event_type (pattern già usato in _edit_vsm_event)
        event_type_map = {
            'vsm_saving': 'Saving',
            'vsm_cost_avoidance': 'Cost Avoidance',
            'vsm_derisking': 'Derisking'
        }
        event_type = event_type_map.get(status)
        
        if not event_type:
            return  # Fail-safe
        
        # Lazy import (come in _edit_vsm_event)
        from ui.dialogs.vsm_event_dialog import VSMEventDialog
        
        try:
            # Apri dialog in modalità CREATE (event_id=None)
            dialog = VSMEventDialog(
                self.root,
                current_username=self.current_username,
                event_type=event_type,
                event_id=None  # CREATE mode
            )
            self.root.wait_window(dialog)
            
            # Refresh se salvato
            if hasattr(dialog, 'result') and dialog.result:
                # Ottieni sheet corrente
                sheet, _ = self.get_current_tree_and_status()
                self._load_vsm_events(event_type, sheet)
                logger.info(f"Nuovo evento VSM {event_type} creato con successo")
        
        except Exception as e:
            logger.error(f"Errore creazione evento VSM: {e}", exc_info=True)
            messagebox.showerror(
                _("Errore"),
                _("Impossibile aprire il form: {}").format(e),
                parent=self.root
            )
```

### MODIFICA 3: Traduzioni (locale/en/LC_MESSAGES/dataflow.po linea ~1712)

**Cambia:**
```po
msgid "➕ Nuova RdO"
msgstr "➕ New RfQ"
```

**In:**
```po
msgid "➕ Nuovo Evento"
msgstr "➕ New Event"
```

### MODIFICA 4: Compila traduzioni

Dopo aver modificato il file .po, esegui:
```bash
python3 compile_translations.py
```

## PATTERN RIUSATI

1. **get_current_tree_and_status()**: Già esistente, usato ovunque per rilevare tab attivo
2. **event_type_map**: Stesso pattern usato in `_edit_vsm_event()` (linea ~4486)
3. **VSMEventDialog**: Già usato in edit mode, ora usato in CREATE mode (event_id=None)
4. **_load_vsm_events()**: Già esistente per refresh sheet VSM
5. **Lazy import**: Stesso pattern di `_edit_vsm_event()` per evitare import circolari

## REGOLE RISPETTATE

✅ NON creare un nuovo pulsante  
✅ NON creare un dropdown in questo step  
✅ RIUTILIZZARE il pulsante toolbar esistente  
✅ NON modificare RFQ se non per il cambio label + routing  
✅ NON duplicare codice  
✅ NON creare nuove architetture  

## VERIFICA FUNZIONALE

### Test RFQ:
1. Vai al tab "RdO Attive"
2. Click su "+ Nuovo Evento"
3. **Atteso**: Si apre il dialog di scelta tipo RdO (Conto Lavoro / Fornitura / etc.)
4. Completa creazione RdO
5. **Atteso**: RdO creata e aperta per editing

### Test VSM Saving:
1. Vai al tab "Saving"
2. Click su "+ Nuovo Evento"
3. **Atteso**: Si apre VSMEventDialog con event_type="Saving" in modalità CREATE
4. Compila form e salva
5. **Atteso**: Nuovo evento appare nella sheet Saving

### Test VSM Cost Avoidance:
1. Vai al tab "Cost Avoidance"
2. Click su "+ Nuovo Evento"
3. **Atteso**: Si apre VSMEventDialog con event_type="Cost Avoidance" in modalità CREATE
4. Compila form e salva
5. **Atteso**: Nuovo evento appare nella sheet Cost Avoidance

### Test VSM Derisking:
1. Vai al tab "Derisking"
2. Click su "+ Nuovo Evento"
3. **Atteso**: Si apre VSMEventDialog con event_type="Derisking" in modalità CREATE
4. Compila form e salva
5. **Atteso**: Nuovo evento appare nella sheet Derisking

## NESSUNA REGRESSIONE

- ✅ RFQ Active/Archived: Funzionano come prima
- ✅ Edit VSM: Già implementato, non toccato
- ✅ Delete VSM: Già implementato, non toccato
- ✅ Duplicate VSM: Già implementato, non toccato
- ✅ Double-click: Già implementato, non toccato
- ✅ Actions menu: Non toccato
- ✅ Export: Non toccato
- ✅ KPI: Non toccato
- ✅ Filtri: Non toccati

## RIEPILOGO

**Linee modificate**: ~10 (3 in pulsante, ~45 in handler, 2 in traduzioni)  
**Linee aggiunte**: ~50 (handler open_new_event)  
**Pattern nuovi**: 0 (tutto riusato)  
**Regressioni potenziali**: 0 (RFQ wrappato in branch, VSM usa dialog esistente)  
**Complessità**: Bassa (routing semplice if/elif)  

## STATO ATTUALE

- ❌ Non implementato
- File pronti per modifiche
- Pattern già testati e funzionanti
- Zero dipendenze esterne da aggiungere
