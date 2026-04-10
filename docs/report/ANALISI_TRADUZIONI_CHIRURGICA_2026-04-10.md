# Analisi Architetturale Sistema Traduzioni (DataFlow)

Data: 10 aprile 2026  
Scope: codebase reale del branch corrente, senza modifiche funzionali al runtime.

---

## A. RICOSTRUZIONE DEL SISTEMA ATTUALE

### A1. Dove e come viene inizializzato gettext

1. Inizializzazione anticipata a import-time in `dataflow.py`:
   - `init_i18n()` chiamata prima degli import UI (`dataflow.py:91`), con commento esplicito sulla dipendenza dall’ordine.
2. Inizializzazione ripetuta anche nel bootstrap runtime:
   - `init_i18n()` richiamata nel blocco `if __name__ == '__main__'` (`dataflow.py:4328`).
3. Implementazione in `utils/i18n_utils.py`:
   - lettura lingua da `config.ini` (`en`/`it`, fallback `en`) (`utils/i18n_utils.py:48-56`),
   - caricamento `.mo` da `locale/<lang>/LC_MESSAGES/dataflow.mo` (`utils/i18n_utils.py:70`),
   - installazione traduttore in `builtins._` via `gettext.translation(...).install()` o `NullTranslations().install()`.

Nota: esiste fallback esplicito `builtins._ = lambda x: x` in `dataflow.py:51-53` e in `ui/kpi_window.py:40-41`.

### A2. Dove e come viene bindata/usata la funzione `_()`

Pattern principali:

1. Pattern corretto (wrapper dinamico):
   - `from utils.i18n_utils import _` in molti moduli.
   - Il wrapper `_` inoltra dinamicamente a `builtins._` (`utils/i18n_utils.py:19-27`).
2. Pattern implicito (fragile):
   - moduli che usano `_()` senza import locale di `_`:
     - `ui/kpi_window.py`
     - `ui/windows/view_request_window.py`
   - Questi moduli dipendono implicitamente da `builtins._` globale.

### A3. Moduli/finestre che dipendono implicitamente dal contesto traduzione

Dipendenza forte da stato globale/traduttore o da configurazione lingua:

1. `ui/kpi_window.py`
   - usa `_()` senza import,
   - ha traduzione a class-load (`_PERIOD_OPTIONS`, vedi fragilità sotto).
2. `ui/windows/view_request_window.py`
   - usa `_()` senza import diretto,
   - mixa `_()` + branch su `get_current_language()`.
3. `ui/windows/purchase_order_window.py`
   - non usa gettext per molte stringhe; usa branch `if get_current_language() == 'it'`.
4. `ui/components/main_dashboard_toolbar.py`
   - placeholder hardcoded inglese (`PLACEHOLDER_TEXT`) fuori gettext.
5. `ui/dialogs/common_dialogs.py`
   - alcune stringhe hardcoded (`OK`, `Preparazione...`, title default copy progress).
6. `services/dashboard_controller.py` e `dataflow.py`
   - confronti logici su stringhe tradotte (`_("Tutte")`, `_("Sì")`, `_(event.action)`).

### A4. Italiano base + inglese tradotto o situazione mista?

Correzione esplicita ipotesi: **la situazione è mista**.

Evidenze:

1. Valori canonici business nel DB sono in italiano (es. `normalize_rfq_type()` usa canonici IT) (`utils/i18n_utils.py:144+`).
2. I cataloghi `.po` non sono a base unica coerente:
   - `locale/en/LC_MESSAGES/dataflow.po`: 876 entry, 0 `msgstr` vuoti, ma 78 `msgstr == msgid` (molte non tradotte).
   - `locale/it/LC_MESSAGES/dataflow.po`: 862 entry, 6 `msgstr` vuoti.
   - differenza set `msgid`: 29 presenti in EN e non in IT, 15 in IT e non in EN.
3. Nel codice coesistono:
   - msgid italiani,
   - msgid inglesi,
   - stringhe hardcoded IT/EN fuori gettext,
   - mapping manuali bilingua.

Quindi non è “solo italiano base + inglese gettext”: è un ecosistema ibrido.

### A5. Punti fragili attuali (sintesi)

1. Dipendenza da `builtins._` in moduli che non importano `_` localmente.
2. Traduzioni valutate troppo presto (class attribute in KPI).
3. Confronti logici basati su stringhe tradotte.
4. Branch manuali `get_current_language()` con stringhe hardcoded.
5. Copertura cataloghi incompleta/disallineata rispetto al codice.
6. Tooling traduzioni fragile (`dev_tools/compile_translations.py` punta a `dev_tools/locale`, path inesistente).

---

## B. ANALISI DELLE CAUSE DI FRAGILITÀ

### B1. Stringhe tradotte a import-time

Presente:

1. `ui/kpi_window.py:145`
   - `KpiWindow._PERIOD_OPTIONS = [..., _("All")]`
   - viene valutato al class load, quindi dipende dall’ordine d’import e dallo stato traduttore già installato.

### B2. Costanti globali o class attributes con `_()`

Presente:

1. `ui/kpi_window.py:145` (`_PERIOD_OPTIONS`).

### B3. Menu/dizionari/liste con stringhe localizzate troppo presto

Parzialmente presente:

1. `ui/kpi_window.py` (class attr `_PERIOD_OPTIONS`).
2. `dataflow.py` / `services/dashboard_controller.py`: filtri con valori tradotti usati come chiavi logiche (`_("Tutte")`, `_("Sì")`).
3. `ui/main_dashboard_builder.py`: valori combobox localizzati e poi confrontati altrove.

### B4. Popup / messagebox / Toplevel / finestre secondarie senza accesso stabile al traduttore

Presente in più forme:

1. Moduli con `_()` senza import locale (dipendenza implicita da builtins):
   - `ui/kpi_window.py`, `ui/windows/view_request_window.py`.
2. Finestre secondarie con branch manuali su lingua (`get_current_language`) e stringhe raw:
   - `ui/windows/purchase_order_window.py`.
3. Message/dialog raw non gettext:
   - `ui/windows/attachment_window.py:327` (`"Errore Database"...`),
   - `ui/dialogs/common_dialogs.py` (`OK`, `Preparazione...`, title default copy).

### B5. Casi hard coded IT che bypassano gettext

Esempi concreti:

1. `ui/windows/purchase_order_window.py`
   - titoli, label, pulsanti, confirm dialog e warning gestiti con `if get_current_language() == 'it'` e stringhe letterali IT/EN.
2. `ui/windows/attachment_window.py:327`
   - `SimpleMessageDialog(self, "Errore Database", f"Impossibile ...", "error")`.
3. `ui/dialogs/common_dialogs.py`
   - `OK`, `Preparazione...`, `title="Copia in corso..."` non passano da gettext.
4. `ui/components/main_dashboard_toolbar.py`
   - `PLACEHOLDER_TEXT = "Search anything..."` hardcoded inglese.

### B6. Pattern incoerenti tra moduli

Molto presenti:

1. Pattern gettext wrapper (`_`) vs branch manuale lingua.
2. Pattern msgid IT vs msgid EN.
3. Pattern runtime gettext vs template-language selection manuale (`LanguagePrompt` + dizionari in `view_request_window.py`).
4. Pattern logica business su stringhe localizzate (`_(event.action)`, `_("Tutte")`).
5. Pattern toolchain/documentazione non allineati alla realtà cataloghi.

---

## C. MAPPATURA DEL RISCHIO

| Problema | Gravità | Prob. regressione | Estensione | File coinvolti (principali) | Rischio UX | Rischio manutentivo |
|---|---|---|---|---|---|---|
| `_()` usato senza import locale, dipendenza da `builtins._` | Alta | Media | Media | `ui/kpi_window.py`, `ui/windows/view_request_window.py` | Finestre con lingua errata o fallback inatteso | Alto (dipendenza implicita globale) |
| Traduzione a import-time (`_PERIOD_OPTIONS`) | Media | Media | Localizzata | `ui/kpi_window.py:145` | Incoerenza lingua KPI / filtri | Medio |
| Branch manuali lingua + hardcoded (IT/EN) | Alta | Alta | Localizzata ma critica (finestre secondarie) | `ui/windows/purchase_order_window.py` | Popup/azioni in lingua diversa dal resto app | Alto |
| Stringhe hardcoded fuori gettext in dialog comuni | Media | Alta | Trasversale | `ui/dialogs/common_dialogs.py`, `ui/windows/attachment_window.py` | Stringhe “fuori tono” rispetto lingua app | Medio/Alto |
| Confronti logici su stringhe tradotte | Alta | Media/Alta | Trasversale VSM/RFQ | `dataflow.py`, `services/dashboard_controller.py`, `ui/kpi_window.py`, `ui/dialogs/vsm_event_dialog.py` | Filtri/comportamenti non coerenti se catalogo cambia | Alto |
| Copertura cataloghi incompleta/disallineata (130 stringhe `_()` mancanti in EN/IT `.po`) | Alta | Alta | Ampia | soprattutto `dataflow.py`, `ui/kpi_window.py`, `ui/windows/*` | In EN compaiono testi IT (fallback msgid) | Alto |
| Entry EN non tradotte (`msgstr==msgid`, 78 casi) | Media | Media | Ampia | `locale/en/LC_MESSAGES/dataflow.po` | Aree in italiano/inglese miste | Medio |
| Tooling compilazione fragile (`compile_translations.py` path errato) | Media | Alta | Dev workflow | `dev_tools/compile_translations.py:12` | `.mo` potenzialmente non aggiornati | Alto |

---

## D. OPZIONI ARCHITETTURALI

### Opzione 1. Hardening dell’attuale gettext/.po/.mo

Vantaggi:

1. Minima invasività.
2. Nessuna nuova dipendenza.
3. Reversibile rapidamente.
4. Compatibile con branch corrente e struttura attuale.

Svantaggi:

1. Rimane il modello ibrido storico (msgid misti IT/EN) se non si disciplina il processo.
2. Serve rigore operativo costante (estrazione/compilazione/review).

Complessità: Bassa/Media.  
Rischio regressioni: Basso/Medio (se step piccoli).  
Impatto file: mirato su finestre critiche + cataloghi + script dev.  
Compatibilità codebase: Alta.  
Sostenibilità futura: Buona se si introducono regole e controlli automatici.

### Opzione 2. Translation service centralizzato (i18n.py/translation_service.py) mantenendo .po/.mo

Vantaggi:

1. Punto unico di accesso alla traduzione.
2. Riduce dipendenze implicite da `builtins._`.
3. Facilita enforcement di policy (no import-time translation, fallback coerenti).

Svantaggi:

1. Richiede toccare più moduli per convergere sull’API.
2. Se spinto troppo presto diventa refactor trasversale non chirurgico.

Complessità: Media.  
Rischio regressioni: Medio.  
Impatto file: medio/ampio (import e chiamate).  
Compatibilità: Alta, se introdotto gradualmente.  
Sostenibilità futura: Molto buona.

### Opzione 3. Refactor strutturato UI/traduzioni

Vantaggi:

1. Coerenza architetturale elevata (se completato).
2. Debito tecnico i18n quasi azzerato.

Svantaggi:

1. Alto impatto e alto rischio su app Tkinter grande.
2. Non allineato al vincolo “basso rischio/reversibile/no overengineering”.

Complessità: Alta.  
Rischio regressioni: Alto.  
Impatto file: ampio.  
Compatibilità: Media (serve migrazione accurata).  
Sostenibilità futura: Alta, ma costo iniziale elevato.

---

## E. RACCOMANDAZIONE

Raccomandazione: **Opzione 1 (hardening gettext attuale), con micro-centralizzazione dentro `utils/i18n_utils.py` senza nuovo sottosistema**.

Perché:

1. Massimizza stabilità e reversibilità.
2. Riduce fragilità reali subito (hardcoded, import-time, cataloghi mancanti).
3. Evita refactor globale.
4. È compatibile con piccoli step verificabili.

In pratica: usare l’infrastruttura già esistente (`i18n_utils`) come “service de-facto”, senza introdurre un framework nuovo.

---

## F. PIANO OPERATIVO CHIRURGICO

### A

Obiettivo: baseline e coerenza cataloghi.

File toccati: nessuno (analisi), poi `locale/*/LC_MESSAGES/dataflow.po`.
Rischio: basso.
Rollback: nessuno.
Verifica: snapshot metriche (`missing msgid`, `msgstr==msgid`, diff set msgid EN/IT).

#### A1

Obiettivo: correggere tooling compilazione traduzioni.

File: `dev_tools/compile_translations.py`.
Rischio: basso.
Rollback: revert file.
Verifica: script trova `locale/en` e `locale/it`, compila `.mo` senza warning path.

#### A2

Obiettivo: riallineare `.po/.mo` ai testi realmente usati.

File: `locale/en/LC_MESSAGES/dataflow.po`, `locale/it/LC_MESSAGES/dataflow.po`, `.mo` correlati.
Rischio: medio (solo UX testi).
Rollback: ripristino `.po/.mo` precedenti.
Verifica: smoke test EN/IT su finestre principali + secondarie.

### B

Obiettivo: eliminare dipendenze implicite dal contesto globale.

File: `ui/kpi_window.py`, `ui/windows/view_request_window.py`.
Rischio: medio-basso.
Rollback: revert mirato.
Verifica: apertura finestre in EN/IT senza NameError e testi coerenti.

#### B1

Obiettivo: import esplicito `_` da `utils.i18n_utils` nei moduli che oggi lo usano implicitamente.

File: `ui/kpi_window.py`, `ui/windows/view_request_window.py`.
Rischio: basso.
Rollback: revert import.
Verifica: apertura KPI e Control Panel con traduzioni corrette.

#### B2

Obiettivo: rimuovere `_()` in class attribute a import-time.

File: `ui/kpi_window.py` (`_PERIOD_OPTIONS`).
Rischio: basso.
Rollback: ripristino attributo statico.
Verifica: cambio lingua + riavvio, KPI period labels corretti.

### C

Obiettivo: ridurre hardcoded in finestre secondarie ad alta incidenza bug UX.

File: priorità `ui/windows/purchase_order_window.py`.
Rischio: medio (finestra operativa).
Rollback: revert file singolo.
Verifica: flusso PO completo in EN/IT (add/delete/save/dialog).

#### C1

Obiettivo: sostituire branch manuali IT/EN con `_()` coerente.

File: `ui/windows/purchase_order_window.py`.
Rischio: medio.
Rollback: revert.
Verifica: tutti i label/dialog PO localizzati senza if lingua.

#### C2

Obiettivo: ripulire stringhe raw residue in dialog comuni.

File: `ui/dialogs/common_dialogs.py`, `ui/windows/attachment_window.py`.
Rischio: basso.
Rollback: revert parziale.
Verifica: nessun testo hardcoded nei popup base.

### D

Obiettivo: hardening logica applicativa indipendente da testi tradotti.

File: `dataflow.py`, `services/dashboard_controller.py`, `ui/dialogs/vsm_event_dialog.py`, `ui/kpi_window.py`.
Rischio: medio.
Rollback: revert per blocchi.
Verifica: filtri e comportamenti invarianti anche se cambia una traduzione in `.po`.

#### D1

Obiettivo: rimuovere confronti tipo `if _(event.action) == ...` e usare valori canonici non localizzati.

Rischio: medio.
Rollback: revert blocchi filtro.
Verifica: filtri VSM corretti in EN/IT con stessi risultati.

#### D2

Obiettivo: eliminare euristiche lingua basate su traduzione (`_("All") != "All"`).

Rischio: basso/medio.
Rollback: revert.
Verifica: grafici/label KPI coerenti con `get_current_language()`.

### E

Obiettivo: guardrail permanenti anti-regressione.

File: script/checklist docs (es. `docs/` e/o `dev_tools/`).
Rischio: basso.
Rollback: rimuovere script.
Verifica: pre-release check ripetibile.

#### E1

Obiettivo: checklist automatica grep/AST per:
- `_()` senza import,
- `_()` a import-time,
- stringhe UI hardcoded fuori gettext,
- `_()` literal non presenti nei `.po`.

#### E2

Obiettivo: smoke test manuale standard EN/IT su finestre secondarie.

---

## G. QUICK WINS (alto beneficio / basso rischio)

1. Fix immediato `dev_tools/compile_translations.py` path locale.
2. Import esplicito `_` in `ui/windows/view_request_window.py` e `ui/kpi_window.py`.
3. Spostare `_PERIOD_OPTIONS` da class attribute a costruzione runtime.
4. Normalizzare `purchase_order_window` a `_()` (eliminare branch lingua ripetuti).
5. Tradurre stringhe raw residue (`OK`, `Preparazione...`, `Errore Database` raw).
6. Eliminare confronto lingua via `_("All") != "All"` in KPI.
7. Prima passata cataloghi: coprire le stringhe `_()` mancanti (scan attuale: 130 non presenti in EN/IT `.po`).

---

## H. REGOLE FUTURE DI SVILUPPO (mini standard)

1. **Dove tradurre**: tutte le stringhe UI visibili all’utente passano da `_()` o helper centralizzato in `i18n_utils`.
2. **Dove NON tradurre**: valori canonici business/db (es. enum persistiti) restano non localizzati; si localizza solo in presentazione.
3. **Popup e finestre figlie**: ogni modulo UI deve importare `_` esplicitamente da `utils.i18n_utils`.
4. **No `_()` a import-time**: vietato in global/class attribute; localizzare in `__init__`/builder runtime.
5. **No logica su stringhe tradotte**: confrontare codici canonici, non label localizzate.
6. **Niente branch lingua hardcoded** (`if get_current_language() == 'it': ...` per testi UI) salvo casi documentati e temporanei.
7. **Cataloghi allineati**: ogni nuova `_('...')` deve avere entry in EN/IT `.po` prima del rilascio.
8. **Build traduzioni obbligatoria**: compilare `.mo` in pipeline/dev step verificato.

---

## I. OUTPUT FINALE

### I1. Diagnosi sintetica

Il problema non è un solo bug puntuale ma un insieme di fragilità strutturali: dipendenza da contesto globale (`builtins._`), pattern misti (gettext + hardcoded + branch manuali), confronti logici su stringhe tradotte e cataloghi `.po` non perfettamente allineati al codice.

### I2. Soluzione raccomandata

Hardening progressivo dell’approccio attuale gettext/.po/.mo (Opzione 1), usando `utils/i18n_utils` come punto unico de-facto e correggendo i pattern fragili più impattanti.

### I3. Piano minimo sicuro

1. Fix tooling compilazione traduzioni.
2. Import esplicito `_` nei moduli impliciti + rimozione `_()` a import-time.
3. Patch mirata finestra PO.
4. Allineamento `.po/.mo` su stringhe realmente usate.
5. Guardrail automatici + smoke test EN/IT.

### I4. Aree da NON toccare subito

1. Refactor globale di tutta la UI.
2. Migrazione massiva dei msgid storici (IT↔EN) in un’unica release.
3. Ristrutturazione del modello dati canonico DB.

### I5. Rischi residui

1. Finché i cataloghi restano ibridi, può persistere qualche testo non uniforme.
2. Le finestre meno usate potrebbero contenere ancora stringhe raw non emerse nei smoke test iniziali.
3. Senza guardrail CI/dev, regressioni i18n tenderanno a rientrare nel tempo.

---

# APPENDICE OPERATIVA

## 1) File prioritari da ispezionare per primi

1. `dataflow.py` (alto numero stringhe `_()` non presenti in `.po` e logica filtro su testi localizzati).
2. `ui/kpi_window.py` (dipendenza implicita da builtins, import-time translation, euristica lingua fragile).
3. `ui/windows/view_request_window.py` (usa `_` senza import, mix branch lingua + gettext).
4. `ui/windows/purchase_order_window.py` (hardcoded massivo IT/EN).
5. `ui/windows/sqdc_analysis_window.py` (diversi msgid mancanti e branch lingua template).
6. `ui/windows/attachment_window.py` (raw dialog non gettext).
7. `ui/dialogs/common_dialogs.py` (stringhe raw condivise da più flussi).
8. `services/dashboard_controller.py` (logica su `_("Tutte")`).
9. `utils/i18n_utils.py` (nodo centrale runtime).
10. `dev_tools/compile_translations.py` (tooling path fragile).
11. `locale/en/LC_MESSAGES/dataflow.po`.
12. `locale/it/LC_MESSAGES/dataflow.po`.

## 2) Elenco grep/ripgrep pattern da cercare

```bash
# Uso di _ senza import esplicito da i18n_utils
rg -n "\b_\(" ui dataflow.py services | cat
rg -n "from utils\.i18n_utils import .*_" ui dataflow.py services

# Traduzioni valutate troppo presto (module/class scope)
# (consigliato script AST dedicato)

# Branch lingua manuali (potenziale bypass gettext)
rg -n "get_current_language\(|==\s*'it'|==\s*\"it\"|==\s*'en'|==\s*\"en\"" ui dataflow.py services

# Dialog/messagebox hardcoded
rg -n "SimpleMessageDialog\(|messagebox\.(show|ask)" ui dataflow.py services

# Confronti logici su stringhe tradotte
rg -n "==\s*_\(|!=\s*_\(|_\(.*\)\s*==|_\(.*\)\s*!=" dataflow.py ui services

# Stringhe interpolate dentro _ (msgid dinamici)
rg -n "_\(f\"|_\(f'" ui dataflow.py services

# Stringhe _() usate nel codice ma non presenti in .po
# (consigliato script AST+PO diff)
```

## 3) Check-list review manuale

1. Avvio app in EN: dashboard, toolbar, filtri, popup base.
2. Apertura finestre secondarie: ViewRequest, PurchaseOrder, Attachment, SQDC, VSM dialog.
3. Verifica messagebox errore/warning/success in ogni finestra.
4. Verifica label dinamiche (bottoni “Aggiungi/Modifica”, stato read-only).
5. Verifica filtri che usano valori tradotti (`Tutte`, `Sì/No`, azioni VSM).
6. Verifica export flow con `LanguagePrompt` e testi file dialog.
7. Ripetere stesso percorso in IT.
8. Cambiare lingua in settings e verificare comportamento prima/dopo riavvio.
9. Ricompilare `.mo` e rieseguire smoke test.
10. Confermare assenza regressioni su DB canonico (tipo RdO, azione, status).

## 4) Proposta patch pilota MINIMA (1 flusso / 1 finestra)

Flusso pilota consigliato: **`ui/windows/purchase_order_window.py`**

Motivo:

1. Alta concentrazione di hardcoded IT/EN.
2. Finestra isolata, basso impatto architetturale.
3. Beneficio UX immediato (dialog e pulsanti coerenti con lingua sessione).

Patch minima proposta:

1. Sostituire stringhe in branch `if get_current_language()` con `_()` unificato.
2. Eliminare duplicazione IT/EN nei dialog (`Campo obbligatorio`, `No Selection`, confirm delete, errori indice).
3. Mantenere invariata logica business e DB.

File toccati:

1. `ui/windows/purchase_order_window.py`
2. `locale/en/LC_MESSAGES/dataflow.po` (solo nuove chiavi mancanti)
3. `locale/it/LC_MESSAGES/dataflow.po` (allineamento)
4. `.mo` ricompilati

Rischio: basso/medio (solo testo UI).  
Rollback: revert file finestra + `.po/.mo`.  
Verifica: test manuale add/delete/save PO in EN e IT, incluso annulla/chiudi e dialog error.

---

## Nota conclusiva operativa

L’ipotesi iniziale “italiano base hardcoded + inglese via gettext” è **parzialmente corretta** ma incompleta: il codice reale è oggi **ibrido** (gettext + hardcoded + branch lingua + msgid misti IT/EN).  
La strategia più sicura è hardening incrementale, non rivoluzione.
