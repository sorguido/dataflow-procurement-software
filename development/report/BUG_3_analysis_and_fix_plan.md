# BUG 3 — Analisi tecnica e piano di fix

## 1) Root Cause Analysis

## 1.1 Distinzione esplicita: falso positivo iniziale vs bug reale

Il bug iniziale ipotizzato ("`New Supplier` salva anche con Cancel") **non spiega** il comportamento osservato in runtime.

Evidenza codice:
- In `ui/dialogs/vsm_event_dialog.py:235` il bind su `Return` chiama `_on_new_supplier_enter`.
- In `ui/dialogs/vsm_event_dialog.py:507-514` `_on_new_supplier_enter` scrive su DB solo su pressione Enter in edit mode.

Questa e' una semantica discutibile, ma non coincide con l'evidenza riportata (Cancel senza persistenza alla prima prova) e **non produce** `KeyError: 'popdown'` + `grab failed`.

Il bug reale osservato e' di **lifecycle/focus/grab UI**.

## 1.2 Root cause primaria (confermata da traceback + codice)

Il componente `SupplierNameSuggestController` ha una race tra callback differita e distruzione widget:

- `ui/components/supplier_name_suggest.py:107-114`
  - su `FocusOut` schedula `self.entry.after(120, self._hide_popup_if_focus_lost)`.
- `ui/components/supplier_name_suggest.py:50-57`
  - in `destroy()` **non cancella** `_hide_after_id` pendente.
- `ui/components/supplier_name_suggest.py:115-129`
  - `_hide_popup_if_focus_lost()` chiama `self.entry.focus_get()` senza guardie.

Quando il dialog viene chiuso (o quando il focus passa a widget interni ttk popdown), il callback puo' eseguire in uno stato non coerente e `focus_get()` puo' sollevare eccezioni (nel caso osservato: `KeyError: 'popdown'`).

## 1.3 Root cause secondaria (fragilita' modal/grab)

Entrambi i dialog coinvolti usano sequenza fragile:
- `grab_set()` **prima** di rendere la finestra visibile.

Evidenza:
- `ui/dialogs/potential_supplier_dialog.py:136-137`
- `ui/dialogs/vsm_event_dialog.py:135-136`

In condizioni normali puo' funzionare, ma con stato focus/event-loop gia' perturbato e' piu' probabile `TclError: grab failed: window not viewable`.

## 1.4 Relazione causale tra i due errori

Sequenza tecnica plausibile (in parte inferenziale, ma fortemente supportata dal codice):
1. `PotentialSupplierDialog` (Derisking) inizializza suggeritore (`ui/dialogs/potential_supplier_dialog.py:304-316`).
2. `FocusOut` sull'entry fornitore pianifica callback differita (`supplier_name_suggest.py:107-114`).
3. Il dialog si chiude (`protocol` su `destroy`, `potential_supplier_dialog.py:302`; destroy controller in `139-143`).
4. Callback differita resta pendente (non cancellata in `destroy()` controller).
5. Callback esegue `focus_get()` su stato focus non piu' valido -> `KeyError: 'popdown'` (traceback reale).
6. Aperture successive di form VSM (Derisking/Saving) entrano in sequenza `grab_set` su finestra non ancora viewable, e falliscono con `grab failed: window not viewable`.

Nota: lo step 6 e' inferenza tecnica coerente con stack/UI behavior, mentre lo step 5 e' confermato direttamente da traceback e codice.

---

## 2) Architettura attuale (semplificata)

## 2.1 Flusso apertura/chiusura dialog VSM / Derisking

- Apertura Derisking (supplier-based):
  - `dataflow.py:1437-1476` (`_on_supplier_sheet_double_click`) apre `PotentialSupplierDialog`.
  - `dataflow.py:2962-2971` (`open_new_event`, ramo `vsm_derisking`) apre `PotentialSupplierDialog`.
- Apertura Saving/Cost Avoidance event:
  - `dataflow.py:1825-1892` (`_edit_vsm_event`) apre `VSMEventDialog`.
  - `dataflow.py:2992-3000` (`open_new_event`, ramo VSM) apre `VSMEventDialog`.

In entrambi i dialog, modalita' impostata con `grab_set()` prima di `deiconify()`.

## 2.2 Ruolo `SupplierNameSuggestController`

- Usato in:
  - `PotentialSupplierDialog` (`ui/dialogs/potential_supplier_dialog.py:304-316`)
  - `EditSuppliersWindow` (`ui/windows/edit_suppliers_window.py:60-73`)
- Pattern:
  - binding su `FocusOut` (`supplier_name_suggest.py:48`)
  - callback differita via `after` (`107-114`)
  - cleanup incompleto (`50-57`: unbind ma nessun `after_cancel` finale)

## 2.3 Perche' il problema si estende anche a Saving

`Saving` non usa supplier suggestion direttamente, ma apre `VSMEventDialog` che ha stessa fragilita' grab/viewability (`vsm_event_dialog.py:135-136`).

Quindi:
- trigger primario nel percorso Derisking+suggestion,
- sintomo successivo anche su altri VSM dialog (Saving) per stato UI/modal non robusto.

---

## 3) Impatti

## 3.1 UX
- L'utente puo' aprire un dialog una volta e poi non riuscire piu' ad aprire form VSM.
- Messaggio bloccante: `Unable to open the form: grab failed: window not viewable`.
- Necessita' di riavvio applicazione per recuperare temporaneamente.

## 3.2 Stabilita' runtime
- Eccezioni non gestite in callback Tk (`KeyError: 'popdown'`) durante event loop.
- Possibile stato focus/grab inconsistente tra finestre modali.

## 3.3 Rischio regressione
- Alto sul modulo VSM (Derisking + Saving).
- Medio su altri punti che usano `SupplierNameSuggestController` (es. RFQ `EditSuppliersWindow`).

## 3.4 Estensione del problema
- Confermato: Derisking.
- Confermato osservazionalmente: Saving (fallimento apertura successiva).
- Probabile coinvolgimento indiretto: tutti i dialog con sequenza `grab_set` pre-visibility.

---

## 4) Piano di Fix (step-by-step, conservativo)

Obiettivo: fix minimo, locale, reversibile, senza refactor globale.

## Step 1 — Hardening lifecycle `SupplierNameSuggestController` (fix principale)

File target: `ui/components/supplier_name_suggest.py`

Interventi minimi:
1. In `destroy()` cancellare `_hide_after_id` se presente (`after_cancel`).
2. In `_hide_popup_if_focus_lost()` proteggere `focus_get()` e chiamate pointer (`winfo_pointerx/y`, `winfo_containing`) con guardie `try/except` per `tk.TclError`/`KeyError`.
3. In caso eccezione focus, fallback sicuro: hide popup e return (no propagate).
4. Opzionale ma consigliato: azzerare riferimenti interni su teardown (`_hide_after_id`, eventuale `_popup` non valido) per evitare callback su stato zombie.

Pro:
- colpisce direttamente la causa del traceback reale.
- impatto locale su un componente unico riusato.

Contro:
- richiede attenzione a non rompere UX suggerimenti (navigazione tastiera/mouse).

## Step 2 — Hardening sequenza modal/grab nei dialog coinvolti (mitigazione critica)

File target:
- `ui/dialogs/potential_supplier_dialog.py`
- `ui/dialogs/vsm_event_dialog.py`

Intervento:
- usare sequenza robusta: `deiconify()` -> `wait_visibility()` -> `grab_set()`.

Pro:
- riduce sensibilita' a race di mapping/focus cross-platform.
- allineato a pattern gia' usato in dialog stabili (`ui/dialogs/common_dialogs.py:69-71`, `559-562`).

Contro:
- modifica lieve comportamento timing di apertura (normalmente trasparente all'utente).

## Step 3 — Logging diagnostico minimo

Aggiungere logging non invasivo solo su:
- cleanup controller (cancel after id),
- eccezioni focus intercettate,
- eventuale fallback grab/viewability.

Pro:
- facilita verifica post-fix e diagnosi regressioni.

Contro:
- nessuno significativo (se livello debug/warning appropriato).

## Alternative (conservative)

1. Solo `try/except` in `_hide_popup_if_focus_lost`.
- Pro: patch minima.
- Contro: lascia callback pendenti non cancellate; mitigazione incompleta.

2. Solo riordino `grab_set` nei dialog.
- Pro: riduce `window not viewable`.
- Contro: non elimina `KeyError: 'popdown'` alla fonte.

Conclusione tecnica: serve combinazione Step 1 + Step 2.

---

## 5) Piano di Test Manuale

## Test 0 — Baseline (pre-fix)
- Riprodurre scenario noto e catturare log/traceback.

## Test 1 — Prima apertura + Cancel (Derisking)
1. Aprire tab Derisking.
2. Aprire `PotentialSupplierDialog`.
3. Interagire con campo Supplier (digitazione breve) e poi Cancel.

Atteso post-fix:
- nessun traceback `KeyError: 'popdown'`.
- dialog si chiude pulitamente.

## Test 2 — Aperture ripetute Derisking
1. Ripetere apertura/chiusura 10+ volte (doppio click riga e nuovo fornitore).

Atteso post-fix:
- nessun `grab failed: window not viewable`.
- nessun blocco progressivo.

## Test 3 — Aperture ripetute Saving
1. Aprire/chiudere `VSMEventDialog` su evento Saving 10+ volte.
2. Alternare con aperture Derisking.

Atteso post-fix:
- nessun errore apertura form.
- comportamento stabile anche dopo interazioni miste.

## Test 4 — Focus edge cases su popup suggerimenti
1. Con popup suggerimenti visibile, cliccare su combobox, poi fuori dialog, poi Cancel rapido.
2. Ripetere con tastiera (`Tab`, `Esc`, `Enter`).

Atteso post-fix:
- nessuna eccezione callback focus.
- popup si nasconde senza errori.

## Test 5 — Verifica non-regressione funzionale
1. Salvare un fornitore in Derisking.
2. Aprire un evento Saving e salvarlo.
3. Verificare che le normali operazioni restino invariate.

Atteso post-fix:
- nessuna regressione nei flussi CRUD standard.

## Test 6 — Linux/Windows smoke
- Ripetere Test 2-3 su entrambi gli OS target.

Atteso post-fix:
- comportamento coerente, nessun errore `window not viewable`.

---

## 6) Strategia di Rollback

Rollback semplice per commit locale del fix:

1. Ripristinare file toccati:
- `ui/components/supplier_name_suggest.py`
- `ui/dialogs/potential_supplier_dialog.py`
- `ui/dialogs/vsm_event_dialog.py`

2. Comando operativo:
- `git restore ui/components/supplier_name_suggest.py ui/dialogs/potential_supplier_dialog.py ui/dialogs/vsm_event_dialog.py`

3. Rieseguire Test 0 per verificare ritorno al comportamento precedente.

Nessuna migrazione DB, nessuna dipendenza nuova: rollback immediato e pulito.

---

## 7) Nota finale

Diagnosi finale:
- il bug reale dipende **principalmente** dal supplier suggestion popup (`SupplierNameSuggestController`) e dal suo lifecycle async (`after` + focus callback non robusta).
- il sintomo `grab failed: window not viewable` e' amplificato da una fragilita' separata ma correlata nella gestione modal/grab dei dialog VSM/Derisking.

Quindi la causa pratica e' **combinata**:
1. callback focus non resiliente + cleanup incompleto,
2. sequenza `grab_set` pre-visibility non robusta.

Questa combinazione spiega la dinamica osservata: prima apertura spesso ok, poi degradazione e fallimenti successivi fino al riavvio.
