# Plan: Add Save-time Email & Web Validation to Supplier Dialog

Introduce minimal, reversible format validation for the `E-mail` and `Web` fields in `PotentialSupplierDialog._on_save()`. No live/dynamic validation. Two regex functions in `validation_utils.py`, a few lines added to `_on_save`, and two string pairs added to the `.po` files + recompiled.

---

## 1. FILE DA TOCCARE

- `utils/validation_utils.py` — aggiungere `is_valid_email(value)` e `is_valid_website(value)`. Nessuna dipendenza esterna, solo `re` (già importato).
- `ui/dialogs/potential_supplier_dialog.py` — modificare `_on_save()`: inserire 2 controlli di validazione dopo il check esistente su `supplier_name` e prima della costruzione del model `PotentialSupplier`.
- `locale/en/LC_MESSAGES/dataflow.po` — aggiungere 2 nuove coppie msgid/msgstr (messaggio errore email + messaggio errore web).
- `locale/it/LC_MESSAGES/dataflow.po` — aggiungere le stesse 2 coppie (identity translation: msgstr = msgid).
- `locale/en/LC_MESSAGES/dataflow.mo` e `locale/it/LC_MESSAGES/dataflow.mo` — ricompilare da `.po` con `polib`.

---

## 2. PIANO A STEP

### STEP A — Validatori in `validation_utils.py`

**STEP A1 — `is_valid_email(value: str) → bool`**
- *Scopo*: validazione formato e-mail
- *File*: `utils/validation_utils.py`
- *Modifica*: aggiungere funzione con regex `^[^\s@]+@[a-zA-Z0-9][a-zA-Z0-9.\-]*\.[a-zA-Z]{2,}$`. Se `value` è vuota stringa, restituisce `True` (campo opzionale).
- *Rischio*: zero — funzione pura, non chiamata da nessuno finché non si fa lo STEP B
- *Rollback*: rimuovere la funzione

**STEP A2 — `is_valid_website(value: str) → bool`**
- *Scopo*: validazione formato web
- *File*: `utils/validation_utils.py`
- *Modifica*: aggiungere funzione con logica a 3 rami (vedi §3). Se `value` è vuota stringa, restituisce `True`.
- *Rischio*: zero — stessa motivazione
- *Rollback*: rimuovere la funzione

---

### STEP B — Locale strings

**STEP B1 — Aggiungere stringhe a `dataflow.po` (EN)**
- *Scopo*: testo messagebox errore email e web in inglese
- *File*: `locale/en/LC_MESSAGES/dataflow.po`
- *Modifica*: appendere al file due blocchi `msgid`/`msgstr`:
  - `"Formato e-mail non valido."` → `"Invalid email format."`
  - `"Formato URL web non valido."` → `"Invalid web URL format."`
- *Rischio*: minimo — append in coda al file, non tocca stringhe esistenti
- *Rollback*: rimuovere i 4 nuovi blocchi

**STEP B2 — Aggiungere stringhe a `dataflow.po` (IT)**
- *Scopo*: identity translation per italiano
- *File*: `locale/it/LC_MESSAGES/dataflow.po`
- *Modifica*: appendere le stesse 2 coppie con msgstr = msgid
- *Rischio*: minimo
- *Rollback*: rimuovere

**STEP B3 — Ricompilare i `.mo`**
- *Scopo*: rendere le nuove stringhe disponibili a runtime
- *File*: `locale/*/LC_MESSAGES/dataflow.mo`
- *Modifica*: script Python one-shot con `polib` per riscrivere i due `.mo` dai rispettivi `.po`
- *Rischio*: basso — i `.mo` precedenti si sovrascrivono; in caso di errore si può rigenerare da `.po` intatti
- *Rollback*: rieseguire la compilazione da `.po` originale (che si conserva sempre)

---

### STEP C — Hook in `_on_save()`

**STEP C1 — Importare i validatori**
- *Scopo*: rendere disponibili le funzioni nel modulo dialog
- *File*: `ui/dialogs/potential_supplier_dialog.py`
- *Modifica*: aggiungere `from utils.validation_utils import is_valid_email, is_valid_website` in testa agli import
- *Rischio*: zero
- *Rollback*: rimuovere la riga import

**STEP C2 — Aggiungere i check in `_on_save()`**
- *Scopo*: bloccare il salvataggio se email o web sono malformate
- *File*: `ui/dialogs/potential_supplier_dialog.py`
- *Modifica*: dopo il blocco di validazione `supplier_name` (riga ~336) e prima di `self._entry_username.configure(state="normal")`, inserire:
  ```python
  email_value = self.var_email.get().strip()
  if not is_valid_email(email_value):
      SimpleMessageDialog(self, _("Validazione"), _("Formato e-mail non valido."), "error")
      self._entry_email.focus_set()
      return

  web_value = self.var_website.get().strip()
  if not is_valid_website(web_value):
      SimpleMessageDialog(self, _("Validazione"), _("Formato URL web non valido."), "error")
      self._entry_website.focus_set()
      return
  ```
- *Rischio*: basso — pattern identico a quello già usato per `supplier_name`; si usa `SimpleMessageDialog` (già importato), i widget `self._entry_email` e `self._entry_website` sono già referenziati nel `_apply_read_only`
- *Rollback*: rimuovere i due blocchi if

---

## 3. LOGICA DI VALIDAZIONE PROPOSTA

**Email:**
- Regex: `^[^\s@]+@[a-zA-Z0-9][a-zA-Z0-9.\-]*\.[a-zA-Z]{2,}$`
- Copre tutti i casi richiesti: esclude spazi (`[^\s@]`), esclude `@` nel local part, forza il dominio a iniziare da alfanumerico (esclude `@.it`), forza TLD ≥ 2 alpha (esclude `@mail`)
- `guido@@mail.it`: il secondo `@` fa sì che la parte dominio sia `@mail.it`, che fallisce `[a-zA-Z0-9]` iniziale → rifiutato ✓

**Web:**
Logica a 3 rami mutuamente esclusivi, tutti required: no spazi nell'intero valore.

| Ramo | Condizione | Regex |
|---|---|---|
| Con schema | inizia con `http://` o `https://` | `^https?://[a-zA-Z0-9][^\s]*\.[a-zA-Z]{2,}` |
| Prefisso www. | inizia con `www.` | `^www\.[a-zA-Z0-9][^\s]*\.[a-zA-Z]{2,}$` |
| Dominio semplice | nessuno dei precedenti | `^[a-zA-Z0-9][a-zA-Z0-9\-]*(\.[a-zA-Z0-9][a-zA-Z0-9\-]*)*\.[a-zA-Z]{2,}$` |

- `https:||guido.it`: schema `https:` trovato ma poi `||` → fallisce il ramo schema (richiede `//`) ✓
- `http//guido.it`: no match `:` dopo schema → regex fallisce ✓
- `guido.` → TLD vuota → fallisce `[a-zA-Z]{2,}` ✓
- `guido` → nessun punto → fallisce ✓

**Campo vuoto → valido:** entrambe le funzioni restituiscono `True` immediatamente se il valore strippato è `""`.

**Ordine dei controlli al Save:**
1. `supplier_name` (già esistente)
2. `email`
3. `website`

Il primo errore blocca e mostra messagebox; gli altri non vengono controllati nella stessa chiamata.

**Messagebox:** `SimpleMessageDialog(self, _("Validazione"), _("messaggio"), "error")` — modale, blocca finché l'utente preme OK, poi il focus va sul campo errato. Identico al pattern già in uso.

---

## 4. IMPATTO UX E COMPATIBILITÀ

- **Linux/Windows**: `SimpleMessageDialog` è Tkinter puro, nessuna dipendenza OS. `re` è stdlib. Comportamento identico su entrambe le piattaforme.
- **Tkinter/ttk**: gli entry `_entry_email` e `_entry_website` sono `ttk.Entry`; `.focus_set()` funziona correttamente su entrambi.
- **Rischio regressione basso**: le modifiche più delicate sono solo 2 blocchi `if` in `_on_save` e 2 nuove funzioni pure in un modulo di utilità. Non tocca nessun'altra finestra, nessuna logica di persistenza, nessun modello dati.
- **Modalità read_only**: in `read_only`, il pulsante Salva non è presente (solo "Chiudi") → i nuovi check non vengono mai raggiunti ✓

---

## 5. CRITICITÀ / DUBBI DA VERIFICARE

1. **`ww3.guido.it` vs `sub.guido.it`**: con il ramo "dominio semplice" entrambi risultano validi (`ww3` è un'etichetta alfanumerica legittima). Per rifiutare `ww3.guido.it` servirebbe una regola esplicita come `if value.startswith("ww") and not value.startswith("www.")`. Valutare se il requisito è davvero stringente su questo o se è solo un esempio indicativo.

2. **Modalità EDIT con dati preesistenti**: se un record ha un'email già salvata in formato old/invalido, aprire in EDIT e premere Salva senza modificarla triggererebbe ora la validazione. Valutare se questo è il comportamento desiderato o se si preferisce saltare la validazione quando il valore non viene modificato (richiederebbe un confronto con il valore originale caricato).
