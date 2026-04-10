# UI Dialog Standardization Log - 2026-04-10

## Step 1 - Avvio attività
- Obiettivo confermato: uniformare tutti i popup/dialog allo standard visivo del dialog reale `Choose Language`.
- Vincoli applicati: modifiche locali, no nuove dipendenze, no cambi logica business.

## Step 2 - Scansione iniziale codice
- Eseguita scansione globale con `rg` su pattern: `messagebox`, `Toplevel`, `SimpleMessageDialog`, `SimpleYesNoDialog`.
- Identificati principali file con dialog non standardizzati (uso diretto `messagebox`):
  - `ui/windows/edit_suppliers_window.py`
  - `ui/windows/edit_reference_window.py`
  - `ui/windows/view_request_window.py`
  - `dataflow.py` (2 occorrenze)
- Confermato che il riferimento stilistico è implementato in `ui/dialogs/common_dialogs.py` (`LanguagePrompt`, `SimpleMessageDialog`, `SimpleYesNoDialog`).

## Step 3 - Strategia operativa
- Introdurre API standard in `ui/dialogs/common_dialogs.py`:
  - `show_info(...)`
  - `show_success(...)`
  - `show_error(...)`
  - `show_warning(...)`
  - `show_confirm(...)`
  - `show_ok_cancel(...)` (per coprire `askokcancel` senza ricadere su messagebox)
- Implementazione wrappers sottili sopra dialog custom già esistenti, per conservare stile del `Choose Language`.
- Sostituzione locale e incrementale degli usi diretti `messagebox` nei file target.

## Step 4 - Centralizzazione API dialog
- Aggiornato `ui/dialogs/common_dialogs.py`:
  - aggiunto `SimpleOkCancelDialog` (stile coerente con `SimpleYesNoDialog`)
  - aggiunte API standard:
    - `show_info(...)`
    - `show_success(...)`
    - `show_error(...)`
    - `show_warning(...)`
    - `show_confirm(...)`
    - `show_ok_cancel(...)`
  - introdotto parser argomenti interno per compatibilità sia con chiamate legacy (`title, message, parent=...`) sia nuovo formato (`parent, title, message`).
  - rimosso uso diretto di `messagebox` anche dentro `UserIdentityDialog`.

## Step 5 - Migrazione chiamate non standard
- `ui/windows/edit_suppliers_window.py`
  - sostituiti `messagebox.showerror/showwarning/showinfo` con API standard (`show_error`, `show_warning`, `show_info`).
- `ui/windows/edit_reference_window.py`
  - sostituiti `messagebox.showerror/showinfo` con API standard (`show_error`, `show_info`).
- `ui/windows/view_request_window.py`
  - sostituite tutte le chiamate:
    - `messagebox.showerror` -> `show_error`
    - `messagebox.showwarning` -> `show_warning`
    - `messagebox.showinfo` -> `show_info`
    - `messagebox.askyesno` -> `show_confirm`
    - `messagebox.askokcancel` -> `show_ok_cancel`
  - centralizzato import da `ui.dialogs.common_dialogs` anche per `LanguagePrompt`.
- `dataflow.py`
  - sostituite le due `messagebox.showerror` residue con `show_error`.

## Step 6 - Bonifica import residui
- Rimossi import `messagebox` non usati da:
  - `ui/windows/purchase_order_window.py`
  - `ui/windows/sqdc_analysis_window.py`
  - `ui/windows/attachment_window.py`

## Step 7 - Verifiche
- Ricerca globale (`rg`) su `ui` + `dataflow.py`:
  - nessuna chiamata runtime a `messagebox.*` residua.
- Verifica sintassi moduli toccati:
  - comando: `python3 -m py_compile ...`
  - esito: OK (nessun errore).

## Step 8 - Note di compatibilità
- Le API standard in `common_dialogs.py` sono backward-compatible sulle signature per minimizzare regressioni.
- Comportamento modale preservato: i dialog custom usano `transient`, `grab_set`, `wait_window` come già in standard esistente.
## Step 9 - Mappatura Toplevel (contesto scope)
- Inventario classi `tk.Toplevel` rilevato in:
  - `dataflow.py` (`SettingsWindow`)
  - `ui/dialogs/common_dialogs.py` (dialog standard + prompt)
  - `ui/dialogs/*.py` (dialog gestionali dedicati)
  - `ui/windows/*.py` (finestre secondarie operative)
- Distinzione applicata:
  - popup/confirm/error/success -> standardizzati tramite `common_dialogs` API
  - finestre operative complete (es. `ViewRequestWindow`, `NotesWindow`, `AttachmentWindow`) mantenute inalterate nel layout principale come da vincoli.
