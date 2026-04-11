# DATAFLOW - Implementazione Feature RFQ PDF (2026-04-11)

## A. FILE TOCCATI
- `ui/windows/view_request_window.py`
  - Inserito pulsante toolbar `RFQ PDF` tra `Export Excel` e `Create SQDC Analysis`.
  - Aggiunto wiring minimo `open_rfq_pdf_export_dialog()` con import lazy del dialog e gestione errore dipendenza.
- `ui/dialogs/rfq_pdf_export_dialog.py` (nuovo)
  - Nuovo dialog modale coerente con stile DataFlow per gestione logo persistente e conferma export PDF.
- `services/rfq_pdf_logo_service.py` (nuovo)
  - Servizio dedicato per validazione/copia/rimozione/logo status persistente in cartella dati utente.
- `services/rfq_pdf_export_service.py` (nuovo)
  - Servizio dedicato generazione PDF A4 multipagina con ReportLab (header, testo fisso, tabella dinamica, wrapping).
- `locale/it/LC_MESSAGES/dataflow.po`
  - Aggiunte nuove stringhe i18n feature RFQ PDF (dialog + PDF content + warning import).
- `locale/en/LC_MESSAGES/dataflow.po`
  - Aggiunte traduzioni EN corrispondenti.
- `locale/it/LC_MESSAGES/dataflow.mo`
  - Rigenerato da `.po`.
- `locale/en/LC_MESSAGES/dataflow.mo`
  - Rigenerato da `.po`.
- `requirements.txt`
  - Aggiunta dipendenza `reportlab==4.2.2`.
- `dataflow.spec`
  - Packaging hardening: aggiunto `collect_data_files("reportlab")`.
- `dataflow_appimage.spec`
  - Packaging hardening: aggiunto `collect_data_files('reportlab')`.

## B. SCELTE ARCHITETTURALI
- Dialog UI in `ui/dialogs/rfq_pdf_export_dialog.py`
  - Coerente con struttura progetto (dialog separati da finestre), evita di caricare logica UI in `ViewRequestWindow`.
- Logica PDF in `services/rfq_pdf_export_service.py`
  - Collocata in `services` come logica applicativa riusabile/testabile, preservando pulizia della Request Window.
- Persistenza logo in `services/rfq_pdf_logo_service.py`
  - Isolata in modulo specifico per responsabilita singola (validazione immagine + storage + config).
  - Riutilizza pattern esistente `config.ini` via `utils.user_utils.get_config_file` e cartelle utente via `services.app_paths`.
- Wiring in `ViewRequestWindow` ridotto al minimo
  - Solo pulsante + apertura dialog (import lazy), nessun refactor della logica RFQ esistente.

## C. FLUSSO UTENTE FINALE
1. Utente apre Request Window RFQ.
2. In toolbar vede nuovo pulsante `🧾 RFQ PDF` tra `📊 Esporta Excel` e pulsante SQDC.
3. Click su `RFQ PDF` apre dialog modale nativo DataFlow.
4. Nel dialog:
   - visualizza stato logo corrente,
   - puo selezionare/sostituire logo,
   - puo rimuovere logo,
   - puo confermare export o annullare.
5. Click `Conferma Export PDF` apre `Save As` PDF.
6. Il servizio genera PDF A4 multipagina con logo opzionale e dati RFQ.
7. Se il logo configurato e invalido/mancante, export continua senza logo con warning chiaro.

## D. DETTAGLI IMPLEMENTATIVI CHIAVE
- Gestione logo
  - Formati accettati: `.png`, `.jpg`, `.jpeg`.
  - Validazione robusta con Pillow (`verify`, dimensioni minime, size max 8MB).
  - Copia in cartella interna utente DataFlow: `.../Assets/RFQ_PDF/company_logo.<ext>`.
  - Persistenza riferimento in `config.ini` (`[Settings] rfq_pdf_logo_file`).
  - Rimozione pulita sia file sia config key.
- Gestione traduzioni
  - Tutte le nuove stringhe UI passano da `tr(...)`.
  - Niente `_()` nei nuovi moduli applicativi.
  - Aggiunte entry `.po` EN/IT e ricompilazione `.mo`.
- Layout PDF
  - ReportLab `SimpleDocTemplate`, formato A4, margini 2 cm.
  - Header con logo (opzionale) in alto a sinistra + titolo/metadati a destra.
  - Metadati header: numero RFQ, issue date, expiry date.
- Gestione wrapping
  - Celle tabella testuali rese con `Paragraph` (Attachment, Description, Raw Attachment, ecc.).
  - Aspect ratio logo preservato (`drawWidth/drawHeight` con ratio), no deformazioni.
- Multipage
  - Tabella con `repeatRows=1` per ripetizione automatica header su ogni pagina.
  - Righe crescono verticalmente in automatico in base al contenuto wrapped.
- Fallback logo assente/non valido
  - Export non si interrompe.
  - Warning esplicito all’utente e PDF generato comunque senza logo.

## E. RISCHI / NOTE
- Nel mio ambiente locale `reportlab` non e installato: l’import del servizio PDF fallisce runtime qui.
  - Mitigazione implementata: import lazy + messaggio utente chiaro (`Funzionalita RFQ PDF non disponibile...`).
  - In ambiente target, installando dipendenza (`requirements.txt`) la feature e operativa.
- Packaging PyInstaller
  - Aggiunto `collect_data_files('reportlab')` in entrambe le spec per evitare missing resources in build.
- Casi limite da testare manualmente
  - Logo molto panoramico / molto verticale.
  - RFQ con descrizioni/allegati molto lunghi su molte pagine.
  - RFQ senza righe materiali.

## F. ROLLBACK
Rollback semplice e reversibile:
1. Rimuovere nuovi file:
   - `services/rfq_pdf_export_service.py`
   - `services/rfq_pdf_logo_service.py`
   - `ui/dialogs/rfq_pdf_export_dialog.py`
2. Ripristinare `ui/windows/view_request_window.py` togliendo:
   - pulsante `🧾 RFQ PDF`
   - metodo `open_rfq_pdf_export_dialog`.
3. Ripristinare `requirements.txt` rimuovendo `reportlab`.
4. Ripristinare `dataflow.spec` e `dataflow_appimage.spec` rimuovendo `collect_data_files('reportlab')`.
5. Ripristinare le modifiche i18n su `.po/.mo`.

## G. VERIFICA MANUALE SUGGERITA
- [ ] In Request Window compare `🧾 RFQ PDF` tra `📊 Esporta Excel` e SQDC.
- [ ] Dialog RFQ PDF ha stile coerente DataFlow (font, padding, bottoni, modalita).
- [ ] Selezione logo PNG/JPG funziona e salva stato persistente.
- [ ] Riapertura dialog mostra logo gia configurato.
- [ ] Sostituzione logo aggiorna il file interno.
- [ ] Rimozione logo elimina stato e non rompe export.
- [ ] Export PDF in lingua IT produce titolo/testi IT.
- [ ] Export PDF in lingua EN produce titolo/testi EN.
- [ ] RFQ Fornitura piena mostra colonne a 5 campi corretti.
- [ ] RFQ Conto lavoro mostra colonne a 7 campi corretti.
- [ ] Celle lunghe fanno wrapping e aumentano altezza riga.
- [ ] Su piu pagine l’header tabella si ripete.
- [ ] Con logo mancante/corrotto export continua senza crash con warning.
