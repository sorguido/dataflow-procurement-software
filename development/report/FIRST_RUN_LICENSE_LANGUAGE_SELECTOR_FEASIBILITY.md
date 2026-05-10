# First-run license language selector - studio di fattibilita

## Executive summary

La modifica e fattibile con scope basso/medio e senza rifattorizzare la sezione lingua esistente in Settings.

Il flusso attuale mostra la finestra obbligatoria `LicenseAcceptanceDialog` prima della creazione del database utente e prima della costruzione della `MainWindow`. La preferenza lingua e gia persistita in `config.ini`, sezione `[Settings]`, chiave `language`, tramite `services.settings_preferences_service.save_language_preference(...)`. Il service i18n centralizzato puo essere reinizializzato richiamando `init_i18n()` dopo il salvataggio della lingua e prima di `MainWindow(root)`.

Raccomandazione: procedere con cautela. La cautela principale riguarda l'ordine di scrittura di `config.ini`: una scrittura con un `ConfigParser` caricato prima del salvataggio lingua puo sovrascrivere accidentalmente la lingua appena salvata.

## Flusso attuale startup/licenza/i18n

Sequenza osservata:

1. `dataflow.py` importa `init_i18n` da `utils.i18n_utils`.
2. `dataflow.py` chiama `init_i18n()` a livello top-level prima degli import UI, per rendere `tr()` disponibile ai moduli UI importati successivamente (`dataflow.py:96-100`).
3. Nel blocco `if __name__ == '__main__'`, `init_i18n()` viene richiamato prima della creazione di `root = tk.Tk()` (`dataflow.py:3218-3227`).
4. `main_task()` legge `config.ini` con `get_config_file()` e controlla `[Settings].license_accepted` (`dataflow.py:3261-3270`).
5. Se la licenza non e accettata, apre `LicenseAcceptanceDialog(root, url=...)` e blocca lo startup con `root.wait_window(license_prompt)` (`dataflow.py:3273-3277`).
6. Se l'utente non accetta, `root.destroy()` e `return` interrompono l'avvio (`dataflow.py:3277-3282`).
7. Se l'utente accetta, viene salvato `[Settings].license_accepted = True` in `config.ini` (`dataflow.py:3284-3290`).
8. Solo dopo il gate licenza vengono gestiti identita utente, database, splash e interfaccia principale.
9. La dashboard viene costruita con `app = MainWindow(root)` solo dopo questi passaggi (`dataflow.py:3362-3368`).

Questa sequenza e favorevole al requisito: la lingua scelta nella finestra licenza puo essere salvata e i18n puo essere reinizializzato prima della costruzione della dashboard.

## File e funzioni rilevanti

### Finestra licenza primo avvio

File: `ui/dialogs/common_dialogs.py`

Classe: `LicenseAcceptanceDialog`

Riferimenti:

- `LicenseAcceptanceDialog.__init__`: costruisce finestra, logo, messaggio e pulsanti (`ui/dialogs/common_dialogs.py:522-561`).
- `_on_license`: apre l'URL licenza con `webbrowser.open(self._url)` (`ui/dialogs/common_dialogs.py:563-564`).
- `_on_accept`: imposta `self.accepted = True`, rilascia il grab e distrugge la finestra (`ui/dialogs/common_dialogs.py:566-569`).
- `_on_exit`: imposta `self.accepted = False`, rilascia il grab e distrugge la finestra (`ui/dialogs/common_dialogs.py:571-574`).

Comportamento attuale:

- Accept: chiude il dialog con `accepted=True`; il salvataggio effettivo avviene in `dataflow.py`.
- Read License: apre il link GitHub della licenza e lascia aperto il dialog.
- Exit o chiusura finestra: chiude il dialog con `accepted=False`; lo startup distrugge root e termina.

Nota: il pulsante License della dashboard non usa questa classe; `ui/window_launchers.py.open_license_window(...)` apre direttamente l'URL licenza (`ui/window_launchers.py:15-16`).

### Persistenza accettazione licenza

File: `dataflow.py`

Funzione/contesto: `main_task()` nel blocco `if __name__ == '__main__'`

Meccanismo:

- Percorso config: `utils.user_utils.get_config_file()`.
- File config su Linux/script: `~/.local/share/DataFlow/config.ini`, oppure `$XDG_DATA_HOME/DataFlow/config.ini`; su Windows frozen: `%USERPROFILE%\AppData\Local\DataFlow\config.ini` (`utils/user_utils.py:11-37`).
- Chiave: `[Settings] license_accepted`.
- Lettura: `config.getboolean('Settings', 'license_accepted', fallback=False)` (`dataflow.py:3266-3270`).
- Scrittura: `config['Settings']['license_accepted'] = 'True'` e `config.write(...)` (`dataflow.py:3284-3290`).

Valutazione semantica:

La selezione lingua puo essere aggiunta senza alterare la semantica dell'accettazione licenza se viene persistita solo nel ramo in cui `license_prompt.accepted` e `True`. Se l'utente seleziona una lingua e poi usa Exit, il ramo di salvataggio non viene eseguito e `license_accepted` resta non salvato, come oggi.

### Persistenza lingua applicazione

File: `services/settings_preferences_service.py`

Funzioni:

- `load_settings_snapshot(config_file)`: legge `[Settings].language`, default `en`, valori ammessi `en`/`it` (`services/settings_preferences_service.py:12-48`).
- `save_language_preference(config_file, selected_language_label)`: converte `English -> en`, ogni altro label atteso `Italiano -> it`, poi salva `[Settings].language` (`services/settings_preferences_service.py:51-68`).

Uso attuale da Settings:

- `SettingsWindow` crea la sezione lingua/valuta con `ttk.Combobox(... values=["English", "Italiano"], state="readonly")` (`dataflow.py:321-333`).
- `SettingsWindow.load_settings()` imposta il valore UI da `load_settings_snapshot(...)` (`dataflow.py:390-398`).
- `SettingsWindow.save_language_currency_settings()` rileva il cambio lingua e chiama `save_language_preference(...)`; poi mostra il prompt di restart attuale (`dataflow.py:457-502`).

Valutazione:

Il meccanismo puo essere riusato in sicurezza dalla finestra licenza perche:

- e gia centralizzato;
- usa la stessa chiave letta da i18n;
- non richiede dipendenze UI;
- non modifica il flusso Settings.

Va pero gestito attentamente l'ordine delle scritture. Se si chiama `save_language_preference(...)` e poi si scrive su disco un oggetto `config` caricato prima, la seconda scrittura puo perdere `[Settings].language`. Mitigazioni possibili:

- scrivere lingua e `license_accepted` nello stesso `ConfigParser` e in un'unica `config.write(...)`; oppure
- salvare prima la licenza, poi chiamare `save_language_preference(...)`, che rilegge il file e preserva la licenza; oppure
- dopo `save_language_preference(...)`, rileggere `config.ini` prima di impostare `license_accepted`.

La soluzione piu pulita per atomicita logica e una scrittura unica di entrambe le chiavi nel ramo Accept. La soluzione piu aderente al riuso dell'helper e salvare licenza e poi chiamare `save_language_preference(...)`, accettando una doppia scrittura ma mantenendo minimo lo scope.

### Inizializzazione i18n

File: `utils/i18n_utils.py`

Funzioni/classi:

- `TranslationService.initialize(language_code="en")`: legge la lingua da `config.ini`, normalizza `en`/`it`, carica `locale/<lang>/LC_MESSAGES/dataflow.mo` e aggiorna il traduttore runtime (`utils/i18n_utils.py:77-113`).
- `init_i18n(language_code="en")`: wrapper pubblico su `TranslationService.initialize(...)` (`utils/i18n_utils.py:141-143`).
- `tr(text)`: usa il traduttore attivo del singleton (`utils/i18n_utils.py:131-133`).
- `get_current_language()`: rilegge la lingua persistita e aggiorna il codice corrente (`utils/i18n_utils.py:115-120`, `utils/i18n_utils.py:146-148`).

Valutazione timing:

i18n puo essere reinizializzato dopo l'accettazione licenza e prima di `MainWindow(root)` perche:

- `TranslationService` e un singleton runtime, non un valore immutabile;
- `tr(...)` delega al traduttore corrente;
- la dashboard viene costruita dopo il dialog licenza (`dataflow.py:3362-3368`);
- `MainWindow.__init__` e `build_main_dashboard(self)` chiamano `tr(...)` durante la costruzione UI (`dataflow.py:1018-1069`).

Dopo aver salvato `[Settings].language = it`, una chiamata a `init_i18n()` prima dello splash/dashboard dovrebbe permettere alla dashboard di aprirsi direttamente in italiano, senza prompt di riavvio.

Nota: `init_i18n(language_code="it")` non forza la lingua se `config.ini` contiene gia `[Settings].language`, perche `initialize(...)` legge sempre il config con default pari all'argomento. Per questo la lingua deve essere persistita prima della reinizializzazione.

## Valutazione di fattibilita

Fattibile.

Motivi:

- Il gate licenza precede `MainWindow`.
- La preferenza lingua e gia una preferenza applicativa persistita.
- i18n e reinvocabile prima della costruzione UI principale.
- Il prompt di restart e confinato a Settings; non e necessario riusarlo nel primo avvio.
- Non servono nuove dipendenze.

Limite da chiarire in implementazione:

- Se si vuole garantire che la finestra licenza sia sempre in inglese anche nel caso anomalo `license_accepted=False` ma `language=it` gia presente in config, non basta l'attuale `init_i18n()`. In quel caso il dialog dovrebbe usare stringhe letterali inglesi per titolo/messaggio/pulsanti, oppure ricevere una modalita dedicata "first-run English". Nei test richiesti, con nessuna preferenza lingua precedente, il comportamento attuale e gia inglese.

## Piano di implementazione minimale raccomandato

### File probabilmente da modificare

Numero minimo consigliato: 2 file.

1. `ui/dialogs/common_dialogs.py`

   Modifica alta quota:

   - aggiungere a `LicenseAcceptanceDialog` una `tk.StringVar(value="English")`;
   - aggiungere un piccolo `ttk.Frame` centrato sotto il messaggio licenza e sopra `btn_frame`;
   - nel frame, aggiungere `ttk.Label(..., text="Interface language:")` e `ttk.Combobox(..., values=["English", "Italiano"], state="readonly", width=20)`;
   - esporre la scelta con un attributo, ad esempio `self.selected_language = "English"` iniziale, aggiornato in `_on_accept`;
   - lasciare `_on_license` e `_on_exit` invariati, senza persistenza.

2. `dataflow.py`

   Modifica alta quota:

   - dopo `root.wait_window(license_prompt)` e solo nel ramo `accepted=True`, leggere `license_prompt.selected_language`;
   - salvare la lingua selezionata usando il meccanismo esistente o la stessa chiave `[Settings].language`;
   - salvare `license_accepted` esattamente come oggi;
   - richiamare `init_i18n()` dopo il salvataggio e prima di identita utente, splash e `MainWindow(root)`;
   - non mostrare prompt di riavvio.

Possibile variante con 3 file:

3. `services/settings_preferences_service.py`

   Da modificare solo se si vuole evitare duplicazione della conversione label/codice in una scrittura unica. Si potrebbe estrarre una piccola funzione, ad esempio `language_label_to_code(selected_language_label)`, usata sia da Settings sia dal primo avvio. Non e necessario per una prima implementazione minimale.

### Evitare modifiche

- Non modificare la UI Settings esistente.
- Non cambiare `SettingsWindow.save_language_currency_settings()`.
- Non spostare l'inizializzazione top-level di i18n in `dataflow.py`, perche l'ordine import/i18n e gia indicato come sensibile.
- Non introdurre nuovi moduli.

## Proposta UI

Nel `LicenseAcceptanceDialog`, inserire tra messaggio e pulsanti:

```python
language_frame = ttk.Frame(frame)
language_frame.pack(pady=(0, 20))

ttk.Label(language_frame, text="Interface language:").pack(side="left", padx=(0, 10))
language_combo = ttk.Combobox(
    language_frame,
    textvariable=self.language_var,
    values=["English", "Italiano"],
    state="readonly",
    width=20,
)
language_combo.pack(side="left")
language_combo.current(0)
```

Motivazione layout:

- `language_frame.pack(...)` senza `fill="x"` mantiene il frame centrato orizzontalmente nel parent.
- Label e combobox affiancati sono centrati come gruppo.
- La posizione e sotto il messaggio licenza e sopra `btn_frame`.
- Usa solo Tkinter/ttk gia presenti.
- Default: `English`.

Se si vuole aderire esattamente al layout desiderato, valutare anche la rimozione delle emoji dai pulsanti solo in questo dialog (`Read License`, `Accept`, `Exit`). Non e funzionalmente necessaria.

## Rischi e mitigazioni

| Rischio | Livello | Dettaglio | Mitigazione |
|---|---:|---|---|
| Timing reinizializzazione i18n | Medio | `init_i18n()` e chiamato gia prima degli import UI e nel main; una reinvocazione troppo tarda lascerebbe la dashboard in inglese. | Richiamare `init_i18n()` immediatamente dopo la persistenza lingua e prima di splash e `MainWindow(root)`. Non spostare l'inizializzazione top-level esistente. |
| Regressione Settings lingua | Basso | Settings usa lo stesso helper e mantiene prompt restart. | Non modificare `SettingsWindow`, salvo nessuna modifica. Riutilizzare solo la persistenza o la chiave config. Test manuale dedicato su Settings. |
| Regressione persistenza licenza | Medio | Se la logica Accept viene alterata, la licenza potrebbe non essere salvata o essere salvata su Exit. | Mantenere il salvataggio dentro il solo ramo `accepted=True`. `_on_exit` deve restare senza side effect. |
| Ordine persistenza config al primo avvio | Medio | Doppie scritture con `ConfigParser` stale possono perdere `[Settings].language` o `license_accepted`. | Scrittura unica di entrambe le chiavi, oppure rilettura esplicita del config prima della seconda scrittura. Coprire con test primo avvio Italiano. |
| Dialog licenza non inglese in config anomalo | Basso/Medio | Se esiste `language=it` ma `license_accepted=False`, il dialog attuale puo tradursi perche usa `tr(...)`. | Per requisito forte "licenza sempre inglese", usare stringhe letterali inglesi nel dialog first-run o aggiungere parametro dedicato. |
| Compatibilita Linux/Windows UI | Basso | `ttk.Combobox` e `ttk.Frame` sono gia usati in Settings e LanguagePrompt. | Usare widget ttk esistenti; evitare dimensioni rigide aggressive; centrare con `pack()` del frame. |
| Prompt restart indesiderato | Basso | Usare direttamente metodi Settings potrebbe mostrare prompt. | Non chiamare `SettingsWindow.save_language_currency_settings()`. Usare solo helper di persistenza o scrittura config diretta. |
| Lingua splash | Basso | Se i18n viene reinizializzato prima dello splash, anche lo splash puo apparire in italiano. | Questo e coerente con la scelta utente; se si vuole solo dashboard italiana ma splash inglese, reinizializzare subito prima di `MainWindow`, ma non e raccomandato. |

## Piano di rollback

Rollback semplice a livello file:

1. Ripristinare `ui/dialogs/common_dialogs.py` rimuovendo:
   - `language_var`;
   - frame label/combobox;
   - attributo `selected_language`;
   - eventuali modifiche ai testi dei pulsanti.
2. Ripristinare `dataflow.py` rimuovendo:
   - lettura di `license_prompt.selected_language`;
   - salvataggio lingua dal ramo Accept;
   - `init_i18n()` aggiuntivo post-licenza.
3. Se modificato, ripristinare `services/settings_preferences_service.py` rimuovendo eventuali helper aggiunti e tornando alla funzione `save_language_preference(...)` attuale.

Il rollback non richiede migrazione dati: una eventuale chiave `[Settings].language` gia salvata e compatibile con il comportamento attuale di Settings/i18n.

## Piano di test manuale

### Test 1 - Primo avvio default

Precondizioni:

- nessuna accettazione licenza precedente;
- nessuna preferenza lingua precedente.

Passi/attese:

1. Avviare DataFlow.
2. Verificare che la finestra licenza si apra in inglese.
3. Verificare che il menu lingua mostri `English`.
4. Cliccare `Accept`.
5. Verificare che l'app si apra in inglese.
6. Verificare in `config.ini` che `license_accepted = True` e, se la futura implementazione salva sempre il default, `language = en`.

### Test 2 - Primo avvio con Italiano

Precondizioni:

- nessuna accettazione licenza precedente;
- nessuna preferenza lingua precedente.

Passi/attese:

1. Avviare DataFlow.
2. Verificare che la finestra licenza si apra in inglese.
3. Selezionare `Italiano`.
4. Cliccare `Accept`.
5. Verificare che la dashboard si apra direttamente in italiano.
6. Chiudere e riaprire DataFlow.
7. Verificare che la finestra licenza non si riapra.
8. Verificare in `config.ini` che `license_accepted = True` e `language = it`.

### Test 3 - Read License

Precondizioni:

- finestra di primo avvio aperta.

Passi/attese:

1. Selezionare `Italiano`.
2. Cliccare `Read License`.
3. Verificare che la licenza si apra come prima nel browser.
4. Verificare che la finestra licenza resti aperta.
5. Verificare che la selezione lingua nel combobox resti `Italiano`.
6. Verificare che `license_accepted` non venga salvato finche non si clicca `Accept`.

### Test 4 - Exit

Precondizioni:

- finestra di primo avvio aperta.

Passi/attese:

1. Selezionare `Italiano`.
2. Cliccare `Exit`.
3. Verificare che l'app esca come prima.
4. Verificare che l'accettazione licenza non venga salvata.
5. Verificare che la finestra licenza si ripresenti al successivo avvio.
6. Verificare che la lingua non venga salvata su Exit, salvo scelta diversa esplicitamente documentata; raccomandazione: non salvarla.

### Test 5 - Settings intatto

Precondizioni:

- licenza accettata e app aperta.

Passi/attese:

1. Aprire Settings/Impostazioni.
2. Verificare che la sezione lingua esistente sia ancora presente.
3. Verificare che il combobox lingua rifletta la lingua persistita.
4. Cambiare lingua da Settings.
5. Salvare.
6. Verificare che il comportamento attuale resti invariato, incluso eventuale prompt di riavvio.
7. Dopo restart, verificare che la lingua scelta da Settings sia applicata.

### Test aggiuntivi consigliati

#### Test 6 - Config anomalo: lingua presente, licenza non accettata

Precondizioni:

- `config.ini` contiene `[Settings] language = it`;
- `license_accepted` assente o `False`.

Attese:

- Se il requisito "finestra licenza sempre inglese" e interpretato rigidamente, il dialog deve comunque essere in inglese e il combobox deve mostrare `English`.
- Se si accetta il comportamento attuale con config anomalo, documentare l'eccezione.

#### Test 7 - Persistenza doppia chiave

Passi/attese:

1. Primo avvio, selezionare `Italiano`, cliccare `Accept`.
2. Ispezionare `config.ini`.
3. Verificare che siano presenti contemporaneamente `license_accepted = True` e `language = it`.
4. Verificare che nessuna altra sezione del config sia stata rimossa.

#### Test 8 - Linux/Windows

Passi/attese:

1. Eseguire su Linux.
2. Eseguire su Windows o build frozen Windows, se disponibile.
3. Verificare centratura del frame lingua, rendering del combobox e corretta scrittura nel percorso config piattaforma.

## Raccomandazione finale

Procedere con cautela.

La modifica e piccola e coerente con l'architettura attuale. Il percorso raccomandato e limitarsi a `LicenseAcceptanceDialog` e al ramo Accept in `main_task()`, salvare la lingua prima della creazione della dashboard e richiamare `init_i18n()` prima di `MainWindow(root)`.

La futura implementazione deve prestare attenzione soprattutto a due aspetti: non toccare il flusso Settings e non usare scritture config in ordine tale da perdere la preferenza lingua o l'accettazione licenza.
