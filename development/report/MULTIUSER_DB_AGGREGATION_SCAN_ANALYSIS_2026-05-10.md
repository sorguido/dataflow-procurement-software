# Multi-user DB Aggregation Scan Analysis — 2026-05-10

## 1. Executive Summary

Durante l'avvio da sorgente con `python dataflow.py`, DataFlow puo bloccarsi nella splash screen mentre carica l'interfaccia. La call chain osservata arriva a `MainWindow.__init__()`, poi a `populate_vsm_username_filter()`, quindi a `DatabaseManager.get_all_vsm_events_aggregated()`.

La causa probabile e una scansione ricorsiva troppo ampia dei database multiutente: il pattern attuale cerca `dataflow_db_*.db` sotto tutta la root condivisa con `**` e `recursive=True`. Nel caso Linux osservato, questa root e `/home/guido`, quindi la ricerca attraversa potenzialmente tutta la home utente prima di completare il caricamento UI.

## 2. Runtime Evidence

Traceback rilevante ottenuto interrompendo con Ctrl+C:

```text
File "dataflow.py", line 3367, in main_task
    app = MainWindow(root)

File "dataflow.py", line 1083, in __init__
    self.populate_vsm_username_filter()

File "dataflow.py", line 1172, in populate_vsm_username_filter
    all_data = db_manager.get_all_vsm_events_aggregated(get_db_path())

File "database_manager.py", line 2337, in get_all_vsm_events_aggregated
    found_files = glob.glob(search_pattern, recursive=True)
```

Path reali osservati:

- DB path: `/home/guido/DataFlow_gsoraru/Database/dataflow_db_gsoraru.db`
- User DataFlow dir: `/home/guido/DataFlow_gsoraru`
- Root shared dir: `/home/guido`
- Pattern attuale: `/home/guido/**/dataflow_db_*.db`

Struttura attesa:

```text
/home/guido/
├── DataFlow_gsoraru/
│   └── Database/dataflow_db_gsoraru.db
├── DataFlow_cpastiscio/
│   └── Database/dataflow_db_cpastiscio.db
```

## 3. Current Behavior

Le funzioni aggregate prendono il path completo del database corrente, lo normalizzano, risalgono alla directory `Database`, poi alla directory utente `DataFlow_<username>`, poi alla directory parent condivisa.

Per un DB corrente come:

```text
/home/guido/DataFlow_gsoraru/Database/dataflow_db_gsoraru.db
```

la derivazione e:

- `my_db_dir`: `/home/guido/DataFlow_gsoraru/Database`
- `user_df_dir`: `/home/guido/DataFlow_gsoraru`
- `root_shared_dir`: `/home/guido`

Da questa root viene costruito:

```python
search_pattern = os.path.join(root_shared_dir, "**", "dataflow_db_*.db")
found_files = glob.glob(search_pattern, recursive=True)
```

Il database locale viene incluso con una query locale o un caricamento locale separato, mentre i DB trovati vengono normalizzati e confrontati con il path locale per saltare il DB gia caricato. I DB esterni vengono poi letti in modalita aggregata: RFQ tramite `ATTACH DATABASE`, VSM e Derisking tramite connessioni SQLite read-only dirette.

## 4. Root Cause Hypothesis

`glob.glob(os.path.join(root_shared_dir, "**", "dataflow_db_*.db"), recursive=True)` non limita la ricerca alla struttura DataFlow standard. Se `root_shared_dir` e `/home/guido`, `C:\Users\<utente>` o una parent directory molto popolata, Python deve visitare ricorsivamente molte sottodirectory prima di restituire i risultati.

Questo puo diventare molto lento in presenza di repository, virtual environment, cache, cloud sync, backup, download, cartelle applicative, Steam, directory nascoste o alberi con moltissimi file. Poiche la chiamata avviene durante `MainWindow.__init__()` e prima della chiusura della splash, l'app appare bloccata in fase di caricamento interfaccia.

Nel traceback osservato il blocco e esattamente dentro `glob.glob(..., recursive=True)` della funzione VSM, chiamata dal popolamento del filtro utenti VSM.

## 5. Functions Reviewed

| Funzione | File | Ruolo | Pattern di ricerca usato | Rischio rilevato | Modifica futura |
| --- | --- | --- | --- | --- | --- |
| `main_task()` | `dataflow.py` | Inizializza licenza, identita, DB, splash e crea `MainWindow` | Nessun glob diretto | La splash resta aperta mentre `MainWindow` esegue caricamenti sincroni | No, salvo eventuale ottimizzazione separata |
| `MainWindow.__init__()` | `dataflow.py` | Costruisce UI e carica dataset iniziali VSM, Derisking e RFQ | Nessun glob diretto | Chiama caricamenti aggregati prima del completamento splash | No diretto per questo fix |
| `populate_vsm_username_filter()` | `dataflow.py` | Popola il filtro utenti VSM leggendo eventi aggregati | Indiretto via `get_all_vsm_events_aggregated()` | Puo bloccare avvio quando la scansione multi-DB e lenta | No diretto, se si corregge il backend aggregato |
| `get_db_path()` | `services/app_paths.py` | Determina il DB corrente da config, `custom_db_path` o struttura standard | Nessun glob | Influenza la root da cui le aggregazioni risalgono alla parent condivisa | No per questo fix |
| `get_user_documents_dataflow_dir()` | `services/app_paths.py` | Costruisce `DataFlow_<username>` in base standard o `dataflow_base_dir` | Nessun glob | Su Linux standard usa `~/DataFlow_<username>`; parent condivisa diventa `~` | No per questo fix |
| `get_all_vsm_events_aggregated()` | `database_manager.py` | Aggrega eventi Saving / Cost Avoidance da DB sibling | `root_shared_dir/**/dataflow_db_*.db`, recursive | Alto: chiamata in avvio e nei filtri VSM | Si, consigliata |
| `get_all_richieste_aggregated()` | `database_manager.py` | Aggrega RFQ multiutente | `root_shared_dir/**/dataflow_db_*.db`, recursive | Medio-alto: dashboard RFQ, filtro utenti RFQ, export e search aggregata | Si, consigliata per coerenza |
| `get_all_potential_suppliers_aggregated()` | `database_manager.py` | Aggrega Derisking / potential suppliers | `root_shared_dir/**/dataflow_db_*.db`, recursive | Medio-alto: caricamento Derisking in avvio e filtri | Si, consigliata per coerenza |
| `get_available_usernames()` | `database_manager.py` | Estrae username da DB DuckDB legacy o non piu centrale | Prima flat `dataflow_db_*.duckdb`, poi fallback `**/dataflow_db_*.duckdb` | Rischio simile ma su `.duckdb` e solo se la prima ricerca non trova file | Da valutare separatamente; non nel path del traceback |
| `detect_username_conflict()` | `services/dataflow_location_service.py` | Verifica conflitti nella cartella destinazione DataFlow | Path diretto `DataFlow_<username>/Database/dataflow_db_<username>.db` | Nessun rischio di scansione ampia | No |

## 6. Cross-platform Analysis

Linux standard:

- `get_user_documents_dataflow_dir()` usa `os.path.expanduser('~')`.
- Il DB corrente diventa tipicamente `~/DataFlow_<username>/Database/dataflow_db_<username>.db`.
- La root condivisa calcolata dalle funzioni aggregate e `~`.
- Il pattern attuale diventa `~/**/dataflow_db_*.db`, quindi scansiona tutta la home.
- Il pattern ristretto diventerebbe `~/DataFlow_*/Database/dataflow_db_*.db`, coerente con la struttura standard.

Windows standard:

- `get_user_documents_dataflow_dir()` usa `~/Documents/DataFlow_<username>`.
- Il DB corrente diventa tipicamente `C:\Users\<utente>\Documents\DataFlow_<username>\Database\dataflow_db_<username>.db`.
- La root condivisa calcolata e `C:\Users\<utente>\Documents`.
- Il pattern attuale scansiona ricorsivamente tutti i Documents.
- Il pattern ristretto diventerebbe `C:\Users\<utente>\Documents\DataFlow_*\Database\dataflow_db_*.db`, coerente con la struttura standard.

Cartella spostata da Settings:

- `get_user_documents_dataflow_dir()` rispetta `Settings.dataflow_base_dir`.
- La cartella selezionata e trattata come parent dove creare `DataFlow_<username>`.
- Se il DB corrente e `<selected_parent_dir>/DataFlow_<username>/Database/dataflow_db_<username>.db`, la root condivisa calcolata e `<selected_parent_dir>`.
- Il pattern ristretto `<selected_parent_dir>/DataFlow_*/Database/dataflow_db_*.db` preserva la logica sibling prevista.

Caso `custom_db_path` legacy:

- `get_db_path()` da priorita a `custom_db_path`.
- Se `custom_db_path` punta fuori dalla struttura `DataFlow_<username>/Database`, la derivazione tramite due `dirname()` puo produrre una root non semanticamente corretta.
- Il vecchio pattern ricorsivo poteva trovare DB in posizioni non standard; il pattern ristretto potrebbe non trovarli.
- Questo e il principale rischio di compatibilita da considerare prima del fix.

## 7. Proposed Fix

Modifica consigliata, non applicata in questa sessione:

- File: `database_manager.py`
- Funzioni: `get_all_vsm_events_aggregated()`, `get_all_richieste_aggregated()`, `get_all_potential_suppliers_aggregated()`
- Vecchio comportamento: ricerca ricorsiva illimitata sotto `root_shared_dir`.
- Nuovo comportamento: ricerca non ricorsiva nella struttura standard dei DB DataFlow sibling.

Diff concettuale:

Da:

```python
search_pattern = os.path.join(root_shared_dir, "**", "dataflow_db_*.db")
found_files = glob.glob(search_pattern, recursive=True)
```

A:

```python
search_pattern = os.path.join(root_shared_dir, "DataFlow_*", "Database", "dataflow_db_*.db")
found_files = glob.glob(search_pattern)
```

Motivazione:

- Limita la scansione alle sole directory DataFlow sibling attese.
- Evita traversal ricorsivi su home, Documents o parent directory molto grandi.
- Preserva il DB locale corrente, perche il codice lo carica gia separatamente o lo salta dopo normalizzazione.
- Preserva i DB sibling standard `DataFlow_<username>/Database/dataflow_db_<username>.db`.
- Mantiene compatibilita con separatori Linux e Windows tramite `os.path.join`.
- Mantiene compatibilita con `dataflow_base_dir`, se la directory scelta e la parent delle cartelle `DataFlow_<username>`.

Rollback:

- Ripristinare le due righe originali con `**` e `recursive=True` nelle funzioni modificate.
- Non sono previste migrazioni dati o modifiche schema, quindi il rollback e solo codice.

Test manuali consigliati: vedere sezione 11.

## 8. Impact Assessment

Avvio applicazione:

- Impatto atteso positivo. La splash dovrebbe completare piu rapidamente perche il popolamento del filtro VSM non scandisce tutta la home o Documents.

Filtro utenti VSM:

- Dovrebbe continuare a mostrare `All users`, l'utente locale e gli utenti sibling se i DB sono nella struttura standard.
- Se un utente dipende da DB collocati manualmente fuori da `DataFlow_*/Database`, tali utenti potrebbero non apparire piu nel filtro aggregato.

Aggregazione multiutente:

- La logica multiutente standard e preservata.
- La vista aggregata diventa piu prevedibile perche considera solo database DataFlow nel layout ufficiale.

RFQ:

- `get_all_richieste_aggregated()` usa lo stesso pattern ampio.
- Dashboard RFQ, filtro utenti RFQ, export senza filtri e ricerca aggregata possono beneficiare della stessa restrizione.

Derisking:

- `get_all_potential_suppliers_aggregated()` usa lo stesso pattern ampio.
- Il caricamento Derisking in avvio e i filtri aggregati dovrebbero beneficiare della stessa restrizione.

Database locale:

- Nessun impatto atteso. Il DB locale viene caricato direttamente dal path corrente e poi escluso dai DB esterni tramite confronto path.

Database sibling:

- I sibling standard `DataFlow_<altro_utente>/Database/dataflow_db_<altro_utente>.db` restano inclusi.
- DB sibling non standard richiedono valutazione legacy o documentazione di migrazione.

## 9. Risks / Edge Cases

- Database legacy non dentro cartelle `DataFlow_*`.
- Database copiati manualmente in path arbitrari sotto la root condivisa.
- Cartelle con nomi diversi da `DataFlow_<username>`.
- DB con nome conforme `dataflow_db_*.db` ma collocati in sottodirectory non `Database`.
- Config `custom_db_path` che punta a una struttura non standard.
- Ambienti Windows in cui la cartella base non e `Documents` ma un percorso custom o sincronizzato.
- Dipendenze implicite dal vecchio comportamento ricorsivo, per esempio raccolta di DB in backup o copie annidate.
- Il nuovo pattern potrebbe ridurre la vista aggregata se oggi l'utente si affida a DB sparsi.
- `get_available_usernames()` contiene un fallback ricorsivo su `.duckdb`; non e nel traceback, ma andrebbe rivisto separatamente se ancora usato in flussi reali.

## 10. Rollback Plan

Il rollback e semplice e reversibile:

1. Ripristinare in `database_manager.py` il pattern:

   ```python
   search_pattern = os.path.join(root_shared_dir, "**", "dataflow_db_*.db")
   found_files = glob.glob(search_pattern, recursive=True)
   ```

2. Applicarlo nelle stesse funzioni eventualmente modificate.
3. Rieseguire i test manuali su avvio, VSM, RFQ e Derisking.
4. Nessuna modifica a database, schema o configurazioni e necessaria.

## 11. Manual Test Plan

1. Avvio DataFlow da `dataflow.py`.
2. Verifica che splash completi il caricamento.
3. Verifica che il filtro utenti VSM mostri:
   - `All users`
   - `gsoraru`
   - `cpastiscio`
4. Verifica che gli eventi Saving / Cost Avoidance siano visibili.
5. Verifica che i dati dell'altro utente siano read-only dove previsto.
6. Verifica RFQ aggregate, se la funzione di aggregazione RFQ usa logica simile.
7. Verifica Derisking aggregate, se la funzione di aggregazione Derisking usa logica simile.
8. Test su Windows o revisione path Windows, se non disponibile ambiente Windows.

Test aggiuntivi consigliati:

1. Ambiente Linux con home popolata da repository/cache: confrontare tempo di avvio prima/dopo.
2. Ambiente con `dataflow_base_dir` custom: verificare che `<selected_parent_dir>/DataFlow_*/Database/dataflow_db_*.db` trovi i sibling.
3. Ambiente con solo DB locale: verificare che l'app funzioni e i filtri includano almeno l'utente corrente.
4. Ambiente con DB legacy fuori standard: verificare se la perdita di aggregazione e accettabile o richiede fallback esplicito.

## 12. Final Recommendation

Il fix e consigliato. La modifica dovrebbe essere applicata in `database_manager.py` alle tre funzioni aggregate SQLite principali: `get_all_vsm_events_aggregated()`, `get_all_richieste_aggregated()` e `get_all_potential_suppliers_aggregated()`.

Non conviene limitarla solo a VSM: il traceback parte da VSM perche VSM viene caricato durante l'avvio, ma RFQ e Derisking condividono lo stesso pattern ricorsivo ampio e possono manifestare lentezza negli stessi ambienti.

La soluzione piu conservativa e sostituire localmente il pattern nelle tre funzioni con `DataFlow_*/Database/dataflow_db_*.db`. Un piccolo helper interno puo essere valutato in seguito per ridurre duplicazione, ma non e necessario per un fix semplice, stabile e reversibile.

Rischio stimato: basso per installazioni standard e per cartella DataFlow spostata tramite Settings; medio per installazioni legacy che dipendono da database collocati manualmente fuori dalla struttura `DataFlow_<username>/Database`.
