"""
services/dashboard_controller.py

Logica di orchestrazione della dashboard principale.
Estratto conservativamente da MainWindow in dataflow.py
come parte del refactoring release 2.1.0.

RESPONSABILITÀ
- search_requests: ricerca RFQ/VSM con filtri multipli
- refresh_data: ricaricamento dati preservando filtri attivi
- clear_filters: reset filtri + reload
- populate_username_filter: aggiornamento combo utenti
- _update_filter_panel_for_current_tab: show/hide pannelli filtro contestuali

NON RESPONSABILE DI
- Costruzione UI (vedi ui/main_dashboard_builder.py)
- Logica VSM CRUD/export
- Logica database/backup/settings
"""

import re
import logging
from ui.dialogs.common_dialogs import SimpleMessageDialog

from database_manager import DatabaseManager, DatabaseError
from services.app_paths import get_db_path
from utils.i18n_utils import _, normalize_rfq_type
from utils.validation_utils import format_date_for_db

logger = logging.getLogger(__name__)


class DashboardController:
    def __init__(self, app):
        self.app = app

    def populate_username_filter(self):
        """Aggiorna la lista degli username disponibili nel filtro usando aggregazione multi-database."""
        if not self.app.user_filter_combo or not self.app.username_filter_var:
            return

        usernames = []

        try:
            # Carica TUTTI gli username da TUTTI i database aggregati
            with DatabaseManager(get_db_path()) as db_manager:
                all_requests = db_manager.get_all_richieste_aggregated(get_db_path())

            # BUG #2 FIX: Validazione robusta per gestire tuple di lunghezza variabile
            # Estrai username unici dalle richieste aggregate (indice 5)
            usernames = list({req[5].strip().lower() for req in all_requests
                            if len(req) > 5 and req[5] and str(req[5]).strip()})

            logger.info(f"[populate_username_filter] Trovati {len(usernames)} utenti: {usernames}")

        except (DatabaseError, IndexError, AttributeError, TypeError) as e:
            # BUG #2 FIX: Fallback robusto con gestione completa eccezioni
            logger.warning(f"Aggregazione multi-DB fallita in populate_username_filter, uso fallback: {e}")
            # Fallback: usa solo il database locale
            try:
                with DatabaseManager(get_db_path()) as db_manager:
                    usernames = db_manager.get_distinct_usernames()
                logger.info(f"Fallback completato: popolato filtro con {len(usernames)} utenti (locale)")
            except DatabaseError as e2:
                logger.error(f"Errore anche nel fallback: {e2}", exc_info=True)

        # Assicurati che l'utente corrente sia nella lista
        if self.app.current_username and self.app.current_username not in usernames:
            usernames.append(self.app.current_username)

        # Ordina e prepara la lista per la combo
        clean_usernames = sorted({u for u in usernames if u})
        values = [self.app.all_users_placeholder] + clean_usernames
        current_value = self.app.username_filter_var.get()
        self.app.user_filter_combo.config(values=values)

        # Resetta al valore corrente se valido, altrimenti all'utente corrente
        if current_value not in values:
            self.app.username_filter_var.set(self.app.current_username or self.app.all_users_placeholder)

    def _update_filter_panel_for_current_tab(self):
        """Aggiorna il contenuto del pannello Advanced Filters in base al tab attivo.

        Esiste un solo pannello Advanced Filters nell'intera dashboard (collapsible_filters).
        Questo metodo mostra il sub-frame corretto (RFQ o VSM) e nasconde l'altro,
        senza spostare il pannello né crearne uno secondo.
        """
        if not hasattr(self.app, 'rfq_filter_subframe') or not hasattr(self.app, 'vsm_filter_subframe'):
            return  # Inizializzazione non ancora completata
        _, status = self.app.get_current_tree_and_status()
        is_vsm = bool(status and status.startswith('vsm_'))
        if is_vsm:
            self.app.rfq_filter_subframe.grid_remove()
            self.app.vsm_filter_subframe.grid()
            # Mostra lo spec frame VSM corretto per il tab attivo, nasconde gli altri
            if hasattr(self.app, '_vsm_spec_frames'):
                show_frame = self.app._vsm_spec_frames.get(status)
                # Opera su frame unici: _vsm_sc_spec è condiviso tra Saving e CA,
                # quindi non iterare sulle chiavi ma sugli oggetti distinti.
                _seen = {}
                for f in self.app._vsm_spec_frames.values():
                    _seen[id(f)] = f
                for f in _seen.values():
                    if f is show_frame:
                        f.grid()
                    else:
                        f.grid_remove()
        else:
            self.app.vsm_filter_subframe.grid_remove()
            self.app.rfq_filter_subframe.grid()

    def refresh_data(self):
        """Ricarica i dati preservando i filtri di ricerca attivi"""
        # BUGFIX: Se ci sono filtri di ricerca attivi, usa search_requests invece di ricaricare tutto
        if self.app._has_active_search_filters():
            logger.info("[refresh_data] Filtri attivi rilevati, riapplico la ricerca")
            self.search_requests()
            return

        # Ottieni il percorso completo del mio DB
        my_path = get_db_path()
        # Chiama get_all_richieste_aggregated per ottenere tutte le richieste aggregate
        try:
            all_requests = self.app.db_manager.get_all_richieste_aggregated(my_path)
            # Salva i dati aggregati per uso successivo
            self.app._all_aggregated_requests = all_requests
        except DatabaseError as e:
            logger.error(f"Errore nel caricamento richieste aggregate: {e}", exc_info=True)
            # Fallback: usa il metodo normale se l'aggregazione fallisce
            self.app._all_aggregated_requests = None

        self.app._load_requests_by_status(self.app.tree_attive, 'attiva'); self.app._load_requests_by_status(self.app.tree_archiviate, 'archiviata')
        self.populate_username_filter()
        self.app.update_button_visibility()

    def search_requests(self):
        tree, status = self.app.get_current_tree_and_status()

        if tree is None:
            return

        # Dispatch al handler del modulo corrente.
        # Per aggiungere un nuovo modulo: aggiungere un elif e creare _search_<modulo>().
        if status.startswith('vsm_'):
            if status == 'vsm_derisking':
                self.app._search_derisking_suppliers(tree)
                return
            self.app._search_vsm_events(tree, status)
            return

        username_filter = self.app._get_active_username_filter()

        # BUG #9 FIX: Validazione lunghezza input per evitare query troppo lente
        MAX_SEARCH_LENGTH = 100
        crit = {k: v.get().strip() for k, v in self.app.search_vars.items()}

        # DEBUG: Verifica criteri di ricerca
        print(f"[search_requests] crit={crit}")
        print(f"[search_requests] search_tipo='{self.app.search_tipo.get()}'")
        print(f"[search_requests] username_filter='{username_filter}'")

        # BUG #9 FIX: Blacklist caratteri pericolosi per SQL injection
        FORBIDDEN_CHARS = re.compile(r"[';\"\\`<>]")

        # Controlla che nessun campo di ricerca sia troppo lungo
        for field_name, value in crit.items():
            if value and len(value) > MAX_SEARCH_LENGTH:
                SimpleMessageDialog(self.app.root, _("Input Troppo Lungo"), _("Il testo di ricerca nel campo '{}' è troppo lungo (max {} caratteri)").format(field_name, MAX_SEARCH_LENGTH), "warning")
                return
                return

            # BUG #5 FIX: Rimuovi caratteri pericolosi per SQL injection
            if value and FORBIDDEN_CHARS.search(value):
                sanitized = FORBIDDEN_CHARS.sub('', value)
                logger.warning(f"Caratteri pericolosi rimossi dal campo '{field_name}': '{value}' -> '{sanitized}'")
                # Aggiorna il campo con il valore sanitizzato
                self.app.search_vars[field_name].set(sanitized)
                crit[field_name] = sanitized

                # Avvisa l'utente una sola volta per tutti i campi
                if not hasattr(self.app, '_sql_injection_warning_shown'):
                    self.app._sql_injection_warning_shown = True
                    SimpleMessageDialog(self.app.root, _("Input Sanitizzato"), _("Alcuni caratteri speciali sono stati rimossi dai campi di ricerca per motivi di sicurezza."), "info")
                    # Reset flag dopo 2 secondi - BUG #48 FIX: cancella timer precedente per evitare memory leak
                    if self.app._sql_warning_after_id is not None:
                        try:
                            self.app.root.after_cancel(self.app._sql_warning_after_id)
                        except Exception as e:
                            logger.warning(f"Impossibile cancellare timer SQL warning: {e}")

                    def reset_flag():
                        if hasattr(self.app, '_sql_injection_warning_shown'):
                            delattr(self.app, '_sql_injection_warning_shown')
                    self.app._sql_warning_after_id = self.app.root.after(2000, reset_flag)

        # Validazione rimossa: ora il numero RdO supporta ricerca parziale come gli altri filtri

        dates = {k: format_date_for_db(v.get().strip()) for k, v in self.app.date_entries.items()}
        base = "SELECT DISTINCT ro.id_richiesta, ro.tipo_rdo, ro.data_emissione, ro.data_scadenza, ro.riferimento, COALESCE(ro.username, '') FROM richieste_offerta ro LEFT JOIN dettagli_richiesta dr ON ro.id_richiesta=dr.id_richiesta LEFT JOIN richiesta_fornitori rf ON ro.id_richiesta=rf.id_richiesta"
        clauses, params = ["ro.stato=?"], [status]

        # Filtri strutturali (tipo, username)
        if self.app.search_tipo.get() != _("Tutte"):
            # Normalizza il valore di ricerca al valore canonico per il confronto nel database
            tipo_canonico = normalize_rfq_type(self.app.search_tipo.get())
            clauses.append("ro.tipo_rdo=?")
            params.append(tipo_canonico)
        if username_filter:
            clauses.append("LOWER(COALESCE(ro.username, '')) = ?"); params.append(username_filter)

        # Global Search: ricerca multi-campo con OR logic (OPZIONE A)
        # Se presente, aggiunge un blocco OR che cerca in 6 campi principali
        # Questo blocco coesiste con i filtri standard (combinazione AND)
        if crit['global']:
            global_query = crit['global']
            global_clauses = [
                "CAST(ro.id_richiesta AS TEXT) LIKE ?",
                "LOWER(ro.riferimento) LIKE LOWER(?)",
                "LOWER(rf.nome_fornitore) LIKE LOWER(?)",
                "LOWER(dr.codice_materiale) LIKE LOWER(?)",
                "LOWER(dr.descrizione_materiale) LIKE LOWER(?)",
                "LOWER(ro.numeri_ordine) LIKE LOWER(?)"
            ]
            clauses.append("(" + " OR ".join(global_clauses) + ")")
            # Aggiungi il parametro global_query per ogni campo OR
            params.extend([f"%{global_query}%"] * len(global_clauses))

            # DEBUG: Verifica costruzione blocco OR
            print(f"[search_requests] global active='{global_query}'")
            print(f"[search_requests] global clauses count={len(global_clauses)}")
            print(f"[search_requests] global params added={len(global_clauses)}")

        # Filtri standard testuali (continuano a funzionare normalmente)
        # Questi si combinano con AND rispetto al blocco global OR
        if crit['num']: clauses.append("CAST(ro.id_richiesta AS TEXT) LIKE ?"); params.append(f"%{crit['num']}%")
        if crit['ref']: clauses.append("LOWER(ro.riferimento) LIKE LOWER(?)"); params.append(f"%{crit['ref']}%")
        if crit['forn']: clauses.append("LOWER(rf.nome_fornitore) LIKE LOWER(?)"); params.append(f"%{crit['forn']}%")
        if crit['cod']: clauses.append("LOWER(dr.codice_materiale) LIKE LOWER(?)"); params.append(f"%{crit['cod']}%")
        if crit['desc']: clauses.append("LOWER(dr.descrizione_materiale) LIKE LOWER(?)"); params.append(f"%{crit['desc']}%")
        if crit['ord']: clauses.append("LOWER(ro.numeri_ordine) LIKE LOWER(?)"); params.append(f"%{crit['ord']}%")
        # --- INIZIO BLOCCO AGGIUNTO ---
        if crit['cod_grezzo']: clauses.append("LOWER(dr.codice_grezzo) LIKE LOWER(?)"); params.append(f"%{crit['cod_grezzo']}%")
        if crit['dis_grezzo']: clauses.append("LOWER(dr.disegno_grezzo) LIKE LOWER(?)"); params.append(f"%{crit['dis_grezzo']}%")
        if crit['mat_cl']: clauses.append("LOWER(dr.materiale_conto_lavoro) LIKE LOWER(?)"); params.append(f"%{crit['mat_cl']}%")
        # --- FINE BLOCCO AGGIUNTO ---
        if dates['emm_da']: clauses.append("ro.data_emissione >= ?"); params.append(dates['emm_da'])
        if dates['emm_a']: clauses.append("ro.data_emissione <= ?"); params.append(dates['emm_a'])
        if dates['scad_da']: clauses.append("ro.data_scadenza >= ?"); params.append(dates['scad_da'])
        if dates['scad_a']: clauses.append("ro.data_scadenza <= ?"); params.append(dates['scad_a'])

        try:
            # Usa DatabaseManager per la ricerca avanzata
            criteria = {
                'global': crit['global'],
                'num': crit['num'],
                'ref': crit['ref'],
                'forn': crit['forn'],
                'cod': crit['cod'],
                'desc': crit['desc'],
                'ord': crit['ord'],
                'cod_grezzo': crit['cod_grezzo'],
                'dis_grezzo': crit['dis_grezzo'],
                'mat_cl': crit['mat_cl']
            }
            date_ranges = {
                'emm_da': dates['emm_da'],
                'emm_a': dates['emm_a'],
                'scad_da': dates['scad_da'],
                'scad_a': dates['scad_a']
            }

            # Gestione tipo RdO
            tipo_rdo = None
            if self.app.search_tipo.get() != _("Tutte"):
                tipo_rdo = normalize_rfq_type(self.app.search_tipo.get())

            # FIX: La ricerca deve usare aggregazione multi-database quando si filtra per altri utenti o "All users"
            # Comportamento:
            # - username_filter = None (All users) → cerca in TUTTI i database
            # - username_filter = altro utente → cerca in TUTTI i database (poi filtra per username)
            # - username_filter = utente corrente → ottimizzazione, cerca solo nel DB locale

            # Ottimizzazione: cerca solo nel DB locale se filtriamo per l'utente corrente
            search_local_only = (username_filter and
                                username_filter.lower() == self.app.current_username.lower() if self.app.current_username else False)

            # DEBUG: Indica quale modalità viene usata
            print(f"[search_requests] aggregated_mode={not search_local_only}")

            if search_local_only:
                # Caso ottimizzato: cerca solo nel database locale
                logger.info(f"[search_requests] Ricerca locale per utente corrente: {username_filter}")

                # DEBUG: Query SQL e params prima dell'esecuzione
                query = base + " WHERE " + " AND ".join(clauses)
                print(f"[search_requests][LOCAL] SQL={query}")
                print(f"[search_requests][LOCAL] params={params}")
                print(f"[search_requests][LOCAL] criteria={criteria}")

                with DatabaseManager(get_db_path()) as db_manager:
                    # BUGFIX: Se c'è global search, usa query SQL diretta perché
                    # search_richieste_advanced() non gestisce il campo 'global'
                    if crit['global']:
                        # Query SQL diretta con blocco OR già costruito
                        query = base + " WHERE " + " AND ".join(clauses) + " ORDER BY ro.id_richiesta DESC"
                        db_manager.cursor.execute(query, params)
                        results = db_manager.cursor.fetchall()
                        print(f"[search_requests][LOCAL][GLOBAL] used direct SQL query")
                    else:
                        # Usa metodo standard per filtri normali
                        results = db_manager.search_richieste_advanced(criteria, date_ranges, status=status, tipo=tipo_rdo, username=username_filter)

                print(f"[search_requests][LOCAL] results count={len(results)}")
            else:
                # Caso generale: usa aggregazione multi-database
                logger.info(f"[search_requests] Ricerca aggregata multi-DB (filtro utente: {username_filter or 'All users'})")
                with DatabaseManager(get_db_path()) as db_manager:
                    # Prima ottieni TUTTE le RdO aggregate
                    all_results = db_manager.get_all_richieste_aggregated(get_db_path())

                # DEBUG: Ramo aggregato
                print(f"[search_requests][AGGREGATED] entered with {len(all_results)} records")
                print(f"[search_requests][AGGREGATED] global='{crit['global']}'")

                # Poi filtra in memoria applicando TUTTI i criteri di ricerca
                # Struttura all_results: [id_richiesta, tipo_rdo, data_emissione, data_scadenza, riferimento, username, stato, is_mine, source_file]
                results = []
                for row in all_results:
                    # Filtro stato (obbligatorio)
                    if row[6] != status:
                        continue

                    # Filtro username (se specificato)
                    if username_filter and (not row[5] or row[5].lower() != username_filter.lower()):
                        continue

                    # Filtro tipo RdO
                    if tipo_rdo and row[1] != tipo_rdo:
                        continue

                    # Global Search (se presente): verifica match OR su campi principali
                    # OPZIONE A: global search + filtri standard coesistono con AND
                    if crit['global']:
                        global_query = crit['global'].lower()

                        # Verifica match su campi immediati (num, ref)
                        num_match = global_query in str(row[0])
                        ref_match = row[4] and global_query in row[4].lower()

                        # Se non matcha campi immediati, verifica campi dettaglio
                        if not (num_match or ref_match):
                            # Controlla forn, cod, desc, ord nel DB source
                            source_db_path = row[8] if len(row) > 8 else 'local'
                            if source_db_path == 'local':
                                source_db_path = get_db_path()
                            try:
                                detail_match = False
                                with DatabaseManager(source_db_path) as source_db_mgr:
                                    # Query SQL per verificare match OR su dettagli
                                    cursor = source_db_mgr.conn.cursor()
                                    detail_sql = """
                                        SELECT 1 FROM richieste_offerta ro
                                        LEFT JOIN dettagli_richiesta dr ON ro.id_richiesta=dr.id_richiesta
                                        LEFT JOIN richiesta_fornitori rf ON ro.id_richiesta=rf.id_richiesta
                                        WHERE ro.id_richiesta=?
                                        AND (
                                            LOWER(COALESCE(rf.nome_fornitore, '')) LIKE LOWER(?)
                                            OR LOWER(COALESCE(dr.codice_materiale, '')) LIKE LOWER(?)
                                            OR LOWER(COALESCE(dr.descrizione_materiale, '')) LIKE LOWER(?)
                                            OR LOWER(COALESCE(ro.numeri_ordine, '')) LIKE LOWER(?)
                                        )
                                        LIMIT 1
                                    """
                                    cursor.execute(detail_sql, (
                                        row[0],
                                        f'%{global_query}%',
                                        f'%{global_query}%',
                                        f'%{global_query}%',
                                        f'%{global_query}%'
                                    ))
                                    detail_match = cursor.fetchone() is not None

                                if not detail_match:
                                    continue
                            except Exception as e:
                                logger.warning(f"Errore verifica global search per RdO {row[0]} su DB {source_db_path}: {e}")
                                continue

                    # Filtri standard testuali (continuano a funzionare normalmente)
                    # Si combinano con AND rispetto al global search
                    if crit['num'] and crit['num'] not in str(row[0]):
                        continue
                    if crit['ref'] and (not row[4] or crit['ref'].lower() not in row[4].lower()):
                        continue

                    # Filtri data emissione
                    if dates['emm_da'] and (not row[2] or row[2] < dates['emm_da']):
                        continue
                    if dates['emm_a'] and (not row[2] or row[2] > dates['emm_a']):
                        continue

                    # Filtri data scadenza
                    if dates['scad_da'] and (not row[3] or row[3] < dates['scad_da']):
                        continue
                    if dates['scad_a'] and (not row[3] or row[3] > dates['scad_a']):
                        continue

                    # Per i filtri su dettagli (fornitore, materiale, ecc.), dobbiamo interrogare il DB specifico
                    # Questo è necessario perché get_all_richieste_aggregated non include questi dettagli
                    if any([crit['forn'], crit['cod'], crit['desc'], crit['ord'],
                           crit['cod_grezzo'], crit['dis_grezzo'], crit['mat_cl']]):
                        # Apri il database di origine per questa RdO
                        # BUG FIX: Se source_file è 'local', usa il percorso del DB corrente
                        source_db_path = row[8] if len(row) > 8 else 'local'
                        if source_db_path == 'local':
                            source_db_path = get_db_path()
                        try:
                            with DatabaseManager(source_db_path) as source_db_mgr:
                                # Verifica i criteri di dettaglio sul DB specifico
                                detail_match = source_db_mgr.check_richiesta_detail_criteria(
                                    row[0],  # id_richiesta
                                    {
                                        'forn': crit['forn'],
                                        'cod': crit['cod'],
                                        'desc': crit['desc'],
                                        'ord': crit['ord'],
                                        'cod_grezzo': crit['cod_grezzo'],
                                        'dis_grezzo': crit['dis_grezzo'],
                                        'mat_cl': crit['mat_cl']
                                    }
                                )
                            if not detail_match:
                                continue
                        except Exception as e:
                            logger.warning(f"Errore verifica criteri dettaglio per RdO {row[0]} su DB {source_db_path}: {e}")
                            continue

                    # Tutti i filtri passati, aggiungi ai risultati
                    # BUGFIX: Passa l'intera tupla con metadati (is_mine, source_file) per update_treeview
                    results.append(row)

                # DEBUG: Risultato finale
                print(f"[search_requests][AGGREGATED] filtered result count={len(results)}")
                logger.info(f"[search_requests] Ricerca aggregata completata: {len(results)} risultati trovati")

            self.app.update_treeview(tree, results)
        except DatabaseError as e:
            logger.error(f"Errore ricerca richieste: {e}", exc_info=True)
            SimpleMessageDialog(self.app.root, _("Errore"), _("Errore ricerca: {}").format(e), "error")

    def clear_filters(self):
        for var in self.app.search_vars.values(): var.set("")
        self.app.search_tipo.set(_("Tutte"))
        for de in self.app.date_entries.values(): de.delete(0, 'end')
        if self.app.username_filter_var:
            self.app.username_filter_var.set(self.app.current_username or self.app.all_users_placeholder)
        if self.app.vsm_username_filter_var:
            self.app.vsm_username_filter_var.set(self.app.current_username or self.app.all_users_placeholder)
        # Resetta filtri VSM avanzati
        for _var in (
            getattr(self.app, 'vsm_action_var', None),
            getattr(self.app, 'vsm_repetitive_var', None),
            getattr(self.app, 'vsm_theoretical_from_var', None),
            getattr(self.app, 'vsm_theoretical_to_var', None),
            getattr(self.app, 'vsm_actual_from_var', None),
            getattr(self.app, 'vsm_actual_to_var', None),
        ):
            if _var:
                _var.set("")
        if getattr(self.app, 'vsm_date_from_entry', None):
            self.app.vsm_date_from_entry.delete(0, 'end')
        if getattr(self.app, 'vsm_date_to_entry', None):
            self.app.vsm_date_to_entry.delete(0, 'end')
        self.refresh_data()
        # refresh_data() ricarica solo i tab RFQ; ricarica esplicitamente anche i tab VSM.
        # Saving e Cost Avoidance usano _load_vsm_events (event-based).
        # Derisking usa _load_potential_suppliers (supplier-based, backend separato).
        for _et, _sh in [
            ("Saving", getattr(self.app, 'sheet_saving', None)),
            ("Cost Avoidance", getattr(self.app, 'sheet_cost_avoidance', None)),
        ]:
            if _sh is not None:
                self.app._load_vsm_events(_et, _sh)
        _sh_dr = getattr(self.app, 'sheet_derisking', None)
        if _sh_dr is not None:
            self.app._load_potential_suppliers(_sh_dr)
