import sqlite3
import os
import glob
from datetime import datetime

# Eccezione personalizzata per isolare dipendenze dal database
class DatabaseError(Exception):
    """Eccezione generica per errori del database.
    
    Questa classe isola il file principale dall'implementazione specifica
    del database (duckdb), permettendo di gestire errori DB in modo generico.
    """
    pass

class DatabaseManager:
    def __init__(self, db_name="dataflow_db.db", read_only=False):
        """
        Inizializza il gestore del database SQLite + WAL.
        Accetta il nome del file del database.
        
        Args:
            db_name: Percorso al file database (estensione .db)
            read_only: Se True, apre il database in modalità sola lettura (per concorrenza multi-utente WAL)
        """
        self.db_name = db_name
        self.read_only = read_only
        self.conn = None
        self.cursor = None
        self.connect()
    
    # BUG #44 FIX: Aggiunti metodi __enter__ e __exit__ per supportare context manager
    def __enter__(self):
        """Context manager entry - ritorna se stesso per usare con 'with' statement"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - garantisce chiusura connessione anche in caso di eccezioni"""
        self.close()
        # Ritorna False per propagare eventuali eccezioni
        return False

    def connect(self):
        """
        Apre la connessione al database SQLite con WAL mode.
        Configurazione ottimale per concorrenza multi-utente su rete.
        """
        try:
            # Apri connessione con URI mode per read-only se richiesto
            if self.read_only:
                # URI mode read-only per WAL concurrent access
                uri = f"file:{self.db_name}?mode=ro"
                self.conn = sqlite3.connect(uri, uri=True, timeout=10.0, check_same_thread=False)
            else:
                # Modalità read-write normale
                self.conn = sqlite3.connect(self.db_name, timeout=10.0, check_same_thread=False, isolation_level=None)
                
                # Configura WAL mode e ottimizzazioni (solo per read-write)
                self.conn.execute("PRAGMA journal_mode=WAL")
                self.conn.execute("PRAGMA synchronous=NORMAL")
                self.conn.execute("PRAGMA wal_autocheckpoint=1000")
                self.conn.execute("PRAGMA cache_size=-64000")  # 64MB cache
                self.conn.execute("PRAGMA temp_store=MEMORY")
            
            # Busy timeout per gestire lock temporanei su rete
            self.conn.execute("PRAGMA busy_timeout=10000")  # 10 secondi
            
            self.cursor = self.conn.cursor()
            # SQLite row_factory per accesso dict-like
            self.conn.row_factory = sqlite3.Row
        except Exception as e:
            raise DatabaseError(f"Errore di connessione al database: {e}") from e
    
    def close(self):
        """
        Chiude la connessione al database in modo sicuro.
        """
        if self.conn:
            self.conn.commit() # Salva eventuali modifiche pendenti
            self.conn.close()

    def get_connection(self):
        """
        Restituisce l'oggetto connessione (utile per casi particolari)
        """
        return self.conn
    
    def _get_last_insert_id(self):
        """
        Helper per ottenere l'ultimo ID inserito da SQLite.
        SQLite supporta nativamente lastrowid.
        """
        try:
            if hasattr(self.cursor, 'lastrowid') and self.cursor.lastrowid is not None:
                return self.cursor.lastrowid
            return None
        except Exception:
            return None

    def create_tables(self):
        """
        Crea tutte le tabelle necessarie per l'applicazione DataFlow.
        Include anche le migrazioni delle colonne esistenti.
        """
        try:
            # SQLite usa AUTOINCREMENT al posto delle sequenze DuckDB
            # Non serve più creare sequenze separate
            
            # Migrazione colonne per richieste_offerta
            try:
                self.cursor.execute("ALTER TABLE richieste_offerta ADD COLUMN stato VARCHAR NOT NULL DEFAULT 'attiva'")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE richieste_offerta ADD COLUMN numeri_ordine VARCHAR")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE richieste_offerta ADD COLUMN tipo_rdo VARCHAR NOT NULL DEFAULT 'Fornitura piena'")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE richieste_offerta ADD COLUMN note_formattate VARCHAR")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE richieste_offerta ADD COLUMN username VARCHAR")
            except Exception:
                pass
            
            # Migrazione colonne per dettagli_richiesta
            try:
                self.cursor.execute("ALTER TABLE dettagli_richiesta ADD COLUMN disegno VARCHAR")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE dettagli_richiesta ADD COLUMN codice_grezzo VARCHAR")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE dettagli_richiesta ADD COLUMN disegno_grezzo VARCHAR")
            except Exception:
                pass
            try:
                self.cursor.execute("ALTER TABLE dettagli_richiesta ADD COLUMN materiale_conto_lavoro VARCHAR")
            except Exception:
                pass
            
            # Creazione tabelle principali
            # SQLite usa INTEGER PRIMARY KEY AUTOINCREMENT per auto-increment
            self.cursor.execute('CREATE TABLE IF NOT EXISTS fornitori (id_fornitore INTEGER PRIMARY KEY AUTOINCREMENT, nome_fornitore VARCHAR NOT NULL UNIQUE)')
            
            self.cursor.execute(''' CREATE TABLE IF NOT EXISTS richieste_offerta (id_richiesta INTEGER PRIMARY KEY AUTOINCREMENT, data_emissione VARCHAR, data_scadenza VARCHAR, riferimento VARCHAR, note_generali VARCHAR, stato VARCHAR NOT NULL DEFAULT 'attiva', numeri_ordine VARCHAR, tipo_rdo VARCHAR NOT NULL DEFAULT 'Fornitura piena', note_formattate VARCHAR, username VARCHAR) ''')
            
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS dettagli_richiesta (id_dettaglio INTEGER PRIMARY KEY AUTOINCREMENT, id_richiesta INTEGER, codice_materiale VARCHAR, descrizione_materiale VARCHAR, quantita VARCHAR, disegno VARCHAR, data_consegna_richiesta VARCHAR, codice_grezzo VARCHAR, disegno_grezzo VARCHAR, materiale_conto_lavoro VARCHAR, FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta))''')
            
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS richiesta_fornitori (id_richiesta INTEGER, nome_fornitore VARCHAR, PRIMARY KEY (id_richiesta, nome_fornitore), FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta))''')
            
            # Migrazione e creazione tabella offerte_ricevute con prezzo_unitario VARCHAR
            try:
                # SQLite usa PRAGMA table_info invece di DESCRIBE
                self.cursor.execute("PRAGMA table_info(offerte_ricevute)")
                cols = self.cursor.fetchall()
                # In SQLite, PRAGMA table_info ritorna: (cid, name, type, notnull, dflt_value, pk)
                prezzo_col_info = next((c for c in cols if c[1] == 'prezzo_unitario'), None)
                if prezzo_col_info and ('DOUBLE' in str(prezzo_col_info[2]).upper() or 'REAL' in str(prezzo_col_info[2]).upper()):
                    print("Avvio migrazione tabella 'offerte_ricevute' per prezzi testuali...")
                    self.cursor.execute("ALTER TABLE offerte_ricevute RENAME TO _offerte_ricevute_old;")
                    self.cursor.execute('''CREATE TABLE offerte_ricevute (id_dettaglio INTEGER, nome_fornitore VARCHAR, prezzo_unitario VARCHAR, PRIMARY KEY (id_dettaglio, nome_fornitore), FOREIGN KEY (id_dettaglio) REFERENCES dettagli_richiesta (id_dettaglio))''')
                    self.cursor.execute("INSERT INTO offerte_ricevute (id_dettaglio, nome_fornitore, prezzo_unitario) SELECT id_dettaglio, nome_fornitore, prezzo_unitario FROM _offerte_ricevute_old;")
                    self.cursor.execute("DROP TABLE _offerte_ricevute_old;")
                    self.conn.commit()
                    print("Migrazione completata.")
            except Exception as e:
                if "Table" not in str(e) or "does not exist" not in str(e):
                    print(f"Nota: impossibile eseguire la migrazione della tabella offerte_ricevute. Errore: {e}")
            
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS offerte_ricevute (id_dettaglio INTEGER, nome_fornitore VARCHAR, prezzo_unitario VARCHAR, PRIMARY KEY (id_dettaglio, nome_fornitore), FOREIGN KEY (id_dettaglio) REFERENCES dettagli_richiesta (id_dettaglio))''')
            
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS allegati_richiesta (id_allegato INTEGER PRIMARY KEY AUTOINCREMENT, id_richiesta INTEGER, nome_file VARCHAR, dati_file BLOB, tipo_allegato VARCHAR, nome_fornitore VARCHAR, percorso_esterno VARCHAR, data_inserimento VARCHAR DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta))''')
            
            # Migrazione colonna percorso_esterno
            try:
                self.cursor.execute("ALTER TABLE allegati_richiesta ADD COLUMN percorso_esterno VARCHAR")
            except Exception:
                pass
            
            # Migrazione colonna data_inserimento con controllo
            try:
                # SQLite usa PRAGMA table_info invece di DESCRIBE
                self.cursor.execute("PRAGMA table_info(allegati_richiesta)")
                columns = [column[1] for column in self.cursor.fetchall()]  # column[1] è il nome in SQLite
                if 'data_inserimento' not in columns:
                    self.cursor.execute("ALTER TABLE allegati_richiesta ADD COLUMN data_inserimento VARCHAR DEFAULT CURRENT_TIMESTAMP")
                    # Aggiorna i record esistenti con la data corrente
                    self.cursor.execute("UPDATE allegati_richiesta SET data_inserimento = CURRENT_TIMESTAMP WHERE data_inserimento IS NULL")
            except Exception:
                pass  # Colonna già esistente
            
            # Commit finale
            self.conn.commit()
            
        except Exception as e:
            raise DatabaseError(f"Errore durante la creazione delle tabelle: {e}") from e

    # ========== METODI INSERT ==========
    
    def insert_allegato_richiesta_link(self, id_richiesta, nome_file, tipo_allegato, nome_fornitore, percorso_esterno):
        """Inserisce un allegato salvato come link esterno."""
        try:
            # SQLite usa lastrowid invece di RETURNING
            self.cursor.execute(
                "INSERT INTO allegati_richiesta (id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore, percorso_esterno) VALUES (?, ?, NULL, ?, ?, ?)",
                (id_richiesta, nome_file, tipo_allegato, nome_fornitore, percorso_esterno)
            )
            self.conn.commit()
            return self._get_last_insert_id()
        except Exception as e:
            print(f"[DB Manager] Errore insert_allegato_richiesta_link: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_allegato_richiesta_blob(self, id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore):
        """Inserisce un allegato salvato come BLOB nel database."""
        try:
            # SQLite usa lastrowid invece di RETURNING
            self.cursor.execute(
                "INSERT INTO allegati_richiesta (id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore) VALUES (?, ?, ?, ?, ?)",
                (id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore)
            )
            self.conn.commit()
            return self._get_last_insert_id()
        except Exception as e:
            print(f"[DB Manager] Errore insert_allegato_richiesta_blob: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_richiesta_fornitore(self, id_richiesta, nome_fornitore):
        """Inserisce un fornitore associato a una richiesta."""
        try:
            self.cursor.execute(
                "INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore) VALUES (?, ?)",
                (id_richiesta, nome_fornitore)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore insert_richiesta_fornitore: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_dettaglio_richiesta(self, id_richiesta, codice_materiale='', disegno='', descrizione_materiale='', 
                                   quantita='', codice_grezzo='', disegno_grezzo='', materiale_conto_lavoro=''):
        """Inserisce un nuovo dettaglio/articolo in una richiesta."""
        try:
            # SQLite usa lastrowid invece di RETURNING
            self.cursor.execute("""
                INSERT INTO dettagli_richiesta 
                (id_richiesta, codice_materiale, disegno, descrizione_materiale, quantita,
                 codice_grezzo, disegno_grezzo, materiale_conto_lavoro) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_richiesta, codice_materiale, disegno, descrizione_materiale, quantita, 
                  codice_grezzo, disegno_grezzo, materiale_conto_lavoro))
            self.conn.commit()
            return self._get_last_insert_id()
        except Exception as e:
            print(f"[DB Manager] Errore insert_dettaglio_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_richiesta_offerta(self, tipo_rdo, stato, data_emissione, username=None):

        """Inserisce una nuova richiesta d'offerta calcolando l'ID basato sull'anno."""

        try:

            username_value = username.strip().lower() if isinstance(username, str) and username.strip() else None

            

            # --- LOGICA YEAR-DRIVEN ID ---

            # 1. Calcola la base per l'anno corrente (es. 2025 -> 2500000)

            yy = int(datetime.now().strftime('%y'))

            min_id_for_year = yy * 100000

            

            # 2. Trova il max ID attuale nel database

            # Usa conn.execute per DuckDB

            res = self.conn.execute("SELECT MAX(id_richiesta) FROM richieste_offerta").fetchone()

            max_id_esistente = res[0] if res and res[0] is not None else 0

            

            # 3. Il nuovo ID è il maggiore tra (Base Anno) e (Max Esistente + 1)

            # Se siamo nel nuovo anno, min_id_for_year vincerà su un vecchio ID basso.

            # Se siamo nello stesso anno, max_id + 1 vincerà.

            next_id = max(min_id_for_year, max_id_esistente + 1)

            

            # 4. Insert Esplicito passando l'ID calcolato

            query = """

                INSERT INTO richieste_offerta 

                (id_richiesta, tipo_rdo, stato, data_emissione, username) 

                VALUES (?, ?, ?, ?, ?) 

            """

            self.cursor.execute(query, (next_id, tipo_rdo, stato, data_emissione, username_value))

            

            self.conn.commit()

            print(f"[DB Manager] Nuova RdO creata con ID Year-Driven: {next_id}")

            return next_id

            

        except Exception as e:

            print(f"[DB Manager] Errore insert_richiesta_offerta: {e}")

            raise DatabaseError(str(e)) from e
    
    def insert_richiesta_offerta_completa(self, columns, values):
        """Inserisce una richiesta d'offerta con colonne personalizzate (per duplicazione)."""
        try:
            placeholders = ', '.join(['?'] * len(columns))
            # SQLite usa lastrowid invece di RETURNING
            sql = f"INSERT INTO richieste_offerta ({', '.join(columns)}) VALUES ({placeholders})"
            self.cursor.execute(sql, values)
            self.conn.commit()
            return self._get_last_insert_id()
        except Exception as e:
            print(f"[DB Manager] Errore insert_richiesta_offerta_completa: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_dettaglio_richiesta_completo(self, id_richiesta, columns, values):
        """Inserisce un dettaglio richiesta con colonne personalizzate (per duplicazione)."""
        try:
            placeholders = ', '.join(['?'] * (len(columns) + 1))
            sql = f"INSERT INTO dettagli_richiesta (id_richiesta, {', '.join(columns)}) VALUES ({placeholders})"
            self.cursor.execute(sql, [id_richiesta, *values])
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore insert_dettaglio_richiesta_completo: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI UPDATE ==========
    
    def update_numeri_ordine(self, id_richiesta, numeri_ordine_json):
        """Aggiorna i numeri ordine di una richiesta (salvati come JSON)."""
        try:
            self.cursor.execute(
                "UPDATE richieste_offerta SET numeri_ordine = ? WHERE id_richiesta = ?",
                (numeri_ordine_json, id_richiesta)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_numeri_ordine: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_riferimento(self, id_richiesta, riferimento):
        """Aggiorna il riferimento di una richiesta."""
        try:
            self.cursor.execute(
                "UPDATE richieste_offerta SET riferimento = ? WHERE id_richiesta = ?",
                (riferimento, id_richiesta)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_riferimento: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_note_formattate(self, id_richiesta, note_formattate):
        """Aggiorna le note formattate di una richiesta."""
        try:
            self.cursor.execute(
                "UPDATE richieste_offerta SET note_formattate = ? WHERE id_richiesta = ?",
                (note_formattate, id_richiesta)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_note_formattate: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_request_username(self, id_richiesta, username):
        """Aggiorna lo username associato a una RdO."""
        try:
            username_value = username.strip().lower() if isinstance(username, str) and username.strip() else None
            self.cursor.execute(
                "UPDATE richieste_offerta SET username = ? WHERE id_richiesta = ?",
                (username_value, id_richiesta)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_request_username: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_all_usernames(self, new_username):
        """Aggiorna lo username in TUTTE le RdO del database.
        Usato quando si cambia identità utente durante spostamento cartella.
        
        Args:
            new_username: Nuovo username da impostare
            
        Returns:
            int: Numero di righe aggiornate
        """
        try:
            username_value = new_username.strip().lower() if isinstance(new_username, str) and new_username.strip() else None
            
            # Prima conta le righe totali
            count_result = self.cursor.execute("SELECT COUNT(*) FROM richieste_offerta").fetchone()
            total_rows = count_result[0] if count_result else 0
            
            # Aggiorna tutte le righe
            self.cursor.execute(
                "UPDATE richieste_offerta SET username = ?",
                (username_value,)
            )
            
            self.conn.commit()
            print(f"[DB Manager] Aggiornate {total_rows} RdO con nuovo username: {username_value}")
            return total_rows
            
        except Exception as e:
            print(f"[DB Manager] Errore update_all_usernames: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_allegato_blob(self, id_allegato, dati_file):
        """Aggiorna i dati BLOB di un allegato esistente."""
        try:
            self.cursor.execute(
                "UPDATE allegati_richiesta SET dati_file = ? WHERE id_allegato = ?",
                (dati_file, id_allegato)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_allegato_blob: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_date_richiesta(self, id_richiesta, data_emissione, data_scadenza):
        """Aggiorna le date di emissione e scadenza di una richiesta."""
        try:
            self.cursor.execute(
                "UPDATE richieste_offerta SET data_emissione = ?, data_scadenza = ? WHERE id_richiesta = ?",
                (data_emissione, data_scadenza, id_richiesta)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_date_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_dettaglio_field(self, id_dettaglio, field_name, value):
        """Aggiorna un campo specifico di un dettaglio richiesta."""
        try:
            sql = f"UPDATE dettagli_richiesta SET {field_name} = ? WHERE id_dettaglio = ?"
            self.cursor.execute(sql, (value, id_dettaglio))
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_dettaglio_field: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_allegato_to_link(self, id_allegato, percorso_esterno):
        """Converte un allegato da BLOB a link esterno."""
        try:
            self.cursor.execute(
                "UPDATE allegati_richiesta SET dati_file = NULL, percorso_esterno = ? WHERE id_allegato = ?",
                (percorso_esterno, id_allegato)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_allegato_to_link: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_stato_richieste(self, params_list):
        """Aggiorna lo stato di multiple richieste in batch."""
        try:
            self.cursor.executemany(
                "UPDATE richieste_offerta SET stato = ? WHERE id_richiesta = ?",
                params_list
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore update_stato_richieste: {e}")
            raise DatabaseError(str(e)) from e
    
    def renumber_richieste(self, old_ids, offset):
        """Rinumera tutte le richieste aggiungendo un offset agli ID."""
        try:
            # DuckDB gestisce automaticamente le foreign keys, non serve PRAGMA
            self.cursor.execute("BEGIN TRANSACTION")
            
            for old_id in old_ids:
                new_id = old_id + offset
                self.cursor.execute("UPDATE richieste_offerta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, old_id))
                self.cursor.execute("UPDATE dettagli_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, old_id))
                self.cursor.execute("UPDATE richiesta_fornitori SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, old_id))
                self.cursor.execute("UPDATE allegati_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, old_id))
            
            # DuckDB non ha sqlite_sequence, gli auto-increment sono gestiti automaticamente
            self.conn.commit()
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore renumber_richieste: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI DELETE ==========
    
    def delete_allegato(self, id_allegato):
        """Elimina un allegato."""
        try:
            self.cursor.execute("DELETE FROM allegati_richiesta WHERE id_allegato = ?", (id_allegato,))
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore delete_allegato: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_offerta_by_dettaglio_fornitore(self, id_dettaglio, nome_fornitore):
        """Elimina un'offerta ricevuta per un dettaglio e fornitore specifici."""
        try:
            self.cursor.execute(
                "DELETE FROM offerte_ricevute WHERE id_dettaglio = ? AND nome_fornitore = ?",
                (id_dettaglio, nome_fornitore)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore delete_offerta_by_dettaglio_fornitore: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_fornitori_by_richiesta(self, id_richiesta):
        """Elimina tutti i fornitori associati a una richiesta."""
        try:
            self.cursor.execute("DELETE FROM richiesta_fornitori WHERE id_richiesta = ?", (id_richiesta,))
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore delete_fornitori_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_offerte_by_dettaglio(self, id_dettaglio):
        """Elimina tutte le offerte associate a un dettaglio."""
        try:
            self.cursor.execute("DELETE FROM offerte_ricevute WHERE id_dettaglio = ?", (id_dettaglio,))
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore delete_offerte_by_dettaglio: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_dettaglio(self, id_dettaglio):
        """Elimina un dettaglio richiesta."""
        try:
            self.cursor.execute("DELETE FROM dettagli_richiesta WHERE id_dettaglio = ?", (id_dettaglio,))
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore delete_dettaglio: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_richiesta_completa(self, id_richiesta):
        """Elimina una richiesta e tutti i dati correlati (offerte, allegati, fornitori, dettagli)."""
        try:
            self.cursor.execute("BEGIN TRANSACTION")
            self.cursor.execute("DELETE FROM offerte_ricevute WHERE id_dettaglio IN (SELECT id_dettaglio FROM dettagli_richiesta WHERE id_richiesta = ?)", (id_richiesta,))
            self.cursor.execute("DELETE FROM allegati_richiesta WHERE id_richiesta = ?", (id_richiesta,))
            self.cursor.execute("DELETE FROM richiesta_fornitori WHERE id_richiesta = ?", (id_richiesta,))
            self.cursor.execute("DELETE FROM dettagli_richiesta WHERE id_richiesta = ?", (id_richiesta,))
            self.cursor.execute("DELETE FROM richieste_offerta WHERE id_richiesta = ?", (id_richiesta,))
            self.conn.commit()
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore delete_richiesta_completa: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI TRANSAZIONALI COMPLESSI ==========
    
    def update_fornitori_richiesta(self, id_richiesta, new_suppliers, detail_ids):
        """
        Aggiorna l'elenco fornitori di una richiesta.
        Elimina offerte dei fornitori rimossi e ricrea la lista fornitori.
        """
        try:
            self.cursor.execute("BEGIN TRANSACTION")
            
            # Ottieni fornitori attuali
            self.cursor.execute("SELECT nome_fornitore FROM richiesta_fornitori WHERE id_richiesta = ?", (id_richiesta,))
            old_suppliers = {row[0] for row in self.cursor.fetchall()}
            removed_suppliers = old_suppliers - set(new_suppliers)
            
            # Elimina offerte dei fornitori rimossi
            for supplier in removed_suppliers:
                for detail_id in detail_ids:
                    self.cursor.execute(
                        "DELETE FROM offerte_ricevute WHERE id_dettaglio = ? AND nome_fornitore = ?",
                        (detail_id, supplier)
                    )
            
            # Elimina TUTTI i fornitori esistenti
            self.cursor.execute("DELETE FROM richiesta_fornitori WHERE id_richiesta = ?", (id_richiesta,))
            
            # Inserisci solo i nuovi fornitori
            for s in new_suppliers:
                self.cursor.execute(
                    "INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore) VALUES (?, ?)",
                    (id_richiesta, s)
                )
            
            self.conn.commit()
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore update_fornitori_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_or_update_allegato_sqdc(self, id_richiesta, sqdc_filename, percorso_esterno):
        """Inserisce o aggiorna un allegato SQDC (Documento Interno) con link esterno.
        
        Args:
            id_richiesta: ID della richiesta
            sqdc_filename: Nome visualizzato del file (es. SQDC_RdO_123.xlsx)
            percorso_esterno: Nome del file fisico nella cartella Attachments (es. RDO123_Interno_ID456.xlsx)
        """
        try:
            # BUG FIX: Controlla se esiste già un SQDC specifico (non qualsiasi Documento Interno)
            # Identifica gli SQDC dal nome file che inizia con "SQDC_"
            self.cursor.execute(
                "SELECT id_allegato FROM allegati_richiesta WHERE id_richiesta = ? AND tipo_allegato = 'Documento Interno' AND nome_file LIKE 'SQDC_%'",
                (id_richiesta,)
            )
            existing = self.cursor.fetchone()
            
            if existing:
                # Aggiorna esistente con nuovo link esterno
                self.cursor.execute(
                    "UPDATE allegati_richiesta SET nome_file = ?, percorso_esterno = ?, dati_file = NULL WHERE id_allegato = ?",
                    (sqdc_filename, percorso_esterno, existing[0])
                )
            else:
                # Inserisci nuovo con link esterno (come tutti gli altri allegati)
                self.cursor.execute(
                    "INSERT INTO allegati_richiesta (id_richiesta, nome_file, dati_file, tipo_allegato, nome_fornitore, percorso_esterno) VALUES (?, ?, NULL, ?, ?, ?)",
                    (id_richiesta, sqdc_filename, "Documento Interno", "Interno", percorso_esterno)
                )
            
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore insert_or_update_allegato_sqdc: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_dettagli_batch(self, detail_ids):
        """Elimina multiple righe di dettaglio e le relative offerte in batch."""
        try:
            # DuckDB: NON usiamo BEGIN TRANSACTION esplicito
            # Lasciamo che DuckDB gestisca automaticamente le transazioni
            # Usa conn.execute() invece di cursor.execute() per garantire persistenza
            
            # Elimina prima i prezzi associati
            for detail_id in detail_ids:
                self.conn.execute("DELETE FROM offerte_ricevute WHERE id_dettaglio = ?", (detail_id,))
            
            # Poi elimina gli articoli
            for detail_id in detail_ids:
                self.conn.execute("DELETE FROM dettagli_richiesta WHERE id_dettaglio = ?", (detail_id,))
            
            # Forza commit esplicito per garantire persistenza
            self.conn.commit()
            print(f"[DB Manager] COMMIT eseguito per eliminazione di {len(detail_ids)} dettagli")
            
            return len(detail_ids)
        except Exception as e:
            # Se c'è un errore, prova a fare rollback
            try:
                self.conn.rollback()
            except Exception as rollback_error:
                # Se il rollback fallisce, non è critico
                print(f"[DB Manager] Nota: Impossibile eseguire rollback: {rollback_error}")
            print(f"[DB Manager] Errore delete_dettagli_batch: {e}")
            raise DatabaseError(str(e)) from e
    
    def import_dettagli_from_list(self, id_richiesta, items_list):
        """Importa una lista di dettagli da Excel in batch."""
        try:
            # DuckDB: NON usiamo BEGIN TRANSACTION esplicito
            # Lasciamo che DuckDB gestisca automaticamente le transazioni
            # Questo dovrebbe garantire che il commit sia persistente
            
            inserted_count = 0
            for cod, allegato, desc, qta, cod_grezzo, dis_grezzo, mat_cl in items_list:
                # Usa conn.execute() invece di cursor.execute() per garantire persistenza
                self.conn.execute("""
                    INSERT INTO dettagli_richiesta 
                    (id_richiesta, codice_materiale, disegno, descrizione_materiale, quantita, 
                     codice_grezzo, disegno_grezzo, materiale_conto_lavoro) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (id_richiesta, cod, allegato, desc, qta, cod_grezzo, dis_grezzo, mat_cl))
                inserted_count += 1
            
            # Forza commit esplicito per garantire persistenza
            self.conn.commit()
            print(f"[DB Manager] COMMIT eseguito per importazione richiesta {id_richiesta}")
            
            # Verifica che i dati siano stati effettivamente salvati nella stessa connessione
            result = self.conn.execute(
                "SELECT COUNT(*) FROM dettagli_richiesta WHERE id_richiesta = ?",
                (id_richiesta,)
            ).fetchone()
            total_count = result[0] if result else 0
            
            print(f"[DB Manager] import_dettagli_from_list: Inseriti {inserted_count} articoli per richiesta {id_richiesta}. Totale articoli nella richiesta: {total_count}")
            
            # IMPORTANTE: Non chiudere la connessione qui - lasciala aperta
            # DuckDB dovrebbe persistere i dati automaticamente dopo il commit
            # Chiudere la connessione qui potrebbe causare problemi di persistenza
            
            return inserted_count
        except Exception as e:
            # Se c'è un errore, prova a fare rollback
            try:
                self.conn.rollback()
            except Exception as rollback_error:
                # Se il rollback fallisce, non è critico
                print(f"[DB Manager] Nota: Impossibile eseguire rollback: {rollback_error}")
            print(f"[DB Manager] Errore import_dettagli_from_list: {e}")
            raise DatabaseError(str(e)) from e
    
    def delete_richieste_batch(self, request_ids):
        """Elimina multiple richieste in batch con tutti i dati correlati."""
        try:
            # DuckDB: NON usiamo BEGIN TRANSACTION esplicito
            # Lasciamo che DuckDB gestisca automaticamente le transazioni
            # Usa conn.execute() invece di cursor.execute() per garantire persistenza
            
            print(f"[DB Manager] delete_richieste_batch: Eliminazione di {len(request_ids)} richieste: {request_ids}")
            
            for req_id in request_ids:
                # Elimina prima i dati correlati (offerte, allegati, fornitori, dettagli)
                self.conn.execute("DELETE FROM offerte_ricevute WHERE id_dettaglio IN (SELECT id_dettaglio FROM dettagli_richiesta WHERE id_richiesta = ?)", (req_id,))
                self.conn.execute("DELETE FROM allegati_richiesta WHERE id_richiesta = ?", (req_id,))
                self.conn.execute("DELETE FROM richiesta_fornitori WHERE id_richiesta = ?", (req_id,))
                self.conn.execute("DELETE FROM dettagli_richiesta WHERE id_richiesta = ?", (req_id,))
                # Elimina infine la richiesta principale
                self.conn.execute("DELETE FROM richieste_offerta WHERE id_richiesta = ?", (req_id,))
                print(f"[DB Manager] Eliminata richiesta {req_id}")
            
            # Forza commit esplicito per garantire persistenza
            self.conn.commit()
            print(f"[DB Manager] COMMIT eseguito per eliminazione di {len(request_ids)} richieste")
            
            # Verifica che le richieste siano state effettivamente eliminate
            for req_id in request_ids:
                result = self.conn.execute("SELECT COUNT(*) FROM richieste_offerta WHERE id_richiesta = ?", (req_id,)).fetchone()
                count = result[0] if result else 0
                if count > 0:
                    print(f"[DB Manager] WARNING: Richiesta {req_id} non è stata eliminata (ancora presente nel database)")
                else:
                    print(f"[DB Manager] Verificato: Richiesta {req_id} eliminata correttamente")
            
            return len(request_ids)
        except Exception as e:
            # Se c'è un errore, prova a fare rollback solo se c'è una transazione attiva
            try:
                self.conn.rollback()
            except Exception as rollback_error:
                # Se il rollback fallisce, non è critico - potrebbe non esserci transazione attiva
                print(f"[DB Manager] Nota: Impossibile eseguire rollback: {rollback_error}")
            print(f"[DB Manager] Errore delete_richieste_batch: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI PER GESTIONE PREZZI/OFFERTE ==========
    
    def insert_or_replace_offerta(self, id_dettaglio, nome_fornitore, prezzo_unitario):
        """Inserisce o sostituisce un'offerta ricevuta (prezzo)."""
        try:
            # SQLite: ON CONFLICT gestisce i duplicati
            self.cursor.execute(
                "INSERT INTO offerte_ricevute (id_dettaglio, nome_fornitore, prezzo_unitario) VALUES (?, ?, ?) ON CONFLICT (id_dettaglio, nome_fornitore) DO UPDATE SET prezzo_unitario = excluded.prezzo_unitario",
                (id_dettaglio, nome_fornitore, prezzo_unitario)
            )
            self.conn.commit()
        except Exception as e:
            print(f"[DB Manager] Errore insert_or_replace_offerta: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI PER ARCHIVIAZIONE ALLEGATI ==========
    
    def get_allegati_to_archive(self):
        """Recupera tutti gli allegati con dati BLOB da archiviare."""
        try:
            self.cursor.execute(
                "SELECT id_allegato, id_richiesta, nome_fornitore, nome_file, dati_file FROM allegati_richiesta WHERE dati_file IS NOT NULL AND LENGTH(dati_file) > 0"
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_allegati_to_archive: {e}")
            raise DatabaseError(str(e)) from e
    
    def archive_allegati_batch(self, allegati_updates):
        """
        Archivia un batch di allegati convertendoli da BLOB a link.
        allegati_updates: lista di tuple (percorso_esterno, id_allegato)
        """
        try:
            self.cursor.execute("BEGIN TRANSACTION")
            
            for percorso_esterno, id_allegato in allegati_updates:
                self.cursor.execute(
                    "UPDATE allegati_richiesta SET dati_file = NULL, percorso_esterno = ? WHERE id_allegato = ?",
                    (percorso_esterno, id_allegato)
                )
            
            self.conn.commit()
            return len(allegati_updates)
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore archive_allegati_batch: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI PER DUPLICAZIONE ==========
    
    def duplicate_richiesta_dettagli(self, original_id, new_request_id, detail_columns, detail_rows):
        """Duplica i dettagli di una richiesta."""
        try:
            # SQLite: conn.execute è equivalente a cursor.execute, entrambi funzionano
            placeholders = ', '.join(['?'] * (len(detail_columns) + 1))
            insert_sql = f"INSERT INTO dettagli_richiesta (id_richiesta, {', '.join(detail_columns)}) VALUES ({placeholders})"
            
            for detail in detail_rows:
                self.conn.execute(insert_sql, [new_request_id, *detail])
            
            # Il commit sarà fatto dal metodo chiamante (duplicate_richiesta_full)
            # Non facciamo commit qui per evitare commit multipli
            print(f"[DB Manager] Duplicati {len(detail_rows)} dettagli per richiesta {new_request_id}")
        except Exception as e:
            print(f"[DB Manager] Errore duplicate_richiesta_dettagli: {e}")
            raise DatabaseError(str(e)) from e
    
    # ========== METODI SELECT (LETTURA) ==========
    
    def get_allegati_by_richiesta(self, id_richiesta, tipo_allegato, has_date_column=True):
        """Recupera allegati per richiesta e tipo."""
        try:
            if has_date_column:
                self.cursor.execute(
                    "SELECT id_allegato, nome_fornitore, nome_file, data_inserimento FROM allegati_richiesta WHERE id_richiesta = ? AND tipo_allegato = ?",
                    (id_richiesta, tipo_allegato)
                )
            else:
                self.cursor.execute(
                    "SELECT id_allegato, nome_fornitore, nome_file FROM allegati_richiesta WHERE id_richiesta = ? AND tipo_allegato = ?",
                    (id_richiesta, tipo_allegato)
                )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_allegati_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_allegato_file_data(self, id_allegato):
        """Recupera nome_file, dati_file e percorso_esterno di un allegato."""
        try:
            self.cursor.execute(
                "SELECT nome_file, dati_file, percorso_esterno FROM allegati_richiesta WHERE id_allegato = ?",
                (id_allegato,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_allegato_file_data: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_max_allegato_id(self):
        """Recupera l'ID massimo degli allegati."""
        try:
            self.cursor.execute("SELECT MAX(id_allegato) FROM allegati_richiesta")
            result = self.cursor.fetchone()
            return result[0] if result and result[0] is not None else 0
        except Exception as e:
            print(f"[DB Manager] Errore get_max_allegato_id: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_fornitori_by_richiesta(self, id_richiesta, order_by=False):
        """Recupera i fornitori associati a una richiesta."""
        try:
            if order_by:
                self.cursor.execute(
                    "SELECT nome_fornitore FROM richiesta_fornitori WHERE id_richiesta = ? ORDER BY nome_fornitore",
                    (id_richiesta,)
                )
            else:
                self.cursor.execute(
                    "SELECT nome_fornitore FROM richiesta_fornitori WHERE id_richiesta = ?",
                    (id_richiesta,)
                )
            results = self.cursor.fetchall()
            print(f"[DB Manager] get_fornitori_by_richiesta: recuperati {len(results)} fornitori per richiesta {id_richiesta}")
            if results:
                print(f"[DB Manager] Fornitori: {[r[0] for r in results]}")
            return results
        except Exception as e:
            print(f"[DB Manager] Errore get_fornitori_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_fornitori_count(self, id_richiesta):
        """Conta i fornitori associati a una richiesta."""
        try:
            self.cursor.execute(
                "SELECT COUNT(*) FROM richiesta_fornitori WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()[0]
        except Exception as e:
            print(f"[DB Manager] Errore get_fornitori_count: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_dettaglio_ids_by_richiesta(self, id_richiesta):
        """Recupera gli ID dettaglio di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT id_dettaglio FROM dettagli_richiesta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_dettaglio_ids_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_numeri_ordine(self, id_richiesta):
        """Recupera i numeri ordine di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT numeri_ordine FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_numeri_ordine: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_riferimento(self, id_richiesta):
        """Recupera il riferimento di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT riferimento FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_riferimento: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_note_formattate(self, id_richiesta):
        """Recupera le note formattate di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT note_formattate FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_note_formattate: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_tipo_rdo(self, id_richiesta):
        """Recupera il tipo RdO di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT tipo_rdo FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_tipo_rdo: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_username_by_richiesta(self, id_richiesta):
        """Recupera lo username associato a una richiesta."""
        try:
            self.cursor.execute(
                "SELECT username FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            result = self.cursor.fetchone()
            return result[0] if result and result[0] else None
        except Exception as e:
            print(f"[DB Manager] Errore get_username_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_richiesta_basic_data(self, id_richiesta):
        """Recupera dati base di una richiesta (riferimento, date)."""
        try:
            self.cursor.execute(
                "SELECT riferimento, data_emissione, data_scadenza FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_richiesta_basic_data: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_richiesta_full_data(self, id_richiesta):
        """Recupera dati completi di una richiesta per export Excel."""
        try:
            self.cursor.execute(
                "SELECT data_emissione, data_scadenza, riferimento, tipo_rdo FROM richieste_offerta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_richiesta_full_data: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_dettagli_by_richiesta(self, id_richiesta):
        """Recupera tutti i dettagli (articoli) di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT id_dettaglio, codice_materiale, disegno, descrizione_materiale, quantita, codice_grezzo, disegno_grezzo, materiale_conto_lavoro FROM dettagli_richiesta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_dettagli_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_offerte_by_richiesta(self, id_richiesta):
        """Recupera tutte le offerte (prezzi) per una richiesta con JOIN."""
        try:
            self.cursor.execute(
                "SELECT o.id_dettaglio, o.nome_fornitore, o.prezzo_unitario FROM offerte_ricevute o JOIN dettagli_richiesta d ON o.id_dettaglio = d.id_dettaglio WHERE d.id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_offerte_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_offerte_count_by_richiesta(self, id_richiesta):
        """Conta le offerte ricevute per una richiesta."""
        try:
            self.cursor.execute(
                "SELECT COUNT(*) FROM offerte_ricevute o JOIN dettagli_richiesta d ON o.id_dettaglio = d.id_dettaglio WHERE d.id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()[0]
        except Exception as e:
            print(f"[DB Manager] Errore get_offerte_count_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_prezzo_quantita_by_fornitore(self, id_richiesta, nome_fornitore):
        """Recupera prezzi e quantità per un fornitore specifico."""
        try:
            self.cursor.execute(
                "SELECT o.prezzo_unitario, d.quantita FROM offerte_ricevute o JOIN dettagli_richiesta d ON o.id_dettaglio = d.id_dettaglio WHERE d.id_richiesta = ? AND o.nome_fornitore = ?",
                (id_richiesta, nome_fornitore)
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_prezzo_quantita_by_fornitore: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_dettagli_count_by_richiesta(self, id_richiesta):
        """Conta i dettagli (articoli) di una richiesta."""
        try:
            self.cursor.execute(
                "SELECT COUNT(*) FROM dettagli_richiesta WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            return self.cursor.fetchone()[0]
        except Exception as e:
            print(f"[DB Manager] Errore get_dettagli_count_by_richiesta: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_all_richieste_by_stato(self, stato, username=None):
        """Recupera tutte le richieste con uno stato specifico, opzionalmente filtrate per username."""
        try:
            query = """
                SELECT id_richiesta, tipo_rdo, data_emissione, data_scadenza, riferimento, COALESCE(username, '')
                FROM richieste_offerta
                WHERE stato = ?
            """
            params = [stato]
            if username:
                query += " AND LOWER(COALESCE(username, '')) = ?"
                params.append(username.strip().lower())
            query += " ORDER BY id_richiesta DESC"
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_all_richieste_by_stato: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_all_richieste_aggregated(self, my_db_full_path):
        """
        Cerca ricorsivamente tutti i database "cugini" partendo dalla cartella padre comune
        e restituisce una UNION di tutte le richieste con flag is_mine e source_file.
        
        IMPORTANTE: SQLite con WAL permette letture concorrenti mentre un altro processo scrive.
        Usa ATTACH per leggere database esterni in modalità READ-ONLY.
        
        Args:
            my_db_full_path: Percorso completo del database corrente
            
        Returns:
            Lista di tuple con tutte le richieste aggregate, ordinate per ID descrescente
        """
        try:
            # Normalizza il percorso per gestire correttamente i separatori
            # NON convertiamo più .duckdb -> .db perché ora supportiamo entrambi i formati
            my_db_full_path = os.path.normpath(os.path.abspath(my_db_full_path))
            
            # 1. Ottieni la cartella del mio DB (es. .../DataFlow_Guido/Database)
            my_db_dir = os.path.dirname(my_db_full_path)
            
            # 2. Ottieni la cartella "DataFlow" utente (es. .../DataFlow_Guido)
            user_df_dir = os.path.dirname(my_db_dir)
            
            # 3. Ottieni la Cartella Radice Condivisa (es. .../ROOT)
            root_shared_dir = os.path.dirname(user_df_dir)
            
            # 4. Usa glob per ricerca ricorsiva partendo dalla Radice Condivisa
            # Cerca tutti i file dataflow_db_*.db (database SQLite degli utenti)
            search_pattern = os.path.join(root_shared_dir, "**", "dataflow_db_*.db")
            found_files = glob.glob(search_pattern, recursive=True)
            
            # LOG: informazioni sui database trovati
            print(f"[DB Aggregation] Root shared dir: {root_shared_dir}")
            print(f"[DB Aggregation] Search pattern: {search_pattern}")
            print(f"[DB Aggregation] Total databases found: {len(found_files)}")
            if found_files:
                print(f"[DB Aggregation] Database files: {[os.path.basename(f) for f in found_files]}")
            
            # Normalizza tutti i percorsi trovati
            found_files = [os.path.normpath(os.path.abspath(f)) for f in found_files]
            
            # 5. Query Base per il database corrente (usa tabella locale)
            union_parts = []
            
            base_query = """
                SELECT 
                    id_richiesta, 
                    tipo_rdo, 
                    data_emissione, 
                    data_scadenza, 
                    riferimento, 
                    COALESCE(username, '') as username,
                    COALESCE(stato, 'attiva') as stato,
                    TRUE as is_mine,
                    'local' as source_file
                FROM richieste_offerta
            """
            union_parts.append(base_query)
            
            # 6. Per ogni altro database, usa ATTACH per lettura concorrente (SQLite WAL)
            attached_aliases = []  # Tieni traccia degli alias attaccati con successo
            my_db_normalized = os.path.normpath(os.path.abspath(my_db_full_path)).lower()
            
            for found_file in found_files:
                # Normalizza il percorso per confronto (case-insensitive su Windows)
                found_file_normalized = os.path.normpath(os.path.abspath(found_file)).lower()
                
                # SALTA il database corrente (già incluso nella query base)
                if found_file_normalized == my_db_normalized:
                    continue
                
                # SQLite ATTACH: percorso con escape singolo quote
                attach_path = found_file.replace('\\', '/')
                attach_path_escaped = attach_path.replace("'", "''")
                source_file_escaped = found_file.replace("'", "''")
                
                alias_name = f"db_{len(attached_aliases) + 1}"
                
                try:
                    # SQLite ATTACH syntax: ATTACH DATABASE 'path' AS alias
                    # WAL mode consente lettura concorrente automaticamente
                    attach_query = f"ATTACH DATABASE '{attach_path_escaped}' AS {alias_name}"
                    self.conn.execute(attach_query)
                    
                    # Se ATTACH ha successo, aggiungi alla UNION
                    union_query = f"""
                        SELECT 
                            id_richiesta, 
                            tipo_rdo, 
                            data_emissione, 
                            data_scadenza, 
                            riferimento, 
                            COALESCE(username, '') as username,
                            COALESCE(stato, 'attiva') as stato,
                            FALSE as is_mine,
                            '{source_file_escaped}' as source_file
                        FROM {alias_name}.richieste_offerta
                    """
                    union_parts.append(union_query)
                    attached_aliases.append(alias_name)
                    
                except Exception:
                    # Se ATTACH fallisce (file non accessibile, errore IO), continua silenziosamente
                    continue
            
            # 7. Esegui UNION di tutte le query disponibili
            if len(union_parts) == 0:
                print("[DB Aggregation] WARNING: No database queries to execute")
                return []
            
            print(f"[DB Aggregation] Executing UNION with {len(union_parts)} database(s)")
            final_query = " UNION ALL ".join(union_parts) + " ORDER BY id_richiesta DESC"
            
            results = []
            try:
                self.cursor.execute(final_query)
                results = self.cursor.fetchall()
                print(f"[DB Aggregation] Retrieved {len(results)} total RfQ records")
            finally:
                # DETACH tutti gli alias attaccati con successo
                for alias_name in attached_aliases:
                    try:
                        self.conn.execute(f"DETACH DATABASE {alias_name}")
                    except Exception:
                        # Alias non esiste o già rimosso, continua
                        pass
            
            return results
            
        except Exception as e:
            print(f"[DB Manager] Errore get_all_richieste_aggregated: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_distinct_usernames(self):
        """Restituisce l'elenco degli username presenti nelle RdO."""
        try:
            self.cursor.execute("""
                SELECT DISTINCT LOWER(username) 
                FROM richieste_offerta 
                WHERE username IS NOT NULL AND username <> '' 
                ORDER BY LOWER(username)
            """)
            rows = self.cursor.fetchall()
            return [row[0] for row in rows if row and row[0]]
        except Exception as e:
            print(f"[DB Manager] Errore get_distinct_usernames: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_all_richiesta_ids(self):
        """Recupera tutti gli ID richiesta ordinati per rinumerazione."""
        try:
            self.cursor.execute("SELECT id_richiesta FROM richieste_offerta ORDER BY id_richiesta DESC")
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_all_richiesta_ids: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_richiesta_columns_data(self, id_richiesta, columns):
        """Recupera dati di una richiesta per colonne specifiche (per duplicazione)."""
        try:
            sql = f"SELECT {', '.join(columns)} FROM richieste_offerta WHERE id_richiesta = ?"
            self.cursor.execute(sql, (id_richiesta,))
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_richiesta_columns_data: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_dettagli_columns_data(self, id_richiesta, columns):
        """Recupera dettagli per colonne specifiche (per duplicazione)."""
        try:
            sql = f"SELECT {', '.join(columns)} FROM dettagli_richiesta WHERE id_richiesta = ?"
            self.cursor.execute(sql, (id_richiesta,))
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_dettagli_columns_data: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_allegato_id_by_filename(self, id_richiesta, nome_file, tipo_allegato):
        """Recupera l'ID di un allegato per nome file."""
        try:
            self.cursor.execute(
                "SELECT id_allegato FROM allegati_richiesta WHERE id_richiesta = ? AND nome_file = ? AND tipo_allegato = ?",
                (id_richiesta, nome_file, tipo_allegato)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_allegato_id_by_filename: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_allegato_dati_file(self, id_richiesta, nome_file, tipo_allegato):
        """Recupera solo i dati BLOB di un allegato."""
        try:
            self.cursor.execute(
                "SELECT dati_file FROM allegati_richiesta WHERE id_richiesta = ? AND nome_file = ? AND tipo_allegato = ?",
                (id_richiesta, nome_file, tipo_allegato)
            )
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_allegato_dati_file: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_table_columns(self, table_name):
        """Recupera le colonne di una tabella (per duplicazione dinamica)."""
        try:
            # SQLite usa PRAGMA table_info invece di DESCRIBE
            self.cursor.execute(f"PRAGMA table_info({table_name})")
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_table_columns: {e}")
            raise DatabaseError(str(e)) from e
    
    def check_table_has_column(self, table_name, column_name):
        """Verifica se una tabella ha una colonna specifica."""
        try:
            # SQLite usa PRAGMA table_info invece di DESCRIBE
            self.cursor.execute(f"PRAGMA table_info({table_name})")
            columns = [column[1] for column in self.cursor.fetchall()]  # SQLite usa colonna[1] per il nome
            return column_name in columns
        except Exception as e:
            print(f"[DB Manager] Errore check_table_has_column: {e}")
            raise DatabaseError(str(e)) from e
    
    def renumber_richieste_with_transaction(self, old_ids, offset):
        """Rinumera tutte le richieste aggiungendo un offset (con transazione completa)."""
        try:
            # DuckDB gestisce automaticamente le foreign keys, non serve PRAGMA
            self.cursor.execute("BEGIN TRANSACTION")
            
            # Mappa vecchi ID -> nuovi ID
            id_map = {old_id: old_id + offset for old_id in old_ids}
            temp_offset = max(old_ids) * 2  # Usa offset temporaneo per evitare collisioni
            
            # Step 1: Sposta tutti a valori temporanei
            for old_id in old_ids:
                temp_id = old_id + temp_offset
                self.cursor.execute("UPDATE richieste_offerta SET id_richiesta = ? WHERE id_richiesta = ?", (temp_id, old_id))
                self.cursor.execute("UPDATE dettagli_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (temp_id, old_id))
                self.cursor.execute("UPDATE richiesta_fornitori SET id_richiesta = ? WHERE id_richiesta = ?", (temp_id, old_id))
                self.cursor.execute("UPDATE allegati_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (temp_id, old_id))
            
            # Step 2: Sposta dai temporanei ai valori finali
            for old_id in old_ids:
                temp_id = old_id + temp_offset
                new_id = id_map[old_id]
                self.cursor.execute("UPDATE richieste_offerta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, temp_id))
                self.cursor.execute("UPDATE dettagli_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, temp_id))
                self.cursor.execute("UPDATE richiesta_fornitori SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, temp_id))
                self.cursor.execute("UPDATE allegati_richiesta SET id_richiesta = ? WHERE id_richiesta = ?", (new_id, temp_id))
            
            self.conn.commit()
            print(f"[DB Manager] Rinumerazione completata: {len(old_ids)} richieste")
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore renumber_richieste_with_transaction: {e}")
            raise DatabaseError(str(e)) from e
    
    def search_richieste_advanced(self, criteria, date_ranges, status=None, tipo=None, username=None):
        """Ricerca avanzata richieste con criteri multipli.
        
        Args:
            criteria: Dict con chiavi num, ref, forn, cod, desc, ord, cod_grezzo, dis_grezzo, mat_cl
            date_ranges: Dict con chiavi emm_da, emm_a, scad_da, scad_a
            status: Stato della richiesta (attiva/archiviata)
            tipo: Tipo RdO (Fornitura piena/Conto lavoro)
            username: Username dell'utente a cui filtrare le RdO
            
        Returns:
            Lista di tuple con risultati della ricerca
        """
        try:
            base = """
                SELECT DISTINCT ro.id_richiesta, ro.tipo_rdo, ro.data_emissione, ro.data_scadenza, ro.riferimento, COALESCE(ro.username, '')
                FROM richieste_offerta ro
                LEFT JOIN richiesta_fornitori rf ON ro.id_richiesta = rf.id_richiesta
                LEFT JOIN dettagli_richiesta dr ON ro.id_richiesta = dr.id_richiesta
            """
            
            clauses = []
            params = []
            
            if status:
                clauses.append("ro.stato = ?")
                params.append(status)
            if tipo:
                clauses.append("ro.tipo_rdo = ?")
                params.append(tipo)
            if username:
                clauses.append("LOWER(COALESCE(ro.username, '')) = ?")
                params.append(username.strip().lower())
            if criteria.get('num'):
                clauses.append("CAST(ro.id_richiesta AS TEXT) LIKE ?")
                params.append(f"%{criteria['num']}%")
            if criteria.get('ref'):
                clauses.append("LOWER(ro.riferimento) LIKE LOWER(?)")
                params.append(f"%{criteria['ref']}%")
            if criteria.get('forn'):
                clauses.append("LOWER(rf.nome_fornitore) LIKE LOWER(?)")
                params.append(f"%{criteria['forn']}%")
            if criteria.get('cod'):
                clauses.append("LOWER(dr.codice_materiale) LIKE LOWER(?)")
                params.append(f"%{criteria['cod']}%")
            if criteria.get('desc'):
                clauses.append("LOWER(dr.descrizione_materiale) LIKE LOWER(?)")
                params.append(f"%{criteria['desc']}%")
            if criteria.get('ord'):
                clauses.append("LOWER(ro.numeri_ordine) LIKE LOWER(?)")
                params.append(f"%{criteria['ord']}%")
            if criteria.get('cod_grezzo'):
                clauses.append("LOWER(dr.codice_grezzo) LIKE LOWER(?)")
                params.append(f"%{criteria['cod_grezzo']}%")
            if criteria.get('dis_grezzo'):
                clauses.append("LOWER(dr.disegno_grezzo) LIKE LOWER(?)")
                params.append(f"%{criteria['dis_grezzo']}%")
            if criteria.get('mat_cl'):
                clauses.append("LOWER(dr.materiale_conto_lavoro) LIKE LOWER(?)")
                params.append(f"%{criteria['mat_cl']}%")
            
            if date_ranges.get('emm_da'):
                clauses.append("ro.data_emissione >= ?")
                params.append(date_ranges['emm_da'])
            if date_ranges.get('emm_a'):
                clauses.append("ro.data_emissione <= ?")
                params.append(date_ranges['emm_a'])
            if date_ranges.get('scad_da'):
                clauses.append("ro.data_scadenza >= ?")
                params.append(date_ranges['scad_da'])
            if date_ranges.get('scad_a'):
                clauses.append("ro.data_scadenza <= ?")
                params.append(date_ranges['scad_a'])
            
            query = base
            if clauses:
                query += " WHERE " + " AND ".join(clauses)
            query += " ORDER BY ro.id_richiesta DESC"
            
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
            
        except Exception as e:
            print(f"[DB Manager] Errore search_richieste_advanced: {e}")
            raise DatabaseError(str(e)) from e
    
    def check_richiesta_detail_criteria(self, richiesta_id, detail_criteria):
        """Verifica se una richiesta soddisfa i criteri di dettaglio specificati.
        
        Args:
            richiesta_id: ID della richiesta da verificare
            detail_criteria: Dict con chiavi forn, cod, desc, ord, cod_grezzo, dis_grezzo, mat_cl
            
        Returns:
            bool: True se la richiesta soddisfa tutti i criteri specificati, False altrimenti
        """
        try:
            # Costruisci query per verificare i criteri
            base = """
                SELECT COUNT(*) 
                FROM richieste_offerta ro
                LEFT JOIN richiesta_fornitori rf ON ro.id_richiesta = rf.id_richiesta
                LEFT JOIN dettagli_richiesta dr ON ro.id_richiesta = dr.id_richiesta
                WHERE ro.id_richiesta = ?
            """
            
            clauses = []
            params = [richiesta_id]
            
            # Aggiungi criteri se specificati (tutti case-insensitive)
            if detail_criteria.get('forn'):
                clauses.append("LOWER(rf.nome_fornitore) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['forn']}%")
            if detail_criteria.get('cod'):
                clauses.append("LOWER(dr.codice_materiale) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['cod']}%")
            if detail_criteria.get('desc'):
                clauses.append("LOWER(dr.descrizione_materiale) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['desc']}%")
            if detail_criteria.get('ord'):
                clauses.append("LOWER(ro.numeri_ordine) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['ord']}%")
            if detail_criteria.get('cod_grezzo'):
                clauses.append("LOWER(dr.codice_grezzo) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['cod_grezzo']}%")
            if detail_criteria.get('dis_grezzo'):
                clauses.append("LOWER(dr.disegno_grezzo) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['dis_grezzo']}%")
            if detail_criteria.get('mat_cl'):
                clauses.append("LOWER(dr.materiale_conto_lavoro) LIKE LOWER(?)")
                params.append(f"%{detail_criteria['mat_cl']}%")
            
            query = base
            if clauses:
                query += " AND " + " AND ".join(clauses)
            
            self.cursor.execute(query, params)
            result = self.cursor.fetchone()
            
            # Ritorna True se almeno una riga soddisfa i criteri
            return result and result[0] > 0
            
        except Exception as e:
            print(f"[DB Manager] Errore check_richiesta_detail_criteria: {e}")
            # In caso di errore, meglio includere la RdO (falso positivo) che escluderla (falso negativo)
            return True
    
    def duplicate_richiesta_full(self, original_id, get_columns_func):
        """Duplica completamente una richiesta (senza fornitori).
        
        Args:
            original_id: ID della richiesta da duplicare
            get_columns_func: Funzione helper per ottenere colonne (exclude set)
            
        Returns:
            ID della nuova richiesta creata
        """
        try:
            # DuckDB: NON usiamo BEGIN TRANSACTION esplicito
            # Lasciamo che DuckDB gestisca automaticamente le transazioni
            # Usa conn.execute() invece di cursor.execute() per garantire persistenza
            
            # Duplica richiesta principale
            request_columns = get_columns_func('richieste_offerta', {'id_richiesta'})
            if not request_columns:
                raise ValueError("Struttura tabella richieste_offerta non valida.")
            
            request_row = self.get_richiesta_columns_data(original_id, request_columns)
            if request_row is None:
                raise ValueError("RdO originale non trovata.")
            
            request_values = list(request_row)
            
            # Imposta stato attiva
            if 'stato' in request_columns:
                request_values[request_columns.index('stato')] = 'attiva'
            else:
                request_columns.append('stato')
                request_values.append('attiva')
            
            # Rimuovi note
            if 'note_formattate' in request_columns:
                request_values[request_columns.index('note_formattate')] = None
            
            # Imposta date
            if 'data_emissione' in request_columns:
                request_values[request_columns.index('data_emissione')] = datetime.now().date().strftime('%Y-%m-%d')
            if 'data_scadenza' in request_columns:
                request_values[request_columns.index('data_scadenza')] = None
            
            # --- CALCOLA IL PROSSIMO ID BASATO SULLA POLICY ANNO (YY00000) ---
            # 1. Calcola la base per l'anno corrente (es. 2025 -> 2500000)
            yy = int(datetime.now().strftime('%y'))
            min_id_for_year = yy * 100000
            
            # 2. Trova il max ID attuale nel database
            res = self.conn.execute("SELECT MAX(id_richiesta) FROM richieste_offerta").fetchone()
            max_id_esistente = res[0] if res and res[0] is not None else 0
            
            # 3. Il nuovo ID è il maggiore tra (Base Anno) e (Max Esistente + 1)
            new_request_id = max(min_id_for_year, max_id_esistente + 1)
            
            # 4. Aggiungi l'ID calcolato alle colonne e valori
            request_columns.insert(0, 'id_richiesta')
            request_values.insert(0, new_request_id)
            
            # Usa conn.execute() con RETURNING per confermare l'ID della nuova richiesta
            result = self.conn.execute(
                f"INSERT INTO richieste_offerta ({', '.join(request_columns)}) VALUES ({', '.join(['?'] * len(request_columns))}) RETURNING id_richiesta",
                request_values
            ).fetchone()
            
            if result:
                new_request_id = result[0]
            else:
                # Fallback: usa l'ID calcolato
                pass
            
            # Duplica dettagli
            detail_columns = get_columns_func('dettagli_richiesta', {'id_dettaglio', 'id_richiesta'})
            if detail_columns:
                detail_rows = self.get_dettagli_columns_data(original_id, detail_columns)
                if detail_rows:
                    self.duplicate_richiesta_dettagli(original_id, new_request_id, detail_columns, detail_rows)
            
            # Forza commit esplicito per garantire persistenza
            self.conn.commit()
            print(f"[DB Manager] COMMIT eseguito per duplicazione richiesta {original_id} -> {new_request_id}")
            print(f"[DB Manager] Duplicazione completata: {original_id} -> {new_request_id}")
            return new_request_id
            
        except Exception as e:
            # Se c'è un errore, prova a fare rollback solo se c'è una transazione attiva
            try:
                self.conn.rollback()
            except Exception as rollback_error:
                # Se il rollback fallisce, non è critico - potrebbe non esserci transazione attiva
                print(f"[DB Manager] Nota: Impossibile eseguire rollback: {rollback_error}")
            print(f"[DB Manager] Errore duplicate_richiesta_full: {e}")
            raise DatabaseError(str(e)) from e
    
    def get_fornitori_ordered_for_request(self, id_richiesta):
        """Recupera fornitori ordinati per una richiesta (per PO window)."""
        try:
            self.cursor.execute(
                "SELECT nome_fornitore FROM richiesta_fornitori WHERE id_richiesta = ? ORDER BY nome_fornitore",
                (id_richiesta,)
            )
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[DB Manager] Errore get_fornitori_ordered_for_request: {e}")
            raise DatabaseError(str(e)) from e
    
    def save_suppliers_with_transaction(self, id_richiesta, new_suppliers, old_suppliers, detail_ids):
        """Salva fornitori con gestione transazione completa e cleanup offerte.
        
        Args:
            id_richiesta: ID della richiesta
            new_suppliers: Lista nuovi fornitori
            old_suppliers: Lista vecchi fornitori
            detail_ids: Lista ID dettagli per cleanup offerte
        """
        try:
            # DuckDB gestisce automaticamente le transazioni, ma possiamo usare BEGIN TRANSACTION esplicito
            # Prova senza BEGIN TRANSACTION esplicito per vedere se risolve il problema di persistenza
            # self.cursor.execute("BEGIN TRANSACTION")
            
            # Identifica fornitori rimossi
            removed_suppliers = set(old_suppliers) - set(new_suppliers)
            
            # Elimina prezzi dei fornitori rimossi
            if removed_suppliers:
                for supplier in removed_suppliers:
                    for detail_id in detail_ids:
                        self.cursor.execute(
                            "DELETE FROM offerte_ricevute WHERE id_dettaglio = ? AND nome_fornitore = ?",
                            (detail_id, supplier)
                        )
            
            # Elimina TUTTI i fornitori esistenti
            self.cursor.execute("DELETE FROM richiesta_fornitori WHERE id_richiesta = ?", (id_richiesta,))
            
            # Inserisci solo i nuovi fornitori
            for supplier in new_suppliers:
                self.cursor.execute(
                    "INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore) VALUES (?, ?)",
                    (id_richiesta, supplier)
                )
                print(f"[DB Manager] Inserito fornitore: {supplier} per richiesta {id_richiesta}")
            
            # Commit esplicito e verificato - IMPORTANTE: DuckDB richiede commit per persistenza
            self.conn.commit()
            print(f"[DB Manager] COMMIT eseguito per richiesta {id_richiesta}")
            
            # BUG FIX: RIMOSSO CHECKPOINT - causava freeze su DuckDB
            # Il WAL viene consolidato automaticamente alla chiusura della connessione
            
            # Verifica che i dati siano stati salvati correttamente
            self.cursor.execute(
                "SELECT COUNT(*) FROM richiesta_fornitori WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            count_result = self.cursor.fetchone()
            saved_count = count_result[0] if count_result else 0
            
            # Verifica anche i nomi dei fornitori salvati
            self.cursor.execute(
                "SELECT nome_fornitore FROM richiesta_fornitori WHERE id_richiesta = ?",
                (id_richiesta,)
            )
            saved_suppliers = [row[0] for row in self.cursor.fetchall()]
            
            print(f"[DB Manager] Fornitori aggiornati per richiesta {id_richiesta}: {saved_count} fornitori salvati (attesi: {len(new_suppliers)})")
            print(f"[DB Manager] Fornitori salvati: {saved_suppliers}")
            
            # Verifica che il numero corrisponda
            if saved_count != len(new_suppliers):
                raise Exception(f"Discrepanza nel salvataggio: salvati {saved_count} fornitori ({saved_suppliers}), attesi {len(new_suppliers)} ({new_suppliers})")
        except Exception as e:
            # Rollback solo se la connessione è ancora aperta
            if self.conn:
                try:
                    self.conn.rollback()
                except:
                    pass
            print(f"[DB Manager] Errore save_suppliers_with_transaction: {e}")
            raise DatabaseError(str(e)) from e
    
    def update_po_numbers_json(self, id_richiesta, supplier, po_json):
        """Aggiorna numeri ordine per fornitore (formato JSON)."""
        try:
            self.cursor.execute(
                "UPDATE richieste_offerta SET numeri_ordine = ? WHERE id_richiesta = ?",
                (po_json, id_richiesta)
            )
            self.conn.commit()
            print(f"[DB Manager] PO numbers aggiornati per richiesta {id_richiesta}, fornitore {supplier}")
        except Exception as e:
            print(f"[DB Manager] Errore update_po_numbers_json: {e}")
            raise DatabaseError(str(e)) from e
    
    def save_sqdc_data_batch(self, id_richiesta, sqdc_data_list):
        """Salva dati SQDC in batch (prezzo, lead time, consegna).
        
        Args:
            id_richiesta: ID della richiesta
            sqdc_data_list: Lista di tuple (id_dettaglio, fornitore, prezzo, lead_time, rispetta_consegna)
        """
        try:
            for id_dettaglio, fornitore, prezzo, lead_time, rispetta_consegna in sqdc_data_list:
                self.cursor.execute("""
                    UPDATE offerte_ricevute 
                    SET prezzo_unitario = ?, lead_time = ?, rispetta_tempi_consegna = ?
                    WHERE id_dettaglio = ? AND nome_fornitore = ?
                """, (prezzo, lead_time, rispetta_consegna, id_dettaglio, fornitore))
            
            self.conn.commit()
            print(f"[DB Manager] SQDC data salvati in batch: {len(sqdc_data_list)} record")
        except Exception as e:
            self.conn.rollback()
            print(f"[DB Manager] Errore save_sqdc_data_batch: {e}")
            raise DatabaseError(str(e)) from e
    
    def insert_or_update_offerta(self, id_dettaglio, fornitore, prezzo):
        """Inserisce o aggiorna un'offerta (per import Excel)."""
        try:
            self.cursor.execute("""
                INSERT INTO offerte_ricevute (id_dettaglio, nome_fornitore, prezzo_unitario)
                VALUES (?, ?, ?)
                ON CONFLICT(id_dettaglio, nome_fornitore) 
                DO UPDATE SET prezzo_unitario = excluded.prezzo_unitario
            """, (id_dettaglio, fornitore, prezzo))
            self.conn.commit()
            print(f"[DB Manager] Offerta inserita/aggiornata: dettaglio {id_dettaglio}, fornitore {fornitore}")
        except Exception as e:
            print(f"[DB Manager] Errore insert_or_update_offerta: {e}")
            raise DatabaseError(str(e)) from e

    def get_dettaglio_row_by_id(self, id_dettaglio):
        """Recupera una singola riga articolo dal database per id_dettaglio.
        
        Returns:
            tuple: (id_dettaglio, codice_materiale, disegno, descrizione_materiale, 
                   quantita, codice_grezzo, disegno_grezzo, materiale_conto_lavoro) o None
        """
        try:
            self.cursor.execute("""
                SELECT id_dettaglio, codice_materiale, disegno, descrizione_materiale, quantita,
                       codice_grezzo, disegno_grezzo, materiale_conto_lavoro
                FROM dettagli_richiesta WHERE id_dettaglio = ?
            """, (id_dettaglio,))
            return self.cursor.fetchone()
        except Exception as e:
            print(f"[DB Manager] Errore get_dettaglio_row_by_id: {e}")
            raise DatabaseError(str(e)) from e

    def get_available_usernames(self, shared_folder_path):
        """Recupera la lista degli username dai file database nella cartella condivisa."""
        try:
            # Cerca pattern ricorsivo se siamo in una root, o semplice se flat
            # Per sicurezza cerchiamo nella cartella specifica passata
            search_pattern = os.path.join(shared_folder_path, "dataflow_db_*.duckdb")
            files = glob.glob(search_pattern)
            
            # Se la logica prevede sottocartelle (es. DataFlow_User/Database/file.duckdb)
            # Aggiungiamo una ricerca più ampia se la prima non trova nulla o per completezza
            if not files:
                 search_pattern_recursive = os.path.join(shared_folder_path, "**", "dataflow_db_*.duckdb")
                 files = glob.glob(search_pattern_recursive, recursive=True)

            usernames = set()
            for f in files:
                filename = os.path.basename(f)
                # Estrae 'username' da 'dataflow_db_username.duckdb'
                # Rimuove prefisso 'dataflow_db_' (12 chars) e suffisso '.duckdb' (7 chars)
                if filename.startswith("dataflow_db_") and filename.endswith(".duckdb"):
                    user = filename[12:-7]
                    if user:
                        usernames.add(user)
            
            return sorted(list(usernames))
        except Exception as e:
            print(f"[DB Manager] Errore scansione username: {e}")
            return []
