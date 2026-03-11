"""
DataFlow Database Migration Tool - Migration Engine
Core migration logic from v1.1.0 to v2.0.0
"""

import sqlite3
import os
import shutil
import time
import subprocess
from datetime import datetime
from id_mapper import IDMapper
from attachment_extractor import AttachmentExtractor


class MigrationError(Exception):
    """Custom exception for migration errors"""
    pass


def normalize_rfq_type(rfq_type):
    """
    Normalize RFQ type to canonical Italian value.
    
    Args:
        rfq_type: Original tipo_rdo value
        
    Returns:
        str: Canonical value ('Fornitura piena' or 'Conto lavoro')
    """
    if not rfq_type:
        return "Fornitura piena"
    
    rfq_type = rfq_type.strip()
    
    # Map all variants to canonical Italian values
    type_map = {
        "Fornitura piena": "Fornitura piena",
        "Conto lavoro": "Conto lavoro",
        "Full Supply": "Fornitura piena",
        "Work Order": "Conto lavoro",
        "fornitura piena": "Fornitura piena",
        "conto lavoro": "Conto lavoro",
        "full supply": "Fornitura piena",
        "work order": "Conto lavoro",
    }
    
    # Exact match first
    if rfq_type in type_map:
        return type_map[rfq_type]
    
    # Case-insensitive match
    rfq_type_lower = rfq_type.lower()
    for key, value in type_map.items():
        if key.lower() == rfq_type_lower:
            return value
    
    # Default fallback
    return "Fornitura piena"


class MigrationEngine:
    """
    Manages the complete migration process from v1.1.0 to v2.0.0.
    """
    
    def __init__(self, source_db_path, target_paths, username, logger, progress_callback=None):
        """
        Initialize migration engine.
        
        Args:
            source_db_path: Path to v1.1.0 database
            target_paths: Dict with 'base_dir', 'db_file', 'attachments_dir' keys
            username: Username to assign to all migrated RfQs
            logger: Logger instance
            progress_callback: Optional callback function(step, message) for UI updates
        """
        self.source_db_path = source_db_path
        self.target_db_path = target_paths['db_file']
        self.target_base_dir = target_paths['base_dir']
        self.target_attachments_dir = target_paths['attachments_dir']
        self.username = username
        self.logger = logger
        self.progress_callback = progress_callback
        
        # Calculate source attachments directories
        # DataFlow 1.1.0 could use either "Allegati" (Italian) or "Attachments" (English)
        # The structure is: Documents\DataFlow\{Allegati|Attachments}\
        source_db_dir = os.path.dirname(source_db_path)
        source_base_dir = os.path.dirname(source_db_dir)  # Go up from Database\ to DataFlow\
        
        # Try both possible attachment folder names
        self.source_attachments_dirs = [
            os.path.join(source_base_dir, 'Allegati'),      # Italian version (v1.1.0 default)
            os.path.join(source_base_dir, 'Attachments'),   # English version or renamed
            os.path.join(source_db_dir, 'Attachments'),     # Next to database (alternative)
        ]
        
        # Log all possible directories and check which exists
        self.logger.info(f"Source database directory: {source_db_dir}")
        self.logger.info(f"Source base directory: {source_base_dir}")
        self.logger.info(f"Checking for attachments in:")
        
        self.source_attachments_dir = None
        for att_dir in self.source_attachments_dirs:
            exists = os.path.exists(att_dir)
            self.logger.info(f"  - {att_dir}: {'EXISTS' if exists else 'not found'}")
            if exists and self.source_attachments_dir is None:
                self.source_attachments_dir = att_dir
        
        if self.source_attachments_dir:
            self.logger.info(f"Using attachments directory: {self.source_attachments_dir}")
        else:
            self.logger.warning(f"No attachments directory found - external files will fail")
            # Use first option as fallback
            self.source_attachments_dir = self.source_attachments_dirs[0]
        
        self.id_mapper = IDMapper()
        self.attachment_extractor = None
        
        self.statistics = {
            'rfqs_migrated': 0,
            'articles_migrated': 0,
            'suppliers_migrated': 0,
            'prices_migrated': 0,
            'attachments_migrated': 0,
            'warnings': [],
            'errors': []
        }
    
    def _update_progress(self, step, message):
        """Update progress via callback if available"""
        self.logger.info(f"[Step {step}] {message}")
        if self.progress_callback:
            self.progress_callback(step, message)
    
    def _prepare_target_folder(self):
        """
        Prepare target folder structure (destructive operation).
        Deletes existing DataFlow_{username} folder and recreates it.
        """
        self._update_progress(1, "Preparing target folder structure...")
        
        # Delete existing folder if it exists
        if os.path.exists(self.target_base_dir):
            self.logger.warning(f"Deleting existing folder: {self.target_base_dir}")
            
            # Close any open database connections to WAL files
            db_file = self.target_db_path
            if os.path.exists(db_file):
                try:
                    # Try to checkpoint and close WAL
                    conn = sqlite3.connect(db_file, timeout=5.0)
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                    conn.close()
                    self.logger.debug("Closed WAL files before deletion")
                    time.sleep(0.5)  # Give Windows time to release locks
                except Exception as e:
                    self.logger.warning(f"Could not checkpoint WAL: {e}")
            
            # Delete WAL and SHM files manually if they exist
            for ext in ['-wal', '-shm']:
                wal_file = db_file + ext
                if os.path.exists(wal_file):
                    try:
                        os.remove(wal_file)
                        self.logger.debug(f"Removed {wal_file}")
                    except Exception as e:
                        self.logger.warning(f"Could not remove {wal_file}: {e}")
            
            # Retry logic for folder deletion (Windows can be slow to release locks)
            max_retries = 3
            deletion_successful = False
            
            for attempt in range(max_retries):
                try:
                    shutil.rmtree(self.target_base_dir)
                    self.logger.info(f"Successfully deleted folder using shutil.rmtree on attempt {attempt + 1}")
                    deletion_successful = True
                    break
                except OSError as e:
                    if attempt < max_retries - 1:
                        self.logger.warning(f"Attempt {attempt + 1} failed: {e}, retrying in 1 second...")
                        time.sleep(1)
                    else:
                        self.logger.warning(f"All shutil.rmtree attempts failed: {e}")
            
            # If standard deletion failed, try PowerShell Remove-Item -Force
            if not deletion_successful:
                self.logger.info("Attempting forced deletion using PowerShell Remove-Item...")
                try:
                    # Use PowerShell's Remove-Item with -Force and -Recurse
                    ps_command = f'Remove-Item -Path "{self.target_base_dir}" -Recurse -Force -ErrorAction Stop'
                    
                    result = subprocess.run(
                        ["powershell", "-Command", ps_command],
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    
                    if result.returncode == 0:
                        self.logger.info("Successfully deleted folder using PowerShell Remove-Item -Force")
                        deletion_successful = True
                        time.sleep(0.5)  # Brief pause to ensure deletion is complete
                    else:
                        self.logger.error(f"PowerShell deletion failed: {result.stderr}")
                        
                except subprocess.TimeoutExpired:
                    self.logger.error("PowerShell deletion timed out after 30 seconds")
                except Exception as e:
                    self.logger.error(f"PowerShell deletion error: {e}")
            
            # Final check - if still not deleted, raise error
            if not deletion_successful:
                raise MigrationError(
                    f"Failed to delete existing folder after all attempts.\n\n"
                    f"The folder may be locked by a system process.\n"
                    f"Troubleshooting steps:\n"
                    f"1. Restart your computer to release all locks\n"
                    f"2. Manually delete the folder: {self.target_base_dir}\n"
                    f"3. Retry the migration"
                )
        
        # Create fresh folder structure
        os.makedirs(self.target_base_dir, exist_ok=True)
        os.makedirs(os.path.dirname(self.target_db_path), exist_ok=True)
        os.makedirs(self.target_attachments_dir, exist_ok=True)
        
        self.logger.info(f"Created target folder: {self.target_base_dir}")
    
    def _create_target_schema(self):
        """Create v2.0.0 database schema in target database"""
        self._update_progress(2, "Creating v2.0.0 database schema...")
        
        try:
            conn = sqlite3.connect(self.target_db_path, timeout=30.0, isolation_level=None)
            cursor = conn.cursor()
            
            # Enable WAL mode and optimizations
            cursor.execute("PRAGMA journal_mode=WAL")
            cursor.execute("PRAGMA synchronous=NORMAL")
            cursor.execute("PRAGMA wal_autocheckpoint=1000")
            cursor.execute("PRAGMA cache_size=-64000")
            cursor.execute("PRAGMA temp_store=MEMORY")
            cursor.execute("PRAGMA busy_timeout=10000")
            
            # Create tables with v2.0.0 schema
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS fornitori (
                    id_fornitore INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome_fornitore VARCHAR NOT NULL UNIQUE
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS richieste_offerta (
                    id_richiesta INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_emissione VARCHAR,
                    data_scadenza VARCHAR,
                    riferimento VARCHAR,
                    note_generali VARCHAR,
                    stato VARCHAR NOT NULL DEFAULT 'attiva',
                    numeri_ordine VARCHAR,
                    tipo_rdo VARCHAR NOT NULL DEFAULT 'Fornitura piena',
                    note_formattate VARCHAR,
                    username VARCHAR
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS dettagli_richiesta (
                    id_dettaglio INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_richiesta INTEGER,
                    codice_materiale VARCHAR,
                    descrizione_materiale VARCHAR,
                    quantita VARCHAR,
                    disegno VARCHAR,
                    data_consegna_richiesta VARCHAR,
                    codice_grezzo VARCHAR,
                    disegno_grezzo VARCHAR,
                    materiale_conto_lavoro VARCHAR,
                    FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta)
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS richiesta_fornitori (
                    id_richiesta INTEGER,
                    nome_fornitore VARCHAR,
                    PRIMARY KEY (id_richiesta, nome_fornitore),
                    FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta)
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS offerte_ricevute (
                    id_dettaglio INTEGER,
                    nome_fornitore VARCHAR,
                    prezzo_unitario VARCHAR,
                    PRIMARY KEY (id_dettaglio, nome_fornitore),
                    FOREIGN KEY (id_dettaglio) REFERENCES dettagli_richiesta (id_dettaglio)
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS allegati_richiesta (
                    id_allegato INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_richiesta INTEGER,
                    nome_file VARCHAR,
                    dati_file BLOB,
                    tipo_allegato VARCHAR,
                    nome_fornitore VARCHAR,
                    percorso_esterno VARCHAR,
                    data_inserimento VARCHAR DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (id_richiesta) REFERENCES richieste_offerta (id_richiesta)
                )
            ''')
            
            conn.commit()
            conn.close()
            self.logger.info("Target database schema created successfully")
            
        except sqlite3.Error as e:
            raise MigrationError(f"Failed to create target schema: {e}")
    
    def _migrate_suppliers(self, source_conn, target_conn):
        """Migrate fornitori table (extract unique supplier names from usage tables)"""
        self._update_progress(3, "Migrating suppliers...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        # First try to get from fornitori table if it exists and has data
        source_cursor.execute("SELECT id_fornitore, nome_fornitore FROM fornitori ORDER BY id_fornitore")
        suppliers = source_cursor.fetchall()
        
        # If fornitori table is empty, extract unique suppliers from usage tables
        if not suppliers:
            self.logger.info("Fornitori table is empty, extracting suppliers from usage tables...")
            
            # Get unique supplier names from richiesta_fornitori and offerte_ricevute
            source_cursor.execute("""
                SELECT DISTINCT nome_fornitore 
                FROM (
                    SELECT nome_fornitore FROM richiesta_fornitori
                    UNION
                    SELECT nome_fornitore FROM offerte_ricevute
                )
                WHERE nome_fornitore IS NOT NULL AND nome_fornitore != ''
                ORDER BY nome_fornitore
            """)
            supplier_names = source_cursor.fetchall()
            
            # Insert into target with auto-incrementing IDs
            for idx, (name,) in enumerate(supplier_names, start=1):
                target_cursor.execute(
                    "INSERT INTO fornitori (id_fornitore, nome_fornitore) VALUES (?, ?)",
                    (idx, name)
                )
            
            target_conn.commit()
            self.statistics['suppliers_migrated'] = len(supplier_names)
            self.logger.info(f"Migrated {len(supplier_names)} suppliers (extracted from usage tables)")
        else:
            # Migrate from fornitori table
            for supplier in suppliers:
                target_cursor.execute(
                    "INSERT INTO fornitori (id_fornitore, nome_fornitore) VALUES (?, ?)",
                    supplier
                )
            
            target_conn.commit()
            self.statistics['suppliers_migrated'] = len(suppliers)
            self.logger.info(f"Migrated {len(suppliers)} suppliers (from fornitori table)")
    
    def _migrate_rfqs(self, source_conn, target_conn):
        """Migrate richieste_offerta table with ID remapping and username injection"""
        self._update_progress(4, "Migrating RfQs with username assignment...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        # Get all RfQs from source
        source_cursor.execute("""
            SELECT id_richiesta, data_emissione, data_scadenza, riferimento, note_generali,
                   stato, numeri_ordine, tipo_rdo, note_formattate
            FROM richieste_offerta
            ORDER BY id_richiesta
        """)
        rfqs = source_cursor.fetchall()
        
        for rfq in rfqs:
            old_id = rfq[0]
            new_id = self.id_mapper.generate_new_id(old_id)
            
            # Normalize tipo_rdo
            tipo_rdo = normalize_rfq_type(rfq[7] if len(rfq) > 7 else None)
            
            # Insert with new ID and username
            target_cursor.execute("""
                INSERT INTO richieste_offerta (
                    id_richiesta, data_emissione, data_scadenza, riferimento, note_generali,
                    stato, numeri_ordine, tipo_rdo, note_formattate, username
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                new_id,
                rfq[1],  # data_emissione
                rfq[2],  # data_scadenza
                rfq[3],  # riferimento
                rfq[4],  # note_generali
                rfq[5] if len(rfq) > 5 else 'attiva',  # stato
                rfq[6] if len(rfq) > 6 else None,  # numeri_ordine
                tipo_rdo,
                rfq[8] if len(rfq) > 8 else None,  # note_formattate
                self.username  # username from config
            ))
        
        target_conn.commit()
        self.statistics['rfqs_migrated'] = len(rfqs)
        self.logger.info(f"Migrated {len(rfqs)} RfQs (ID remapped, username '{self.username}' assigned)")
    
    def _migrate_articles(self, source_conn, target_conn):
        """Migrate dettagli_richiesta table with FK remapping"""
        self._update_progress(5, "Migrating articles...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        source_cursor.execute("""
            SELECT id_dettaglio, id_richiesta, codice_materiale, descrizione_materiale, quantita,
                   disegno, data_consegna_richiesta, codice_grezzo, disegno_grezzo, materiale_conto_lavoro
            FROM dettagli_richiesta
            ORDER BY id_dettaglio
        """)
        articles = source_cursor.fetchall()
        
        for article in articles:
            old_rfq_id = article[1]
            new_rfq_id = self.id_mapper.get_mapping(old_rfq_id)
            
            if new_rfq_id is None:
                self.logger.warning(f"Article {article[0]} references unknown RfQ ID {old_rfq_id}, skipping")
                self.statistics['warnings'].append(f"Article {article[0]}: unknown RfQ ID {old_rfq_id}")
                continue
            
            # Keep original id_dettaglio (no remapping for detail IDs)
            target_cursor.execute("""
                INSERT INTO dettagli_richiesta (
                    id_dettaglio, id_richiesta, codice_materiale, descrizione_materiale, quantita,
                    disegno, data_consegna_richiesta, codice_grezzo, disegno_grezzo, materiale_conto_lavoro
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                article[0],  # id_dettaglio (keep original)
                new_rfq_id,  # id_richiesta (remapped)
                article[2],  # codice_materiale
                article[3],  # descrizione_materiale
                article[4],  # quantita
                article[5] if len(article) > 5 else None,  # disegno
                article[6] if len(article) > 6 else None,  # data_consegna_richiesta
                article[7] if len(article) > 7 else None,  # codice_grezzo
                article[8] if len(article) > 8 else None,  # disegno_grezzo
                article[9] if len(article) > 9 else None   # materiale_conto_lavoro
            ))
        
        target_conn.commit()
        self.statistics['articles_migrated'] = len(articles)
        self.logger.info(f"Migrated {len(articles)} articles")
    
    def _migrate_rfq_suppliers(self, source_conn, target_conn):
        """Migrate richiesta_fornitori table with FK remapping"""
        self._update_progress(6, "Migrating RfQ-supplier associations...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        source_cursor.execute("SELECT id_richiesta, nome_fornitore FROM richiesta_fornitori")
        associations = source_cursor.fetchall()
        
        migrated = 0
        for assoc in associations:
            old_rfq_id = assoc[0]
            new_rfq_id = self.id_mapper.get_mapping(old_rfq_id)
            
            if new_rfq_id is None:
                self.logger.warning(f"RfQ-supplier association references unknown RfQ ID {old_rfq_id}, skipping")
                continue
            
            target_cursor.execute(
                "INSERT INTO richiesta_fornitori (id_richiesta, nome_fornitore) VALUES (?, ?)",
                (new_rfq_id, assoc[1])
            )
            migrated += 1
        
        target_conn.commit()
        self.logger.info(f"Migrated {migrated} RfQ-supplier associations")
    
    def _migrate_prices(self, source_conn, target_conn):
        """Migrate offerte_ricevute table"""
        self._update_progress(7, "Migrating supplier prices...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        source_cursor.execute("SELECT id_dettaglio, nome_fornitore, prezzo_unitario FROM offerte_ricevute")
        prices = source_cursor.fetchall()
        
        for price in prices:
            target_cursor.execute(
                "INSERT INTO offerte_ricevute (id_dettaglio, nome_fornitore, prezzo_unitario) VALUES (?, ?, ?)",
                price
            )
        
        target_conn.commit()
        self.statistics['prices_migrated'] = len(prices)
        self.logger.info(f"Migrated {len(prices)} price entries")
    
    def _migrate_attachments(self, source_conn, target_conn):
        """Migrate allegati_richiesta table and extract BLOB/external files"""
        self._update_progress(8, "Migrating attachments (extracting BLOBs to filesystem)...")
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        # Initialize attachment extractor with source and target directories
        self.attachment_extractor = AttachmentExtractor(
            self.target_attachments_dir, 
            self.source_attachments_dir,
            self.logger
        )
        
        # Check if percorso_esterno and data_inserimento columns exist
        source_cursor.execute("PRAGMA table_info(allegati_richiesta)")
        columns = [col[1] for col in source_cursor.fetchall()]
        has_percorso_esterno = 'percorso_esterno' in columns
        has_data_inserimento = 'data_inserimento' in columns
        has_dati_file = 'dati_file' in columns
        
        # Build SELECT query based on available columns
        select_cols = ['id_allegato', 'id_richiesta', 'nome_file', 'tipo_allegato', 'nome_fornitore']
        if has_dati_file:
            select_cols.append('dati_file')
        if has_percorso_esterno:
            select_cols.append('percorso_esterno')
        if has_data_inserimento:
            select_cols.append('data_inserimento')
        
        query = f"SELECT {', '.join(select_cols)} FROM allegati_richiesta ORDER BY id_allegato"
        source_cursor.execute(query)
        attachments = source_cursor.fetchall()
        
        for idx, attachment in enumerate(attachments, 1):
            old_rfq_id = attachment[1]
            new_rfq_id = self.id_mapper.get_mapping(old_rfq_id)
            
            if new_rfq_id is None:
                self.logger.warning(f"Attachment {attachment[0]} references unknown RfQ ID {old_rfq_id}, skipping")
                continue
            
            # Build attachment data dict
            att_data = {
                'id_allegato': attachment[0],
                'nome_file': attachment[2],
                'tipo_allegato': attachment[3],
                'nome_fornitore': attachment[4],
                'dati_file': attachment[5] if has_dati_file and len(attachment) > 5 else None,
                'percorso_esterno': attachment[6] if has_percorso_esterno and len(attachment) > 6 else None
            }
            
            # Extract attachment to filesystem
            new_attachment_id = attachment[0]  # Keep original attachment ID
            relative_path = self.attachment_extractor.extract_attachment(att_data, new_rfq_id, new_attachment_id)
            
            # Determine data_inserimento value
            if has_data_inserimento and len(attachment) > 7 and attachment[7]:
                data_inserimento = attachment[7]
            else:
                # Use NULL (will default to CURRENT_TIMESTAMP)
                data_inserimento = None
            
            # Insert into target database (BLOB set to NULL, percorso_esterno set)
            target_cursor.execute("""
                INSERT INTO allegati_richiesta (
                    id_allegato, id_richiesta, nome_file, dati_file, tipo_allegato,
                    nome_fornitore, percorso_esterno, data_inserimento
                ) VALUES (?, ?, ?, NULL, ?, ?, ?, ?)
            """, (
                new_attachment_id,
                new_rfq_id,
                attachment[2],  # nome_file (original)
                attachment[3],  # tipo_allegato
                attachment[4],  # nome_fornitore
                relative_path,  # percorso_esterno (new file path)
                data_inserimento
            ))
            
            # Update progress
            if idx % 10 == 0:
                self._update_progress(8, f"Migrating attachments ({idx}/{len(attachments)})...")
        
        target_conn.commit()
        
        # Get extraction statistics
        att_stats = self.attachment_extractor.get_statistics()
        self.statistics['attachments_migrated'] = att_stats['extracted']
        self.statistics['warnings'].extend(att_stats['warnings'])
        
        self.logger.info(f"Migrated {att_stats['extracted']} attachments ({att_stats['failed']} failed)")
    
    def _verify_migration(self, target_conn):
        """Verify migration integrity"""
        self._update_progress(9, "Verifying migration integrity...")
        
        cursor = target_conn.cursor()
        
        # Check username consistency
        cursor.execute("SELECT COUNT(*) FROM richieste_offerta WHERE username IS NULL OR username = ''")
        null_usernames = cursor.fetchone()[0]
        if null_usernames > 0:
            self.logger.warning(f"Found {null_usernames} RfQs with NULL username, correcting...")
            cursor.execute("UPDATE richieste_offerta SET username = ? WHERE username IS NULL OR username = ''", (self.username,))
            target_conn.commit()
            self.statistics['warnings'].append(f"Corrected {null_usernames} NULL usernames")
        
        # Check foreign key integrity (SQLite)
        cursor.execute("PRAGMA foreign_key_check")
        fk_errors = cursor.fetchall()
        if fk_errors:
            self.logger.error(f"Foreign key integrity errors: {fk_errors}")
            self.statistics['errors'].extend([f"FK error: {e}" for e in fk_errors])
        
        # Get final counts
        cursor.execute("SELECT COUNT(*) FROM richieste_offerta")
        rfq_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM dettagli_richiesta")
        article_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM allegati_richiesta")
        attachment_count = cursor.fetchone()[0]
        
        self.logger.info(f"Verification complete: {rfq_count} RfQs, {article_count} articles, {attachment_count} attachments")
    
    def execute_migration(self):
        """
        Execute complete migration process.
        
        Returns:
            dict: Migration statistics and results
            
        Raises:
            MigrationError: If migration fails
        """
        self.logger.info("=" * 70)
        self.logger.info("STARTING DATABASE MIGRATION")
        self.logger.info(f"Source: {self.source_db_path}")
        self.logger.info(f"Target: {self.target_db_path}")
        self.logger.info(f"Username: {self.username}")
        self.logger.info("=" * 70)
        
        start_time = datetime.now()
        
        try:
            # Step 1: Prepare target folder (destructive)
            self._prepare_target_folder()
            
            # Step 2: Create target schema
            self._create_target_schema()
            
            # Open connections
            source_conn = sqlite3.connect(self.source_db_path, timeout=30.0)
            target_conn = sqlite3.connect(self.target_db_path, timeout=30.0)
            
            try:
                # Step 3-9: Migrate data
                self._migrate_suppliers(source_conn, target_conn)
                self._migrate_rfqs(source_conn, target_conn)
                self._migrate_articles(source_conn, target_conn)
                self._migrate_rfq_suppliers(source_conn, target_conn)
                self._migrate_prices(source_conn, target_conn)
                self._migrate_attachments(source_conn, target_conn)
                self._verify_migration(target_conn)
                
            finally:
                source_conn.close()
                target_conn.close()
            
            # Calculate duration
            duration = (datetime.now() - start_time).total_seconds()
            self.statistics['duration_seconds'] = duration
            
            self._update_progress(10, f"Migration completed successfully in {duration:.1f} seconds!")
            
            self.logger.info("=" * 70)
            self.logger.info("MIGRATION COMPLETED SUCCESSFULLY")
            self.logger.info(f"Duration: {duration:.1f} seconds")
            self.logger.info(f"RfQs migrated: {self.statistics['rfqs_migrated']}")
            self.logger.info(f"Articles migrated: {self.statistics['articles_migrated']}")
            self.logger.info(f"Attachments migrated: {self.statistics['attachments_migrated']}")
            self.logger.info(f"Warnings: {len(self.statistics['warnings'])}")
            self.logger.info("=" * 70)
            
            return self.statistics
            
        except Exception as e:
            self.logger.error(f"Migration failed: {e}", exc_info=True)
            self.statistics['errors'].append(str(e))
            raise MigrationError(f"Migration failed: {e}")
