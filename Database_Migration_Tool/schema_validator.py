"""
DataFlow Database Migration Tool - Schema Validator
Validates source database schema to ensure compatibility with v1.1.0
"""

import sqlite3


class SchemaError(Exception):
    """Custom exception for schema validation errors"""
    pass


def validate_v1_schema(db_path):
    """
    Validate that source database matches expected v1.1.0 schema.
    
    Checks for presence of required tables and columns.
    
    Args:
        db_path: Path to source database file
        
    Returns:
        dict: Validation results with keys:
            - valid (bool): True if schema is valid
            - tables (dict): Found tables and their columns
            - warnings (list): Non-critical issues
            
    Raises:
        SchemaError: If database is incompatible or corrupt
    """
    try:
        conn = sqlite3.connect(db_path, timeout=10.0)
        cursor = conn.cursor()
    except sqlite3.Error as e:
        raise SchemaError(f"Cannot open database: {e}")
    
    try:
        # Define required tables and their mandatory columns
        required_schema = {
            'fornitori': ['id_fornitore', 'nome_fornitore'],
            'richieste_offerta': ['id_richiesta', 'data_emissione', 'riferimento', 'stato'],
            'dettagli_richiesta': ['id_dettaglio', 'id_richiesta', 'codice_materiale', 'descrizione_materiale', 'quantita'],
            'richiesta_fornitori': ['id_richiesta', 'nome_fornitore'],
            'offerte_ricevute': ['id_dettaglio', 'nome_fornitore', 'prezzo_unitario'],
            'allegati_richiesta': ['id_allegato', 'id_richiesta', 'nome_file', 'tipo_allegato']
        }
        
        # Optional columns that may or may not exist in v1.1.0
        optional_columns = {
            'richieste_offerta': ['numeri_ordine', 'tipo_rdo', 'note_formattate', 'username'],
            'dettagli_richiesta': ['disegno', 'codice_grezzo', 'disegno_grezzo', 'materiale_conto_lavoro', 'data_consegna_richiesta'],
            'allegati_richiesta': ['percorso_esterno', 'data_inserimento', 'nome_fornitore', 'dati_file']
        }
        
        results = {
            'valid': True,
            'tables': {},
            'warnings': []
        }
        
        # Get list of all tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        existing_tables = [row[0] for row in cursor.fetchall()]
        
        # Check each required table
        for table_name, required_cols in required_schema.items():
            if table_name not in existing_tables:
                raise SchemaError(
                    f"Missing required table: {table_name}\n"
                    f"This database does not appear to be a valid DataFlow v1.1.0 database."
                )
            
            # Get columns for this table
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns_info = cursor.fetchall()
            existing_cols = [col[1] for col in columns_info]  # col[1] is column name
            
            results['tables'][table_name] = existing_cols
            
            # Check required columns
            missing_cols = [col for col in required_cols if col not in existing_cols]
            if missing_cols:
                raise SchemaError(
                    f"Table '{table_name}' is missing required columns: {', '.join(missing_cols)}\n"
                    f"This database does not appear to be a valid DataFlow v1.1.0 database."
                )
            
            # Check optional columns and note which are missing
            if table_name in optional_columns:
                for opt_col in optional_columns[table_name]:
                    if opt_col not in existing_cols:
                        results['warnings'].append(
                            f"Table '{table_name}' is missing optional column '{opt_col}' "
                            f"(will be handled during migration)"
                        )
        
        # Get row counts for summary
        results['row_counts'] = {}
        for table_name in required_schema.keys():
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            results['row_counts'][table_name] = count
        
        conn.close()
        return results
        
    except sqlite3.Error as e:
        conn.close()
        raise SchemaError(f"Database validation failed: {e}")


def get_database_summary(db_path):
    """
    Get a human-readable summary of database contents.
    
    Args:
        db_path: Path to database file
        
    Returns:
        dict: Summary with keys:
            - rfqs (int): Number of RfQs
            - articles (int): Number of articles
            - suppliers (int): Number of suppliers
            - attachments (int): Number of attachments
            - attachment_types (dict): Count by type (BLOB vs external)
    """
    try:
        conn = sqlite3.connect(db_path, timeout=10.0)
        cursor = conn.cursor()
        
        summary = {}
        
        # Count RfQs
        cursor.execute("SELECT COUNT(*) FROM richieste_offerta")
        summary['rfqs'] = cursor.fetchone()[0]
        
        # Count articles
        cursor.execute("SELECT COUNT(*) FROM dettagli_richiesta")
        summary['articles'] = cursor.fetchone()[0]
        
        # Count unique suppliers associated with RfQs (not total in fornitori table)
        cursor.execute("SELECT COUNT(DISTINCT nome_fornitore) FROM richiesta_fornitori")
        summary['suppliers'] = cursor.fetchone()[0]
        
        # Count attachments
        cursor.execute("SELECT COUNT(*) FROM allegati_richiesta")
        summary['attachments'] = cursor.fetchone()[0]
        
        # Count BLOB vs external attachments
        summary['attachment_types'] = {'blob': 0, 'external': 0, 'both': 0}
        
        cursor.execute("""
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN dati_file IS NOT NULL THEN 1 ELSE 0 END) as has_blob,
                SUM(CASE WHEN percorso_esterno IS NOT NULL AND percorso_esterno != '' THEN 1 ELSE 0 END) as has_external
            FROM allegati_richiesta
        """)
        row = cursor.fetchone()
        if row and row[0] > 0:
            # Count purely BLOB, purely external, and hybrid
            cursor.execute("""
                SELECT 
                    CASE 
                        WHEN dati_file IS NOT NULL AND (percorso_esterno IS NULL OR percorso_esterno = '') THEN 'blob'
                        WHEN (dati_file IS NULL OR dati_file = '') AND percorso_esterno IS NOT NULL AND percorso_esterno != '' THEN 'external'
                        WHEN dati_file IS NOT NULL AND percorso_esterno IS NOT NULL AND percorso_esterno != '' THEN 'both'
                        ELSE 'unknown'
                    END as type,
                    COUNT(*) as count
                FROM allegati_richiesta
                GROUP BY type
            """)
            for row in cursor.fetchall():
                att_type, count = row
                if att_type in summary['attachment_types']:
                    summary['attachment_types'][att_type] = count
        
        conn.close()
        return summary
        
    except sqlite3.Error as e:
        raise SchemaError(f"Failed to get database summary: {e}")
