"""
DataFlow Database Migration Tool - Attachment Extractor
Extracts attachments from v1.1.0 BLOBs and external files to v2.0.0 filesystem
"""

import os
import re
import shutil


class AttachmentError(Exception):
    """Custom exception for attachment extraction errors"""
    pass


def sanitize_filename(name):
    """
    Sanitize filename by removing invalid characters.
    
    Args:
        name: Original filename
        
    Returns:
        str: Sanitized filename safe for filesystem
    """
    if not name:
        return "unknown"
    
    # Remove or replace invalid characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = name.strip()
    
    if not name:
        return "unknown"
    
    return name


def sanitize_supplier_name(supplier):
    """
    Sanitize supplier name for use in filename.
    
    Args:
        supplier: Supplier name
        
    Returns:
        str: Sanitized supplier name
    """
    if not supplier or supplier.lower() == 'interno':
        return None
    
    # Remove special characters, keep only alphanumeric
    sanitized = re.sub(r'[^a-zA-Z0-9]', '', supplier)
    
    if not sanitized:
        return "Supplier"
    
    return sanitized


def get_file_extension(filename):
    """
    Extract file extension from filename.
    
    Args:
        filename: Original filename
        
    Returns:
        str: Extension with dot (e.g., '.pdf') or empty string
    """
    if not filename:
        return ''
    
    _, ext = os.path.splitext(filename)
    return ext.lower()


def generate_attachment_filename(rfq_id, attachment_id, supplier_name, original_filename, tipo_allegato):
    """
    Generate v2.0.0 compliant attachment filename.
    
    Pattern:
    - Supplier attachments: RfQ{rfq_id}_{supplier}_ID{attachment_id}.{ext}
    - Internal docs: RfQ{rfq_id}_ID{attachment_id}.{ext}
    
    Args:
        rfq_id: New v2.0.0 RfQ ID
        attachment_id: New attachment ID
        supplier_name: Supplier name (sanitized) or None for internal
        original_filename: Original file name (for extension)
        tipo_allegato: Attachment type ('Offerta Fornitore' or 'Documento Interno')
        
    Returns:
        str: Generated filename
    """
    ext = get_file_extension(original_filename) or '.bin'
    
    # Sanitize supplier name
    sanitized_supplier = sanitize_supplier_name(supplier_name)
    
    if sanitized_supplier and tipo_allegato == 'Offerta Fornitore':
        # Supplier attachment
        filename = f"RfQ{rfq_id}_{sanitized_supplier}_ID{attachment_id}{ext}"
    else:
        # Internal document
        filename = f"RfQ{rfq_id}_ID{attachment_id}{ext}"
    
    return filename


def extract_blob_to_file(blob_data, target_path):
    """
    Write BLOB data to file.
    
    Args:
        blob_data: Binary data from database
        target_path: Full path where to save file
        
    Returns:
        bool: True if successful
        
    Raises:
        AttachmentError: If write fails
    """
    if not blob_data:
        raise AttachmentError("BLOB data is empty")
    
    try:
        # Ensure parent directory exists
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        
        # Write binary data
        with open(target_path, 'wb') as f:
            f.write(blob_data)
        
        return True
        
    except OSError as e:
        raise AttachmentError(f"Failed to write file {target_path}: {e}")


def copy_external_file(source_path, target_path):
    """
    Copy external attachment file from old location to new location.
    
    Args:
        source_path: Original file path
        target_path: New file path in Attachments folder
        
    Returns:
        bool: True if successful, False if source not found
        
    Raises:
        AttachmentError: If copy fails
    """
    if not os.path.exists(source_path):
        return False
    
    try:
        # Ensure parent directory exists
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        
        # Copy file
        shutil.copy2(source_path, target_path)
        
        return True
        
    except OSError as e:
        raise AttachmentError(f"Failed to copy file from {source_path} to {target_path}: {e}")


class AttachmentExtractor:
    """
    Manages extraction of attachments during migration.
    """
    
    def __init__(self, target_attachments_dir, source_attachments_dir, logger):
        """
        Initialize attachment extractor.
        
        Args:
            target_attachments_dir: Path to v2.0.0 Attachments folder
            source_attachments_dir: Path to v1.1.0 Attachments folder (where external files are stored)
            logger: Logger instance
        """
        self.target_dir = target_attachments_dir
        self.source_dir = source_attachments_dir
        self.logger = logger
        self.extracted_count = 0
        self.failed_count = 0
        self.warnings = []
    
    def extract_attachment(self, attachment_data, new_rfq_id, new_attachment_id):
        """
        Extract a single attachment (from BLOB or external file).
        
        Args:
            attachment_data: Dict with keys:
                - id_allegato: Original attachment ID
                - nome_file: Original filename
                - dati_file: BLOB data (may be None)
                - percorso_esterno: External path (may be None)
                - tipo_allegato: Attachment type
                - nome_fornitore: Supplier name
            new_rfq_id: New v2.0.0 RfQ ID
            new_attachment_id: New attachment ID to use
            
        Returns:
            str: Relative path to saved file (for percorso_esterno column) or None if failed
        """
        original_filename = attachment_data.get('nome_file') or 'unknown'
        blob_data = attachment_data.get('dati_file')
        external_path = attachment_data.get('percorso_esterno')
        tipo_allegato = attachment_data.get('tipo_allegato', 'Documento Interno')
        supplier = attachment_data.get('nome_fornitore')
        
        # Generate new filename
        new_filename = generate_attachment_filename(
            new_rfq_id,
            new_attachment_id,
            supplier,
            original_filename,
            tipo_allegato
        )
        
        target_path = os.path.join(self.target_dir, new_filename)
        
        # Try to extract from external file first (if exists)
        if external_path and external_path.strip():
            # Log original path for debugging
            self.logger.debug(f"Processing external_path: '{external_path}', source_dir: '{self.source_dir}'")
            
            # Normalize path separators
            external_path_normalized = external_path.replace('\\', os.sep).replace('/', os.sep)
            
            # Build absolute path from source Attachments directory
            if os.path.isabs(external_path_normalized):
                # Already absolute path
                source_file_path = external_path_normalized
                self.logger.debug(f"Absolute path detected: {source_file_path}")
            else:
                # Check if path includes attachment folder name (Attachments or Allegati)
                path_parts = external_path_normalized.split(os.sep)
                self.logger.debug(f"Path parts: {path_parts}")
                
                if len(path_parts) > 1 and path_parts[0].lower() in ['attachments', 'allegati']:
                    # Path like "Attachments\file.xlsx" or "Allegati\file.xlsx"
                    # Extract just the filename (skip folder name)
                    filename = os.sep.join(path_parts[1:])
                    source_file_path = os.path.join(self.source_dir, filename)
                    self.logger.debug(f"Path includes folder prefix, using filename: {filename}, result: {source_file_path}")
                else:
                    # Just filename - combine with source Attachments directory
                    source_file_path = os.path.join(self.source_dir, external_path_normalized)
                    self.logger.debug(f"Simple filename, combining with source_dir: {source_file_path}")
            
            # Normalize the final path
            source_file_path = os.path.normpath(source_file_path)
            self.logger.debug(f"Final normalized path: {source_file_path}, exists: {os.path.exists(source_file_path)}")
            
            # List files in directory for debugging if file not found
            if not os.path.exists(source_file_path):
                check_dir = os.path.dirname(source_file_path)
                if os.path.exists(check_dir):
                    files = os.listdir(check_dir)
                    self.logger.debug(f"Files in {check_dir}: {files[:10]}")  # First 10 files
                else:
                    self.logger.debug(f"Directory does not exist: {check_dir}")
            
            try:
                if copy_external_file(source_file_path, target_path):
                    self.logger.info(f"✓ Copied external file: {os.path.basename(source_file_path)}")
                    self.extracted_count += 1
                    return new_filename
                else:
                    self.logger.warning(f"External file not found: {source_file_path} (original: {external_path}), trying BLOB...")
                    self.warnings.append(f"External file not found: {source_file_path}")
            except AttachmentError as e:
                self.logger.warning(f"Failed to copy external file: {e}, trying BLOB...")
                self.warnings.append(str(e))
        
        # Try to extract from BLOB
        if blob_data:
            try:
                extract_blob_to_file(blob_data, target_path)
                self.logger.debug(f"Extracted BLOB to: {new_filename}")
                self.extracted_count += 1
                return new_filename
            except AttachmentError as e:
                self.logger.error(f"Failed to extract BLOB for attachment ID {attachment_data.get('id_allegato')}: {e}")
                self.warnings.append(f"Failed to extract attachment ID {attachment_data.get('id_allegato')}: {e}")
                self.failed_count += 1
                return None
        
        # No data available
        self.logger.warning(f"Attachment ID {attachment_data.get('id_allegato')} has no BLOB data and no external file")
        self.warnings.append(f"Attachment ID {attachment_data.get('id_allegato')} ({original_filename}): No data available")
        self.failed_count += 1
        return None
    
    def get_statistics(self):
        """
        Get extraction statistics.
        
        Returns:
            dict: Statistics with extracted_count, failed_count, warnings
        """
        return {
            'extracted': self.extracted_count,
            'failed': self.failed_count,
            'warnings': self.warnings.copy()
        }
