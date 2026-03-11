# DataFlow Database Migration Tool

## Version 1.0.0

A specialized tool to migrate DataFlow v1.1.0 databases to DataFlow v2.0.0 format with full schema transformation, ID remapping, and attachment extraction.

---

## Overview

This migration tool automates the complete transition from DataFlow v1.1.0 (single-user, BLOB-based attachments) to DataFlow v2.0.0 (multi-user, filesystem-based attachments with year-driven ID format).

### What This Tool Does

1. **Reads Configuration**: Extracts username and target paths from DataFlow 2.0.0 `config.ini`
2. **Validates Source Database**: Ensures source database is a valid v1.1.0 schema
3. **Shows Migration Summary**: Displays detailed preview of migration operations
4. **Confirms Destructive Operation**: Requires explicit user confirmation (typing "DELETE")
5. **Migrates Data**: 
   - Creates fresh v2.0.0 database with WAL mode
   - Remaps RfQ IDs to year-based format (e.g., 2500001 for 2025)
   - Assigns username to all migrated RfQs
   - Extracts BLOB attachments to filesystem
   - Updates all foreign key relationships
6. **Generates Report**: Creates detailed log and statistics summary

---

## Prerequisites

- **Python 3.10 or higher**
- **Windows 10/11** (tool is Windows-specific due to path conventions)
- **DataFlow 2.0.0** installed and configured (at least one launch to create `config.ini`)
- **DataFlow 1.1.0 database** to migrate

### Required Python Packages

**All dependencies are built-in Python modules!** No external packages needed:
- `sqlite3` (built-in)
- `tkinter` (built-in on Windows)
- `configparser`, `os`, `shutil`, `glob`, `re` (built-in)
- `datetime`, `json`, `logging` (built-in)

---

## Installation

### Step 1: Create Virtual Environment

Open PowerShell in the `Database_Migration_Tool` folder and run:

```powershell
python -m venv dmt
```

### Step 2: Activate Virtual Environment

**PowerShell:**
```powershell
.\dmt\Scripts\Activate.ps1
```

**Command Prompt:**
```cmd
.\dmt\Scripts\activate.bat
```

### Step 3: Install Dependencies

Since all dependencies are built-in, no installation needed! But you can run:

```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

(This will essentially be a no-op, but ensures pip is up to date)

### Alternative: Automated Setup

Run the provided setup script:

**PowerShell:**
```powershell
.\setup_venv.ps1
```

**Command Prompt:**
```cmd
.\setup_venv.bat
```

---

## Usage

### Quick Start

1. **Activate virtual environment** (see Installation Step 2)
2. **Run the migration tool**:
   ```powershell
   python Database_Migration_Tool.py
   ```
3. **Follow the GUI prompts**:
   - Select `config.ini` from DataFlow 2.0.0
   - Select source v1.1.0 database file
   - Review migration summary
   - Type `DELETE` to confirm
   - Wait for migration to complete

### Step-by-Step Guide

#### Step 1: Locate DataFlow 2.0.0 Config File

The tool will prompt you to select `config.ini`. Default location:
```
C:\Users\[YourName]\AppData\Local\DataFlow\config.ini
```

**Important:** This file must contain valid user identity (`first_name`, `last_name`, or `username`). If not, run DataFlow 2.0.0 at least once to configure your identity.

#### Step 2: Select Source Database

Choose your DataFlow 1.1.0 database file (typically named `dataflow_db.db` or similar). The tool will validate that it's a compatible v1.1.0 database.

#### Step 3: Review Migration Summary

The tool displays:
- Source database statistics (RfQs, articles, suppliers, attachments)
- Target location and username
- **WARNING if target folder exists** (will be deleted!)
- Complete list of operations to be performed

Click **"⚠️ I UNDERSTAND - PROCEED"** to continue or **"Cancel"** to abort.

#### Step 4: Final Confirmation

For safety, you must type `DELETE` (case-sensitive) to confirm that the target folder will be permanently deleted. This is **irreversible**.

#### Step 5: Migration Execution

A progress window shows real-time migration status:
- Step 1: Preparing target folder (destructive deletion)
- Step 2: Creating v2.0.0 database schema
- Step 3: Migrating suppliers
- Step 4: Migrating RfQs with username assignment
- Step 5: Migrating articles
- Step 6: Migrating RfQ-supplier associations
- Step 7: Migrating prices
- Step 8: Extracting attachments from BLOBs to filesystem
- Step 9: Verifying migration integrity
- Step 10: Completion

#### Step 6: Review Results

After migration, a completion dialog shows:
- Migration duration
- Data migrated (RfQs, articles, suppliers, prices, attachments)
- Warnings (if any)
- Errors (if any)

Click **"📁 Open Database Folder"** to view the migrated database in File Explorer.

---

## Migration Behavior

### ⚠️ Destructive Operation

**CRITICAL:** This tool performs a **destructive operation** on the target folder.

If the folder `C:\Users\[YourName]\Documents\DataFlow_[username]\` exists, it will be **completely deleted** including:
- Existing database file
- All attachment files
- Any other files in that folder

The source v1.1.0 database is **never modified** (read-only).

### ID Remapping

v1.1.0 uses sequential IDs starting from 1: `1, 2, 3, ...`

v2.0.0 uses year-based IDs: `YYXXXXX` (e.g., `2500001, 2500002, ...` for 2025)

The tool automatically remaps all IDs and updates foreign key relationships in:
- `dettagli_richiesta.id_richiesta`
- `richiesta_fornitori.id_richiesta`
- `allegati_richiesta.id_richiesta`

### Username Assignment

All migrated RfQs will have the `username` field populated with the username from `config.ini`.

**Example:** If your `config.ini` has username `mrossi`, all RfQs in the migrated database will have `username = 'mrossi'`.

### Attachment Handling

v1.1.0 attachments are stored in two ways:
1. **BLOB in database** (`dati_file` column)
2. **External files** (`percorso_esterno` column)

The migration tool:
1. Extracts BLOB data to filesystem files
2. Copies external files (if they exist)
3. Generates v2.0.0-compliant filenames:
   - Supplier: `RfQ2500001_SupplierName_ID123.pdf`
   - Internal: `RfQ2500001_ID123.pdf`
4. Updates database with `percorso_esterno` (relative path)
5. Sets `dati_file = NULL` (BLOBs no longer used in v2.0.0)

**Target location:** `C:\Users\[YourName]\Documents\DataFlow_[username]\Attachments\`

### Data Normalization

The tool automatically normalizes:
- **RfQ Types**: Converts any language variant to canonical Italian
  - `"Full Supply"` → `"Fornitura piena"`
  - `"Work Order"` → `"Conto lavoro"`
- **Timestamps**: Preserves original `data_inserimento` or uses NULL (defaults to current timestamp in DB)

---

## Folder Structure

After migration, your target folder will have this structure:

```
C:\Users\[YourName]\Documents\DataFlow_[username]\
├── Database\
│   └── dataflow_db_[username].db    # Migrated database
└── Attachments\
    ├── RfQ2500001_SupplierA_ID45.pdf
    ├── RfQ2500001_ID46.docx
    └── ... (all extracted attachments)
```

---

## Logging

### Console Output

The tool prints high-level progress to the console (INFO level).

### Detailed Log File

A comprehensive log file is saved to:
```
C:\Users\[YourName]\Documents\DataFlow\Migration_Logs\migration_YYYYMMDD_HHMMSS.log
```

This file contains:
- DEBUG-level details (every SQL query, every file operation)
- Complete traceback for any errors
- Warnings for soft errors (skipped attachments, missing columns, etc.)

**Tip:** If migration fails, check this log file for diagnostic information.

---

## Troubleshooting

### Error: "Config file not found"

**Solution:** Ensure you've run DataFlow 2.0.0 at least once to generate `config.ini`. The default location is:
```
C:\Users\[YourName]\AppData\Local\DataFlow\config.ini
```

### Error: "Username is empty"

**Solution:** Open DataFlow 2.0.0, go to Settings, and configure your user identity (First Name and Last Name). The tool will generate username from these fields.

### Error: "Database validation failed: Missing required table"

**Solution:** The selected database is not a valid DataFlow v1.1.0 database. Ensure you're selecting the correct file. The database must contain tables: `fornitori`, `richieste_offerta`, `dettagli_richiesta`, `richiesta_fornitori`, `offerte_ricevute`, `allegati_richiesta`.

### Warning: "Attachment ID X has no BLOB data and no external file"

**Solution:** This is a soft error. The attachment metadata exists in the database but the actual file is missing. The tool will skip this attachment and continue migration. Check the log file for details.

### Error: "Failed to delete existing folder"

**Solution:** The target folder may be open in File Explorer or another program has locked it. Close all programs accessing that folder and retry.

### Error: "Foreign key integrity errors"

**Solution:** This indicates data corruption in the source database (orphaned records with invalid foreign keys). Review the log file to identify problematic records. You may need to manually clean the source database before migration.

---

## Advanced Configuration

### Custom Database Location

If your DataFlow 2.0.0 uses a custom database location (configured via `dataflow_base_dir` in `config.ini`), the migration tool will automatically use that location.

### Batch Migration

To migrate multiple databases, run the tool multiple times. Each run will overwrite the target folder for the specified username.

**Note:** This tool does NOT support merging multiple v1.1.0 databases into one v2.0.0 database. Each migration is one-to-one.

---

## Migration Report

After completion, the tool displays a summary with:

### Statistics
- **Duration**: Total time taken (seconds)
- **RfQs migrated**: Number of RfQs (with ID remapping)
- **Articles migrated**: Number of articles
- **Suppliers migrated**: Number of suppliers
- **Prices migrated**: Number of price entries
- **Attachments migrated**: Number of successfully extracted attachments

### Warnings (if any)
- Missing optional columns in source database
- Skipped attachments (no data available)
- External files not found
- Other non-critical issues

### Errors (if any)
- Foreign key integrity violations
- Critical failures during migration

---

## Safety Features

1. **Read-Only Source**: Source database is opened in read-only mode and never modified
2. **Two-Step Confirmation**: Summary dialog + DELETE typing requirement
3. **Detailed Warnings**: Clear explanation of destructive operations
4. **Transaction Safety**: All database operations use transactions (rollback on error)
5. **Comprehensive Logging**: Full audit trail in log file
6. **Integrity Verification**: Automatic foreign key checks after migration

---

## Technical Details

### Database Schema Changes

**v1.1.0 → v2.0.0 Key Differences:**

| Feature | v1.1.0 | v2.0.0 |
|---------|--------|--------|
| RfQ ID Format | Sequential (1, 2, 3...) | Year-based (2500001, 2500002...) |
| Username Column | Not present | Present (all RfQs) |
| Attachments Storage | BLOB in database | Files in Attachments folder |
| Journal Mode | DELETE (default) | WAL (concurrent access) |
| Column Types | TEXT | VARCHAR |

### ID Mapping Logic

```python
# v1.1.0 ID → v2.0.0 ID
old_id = 105
year = 2025
new_id = (year % 100) * 100000 + sequence
# Result: 2500001, 2500002, etc.
```

All foreign key references are updated accordingly.

### Attachment Filename Pattern

**Supplier Attachment:**
```
RfQ{new_id}_{sanitized_supplier}_ID{attachment_id}.{ext}
Example: RfQ2500001_ACMECorp_ID123.pdf
```

**Internal Document:**
```
RfQ{new_id}_ID{attachment_id}.{ext}
Example: RfQ2500001_ID124.docx
```

---

## Frequently Asked Questions

### Q: Can I migrate multiple times?

**A:** Yes, but each migration overwrites the target folder. If you need to preserve previous migrations, manually backup the target folder before re-running.

### Q: What happens if migration fails halfway?

**A:** The target folder will be in an incomplete state (partial data). The source database remains untouched. Simply delete the target folder and retry migration.

### Q: Can I cancel migration after typing DELETE?

**A:** No, once migration starts, it cannot be cancelled. The operation will complete (or fail with error). **Do not close the progress window**.

### Q: Why are attachment IDs not remapped?

**A:** Attachment IDs are kept original to simplify migration logic. Only RfQ IDs are remapped to the year-based format. This doesn't affect functionality.

### Q: Can I use the old and new databases simultaneously?

**A:** **Not recommended.** The databases use different folder structures and ID formats. Use only the migrated v2.0.0 database with DataFlow 2.0.0.

### Q: How do I verify migration success?

**A:** Open DataFlow 2.0.0, load the migrated database, and verify:
1. All RfQs are present
2. Attachments open correctly
3. No error messages

Also check the completion dialog for warnings/errors.

---

## Support

If you encounter issues not covered in this README:

1. Check the **detailed log file** in `Documents\DataFlow\Migration_Logs\`
2. Review the **Troubleshooting** section above
3. Ensure you're using **Python 3.10+** and **DataFlow 2.0.0**

---

## Version History

### Version 1.0.0 (December 2025)
- Initial release
- Full v1.1.0 → v2.0.0 migration support
- GUI-based workflow with Tkinter
- Dual logging (console + file)
- Comprehensive error handling and validation
- ID remapping with foreign key updates
- BLOB extraction to filesystem
- Username assignment to all RfQs

---

## License

This tool is part of the DataFlow ecosystem.

---

## Acknowledgments

Developed specifically for DataFlow users migrating from v1.1.0 to v2.0.0.

**Enjoy your upgraded DataFlow experience!** 🚀
