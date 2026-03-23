# VSM Persistence Layer Implementation Plan (Step 3)

## Context
Implementing database persistence for VSM events and impacts following the DELETE-REGENERATE-SAVE pattern to ensure idempotence and data integrity.

## Completed Prerequisites
- ✅ VSM Engine (services/vsm_engine.py) - 24/24 tests passing
- ✅ Business logic corrections (one-shot events, event_id=None handling)
- ✅ Models defined (VSMEvent, VSMImpact with Optional[int] for event_id)
- ✅ Comprehensive test suite (tests/test_vsm_engine.py)
- ✅ Manual verification script (test_vsm_manual.py)

## Core Requirements

### Mandatory Pattern: DELETE-REGENERATE-SAVE
Impact records must NEVER be updated in-place. Always:
1. DELETE old impacts for event_id
2. REGENERATE impacts from event using VSM Engine
3. SAVE new impacts with transaction wrapping

This ensures:
- No duplicate impacts
- Idempotent updates
- Data consistency
- Simplified debugging

## Implementation Plan

### File 1: services/vsm_persistence.py (NEW)
**Purpose**: High-level persistence layer interfacing between UI and database

**Key Functions**:
```python
def save_event_with_impacts(db_manager, event: VSMEvent) -> int:
    """
    Save new event + generate and save impacts.
    Returns: event_id (from lastrowid)
    Steps:
    1. INSERT vsm_event (without event_id)
    2. Get lastrowid → event_id
    3. Generate impacts with VSM Engine
    4. Batch INSERT impacts with transaction
    """

def update_event_with_impacts(db_manager, event: VSMEvent) -> None:
    """
    Update existing event + regenerate impacts.
    Steps:
    1. UPDATE vsm_event (event must have event_id)
    2. DELETE old impacts for event_id
    3. REGENERATE impacts with VSM Engine
    4. Batch INSERT new impacts with transaction
    """

def delete_event_and_impacts(db_manager, event_id: int) -> None:
    """
    Delete event and all related impacts.
    Steps:
    1. DELETE impacts for event_id (children first)
    2. DELETE event record (parent second)
    """
```

**Design Decisions**:
- Use existing DatabaseManager methods (to be added in File 2)
- Wrap batch operations in BEGIN TRANSACTION/COMMIT/ROLLBACK
- Log all operations with 'DataFlow.VSMPersistence' logger
- Raise VSMError on validation failures, DatabaseError on DB issues
- Import generate_impacts_for_event from services.vsm_engine

### File 2: database_manager.py (MODIFICATIONS)
**Purpose**: Add VSM tables and CRUD methods to existing DatabaseManager

**Section 2.1: Database Schema** (add to create_tables() method)
```sql
-- Table: vsm_events
CREATE TABLE IF NOT EXISTS vsm_events (
    event_id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL,
    data_evento TEXT NOT NULL,  -- ISO format YYYY-MM-DD
    voce_vsm TEXT NOT NULL,
    value_driver TEXT NOT NULL,
    value_stream TEXT NOT NULL,
    opex_ripetitivo INTEGER NOT NULL,  -- 0 or 1 (boolean)
    num_mesi_ripetizione INTEGER,
    -- ... (20+ fields total matching VSMEvent dataclass)
    FOREIGN KEY (username) REFERENCES utenti(username)
);

-- Table: vsm_impacts
CREATE TABLE IF NOT EXISTS vsm_impacts (
    impact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL,  -- Sempre richiesto per impacts persistiti
    username TEXT NOT NULL,
    anno INTEGER NOT NULL,
    mese INTEGER NOT NULL,
    tipo_valore TEXT NOT NULL,  -- 'Teorico' or 'Effettivo'
    valore_teorico REAL NOT NULL,
    valore_effettivo REAL NOT NULL,
    FOREIGN KEY (event_id) REFERENCES vsm_events(event_id),
    FOREIGN KEY (username) REFERENCES utenti(username)
);

-- Indices for performance
CREATE INDEX IF NOT EXISTS idx_vsm_impacts_event_id ON vsm_impacts(event_id);
CREATE INDEX IF NOT EXISTS idx_vsm_impacts_period ON vsm_impacts(anno, mese);
CREATE INDEX IF NOT EXISTS idx_vsm_impacts_username ON vsm_impacts(username);
CREATE INDEX IF NOT EXISTS idx_vsm_events_username ON vsm_events(username);
```

**Section 2.2: CRUD Methods** (add to DatabaseManager class)
```python
# Event CRUD
def insert_vsm_event(self, event: VSMEvent) -> int:
    """Insert event, return event_id"""
    
def update_vsm_event(self, event: VSMEvent) -> None:
    """Update existing event (requires event_id)"""
    
def delete_vsm_event(self, event_id: int) -> None:
    """Delete event by ID"""
    
def get_vsm_event_by_id(self, event_id: int) -> Optional[VSMEvent]:
    """Retrieve event by ID, return None if not found"""

# Impact CRUD
def insert_vsm_impacts_batch(self, impacts: List[VSMImpact]) -> None:
    """Batch insert impacts with transaction"""
    
def delete_vsm_impacts_by_event_id(self, event_id: int) -> None:
    """Delete all impacts for given event_id"""
    
def get_vsm_impacts_by_event_id(self, event_id: int) -> List[VSMImpact]:
    """Retrieve all impacts for event"""
    
def get_vsm_impacts_by_period(self, year: int, month: int, username: str) -> List[VSMImpact]:
    """Retrieve impacts for specific period"""
```

**Implementation Notes**:
- Follow existing pattern: INSERT returns self._get_last_insert_id()
- Use self.conn.commit() after write operations
- Wrap exceptions in DatabaseError
- Use logger.getLogger('DataFlow') for consistency
- Batch operations use BEGIN TRANSACTION/COMMIT/ROLLBACK

### File 3: tests/test_vsm_persistence.py (NEW)
**Purpose**: Comprehensive unit tests for persistence layer

**Test Coverage**:
```python
class TestVSMPersistence(unittest.TestCase):
    def setUp(self):
        """Create in-memory database with schema"""
    
    def test_save_event_with_impacts(self):
        """Verify event saved, impacts created, event_id returned"""
    
    def test_update_event_with_impacts(self):
        """Verify old impacts deleted, new impacts correct"""
    
    def test_delete_event_and_impacts(self):
        """Verify explicit deletion works (no CASCADE)"""
    
    def test_update_twice_no_duplication(self):
        """Critical: verify DELETE-REGENERATE-SAVE prevents duplicates"""
    
    def test_one_shot_event_persistence(self):
        """Verify one-shot event creates single impact"""
    
    def test_repetitive_event_persistence(self):
        """Verify repetitive event creates 24 impacts with pro-rata"""
```

**Testing Strategy**:
- Use in-memory SQLite database (:memory:)
- Create schema in setUp() using DatabaseManager
- After each save/update: query database to verify impact count
- Use assertEqual to verify: len(impacts), sum(values), first month coefficient
- Test error cases: update without event_id, delete non-existent event

## Verification Steps

### Step 1: Code Review
- [ ] Verify services/vsm_persistence.py follows DELETE-REGENERATE-SAVE pattern
- [ ] Verify database_manager.py has all 8 CRUD methods
- [ ] Verify SQL schema matches VSMEvent and VSMImpact fields
- [ ] Verify indices created for performance

### Step 2: Unit Tests
```bash
python3 -m unittest tests.test_vsm_persistence -v
```
Expected: All tests pass, no duplicates in any scenario

### Step 3: Manual Integration Test
```python
# Create database manager
db = DatabaseManager(":memory:")

# Test scenario 1: Save new event
event = VSMEvent(...)  # repetitive event
event_id = save_event_with_impacts(db, event)
impacts = db.get_vsm_impacts_by_event_id(event_id)
assert len(impacts) == 24  # repetitive event

# Test scenario 2: Update event
event.event_id = event_id
event.num_mesi_ripetizione = 12  # change duration
update_event_with_impacts(db, event)
impacts = db.get_vsm_impacts_by_event_id(event_id)
assert len(impacts) == 12  # updated duration
# SQL verification: SELECT COUNT(*) should show EXACTLY 12, not 36

# Test scenario 3: Delete event
delete_event_and_impacts(db, event_id)
impacts = db.get_vsm_impacts_by_event_id(event_id)
assert len(impacts) == 0  # cascade deletion worked
```

### Step 4: SQL Verification Queries
```sql
-- After update: verify no duplicate impacts
SELECT event_id, anno, mese, COUNT(*) as cnt
FROM vsm_impacts
WHERE event_id = ?
GROUP BY event_id, anno, mese
HAVING cnt > 1;
-- Expected: 0 rows

-- Verify FK constraints work
PRAGMA foreign_keys;  -- Should be ON
```

## Success Criteria

### Functional Requirements
- [x] Save new event → impacts generated automatically
- [x] Update event → old impacts deleted, new impacts regenerated
- [x] Delete event → cascade deletion of impacts
- [x] No duplicate impacts in any scenario
- [x] event_id=None supported for non-persisted impacts

### Non-Functional Requirements
- [x] Transaction safety (ROLLBACK on error)
- [x] Performance: batch insert for impacts (single transaction)
- [x] Logging: all operations logged with DataFlow.VSMPersistence
- [x] Error handling: VSMError for business logic, DatabaseError for DB issues
- [x] Code style: follows existing DatabaseManager patterns

### Test Requirements
- [x] All unit tests pass (tests/test_vsm_persistence.py)
- [x] Manual integration test successful
- [x] SQL verification shows no duplicates
- [x] Edge cases tested: one-shot events, event_id=None, double updates

## Next Steps After Step 3

### Step 4: UI Integration (Future)
- Add "Gestione Eventi VSM" window to UI
- Create form for VSMEvent input (all 20+ fields)
- Add table view for existing events with edit/delete actions
- Add dashboard widget showing monthly KPIs from vsm_impacts

### Step 5: Reporting & Analytics (Future)
- Monthly summary: SUM(valore_teorico), SUM(valore_effettivo) per username
- Trend charts: VSM value over time
- Comparison: Teorico vs Effettivo gap analysis
- Export to Excel/CSV functionality

## Technical Debt & Risks

### Known Limitations
- No soft delete (permanent deletion)
- No audit trail (who changed what when)
- No bulk import from CSV
- No validation for overlapping events

### Mitigation Strategies
- For audit: add created_at, modified_at, modified_by columns (future)
- For soft delete: add is_deleted flag (future)
- For bulk import: add batch wrapper in persistence layer (future)
- For overlap validation: add business rule check in persistence layer (future)

## Dependencies

### Internal
- services/vsm_engine.py (generate_impacts_for_event)
- models/vsm_event.py (VSMEvent dataclass)
- models/vsm_impact.py (VSMImpact dataclass)
- database_manager.py (DatabaseManager class)
- constants.py (existing project constants)

### External
- None (stdlib only: sqlite3, logging, datetime, dataclasses)

## Rollback Plan
If Step 3 fails or needs to be reverted:
1. Remove services/vsm_persistence.py
2. Remove VSM methods from database_manager.py
3. Remove VSM tables from database (or leave empty)
4. services/vsm_engine.py remains intact (can work standalone)
5. No UI changes yet, so nothing to revert in UI layer

## Documentation Updates
After successful implementation:
- Update docs/VSM_ENGINE_IMPLEMENTATION_SUMMARY.md (add Step 3 section)
- Add docs/VSM_PERSISTENCE_API.md (API reference for persistence layer)
- Update README.md (add VSM module to features list)
