# Dead-Code Audit — dataflow.py Candidate Blocks (2026-03-30)

## Key Scoping Rule

Python resolves names sequentially at module-load time. When `class X:` appears AFTER
`from module import X` in the same file, the later definition **silently overwrites** the
imported name for all code that runs after the redefinition point. This affects:
`LicenseWindow`, `NewRdOTypeDialog`, `LanguagePrompt`, `UserIdentityDialog`,
`CopyProgressWindow`, `SplashScreen` — all imported at lines 118–131, then redefined later.
Deletion removes the redefinition and lets the already-imported (identical) version take over.

---

## Complete Audit Table

| Class | Lines in dataflow.py | Extracted to | Instantiation sites | Runtime active version | Classification |
|---|---|---|---|---|---|
| `AttachmentWindow` | 149–792 | ui/windows/attachment_window.py | view_request_window.py:1331 | ui/windows version (never called from dataflow.py) | **SAFE TO DELETE** |
| `PurchaseOrderWindow` | 793–1234 | ui/windows/purchase_order_window.py | view_request_window.py:549 | ui/windows version | **SAFE TO DELETE** |
| `EditSuppliersWindow` | 1235–1325 | ui/windows/edit_suppliers_window.py | view_request_window.py:763 | ui/windows version | **SAFE TO DELETE** |
| `EditReferenceWindow` | 1326–1370 | ui/windows/edit_reference_window.py | view_request_window.py:533 | ui/windows version | **SAFE TO DELETE** |
| `LanguagePrompt` | 1371–1437 | ui/dialogs/common_dialogs.py:111 | dataflow.py:6030, 6219 | ⚠️ dataflow.py version (overwrites import at line 119) — both versions identical | **SAFE TO DELETE** |
| `NotesWindow` | 1438–1609 | ui/windows/notes_window.py | view_request_window.py:316 | ui/windows version | **SAFE TO DELETE** |
| `SQDCAnalysisWindow` | 1610–2529 | ui/windows/sqdc_analysis_window.py | view_request_window.py:491 | ui/windows version | **SAFE TO DELETE** |
| `SettingsWindow` | 2530–3390 | ❌ NOT extracted | dataflow.py:4140 | dataflow.py (live, only definition) | **DO NOT TOUCH** |
| `LicenseWindow` | 3391–3499 | ui/license_window.py | dataflow.py:3632, 3645, 6654 | ⚠️ dataflow.py version (overwrites import at line 118) — both identical | **SAFE TO DELETE** |
| `NewRdOTypeDialog` | 3500–3556 | ui/dialogs/common_dialogs.py:176 | dataflow.py:5858 | ⚠️ dataflow.py version (overwrites import at line 119) — both identical | **SAFE TO DELETE** |
| `UserIdentityDialog` | 6368–6468 | ui/dialogs/common_dialogs.py:233 | dataflow.py:3148, 3695, 6677 | ⚠️ dataflow.py version (overwrites import at line 119) | **NEEDS MANUAL REVIEW** ⚠️ |
| `CopyProgressWindow` | 6469–6525 | ui/dialogs/common_dialogs.py:332 | dataflow.py:3196 (in SettingsWindow) | ⚠️ dataflow.py version (overwrites import) — both identical | **SAFE TO DELETE** |
| `SplashScreen` | 6527–6596 | ui/dialogs/common_dialogs.py:387 | dataflow.py:6742 (startup) | ⚠️ dataflow.py version (overwrites import) — both identical | **SAFE TO DELETE** |

---

## 1. Blocks SAFE TO DELETE (11 of 13)

**Group A — Pure dead stubs** (extracted to ui/windows/, call chain goes entirely through
view_request_window.py; dataflow.py never imports or instantiates any of them):

- `AttachmentWindow` 149–792
- `PurchaseOrderWindow` 793–1234
- `EditSuppliersWindow` 1235–1325
- `EditReferenceWindow` 1326–1370
- `NotesWindow` 1438–1609
- `SQDCAnalysisWindow` 1610–2529

**Group B — Shadowing redefinitions of identical extracted copies:**

- `LanguagePrompt` 1371–1437 — after deletion, usages at lines 6030/6219 fall back to
  the already-imported common_dialogs version (identical)
- `LicenseWindow` 3391–3499 — after deletion, usages at 3632/3645/6654 use
  ui/license_window.py (identical)
- `NewRdOTypeDialog` 3500–3556 — after deletion, usage at 5858 uses common_dialogs version
  (identical)
- `CopyProgressWindow` 6469–6525 — after deletion, usage at 3196 uses common_dialogs version
  (BUG #51 fix already present in both)
- `SplashScreen` 6527–6596 — after deletion, startup usage at 6742 uses common_dialogs version
  (identical)

**Total safe-to-delete: ~2,879 lines across 11 classes.**

---

## 2. DO NOT TOUCH pre-release (1 of 13)

**`SettingsWindow`** (lines 2530–3390, 860 lines):
- **Live definition — no extracted version exists** (`ui/windows/settings_window.py` does not exist)
- Only instantiated at `dataflow.py:4140` as `SettingsWindow(self.root, self)`
- Tightly coupled to `MainWindow` via `self.main_app` parameter — accesses
  `self.main_app.db_manager`, `self.main_app.root`, etc. throughout
- Contains `select_standard_dataflow_location` (456 lines — directory copy, DB migration,
  restart): extreme risk on its own
- **Needs a dedicated extraction plan with MainWindow decoupling — post-release**

---

## 3. NEEDS MANUAL REVIEW (1 of 13)

**`UserIdentityDialog`** (lines 6368–6468):
- Extracted copy exists in `ui/dialogs/common_dialogs.py:233`
- **Behavioral difference in `_prevent_close`:**
  - `dataflow.py:6439` (current active): `SimpleMessageDialog(self, …, "warning")` — styled
  - `common_dialogs.py:316` (would become active after deletion): `messagebox.showwarning(…)` — native
- **Prerequisite fix:** Update `common_dialogs.py` line 316 to use `SimpleMessageDialog`
  (one-line change). Then the dataflow.py redefinition is safe to delete.
- Risk if fix is skipped: `_prevent_close` silently regresses to native messagebox.

---

## 4. Recommended Cleanup Sequence

**Step 1 — Zero risk (Group A, lines 149–2529):**
Delete the 6 pure dead stubs. No code anywhere calls the dataflow.py versions.
Only verification needed: app starts and view_request_window features work.

**Step 2 — Low risk (Group B + `SettingsWindow`-scoped classes):**
Delete `LanguagePrompt`, `LicenseWindow`, `NewRdOTypeDialog`, `CopyProgressWindow`,
`SplashScreen`. Each deletion shifts a usage to the already-imported, identical copy.
Smoke test required: first-run license dialog, new RdO type picker, settings DB-move dialog,
app startup splash, language picker in export flow.

**Step 2b — Needs one-line fix first:**
1. Fix `common_dialogs.py:316` → `SimpleMessageDialog` (mirrors dataflow.py behaviour)
2. Delete `UserIdentityDialog` (6368–6468)
3. Smoke test: startup identity prompt, SettingsWindow identity prompt, ensure_user_identity flow

**Post-release:**
Extract `SettingsWindow` (860 lines) with MainWindow decoupling refactor.

---

## 5. Additional Confirmed Facts

- `dataflow.py` is **never imported** by any other module in the project (no `import dataflow`
  or `from dataflow import` anywhere)
- No dynamic lookups (`getattr`, `globals()`, `importlib`) reference any of these classes
- All Group A call chains go: `MainWindow` → `ViewRequestWindow` → `ui/windows/<class>.py`
  (view_request_window.py imports directly from ui/windows/, not from dataflow.py)
- `LanguagePrompt` in `ui/kpi_window.py:20` and `view_request_window.py:618` both import
  from `common_dialogs` directly — unaffected by any dataflow.py change
