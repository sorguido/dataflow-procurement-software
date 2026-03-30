# Assessment: Conservative Extraction — dataflow.py (2026-03-30)

## File structure snapshot
- Total: 6,777 lines
- MainWindow class: 3560–6367 (~2,808 lines)
- Startup block: 6597–6777 (181 lines)

## Dead class stubs (confirmed dead, never instantiated by MainWindow)
- Lines 149–2529: AttachmentWindow, PurchaseOrderWindow, EditSuppliersWindow,
  EditReferenceWindow, LanguagePrompt, NotesWindow, SQDCAnalysisWindow (~2,380 lines)
  Note: definitions SHADOW imports, but MainWindow never instantiates them directly
- Lines 3391–3559: LicenseWindow + NewRdOTypeDialog (~169 lines, need confirmation)
- Lines 6368–6527: UserIdentityDialog, CopyProgressWindow, SplashScreen (~160 lines)
  These ARE used by startup code directly below — do not touch without careful check

## Active candidates measured
- `_export_vsm_excel`: 6202–6363, 162 lines
  Deps: self.root, self.current_username + 2 passed params. 1 DB query. 0 MainWindow method calls.
  Extraction target: services/vsm_excel_export.py
  Stub in MainWindow: self._export_vsm_excel(status, current_tree) → 1 line dispatch

- `mega_export_excel`: 5894–6201, 308 lines — HIGH coupling, do NOT extract pre-release

- `perform_autobackup`: 4001–4138, 138 lines — file I/O, retry, timer — penalized

- `restart_program`: 3832–3970, 139 lines — subprocess/os._exit — penalized

- `_load_vsm_events`: 4338–4367, ~30 lines — orchestrator, god-param risk

- `select_standard_dataflow_location` (in SettingsWindow): 2935–3390, 456 lines — EXTREME risk

## Decision
Best extraction: `_export_vsm_excel` → services/vsm_excel_export.py
Highest-leverage cleanup: delete dead stubs 149–2529 (pure deletion, ~2,380 lines)
Avoid pre-release: everything involving file I/O, subprocess, timer, config I/O
