# 13 – Logs and Diagnostics

## Purpose of the Log File

DataFlow automatically records significant operations and errors in a log file. The log is useful for:

- Diagnosing errors that occur at startup or during normal use.
- Technical support: the log file is the first piece of information requested when submitting a bug report.
- Verifying that a scheduled operation (such as the automatic backup) completed successfully.

---

## Log File Location

| Operating System | Path |
|------------------|------|
| **Windows** | `C:\Users\[username]\AppData\Local\DataFlow\dataflow.log` |
| **Linux** | `~/.local/share/DataFlow/dataflow.log` |

On Windows, the `AppData\Local` folder is hidden by default. To access it, type `%LOCALAPPDATA%\DataFlow\` into the File Explorer address bar.

---

## Automatic Log Rotation

The log file has a maximum size of **5 MB**. When it reaches the limit, it is renamed to `dataflow.log.1` and a new file is created. DataFlow keeps a maximum of **3 backup files**:

```
dataflow.log       ← most recent
dataflow.log.1     ← backup 1
dataflow.log.2     ← backup 2
dataflow.log.3     ← oldest backup
```

Files beyond the third backup are deleted automatically.

---

## Reading the Log File

The log file is a plain text file, openable with Notepad or any text editor. Each line follows this format:

```
2026-04-04 09:15:32,123 - INFO - [module] - Descriptive message
2026-04-04 09:15:33,456 - ERROR - [module] - Error message
```

### Severity Levels

| Level | Meaning |
|-------|---------|
| `INFO` | Normal operation recorded for informational purposes |
| `WARNING` | Unusual situation that did not prevent the operation from completing |
| `ERROR` | An error that prevented an operation; requires attention |
| `CRITICAL` | A severe error that caused the application to shut down |

---

## Operations Logged

The log file records (non-exhaustive list):

- Application startup and version number
- Database open (path and outcome)
- RFQ creation, modification, and deletion
- Value Stream Mapping event creation, modification, and deletion
- Backup operations (manual and automatic) with outcome
- Data validation errors
- Database write errors
- Attachment open and save operations
- Interface language changes
- Database path changes

---

## Automatic Temporary File Cleanup

At startup, DataFlow silently removes:

- `_MEI*` temporary files left behind by previous PyInstaller sessions that did not close cleanly.
- Files with the prefix `tmp*` that are older than 24 hours in the system temporary folder.

This process runs silently in the background and does not affect startup time.

---

## Reporting a Problem Using the Logs

When submitting a bug report to technical support:

1. Open the log folder (see the path table above).
2. Open `dataflow.log` with a text editor.
3. Look for lines containing `ERROR` or `CRITICAL` around the time the problem occurred.
4. Attach the entire `dataflow.log` file to the report — do not paste only a few lines, as the full context is necessary for diagnosis.

See the [Support](14-Support.md) section for where to submit reports.
