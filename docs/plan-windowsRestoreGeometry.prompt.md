# Plan: Windows Restore Geometry Fix — MainWindow

## Root Cause

The startup sequence in `MainWindow`:

1. `MainWindow.__init__` maximizes the window (`state("zoomed")` on Windows).
2. `main_task` calls `calculate_center_position(root)` which invokes `root.update()` then `winfo_reqwidth() / winfo_reqheight()` — **while in zoomed state these return the natural minimum widget size**, not the screen size.
3. `main_task` calls `root.geometry(small_string)` **while the window is still zoomed**.
4. On Windows this geometry call doesn't change the visible appearance (still maximized), but **updates Windows' internal "restore geometry"** with the tiny calculated dimensions.
5. User drags from title bar → Windows restores the window to the tiny restore geometry.

On Linux the window manager handles the restore independently of the `geometry()` call (the `-zoomed` attribute is managed separately), so it is unaffected.

## Current State

- No `<Configure>` bind, no save/restore logic, no `minsize`/`maxsize` in `MainWindow`.
- Only window-state logic: the startup maximize block at lines 1015–1025.
- `calculate_center_position` is in `utils/window_utils.py` (out of scope).

## Minimal Fix

Add ~10 lines to `MainWindow.__init__`, Windows-only, after the maximize block.

### File to touch

- `dataflow.py` — `MainWindow.__init__`, immediately after the maximize block (after line ~1025)

### Method(s) to add / modify

No new methods. Add a Windows-only `after(0)` one-shot callback inside `__init__`:

```python
# Windows-only: overwrite the bad "restore geometry" written by startup's geometry() call.
# after(0) fires after main_task completes and returns control to the event loop,
# so it runs after root.geometry() has already set the bad value.
if sys.platform == 'win32':
    def _fix_restore_geometry():
        try:
            if not self.root.winfo_exists():
                return
            if self.root.state() != 'zoomed':
                return
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            w = max(1200, int(sw * 0.75))
            h = max(768, int(sh * 0.75))
            x = (sw - w) // 2
            y = (sh - h) // 2
            self.root.geometry(f'{w}x{h}+{x}+{y}')
        except Exception:
            pass
    self.root.after(0, _fix_restore_geometry)
```

### Logic

- Condition `state() == 'zoomed'`: skip if the window is already in normal state (user un-maximized before the callback ran — effectively impossible during splash, but safe).
- Geometry: 75% of screen, minimum 1200×768, centered. Sensible starting point for any resolution.
- `after(0)`: deferred to next event-loop iteration, which is **after** `main_task` finishes (including its `root.geometry(bad)` call).
- `sys.platform == 'win32'`: entire block is no-op on Linux/macOS; Linux behavior unchanged.
- `try/except`: silently swallowed; no UX impact if tkinter is in a transient state.

### Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|-----------|
| User un-maximizes during splash (~200 ms gap) | Near zero | Callback is no-op (`state() != 'zoomed'` guard) | Guard already in place |
| tkinter internal error on exotic WM | Very low | Silent fail via `try/except` | Guard already in place |
| Linux unintended activation | None | `sys.platform == 'win32'` prevents it | Platform check |
| Geometry not persisted across sessions | By design | Out of scope | Future extension if needed |

### Rollback

Remove the `if sys.platform == 'win32':` block (~12 lines) from `MainWindow.__init__`. No other files touched.

## Exact Insertion Point

[dataflow.py](dataflow.py#L1025) — after the closing `pass` of the maximize try/except block, before `self.all_users_placeholder = ...`
