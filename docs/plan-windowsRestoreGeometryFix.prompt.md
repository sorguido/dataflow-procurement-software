# Plan: Windows Restore Geometry Fix — MainWindow (revised, no after())

## Root Cause

Exact code flow in `dataflow.py`:

| Step | Line | What happens |
|------|------|--------------|
| 1 | ~1022 | `self.root.state("zoomed")` — window maximized on Windows |
| 2 | 4490 | `MainWindow(root)` returns, window still in `zoomed` state |
| 3 | 4491–4494 | two `time.sleep()` |
| 4 | **4497** | `calculate_center_position(root)` calls `root.update()` then `winfo_reqwidth/height()` → returns **minimum widget size**, not screen size |
| 5 | **4498** | `root.geometry(small_string)` → on Windows **this updates the internal "restore geometry"** while window is maximized; Windows saves this as the restore target |
| 6 | 4499 | `root.deiconify()` — window appears maximized (visually correct) |
| 7 | — | user drags from title bar → Windows restores to the tiny restore geometry |

On Linux: the window manager handles restore via the `-zoomed` attribute independently of the `geometry()` call. The WM does not use the `geometry()` value as a restore target. Behavior is invariant.

## Fix

Single change to `main_task`, lines 4497–4498.

Replace:
```python
geometry = calculate_center_position(root)
root.geometry(geometry)
```

With:
```python
if sys.platform == 'win32' and root.state() == 'zoomed':
    _sw = root.winfo_screenwidth()
    _sh = root.winfo_screenheight()
    _w = max(1200, int(_sw * 0.75))
    _h = max(768, int(_sh * 0.75))
    _x = (_sw - _w) // 2
    _y = (_sh - _h) // 2
    root.geometry(f'{_w}x{_h}+{_x}+{_y}')
else:
    geometry = calculate_center_position(root)
    root.geometry(geometry)
```

## File to touch

- `dataflow.py` — lines 4497–4498, inside `main_task`

## Linux guarantee

`sys.platform == 'win32'` is never `True` on Linux → Linux always executes the `else` branch, which is byte-for-byte identical to the current code. No behavioral difference, no flicker, no extra geometry call.

## Verification

1. Windows: launch app → maximized → drag from title bar → window restores to ~75% of screen, centered
2. Linux: same scenario → behavior unchanged vs pre-fix
3. Windows: if window somehow starts non-maximized (`state() != 'zoomed'`) → `else` branch fires, `calculate_center_position` called normally

## Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|-----------|
| `root.state()` returns unexpected value on Wine/VM | Low | Falls into `else` (original code) — no damage | Guard `and root.state() == 'zoomed'` |
| `winfo_screenwidth()` returns 0 on exotic virtual display | Very low | Geometry `0x0+…` → tkinter clamp → small window, no crash | Same risk as original code |
| Linux unintended activation | **Zero** | — | `sys.platform == 'win32'` guarantees total isolation |

## Rollback

Restore the two original lines at 4497–4498:
```python
geometry = calculate_center_position(root)
root.geometry(geometry)
```
No other file touched. Minimal diff.
