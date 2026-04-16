"""
Controller UI riusabile per suggerimenti nome fornitore su ttk.Entry.

Overlay non invasivo (Toplevel + Listbox) per evitare layout shift.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class SupplierNameSuggestController:
    def __init__(
        self,
        parent,
        entry: ttk.Entry,
        suggestion_provider,
        *,
        get_query_text=None,
        apply_suggestion=None,
        min_chars: int = 2,
        max_items: int = 8,
    ):
        self.parent = parent
        self.entry = entry
        self.suggestion_provider = suggestion_provider
        self.get_query_text = get_query_text or self.entry.get
        self.apply_suggestion = apply_suggestion or self._default_apply
        self.min_chars = max(1, int(min_chars))
        self.max_items = max(1, int(max_items))

        self._popup = None
        self._listbox = None
        self._suggestions: list[str] = []
        self._suspend = False
        self._hide_after_id = None

        self._bind_ids = []
        self._bind_events()

    def _bind_events(self):
        self._bind_ids.append(("<KeyRelease>", self.entry.bind("<KeyRelease>", self._on_key_release, add="+")))
        self._bind_ids.append(("<Down>", self.entry.bind("<Down>", self._on_down_key, add="+")))
        self._bind_ids.append(("<Up>", self.entry.bind("<Up>", self._on_up_key, add="+")))
        self._bind_ids.append(("<Return>", self.entry.bind("<Return>", self._on_return_key, add="+")))
        self._bind_ids.append(("<Escape>", self.entry.bind("<Escape>", self._on_escape_key, add="+")))
        self._bind_ids.append(("<FocusOut>", self.entry.bind("<FocusOut>", self._on_focus_out, add="+")))

    def destroy(self):
        self._hide_popup()
        for event_name, bind_id in self._bind_ids:
            try:
                self.entry.unbind(event_name, bind_id)
            except Exception:
                pass
        self._bind_ids = []

    def refresh(self):
        self._update_suggestions()

    def _on_key_release(self, event=None):
        if self._suspend:
            return
        keys_to_ignore = {
            "Up", "Down", "Return", "Escape", "Tab",
            "Shift_L", "Shift_R", "Control_L", "Control_R",
            "Alt_L", "Alt_R",
        }
        if event is not None and getattr(event, "keysym", "") in keys_to_ignore:
            return
        self._update_suggestions()

    def _on_down_key(self, event=None):
        if not self._listbox or not self._suggestions:
            self._update_suggestions()
        if not self._listbox or not self._suggestions:
            return None
        sel = self._listbox.curselection()
        idx = (sel[0] + 1) if sel else 0
        idx = min(idx, len(self._suggestions) - 1)
        self._select_index(idx)
        return "break"

    def _on_up_key(self, event=None):
        if not self._listbox or not self._suggestions:
            return None
        sel = self._listbox.curselection()
        idx = (sel[0] - 1) if sel else 0
        idx = max(idx, 0)
        self._select_index(idx)
        return "break"

    def _on_return_key(self, event=None):
        if not self._listbox or not self._suggestions:
            return None
        sel = self._listbox.curselection()
        if not sel:
            return None
        self._apply_index(sel[0])
        return "break"

    def _on_escape_key(self, event=None):
        self._hide_popup()
        return None

    def _on_focus_out(self, event=None):
        if self._hide_after_id is not None:
            try:
                self.entry.after_cancel(self._hide_after_id)
            except Exception:
                pass
        self._hide_after_id = self.entry.after(120, self._hide_popup_if_focus_lost)

    def _hide_popup_if_focus_lost(self):
        self._hide_after_id = None
        if self._popup is None:
            return
        focused = self.entry.focus_get()
        if focused is self.entry or focused is self._listbox:
            return
        pointer_widget = self.entry.winfo_containing(
            self.entry.winfo_pointerx(),
            self.entry.winfo_pointery(),
        )
        if pointer_widget is self._listbox:
            return
        self._hide_popup()

    def _update_suggestions(self):
        query = (self.get_query_text() or "").strip()
        if len(query) < self.min_chars:
            self._hide_popup()
            return

        suggestions = self.suggestion_provider(query) or []
        suggestions = suggestions[: self.max_items]
        if not suggestions:
            self._hide_popup()
            return

        self._suggestions = suggestions
        self._ensure_popup()
        self._listbox.delete(0, tk.END)
        for item in suggestions:
            self._listbox.insert(tk.END, item)
        self._select_index(0)
        self._place_popup()

    def _ensure_popup(self):
        if self._popup is not None:
            return
        self._popup = tk.Toplevel(self.parent)
        self._popup.withdraw()
        self._popup.overrideredirect(True)
        try:
            self._popup.attributes("-topmost", True)
        except Exception:
            pass

        frame = ttk.Frame(self._popup, borderwidth=1, relief="solid")
        frame.pack(fill="both", expand=True)

        self._listbox = tk.Listbox(
            frame,
            activestyle="none",
            exportselection=False,
            selectmode=tk.SINGLE,
        )
        self._listbox.pack(fill="both", expand=True)
        self._listbox.bind("<Button-1>", self._on_click_select)
        self._listbox.bind("<Double-Button-1>", self._on_click_select)
        self._listbox.bind("<Return>", self._on_listbox_return)
        self._listbox.bind("<Escape>", self._on_listbox_escape)

    def _place_popup(self):
        if self._popup is None or not self.entry.winfo_ismapped():
            return
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width_px = max(self.entry.winfo_width(), 240)
        width_chars = max(24, int(width_px / 7))
        self._listbox.configure(width=width_chars, height=min(len(self._suggestions), self.max_items))
        self._popup.geometry(f"+{x}+{y}")
        self._popup.deiconify()
        self._popup.lift()

    def _hide_popup(self):
        if self._popup is None:
            return
        try:
            self._popup.withdraw()
        except Exception:
            pass
        self._suggestions = []

    def _select_index(self, index: int):
        if self._listbox is None:
            return
        self._listbox.selection_clear(0, tk.END)
        self._listbox.selection_set(index)
        self._listbox.activate(index)
        self._listbox.see(index)

    def _on_click_select(self, event=None):
        if self._listbox is None:
            return "break"
        idx = None
        if event is not None and getattr(event, "y", None) is not None:
            idx = self._listbox.nearest(event.y)
        if idx is None:
            sel = self._listbox.curselection()
            if not sel:
                return "break"
            idx = sel[0]
        if idx < 0:
            return "break"
        self._apply_index(idx)
        return "break"

    def _on_listbox_return(self, event=None):
        self._on_click_select(event)
        return "break"

    def _on_listbox_escape(self, event=None):
        self._hide_popup()
        self.entry.focus_set()
        return "break"

    def _apply_index(self, idx: int):
        if idx < 0:
            return
        suggestion = None
        if self._listbox is not None and idx < self._listbox.size():
            suggestion = self._listbox.get(idx)
        elif idx < len(self._suggestions):
            suggestion = self._suggestions[idx]
        if not suggestion:
            return
        self._suspend = True
        try:
            self.apply_suggestion(suggestion)
        finally:
            self._suspend = False
        self._hide_popup()
        self.entry.focus_set()

    def _default_apply(self, suggestion: str):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, suggestion)
