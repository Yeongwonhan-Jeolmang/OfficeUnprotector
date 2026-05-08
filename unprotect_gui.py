#!/usr/bin/env python3
"""
unprotect_gui.py — Full-featured GUI for unprotect.py
All CLI features exposed: password/password-file/password-list, output/output-dir/in-place,
backup, no-overwrite, fail-fast, check/dry-run, JSON export, glob/recursive input,
verbose/quiet log, per-file status.

Requires:  pip install customtkinter pypdf msoffcrypto-tool lxml
Place in the same directory as unprotect.py and run:  python unprotect_gui.py
"""

from __future__ import annotations

import glob
import json
import os
import queue
import sys
import threading
import time
from pathlib import Path
from tkinter import BooleanVar, StringVar, filedialog, messagebox
import tkinter as tk

try:
    import customtkinter as ctk
except ImportError:
    print("Missing: pip install customtkinter")
    sys.exit(1)

# ── Import unprotect library ─────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from unprotect import (
        UnprotectError, check_file, unprotect_file,
        _try_password_list, _load_password_file,
    )
    _HAVE_LIB = True
except ImportError:
    _HAVE_LIB = False

# ── Palette ──────────────────────────────────────────────────────────────────
BG       = "#0b0d12"
SURFACE  = "#13161e"
CARD     = "#1a1e2a"
CARD2    = "#1f2435"
BORDER   = "#252c3f"
ACCENT   = "#4f7ef8"
ACCENT_H = "#3563d4"
SUCCESS  = "#2ecc8f"
ERROR    = "#e05c5c"
WARN     = "#e8a838"
INFO     = "#7eb8f7"
TEXT     = "#dde1ef"
TEXT_DIM = "#5a627f"
MONO     = ("Courier New", 11)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Status colours ───────────────────────────────────────────────────────────
STATUS_COLOR = {
    "unlocked":  SUCCESS,
    "open":      INFO,
    "protected": WARN,
    "failed":    ERROR,
    "skipped":   TEXT_DIM,
    "pending":   TEXT_DIM,
    "running":   ACCENT,
}

EXT_LABEL = {
    "pdf": "PDF", "xlsx": "XLS", "xlsm": "XLS",
    "xls": "XLS", "docx": "DOC", "doc": "DOC",
    "pptx": "PPT", "ppt": "PPT",
}

def _badge(path: str) -> str:
    return EXT_LABEL.get(Path(path).suffix.lstrip(".").lower(), "???")


# ─────────────────────────────────────────────────────────────────────────────
# Widgets
# ─────────────────────────────────────────────────────────────────────────────

class SectionLabel(ctk.CTkLabel):
    def __init__(self, master, text, **kw):
        super().__init__(master, text=text,
                         font=ctk.CTkFont("Courier New", 9, "bold"),
                         text_color=TEXT_DIM, **kw)


class FileRow(ctk.CTkFrame):
    def __init__(self, master, path: str, on_remove, **kw):
        super().__init__(master, fg_color=CARD, corner_radius=7,
                         border_width=1, border_color=BORDER, **kw)
        self.path = path
        self._status = "pending"
        self._status_var = StringVar(value="pending")
        self._msg_var    = StringVar(value="waiting...")
        self.columnconfigure(1, weight=1)

        badge = ctk.CTkLabel(self, text=_badge(path), width=40, height=26,
                             font=ctk.CTkFont("Courier New", 10, "bold"),
                             fg_color=BORDER, corner_radius=5, text_color=ACCENT)
        badge.grid(row=0, column=0, padx=(9, 7), pady=8)

        info = ctk.CTkFrame(self, fg_color="transparent")
        info.grid(row=0, column=1, sticky="ew")
        info.columnconfigure(0, weight=1)
        ctk.CTkLabel(info, text=Path(path).name,
                     font=ctk.CTkFont("Courier New", 11, "bold"),
                     text_color=TEXT, anchor="w").grid(row=0, column=0, sticky="ew")
        self._msg_lbl = ctk.CTkLabel(info, textvariable=self._msg_var,
                                      font=ctk.CTkFont("Courier New", 9),
                                      text_color=TEXT_DIM, anchor="w")
        self._msg_lbl.grid(row=1, column=0, sticky="ew")

        self._pill = ctk.CTkLabel(self, textvariable=self._status_var,
                                   width=76, height=22,
                                   font=ctk.CTkFont("Courier New", 9, "bold"),
                                   fg_color=BORDER, corner_radius=11,
                                   text_color=TEXT_DIM)
        self._pill.grid(row=0, column=2, padx=6)

        ctk.CTkButton(self, text="X", width=26, height=26,
                      fg_color="transparent", hover_color=BORDER,
                      text_color=TEXT_DIM, font=ctk.CTkFont("Courier New", 11),
                      command=lambda: on_remove(self)).grid(row=0, column=3, padx=(0, 7))

    def set_status(self, status: str, msg: str = ""):
        self._status = status
        self._status_var.set(status)
        self._msg_var.set(msg)
        c = STATUS_COLOR.get(status, TEXT_DIM)
        self._pill.configure(text_color=c)
        self._msg_lbl.configure(text_color=c if status != "pending" else TEXT_DIM)


class LogPanel(ctk.CTkFrame):
    LEVELS = {"debug": TEXT_DIM, "info": TEXT, "ok": SUCCESS,
              "warn": WARN, "err": ERROR, "json": INFO}

    def __init__(self, master, **kw):
        super().__init__(master, fg_color=SURFACE, corner_radius=8,
                         border_width=1, border_color=BORDER, **kw)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(8, 2))
        SectionLabel(top, text="OUTPUT LOG").pack(side="left")
        ctk.CTkButton(top, text="Copy", width=44, height=20,
                      fg_color="transparent", hover_color=BORDER,
                      font=ctk.CTkFont("Courier New", 9), text_color=TEXT_DIM,
                      command=self._copy).pack(side="right")
        ctk.CTkButton(top, text="Clear", width=44, height=20,
                      fg_color="transparent", hover_color=BORDER,
                      font=ctk.CTkFont("Courier New", 9), text_color=TEXT_DIM,
                      command=self.clear).pack(side="right", padx=(0, 2))

        self._text = tk.Text(self, bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                             font=MONO, relief="flat", bd=0, wrap="word",
                             state="disabled", height=9,
                             selectbackground=ACCENT, selectforeground=BG)
        self._text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        for name, color in self.LEVELS.items():
            self._text.tag_configure(name, foreground=color)
        self._text.tag_configure("ts", foreground=TEXT_DIM)

    def append(self, text: str, level: str = "info"):
        self._text.configure(state="normal")
        ts = time.strftime("%H:%M:%S")
        self._text.insert("end", f"[{ts}] ", "ts")
        self._text.insert("end", text + "\n", level)
        self._text.see("end")
        self._text.configure(state="disabled")

    def clear(self):
        self._text.configure(state="normal")
        self._text.delete("1.0", "end")
        self._text.configure(state="disabled")

    def _copy(self):
        content = self._text.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(content)


# ─────────────────────────────────────────────────────────────────────────────
# Tab: Files
# ─────────────────────────────────────────────────────────────────────────────

class InputTab(ctk.CTkFrame):
    def __init__(self, master, app, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._build()

    def _build(self):
        # Glob entry
        glob_frame = ctk.CTkFrame(self, fg_color=CARD, corner_radius=8,
                                  border_width=1, border_color=BORDER)
        glob_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        glob_frame.columnconfigure(1, weight=1)

        SectionLabel(glob_frame, text="GLOB PATTERN / PATH").grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 4))

        self._glob_var = StringVar()
        ctk.CTkEntry(glob_frame, textvariable=self._glob_var,
                     placeholder_text="e.g.  *.xlsx  or  /docs/**/*.pdf",
                     font=ctk.CTkFont("Courier New", 11),
                     fg_color=CARD2, border_color=BORDER, text_color=TEXT,
                     ).grid(row=1, column=0, columnspan=2, sticky="ew",
                            padx=(12, 6), pady=(0, 8))

        self._recursive_var = BooleanVar()
        ctk.CTkCheckBox(glob_frame, text="Recursive (**)",
                        variable=self._recursive_var,
                        font=ctk.CTkFont("Courier New", 10),
                        text_color=TEXT, fg_color=ACCENT, width=16,
                        ).grid(row=1, column=2, padx=(0, 6))

        ctk.CTkButton(glob_frame, text="Expand & Add", height=30, width=110,
                      font=ctk.CTkFont("Courier New", 10, "bold"),
                      fg_color=ACCENT, hover_color=ACCENT_H,
                      command=self._expand_glob,
                      ).grid(row=1, column=3, padx=(0, 12), pady=(0, 8))

        # File list
        list_card = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=8,
                                 border_width=1, border_color=BORDER)
        list_card.grid(row=1, column=0, sticky="nsew")
        list_card.columnconfigure(0, weight=1)
        list_card.rowconfigure(1, weight=1)

        hdr = ctk.CTkFrame(list_card, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        SectionLabel(hdr, text="FILES").pack(side="left")
        ctk.CTkButton(hdr, text="Browse...", height=26, width=70,
                      font=ctk.CTkFont("Courier New", 10),
                      fg_color=ACCENT, hover_color=ACCENT_H,
                      command=self._browse).pack(side="right")
        ctk.CTkButton(hdr, text="Clear all", height=26, width=66,
                      fg_color="transparent", hover_color=BORDER,
                      font=ctk.CTkFont("Courier New", 10), text_color=TEXT_DIM,
                      command=self.clear_all).pack(side="right", padx=(0, 4))

        self._scroll = ctk.CTkScrollableFrame(list_card, fg_color="transparent",
                                              scrollbar_button_color=BORDER,
                                              scrollbar_button_hover_color=ACCENT)
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self._scroll.columnconfigure(0, weight=1)

        self._rows: list[FileRow] = []
        self._empty_lbl = ctk.CTkLabel(self._scroll,
            text="No files yet. Browse or enter a glob pattern above.",
            font=ctk.CTkFont("Courier New", 10), text_color=TEXT_DIM)
        self._empty_lbl.grid(row=0, column=0, pady=30)

    def _browse(self):
        paths = filedialog.askopenfilenames(
            title="Select files",
            filetypes=[("Supported", "*.pdf *.xlsx *.xlsm *.xls *.docx *.doc *.pptx *.ppt"),
                       ("All", "*.*")])
        for p in paths:
            self.add_file(p)

    def _expand_glob(self):
        pattern = self._glob_var.get().strip()
        if not pattern:
            return
        expanded = glob.glob(pattern, recursive=self._recursive_var.get())
        if not expanded:
            messagebox.showwarning("No matches", f"No files matched:\n{pattern}")
            return
        added = 0
        for p in expanded:
            if os.path.isfile(p):
                self.add_file(p)
                added += 1
        self.app.log(f"Glob '{pattern}' -> {added} file(s) added.", "info")
        self._glob_var.set("")

    def add_file(self, path: str):
        if any(r.path == path for r in self._rows):
            return
        if hasattr(self, '_empty_lbl') and self._empty_lbl.winfo_exists():
            self._empty_lbl.grid_forget()
        row = FileRow(self._scroll, path, on_remove=self._remove)
        row.grid(row=len(self._rows), column=0, sticky="ew", pady=(0, 5))
        self._rows.append(row)

    def _remove(self, row: FileRow):
        row.destroy()
        self._rows.remove(row)
        for i, r in enumerate(self._rows):
            r.grid(row=i, column=0, sticky="ew", pady=(0, 5))
        if not self._rows:
            self._empty_lbl = ctk.CTkLabel(self._scroll,
                text="No files yet. Browse or enter a glob pattern above.",
                font=ctk.CTkFont("Courier New", 10), text_color=TEXT_DIM)
            self._empty_lbl.grid(row=0, column=0, pady=30)

    def clear_all(self):
        for r in list(self._rows):
            self._remove(r)

    @property
    def file_rows(self) -> list[FileRow]:
        return self._rows


# ─────────────────────────────────────────────────────────────────────────────
# Tab: Password
# ─────────────────────────────────────────────────────────────────────────────

class PasswordTab(ctk.CTkFrame):
    """Mutually exclusive: no password | direct | password-file | wordlist brute-force."""

    def __init__(self, master, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)
        self._mode = StringVar(value="none")
        self._build()

    def _build(self):
        def _card():
            f = ctk.CTkFrame(self, fg_color=CARD, corner_radius=8,
                             border_width=1, border_color=BORDER)
            f.columnconfigure(1, weight=1)
            return f

        # Mode selector
        mode_card = _card()
        mode_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(mode_card, text="PASSWORD SOURCE").grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 6))

        for col, (val, lbl) in enumerate([("none",   "No password"),
                                          ("direct", "Type password"),
                                          ("file",   "Password file"),
                                          ("list",   "Wordlist (brute-force)")]):
            ctk.CTkRadioButton(mode_card, text=lbl, variable=self._mode,
                               value=val, command=self._refresh,
                               font=ctk.CTkFont("Courier New", 11),
                               text_color=TEXT, fg_color=ACCENT,
                               ).grid(row=1, column=col, padx=(12, 0), pady=(0, 10))

        # Direct password
        self._pw_card = _card()
        self._pw_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._pw_card, text="PASSWORD").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 4))

        self._pw_var = StringVar()
        self._pw_entry = ctk.CTkEntry(self._pw_card, textvariable=self._pw_var,
                                      show="*", placeholder_text="enter password",
                                      font=ctk.CTkFont("Courier New", 11),
                                      fg_color=CARD2, border_color=BORDER,
                                      text_color=TEXT, width=260)
        self._pw_entry.grid(row=1, column=0, padx=(12, 6), pady=(0, 10))

        self._show_var = BooleanVar()
        ctk.CTkCheckBox(self._pw_card, text="Show", variable=self._show_var,
                        font=ctk.CTkFont("Courier New", 10),
                        text_color=TEXT_DIM, fg_color=ACCENT,
                        command=lambda: self._pw_entry.configure(
                            show="" if self._show_var.get() else "*"),
                        ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        # Password file
        self._pf_card = _card()
        self._pf_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._pf_card, text="PASSWORD FILE  (first line used)").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 4))

        self._pf_var = StringVar()
        ctk.CTkEntry(self._pf_card, textvariable=self._pf_var,
                     placeholder_text="path/to/password.txt",
                     font=ctk.CTkFont("Courier New", 11),
                     fg_color=CARD2, border_color=BORDER,
                     text_color=TEXT).grid(row=1, column=0, sticky="ew",
                                           padx=(12, 6), pady=(0, 10))
        ctk.CTkButton(self._pf_card, text="...", width=30,
                      fg_color=BORDER, hover_color=ACCENT,
                      command=self._pick_pw_file,
                      ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        # Wordlist
        self._wl_card = _card()
        self._wl_card.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._wl_card, text="WORDLIST  (one password per line — brute-force)").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 4))

        self._wl_var = StringVar()
        ctk.CTkEntry(self._wl_card, textvariable=self._wl_var,
                     placeholder_text="path/to/wordlist.txt",
                     font=ctk.CTkFont("Courier New", 11),
                     fg_color=CARD2, border_color=BORDER,
                     text_color=TEXT).grid(row=1, column=0, sticky="ew",
                                           padx=(12, 6), pady=(0, 10))
        ctk.CTkButton(self._wl_card, text="...", width=30,
                      fg_color=BORDER, hover_color=ACCENT,
                      command=self._pick_wl_file,
                      ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        self._refresh()

    def _pick_pw_file(self):
        p = filedialog.askopenfilename(title="Password file",
                                       filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            self._pf_var.set(p)

    def _pick_wl_file(self):
        p = filedialog.askopenfilename(title="Wordlist file",
                                       filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            self._wl_var.set(p)

    def _refresh(self):
        m = self._mode.get()
        self._pw_card.configure(border_color=ACCENT if m == "direct" else BORDER)
        self._pf_card.configure(border_color=ACCENT if m == "file"   else BORDER)
        self._wl_card.configure(border_color=ACCENT if m == "list"   else BORDER)

    @property
    def mode(self) -> str:
        return self._mode.get()

    def get_password(self) -> str | None:
        m = self.mode
        if m == "none":
            return None
        if m == "direct":
            return self._pw_var.get() or None
        if m == "file":
            return _load_password_file(self._pf_var.get())
        return None  # wordlist handled separately in worker

    @property
    def wordlist_path(self) -> str | None:
        return self._wl_var.get().strip() or None


# ─────────────────────────────────────────────────────────────────────────────
# Tab: Output
# ─────────────────────────────────────────────────────────────────────────────

class OutputTab(ctk.CTkFrame):
    """Mutually exclusive output modes: output-dir | in-place | explicit path (single file)."""

    def __init__(self, master, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)
        self._mode = StringVar(value="dir")
        self._build()

    def _build(self):
        def _card():
            f = ctk.CTkFrame(self, fg_color=CARD, corner_radius=8,
                             border_width=1, border_color=BORDER)
            f.columnconfigure(1, weight=1)
            return f

        # Mode
        mode_card = _card()
        mode_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(mode_card, text="OUTPUT MODE").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 6))
        for col, (val, lbl) in enumerate([("dir",     "Output directory"),
                                          ("inplace", "In-place (overwrite original)"),
                                          ("file",    "Explicit path  [single file only]")]):
            ctk.CTkRadioButton(mode_card, text=lbl, variable=self._mode,
                               value=val, command=self._refresh,
                               font=ctk.CTkFont("Courier New", 11),
                               text_color=TEXT, fg_color=ACCENT,
                               ).grid(row=1, column=col, padx=(12, 0), pady=(0, 10))

        # Output directory
        self._dir_card = _card()
        self._dir_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._dir_card,
                     text="OUTPUT DIRECTORY  (blank = unlocked_<name>.ext in same folder)").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))
        self._dir_var = StringVar()
        ctk.CTkEntry(self._dir_card, textvariable=self._dir_var,
                     placeholder_text="leave blank for same folder",
                     font=ctk.CTkFont("Courier New", 11),
                     fg_color=CARD2, border_color=BORDER, text_color=TEXT,
                     ).grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=(0, 10))
        ctk.CTkButton(self._dir_card, text="...", width=30,
                      fg_color=BORDER, hover_color=ACCENT,
                      command=lambda: self._dir_var.set(
                          filedialog.askdirectory(title="Output directory") or self._dir_var.get()),
                      ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        # In-place warning
        self._ip_card = _card()
        self._ip_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._ip_card, text="IN-PLACE").grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self._ip_card,
                     text="The original file will be overwritten. Enable Backup (below) to keep a .bak copy.",
                     font=ctk.CTkFont("Courier New", 10), text_color=WARN,
                     wraplength=500, justify="left",
                     ).grid(row=1, column=0, padx=12, pady=(0, 10), sticky="w")

        # Explicit output file (single-file mode)
        self._file_card = _card()
        self._file_card.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(self._file_card,
                     text="OUTPUT FILE PATH  (single-file mode only)").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))
        self._file_var = StringVar()
        ctk.CTkEntry(self._file_card, textvariable=self._file_var,
                     placeholder_text="e.g.  /home/user/unlocked_report.pdf",
                     font=ctk.CTkFont("Courier New", 11),
                     fg_color=CARD2, border_color=BORDER, text_color=TEXT,
                     ).grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=(0, 10))
        ctk.CTkButton(self._file_card, text="...", width=30,
                      fg_color=BORDER, hover_color=ACCENT,
                      command=lambda: self._file_var.set(
                          filedialog.asksaveasfilename(title="Output file") or self._file_var.get()),
                      ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        # Safety options
        safe_card = _card()
        safe_card.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(safe_card, text="SAFETY OPTIONS").grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 6))
        self._backup_var = BooleanVar()
        self._noover_var = BooleanVar()
        ctk.CTkCheckBox(safe_card,
                        text="Backup (.bak) before overwriting  [requires In-place mode]",
                        variable=self._backup_var,
                        font=ctk.CTkFont("Courier New", 11),
                        text_color=TEXT, fg_color=ACCENT,
                        ).grid(row=1, column=0, padx=(12, 20), pady=(0, 10))
        ctk.CTkCheckBox(safe_card,
                        text="Skip if output file already exists",
                        variable=self._noover_var,
                        font=ctk.CTkFont("Courier New", 11),
                        text_color=TEXT, fg_color=ACCENT,
                        ).grid(row=1, column=1, padx=(0, 12), pady=(0, 10))

        self._refresh()

    def _refresh(self):
        m = self._mode.get()
        self._dir_card.configure( border_color=ACCENT if m == "dir"     else BORDER)
        self._ip_card.configure(  border_color=ACCENT if m == "inplace" else BORDER)
        self._file_card.configure(border_color=ACCENT if m == "file"    else BORDER)

    @property
    def mode(self) -> str:
        return self._mode.get()

    @property
    def output_dir(self) -> str | None:
        return self._dir_var.get().strip() or None

    @property
    def output_file(self) -> str | None:
        return self._file_var.get().strip() or None

    @property
    def in_place(self) -> bool:
        return self._mode.get() == "inplace"

    @property
    def backup(self) -> bool:
        return self._backup_var.get()

    @property
    def no_overwrite(self) -> bool:
        return self._noover_var.get()


# ─────────────────────────────────────────────────────────────────────────────
# Tab: Advanced
# ─────────────────────────────────────────────────────────────────────────────

class AdvancedTab(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        def _card():
            f = ctk.CTkFrame(self, fg_color=CARD, corner_radius=8,
                             border_width=1, border_color=BORDER)
            f.columnconfigure(0, weight=1)
            return f

        # Fail-fast
        err_card = _card()
        err_card.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(err_card, text="ERROR HANDLING").grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 6))
        self._failfast_var = BooleanVar()
        ctk.CTkCheckBox(err_card, text="Fail fast  -- stop immediately on the first error",
                        variable=self._failfast_var,
                        font=ctk.CTkFont("Courier New", 11),
                        text_color=TEXT, fg_color=ACCENT,
                        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))

        # Verbosity
        log_card = _card()
        log_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(log_card, text="LOG VERBOSITY").grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 6))
        self._verbosity = StringVar(value="normal")
        vf = ctk.CTkFrame(log_card, fg_color="transparent")
        vf.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
        for val, lbl in [("quiet",   "Quiet  (errors only)"),
                         ("normal",  "Normal"),
                         ("verbose", "Verbose  (debug detail)")]:
            ctk.CTkRadioButton(vf, text=lbl, variable=self._verbosity, value=val,
                               font=ctk.CTkFont("Courier New", 11),
                               text_color=TEXT, fg_color=ACCENT,
                               ).pack(side="left", padx=(0, 20))

        # JSON export
        json_card = _card()
        json_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        SectionLabel(json_card,
                     text="JSON EXPORT  (Inspect/Check mode only -- mirrors --check --json)").grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 6))
        self._json_var = BooleanVar()
        ctk.CTkCheckBox(json_card,
                        text="Emit machine-readable JSON results and offer to save to file",
                        variable=self._json_var,
                        font=ctk.CTkFont("Courier New", 11),
                        text_color=TEXT, fg_color=ACCENT,
                        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))

    @property
    def fail_fast(self) -> bool:
        return self._failfast_var.get()

    @property
    def verbosity(self) -> str:
        return self._verbosity.get()

    @property
    def json_export(self) -> bool:
        return self._json_var.get()


# ─────────────────────────────────────────────────────────────────────────────
# Main Application
# ─────────────────────────────────────────────────────────────────────────────

class UnprotectApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Unprotect")
        self.geometry("860x820")
        self.minsize(720, 640)
        self.configure(fg_color=BG)

        self._q: queue.Queue = queue.Queue()
        self._running = False

        self._build()
        self._poll()

        if not _HAVE_LIB:
            self.log("unprotect.py not found in the same directory.", "warn")
            self.log("Place unprotect.py alongside this file and restart.", "warn")

    # ── Layout ────────────────────────────────────────────────────────────

    def _build(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="UNPROTECT",
                     font=ctk.CTkFont("Courier New", 17, "bold"),
                     text_color=TEXT).pack(side="left", padx=18)
        ctk.CTkLabel(hdr, text="PDF  XLSX  DOCX  PPTX",
                     font=ctk.CTkFont("Courier New", 9),
                     text_color=TEXT_DIM).pack(side="left", padx=4)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        # Tabview
        self._tabs = ctk.CTkTabview(
            body,
            fg_color=SURFACE,
            segmented_button_fg_color=CARD,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_H,
            segmented_button_unselected_color=CARD,
            segmented_button_unselected_hover_color=BORDER,
            text_color=TEXT,
            border_width=1, border_color=BORDER,
            corner_radius=10,
        )
        self._tabs.grid(row=0, column=0, sticky="nsew")

        for name in ("Files", "Password", "Output", "Advanced"):
            self._tabs.add(name)
            self._tabs.tab(name).configure(fg_color="transparent")
            self._tabs.tab(name).columnconfigure(0, weight=1)
            self._tabs.tab(name).rowconfigure(0, weight=1)

        self._input_tab    = InputTab(   self._tabs.tab("Files"),    app=self)
        self._password_tab = PasswordTab(self._tabs.tab("Password"))
        self._output_tab   = OutputTab(  self._tabs.tab("Output"))
        self._advanced_tab = AdvancedTab(self._tabs.tab("Advanced"))

        for tab in (self._input_tab, self._password_tab,
                    self._output_tab, self._advanced_tab):
            tab.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        # Action bar
        act = ctk.CTkFrame(body, fg_color=SURFACE, corner_radius=8,
                           border_width=1, border_color=BORDER)
        act.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        act.columnconfigure(3, weight=1)

        self._check_btn = ctk.CTkButton(
            act, text="Inspect / Check",
            width=148, height=40,
            font=ctk.CTkFont("Courier New", 11, "bold"),
            fg_color=CARD, hover_color=BORDER,
            border_color=BORDER, border_width=1,
            text_color=TEXT,
            command=self._run_check,
        )
        self._check_btn.pack(side="left", padx=(10, 6), pady=10)

        self._run_btn = ctk.CTkButton(
            act, text="Unprotect",
            width=140, height=40,
            font=ctk.CTkFont("Courier New", 12, "bold"),
            fg_color=ACCENT, hover_color=ACCENT_H,
            text_color="#fff",
            command=self._run_unprotect,
        )
        self._run_btn.pack(side="left", padx=(0, 10), pady=10)

        self._progress = ctk.CTkProgressBar(act, width=160, height=7,
                                            fg_color=BORDER, progress_color=ACCENT)
        self._progress.pack(side="left", pady=10)
        self._progress.set(0)

        self._prog_lbl = ctk.CTkLabel(act, text="",
                                      font=ctk.CTkFont("Courier New", 9),
                                      text_color=TEXT_DIM)
        self._prog_lbl.pack(side="left", padx=8)

        # Log
        self._log_panel = LogPanel(body)
        self._log_panel.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    # ── Logging ───────────────────────────────────────────────────────────

    def log(self, text: str, level: str = "info"):
        verbosity = self._advanced_tab.verbosity if hasattr(self, "_advanced_tab") else "normal"
        if verbosity == "quiet" and level not in ("err", "warn"):
            return
        if verbosity != "verbose" and level == "debug":
            return
        self._log_panel.append(text, level)

    # ── Busy state ────────────────────────────────────────────────────────

    def _set_busy(self, busy: bool):
        self._running = busy
        s = "disabled" if busy else "normal"
        self._run_btn.configure(state=s)
        self._check_btn.configure(state=s)
        if not busy:
            self._progress.set(0)
            self._prog_lbl.configure(text="")

    # ── Validation ────────────────────────────────────────────────────────

    def _validate(self, check_only: bool = False) -> bool:
        if not _HAVE_LIB:
            messagebox.showerror("Missing library", "unprotect.py not found. See log.")
            return False

        rows = self._input_tab.file_rows
        if not rows:
            messagebox.showwarning("No files", "Add at least one file first.")
            return False

        out = self._output_tab
        pw  = self._password_tab
        adv = self._advanced_tab

        if out.backup and not out.in_place:
            messagebox.showwarning(
                "Config error",
                "Backup (.bak) only works with In-place mode.\n"
                "Switch Output mode to 'In-place' or uncheck Backup.")
            return False

        if out.mode == "file" and len(rows) > 1:
            messagebox.showwarning(
                "Config error",
                "Explicit output path is single-file mode only.\n"
                "Either remove extra files or switch to 'Output directory' mode.")
            return False

        if pw.mode == "file" and not self._password_tab._pf_var.get().strip():
            messagebox.showwarning("Config error",
                                   "Password file mode selected but no file path entered.")
            return False

        if pw.mode == "list" and not pw.wordlist_path:
            messagebox.showwarning("Config error",
                                   "Wordlist mode selected but no wordlist file chosen.")
            return False

        if adv.json_export and not check_only:
            messagebox.showwarning(
                "Config error",
                "JSON export is only available with 'Inspect / Check' mode,\n"
                "not with Unprotect.")
            return False

        return True

    # ── Run Check ─────────────────────────────────────────────────────────

    def _run_check(self):
        if not self._validate(check_only=True):
            return
        self._set_busy(True)
        for r in self._input_tab.file_rows:
            r.set_status("running", "inspecting...")
        threading.Thread(target=self._worker_check, daemon=True).start()

    def _worker_check(self):
        rows      = list(self._input_tab.file_rows)
        total     = len(rows)
        do_json   = self._advanced_tab.json_export
        fail_fast = self._advanced_tab.fail_fast
        json_acc: list[dict] = []

        for i, row in enumerate(rows):
            self._q.put(("progress", i / total, f"{i}/{total}"))
            try:
                result = check_file(row.path)
                self._q.put(("row", row, result.status, result.message))
                level = {"open": "ok", "protected": "warn", "failed": "err"}.get(result.status, "info")
                self._q.put(("log", f"{Path(row.path).name}: [{result.status}] {result.message}", level))
                if do_json:
                    json_acc.append({"file": row.path, "status": result.status,
                                     "layers": result.layers, "message": result.message})
            except Exception as e:
                self._q.put(("row", row, "failed", str(e)))
                self._q.put(("log", f"{Path(row.path).name}: {e}", "err"))
                if fail_fast:
                    self._q.put(("log", "Fail-fast: stopping after first error.", "warn"))
                    break

        if do_json and json_acc:
            blob = json.dumps(json_acc, indent=2)
            self._q.put(("log", "JSON output:", "info"))
            for line in blob.splitlines():
                self._q.put(("log", line, "json"))
            self._q.put(("json_save", blob))

        self._q.put(("progress", 1.0, f"{total}/{total}"))
        self._q.put(("done", None))

    # ── Run Unprotect ─────────────────────────────────────────────────────

    def _run_unprotect(self):
        if not self._validate(check_only=False):
            return
        self._set_busy(True)
        for r in self._input_tab.file_rows:
            r.set_status("pending", "queued...")
        threading.Thread(target=self._worker_unprotect, daemon=True).start()

    def _worker_unprotect(self):
        rows      = list(self._input_tab.file_rows)
        total     = len(rows)
        out       = self._output_tab
        pw        = self._password_tab
        adv       = self._advanced_tab
        fail_fast = adv.fail_fast

        # Resolve password once (unless wordlist mode, resolved per-file)
        password: str | None = None
        if pw.mode != "list" and pw.mode != "none":
            try:
                password = pw.get_password()
            except (UnprotectError, OSError) as e:
                self._q.put(("log", f"Password error: {e}", "err"))
                self._q.put(("done", None))
                return

        counts = {"unlocked": 0, "open": 0, "skipped": 0, "failed": 0}

        for i, row in enumerate(rows):
            self._q.put(("progress", i / total, f"{i}/{total}"))
            self._q.put(("row", row, "running", "processing..."))
            try:
                if pw.mode == "list":
                    result = _try_password_list(
                        input_path=row.path,
                        wordlist_path=pw.wordlist_path,
                        output_path=out.output_file if out.mode == "file" else None,
                        in_place=out.in_place,
                        output_dir=out.output_dir if out.mode == "dir" else None,
                        backup=out.backup,
                        no_overwrite=out.no_overwrite,
                    )
                else:
                    result = unprotect_file(
                        input_path=row.path,
                        password=password,
                        output_path=out.output_file if out.mode == "file" else None,
                        in_place=out.in_place,
                        output_dir=out.output_dir if out.mode == "dir" else None,
                        backup=out.backup,
                        no_overwrite=out.no_overwrite,
                    )

                self._q.put(("row", row, result.status, result.message))
                level = {"unlocked": "ok", "skipped": "warn",
                         "failed": "err", "open": "info"}.get(result.status, "info")
                self._q.put(("log", f"{Path(row.path).name}: {result.message}", level))
                key = result.status if result.status in counts else "failed"
                counts[key] += 1

            except UnprotectError as e:
                self._q.put(("row", row, "failed", str(e)))
                self._q.put(("log", f"{Path(row.path).name}: {e}", "err"))
                counts["failed"] += 1
                if fail_fast:
                    self._q.put(("log", "Fail-fast: stopping after first error.", "warn"))
                    break
            except Exception as e:
                self._q.put(("row", row, "failed", str(e)))
                self._q.put(("log", f"{Path(row.path).name}: unexpected error -- {e}", "err"))
                counts["failed"] += 1
                if fail_fast:
                    self._q.put(("log", "Fail-fast: stopping after first error.", "warn"))
                    break

        # Batch summary (mirrors CLI)
        if total > 1:
            parts = []
            if counts["unlocked"]: parts.append(f"{counts['unlocked']} unlocked")
            if counts["open"]:     parts.append(f"{counts['open']} already open")
            if counts["skipped"]:  parts.append(f"{counts['skipped']} skipped")
            if counts["failed"]:   parts.append(f"{counts['failed']} failed")
            self._q.put(("log", "Done: " + (", ".join(parts) if parts else "nothing processed"), "info"))

        self._q.put(("progress", 1.0, f"{total}/{total}"))
        self._q.put(("done", None))

    # ── Queue poll ────────────────────────────────────────────────────────

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self.log(msg[1], msg[2])
                elif kind == "row":
                    _, row, status, message = msg
                    if row.winfo_exists():
                        row.set_status(status, message)
                elif kind == "progress":
                    self._progress.set(msg[1])
                    self._prog_lbl.configure(text=msg[2])
                elif kind == "json_save":
                    self._offer_json_save(msg[1])
                elif kind == "done":
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.after(40, self._poll)

    def _offer_json_save(self, blob: str):
        path = filedialog.asksaveasfilename(
            title="Save JSON results",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")])
        if path:
            with open(path, "w") as f:
                f.write(blob)
            self.log(f"JSON saved -> {path}", "ok")


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = UnprotectApp()
    app.mainloop()