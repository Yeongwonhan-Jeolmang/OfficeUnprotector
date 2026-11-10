#!/usr/bin/env python3
"""
unprotect_gui.py — PyQt6 desktop front-end for unprotect.py

Exposes every CLI feature: password / password-file / wordlist, output-dir /
in-place / explicit-path, backup, no-overwrite, fail-fast, check (dry-run),
JSON export, glob / recursive input, verbosity, parallel jobs (--jobs) and
legacy conversion (--convert).

Requires:  pip install PyQt6 pypdf msoffcrypto-tool lxml
Place alongside unprotect.py and run:  python unprotect_gui.py
"""

from __future__ import annotations

import glob
import json
import os
import shutil
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from PyQt6.QtCore import Qt, QObject, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QGuiApplication
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QLineEdit, QCheckBox, QSpinBox, QComboBox,
    QFileDialog, QMessageBox, QScrollArea, QSplitter, QGroupBox,
    QHBoxLayout, QVBoxLayout, QGridLayout, QFormLayout, QPlainTextEdit,
    QProgressBar, QSizePolicy,
)

# ── Import the unprotect library ─────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from unprotect import (
        UnprotectError, check_file, unprotect_file,
        _try_password_list, _load_password_file,
    )
    _HAVE_LIB = True
except ImportError:
    _HAVE_LIB = False

    class UnprotectError(Exception):  # fallback so the window still builds
        code = 1


# ── Palette ──────────────────────────────────────────────────────────────────
BG       = "#15171c"
SURFACE  = "#1c1f27"
ELEVATED = "#232733"
BORDER   = "#2f3340"
ACCENT   = "#5b8cff"
ACCENT_H = "#6f9bff"
ACCENT_D = "#3f6ad6"
TEXT     = "#e7e9f0"
MUTED    = "#8b92a6"
FAINT    = "#5c6378"
SUCCESS  = "#4fc38a"
WARN     = "#e3a93c"
ERROR    = "#e86374"
INFO     = "#73a9f2"

STATUS_COLOR = {
    "unlocked":  SUCCESS,
    "open":      INFO,
    "protected": WARN,
    "failed":    ERROR,
    "skipped":   MUTED,
    "pending":   FAINT,
    "running":   ACCENT,
}

EXT_LABEL = {
    "pdf": "PDF", "xlsx": "XLSX", "xlsm": "XLSM", "xls": "XLS",
    "docx": "DOCX", "doc": "DOC", "pptx": "PPTX", "ppt": "PPT",
}
SUPPORTED_EXTS = {".pdf", ".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pptx", ".ppt"}


def _badge(path: str) -> str:
    return EXT_LABEL.get(Path(path).suffix.lstrip(".").lower(), "FILE")


# ── Global stylesheet ────────────────────────────────────────────────────────
STYLESHEET = f"""
QWidget {{
    background-color: {BG};
    color: {TEXT};
    font-family: "Segoe UI", "Inter", "Helvetica Neue", Arial, sans-serif;
    font-size: 13px;
}}
QLabel#Title {{ font-size: 18px; font-weight: 700; letter-spacing: 0.5px; }}
QLabel#Subtitle {{ color: {MUTED}; font-size: 11px; letter-spacing: 1px; }}
QLabel#Hint {{ color: {MUTED}; font-size: 11px; }}
QLabel#Warn {{ color: {WARN}; font-size: 11px; }}

QFrame#Header, QFrame#ActionBar {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}

QGroupBox {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
    margin-top: 14px;
    padding: 10px 12px 12px 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 4px;
    color: {MUTED};
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
}}

QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
    background-color: {ELEVATED};
    border: 1px solid {BORDER};
    border-radius: 7px;
    padding: 6px 9px;
    selection-background-color: {ACCENT};
    selection-color: {BG};
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{ border: 1px solid {ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background-color: {ELEVATED};
    border: 1px solid {BORDER};
    selection-background-color: {ACCENT};
    selection-color: {BG};
    outline: none;
}}

QPushButton {{
    background-color: {ELEVATED};
    border: 1px solid {BORDER};
    border-radius: 7px;
    padding: 7px 14px;
    color: {TEXT};
}}
QPushButton:hover {{ border: 1px solid {ACCENT}; }}
QPushButton:pressed {{ background-color: {BORDER}; }}
QPushButton:disabled {{ color: {FAINT}; border-color: {BORDER}; }}

QPushButton#Primary {{
    background-color: {ACCENT};
    border: none;
    color: #ffffff;
    font-weight: 700;
}}
QPushButton#Primary:hover {{ background-color: {ACCENT_H}; }}
QPushButton#Primary:pressed {{ background-color: {ACCENT_D}; }}
QPushButton#Primary:disabled {{ background-color: {BORDER}; color: {FAINT}; }}

QPushButton#Ghost {{
    background-color: transparent;
    border: none;
    color: {MUTED};
    padding: 4px 8px;
}}
QPushButton#Ghost:hover {{ color: {TEXT}; }}

QCheckBox, QRadioButton {{ spacing: 7px; }}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {BORDER};
    background-color: {ELEVATED};
}}
QCheckBox::indicator {{ border-radius: 4px; }}
QRadioButton::indicator {{ border-radius: 9px; }}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
    background-color: {ACCENT};
    border: 1px solid {ACCENT};
}}
QRadioButton:disabled, QCheckBox:disabled {{ color: {FAINT}; }}

QProgressBar {{
    background-color: {ELEVATED};
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{ background-color: {ACCENT}; border-radius: 4px; }}

QScrollArea {{ border: none; background: transparent; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {BORDER}; border-radius: 5px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {ACCENT_D}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {BORDER}; border-radius: 5px; min-width: 30px; }}

QSplitter::handle {{ background: transparent; }}
"""


# ─────────────────────────────────────────────────────────────────────────────
# File row
# ─────────────────────────────────────────────────────────────────────────────
class FileRow(QFrame):
    remove_requested = pyqtSignal(object)

    def __init__(self, path: str):
        super().__init__()
        self.path = path
        self._status = "pending"
        self.setObjectName("FileRow")
        self.setStyleSheet(
            f"#FileRow {{ background-color: {ELEVATED}; border: 1px solid {BORDER};"
            f" border-radius: 8px; }}"
        )

        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 8, 8, 8)
        lay.setSpacing(10)

        self._badge = QLabel(_badge(path))
        self._badge.setFixedWidth(46)
        self._badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._badge.setStyleSheet(
            f"background-color: {BORDER}; color: {ACCENT}; border-radius: 5px;"
            f" padding: 4px 0; font-weight: 700; font-size: 10px;"
        )
        lay.addWidget(self._badge)

        info = QVBoxLayout()
        info.setSpacing(1)
        self._name = QLabel(Path(path).name)
        self._name.setStyleSheet("font-weight: 600;")
        self._name.setToolTip(path)
        self._msg = QLabel("waiting")
        self._msg.setStyleSheet(f"color: {FAINT}; font-size: 11px;")
        info.addWidget(self._name)
        info.addWidget(self._msg)
        lay.addLayout(info, 1)

        self._pill = QLabel("pending")
        self._pill.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._pill.setFixedHeight(22)
        self._pill.setMinimumWidth(78)
        self._style_pill(FAINT)
        lay.addWidget(self._pill)

        close = QPushButton("✕")
        close.setObjectName("Ghost")
        close.setFixedSize(26, 26)
        close.setCursor(Qt.CursorShape.PointingHandCursor)
        close.clicked.connect(lambda: self.remove_requested.emit(self))
        lay.addWidget(close)

    def _style_pill(self, color: str):
        self._pill.setStyleSheet(
            f"color: {color}; border: 1px solid {color}; border-radius: 11px;"
            f" padding: 2px 10px; font-size: 10px; font-weight: 700;"
        )

    def set_status(self, status: str, msg: str = ""):
        self._status = status
        self._pill.setText(status)
        c = STATUS_COLOR.get(status, FAINT)
        self._style_pill(c)
        if msg:
            self._msg.setText(msg)
        self._msg.setStyleSheet(
            f"color: {c if status not in ('pending', 'running') else FAINT};"
            f" font-size: 11px;"
        )

    @property
    def status(self) -> str:
        return self._status


# ─────────────────────────────────────────────────────────────────────────────
# Input panel (files)
# ─────────────────────────────────────────────────────────────────────────────
class InputPanel(QWidget):
    changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._rows: list[FileRow] = []
        self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(10)

        # Glob row
        glob_box = QGroupBox("Add files")
        gl = QGridLayout(glob_box)
        gl.setContentsMargins(12, 16, 12, 12)
        gl.setSpacing(8)

        self._glob = QLineEdit()
        self._glob.setPlaceholderText("glob or path, e.g.  *.xlsx  —  press Enter to add")
        self._glob.returnPressed.connect(self._expand_glob)
        gl.addWidget(self._glob, 0, 0, 1, 2)

        self._recursive = QCheckBox("Recursive (**)")
        gl.addWidget(self._recursive, 0, 2)

        add_btn = QPushButton("Expand")
        add_btn.clicked.connect(self._expand_glob)
        gl.addWidget(add_btn, 0, 3)

        browse = QPushButton("Browse…")
        browse.clicked.connect(self._browse)
        gl.addWidget(browse, 1, 2)

        clear = QPushButton("Clear all")
        clear.setObjectName("Ghost")
        clear.clicked.connect(self.clear_all)
        gl.addWidget(clear, 1, 3)
        gl.setColumnStretch(0, 1)
        outer.addWidget(glob_box)

        # File list (drop target)
        list_box = QGroupBox("Queue")
        lv = QVBoxLayout(list_box)
        lv.setContentsMargins(10, 16, 10, 10)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setAcceptDrops(True)
        self._scroll.installEventFilter(self)  # not strictly needed; DnD on self

        self._list_host = QWidget()
        self._list_lay = QVBoxLayout(self._list_host)
        self._list_lay.setContentsMargins(2, 2, 2, 2)
        self._list_lay.setSpacing(6)
        self._list_lay.addStretch(1)

        self._empty = QLabel("Drop files here, browse, or expand a glob pattern.")
        self._empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._empty.setStyleSheet(f"color: {FAINT}; padding: 40px;")
        self._list_lay.insertWidget(0, self._empty)

        self._scroll.setWidget(self._list_host)
        lv.addWidget(self._scroll)
        outer.addWidget(list_box, 1)

        # enable drag & drop on the whole panel
        self.setAcceptDrops(True)

    # ── Drag & drop ──────────────────────────────────────────────────────────
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        added = 0
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p and os.path.isfile(p):
                before = len(self._rows)
                self.add_file(p)
                added += len(self._rows) - before
        if added:
            e.acceptProposedAction()

    # ── Actions ──────────────────────────────────────────────────────────────
    def _browse(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select files", "",
            "Supported (*.pdf *.xlsx *.xlsm *.xls *.docx *.doc *.pptx *.ppt);;All files (*.*)")
        for p in paths:
            self.add_file(p)

    def _expand_glob(self):
        pattern = self._glob.text().strip()
        if not pattern:
            return
        matches = glob.glob(pattern, recursive=self._recursive.isChecked())
        files = [p for p in matches if os.path.isfile(p)]
        if not files:
            QMessageBox.warning(self, "No matches", f"No files matched:\n{pattern}")
            return
        for p in files:
            self.add_file(p)
        self._glob.clear()

    def add_file(self, path: str):
        if any(r.path == path for r in self._rows):
            return
        self._empty.hide()
        row = FileRow(path)
        row.remove_requested.connect(self._remove)
        # insert before the trailing stretch
        self._list_lay.insertWidget(self._list_lay.count() - 1, row)
        self._rows.append(row)
        self.changed.emit()

    def _remove(self, row: FileRow):
        if row not in self._rows:
            return
        self._rows.remove(row)
        row.setParent(None)
        row.deleteLater()
        if not self._rows:
            self._empty.show()
        self.changed.emit()

    def clear_all(self):
        for r in list(self._rows):
            self._remove(r)

    @property
    def file_rows(self) -> list[FileRow]:
        return self._rows


# ─────────────────────────────────────────────────────────────────────────────
# Password panel
# ─────────────────────────────────────────────────────────────────────────────
class PasswordPanel(QGroupBox):
    def __init__(self):
        super().__init__("Password")
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 18, 12, 12)
        lay.setSpacing(8)

        self._mode = QComboBox()
        self._mode.addItem("No password", "none")
        self._mode.addItem("Type a password", "direct")
        self._mode.addItem("Password file (first line)", "file")
        self._mode.addItem("Wordlist — brute force", "list")
        self._mode.currentIndexChanged.connect(self._refresh)
        lay.addWidget(self._mode)

        # direct
        self._pw = QLineEdit()
        self._pw.setEchoMode(QLineEdit.EchoMode.Password)
        self._pw.setPlaceholderText("password")
        self._show = QCheckBox("Show")
        self._show.toggled.connect(
            lambda on: self._pw.setEchoMode(
                QLineEdit.EchoMode.Normal if on else QLineEdit.EchoMode.Password))
        row = QHBoxLayout()
        row.addWidget(self._pw, 1)
        row.addWidget(self._show)
        self._pw_wrap = self._wrap(row)
        lay.addWidget(self._pw_wrap)

        # password file
        self._pf = QLineEdit()
        self._pf.setPlaceholderText("path/to/password.txt")
        self._pf.setMinimumWidth(60)
        self._pf_wrap = self._file_row(self._pf, "Password file")
        lay.addWidget(self._pf_wrap)

        # wordlist
        self._wl = QLineEdit()
        self._wl.setPlaceholderText("path/to/wordlist.txt")
        self._wl.setMinimumWidth(60)
        self._wl_wrap = self._file_row(self._wl, "Wordlist")
        lay.addWidget(self._wl_wrap)

        self._refresh()

    def _wrap(self, inner_layout) -> QWidget:
        w = QWidget()
        w.setLayout(inner_layout)
        inner_layout.setContentsMargins(0, 0, 0, 0)
        return w

    def _file_row(self, line: QLineEdit, title: str) -> QWidget:
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(line, 1)
        btn = QPushButton("…")
        btn.setFixedWidth(34)
        btn.clicked.connect(lambda: self._pick(line, title))
        row.addWidget(btn)
        return self._wrap(row)

    def _pick(self, line: QLineEdit, title: str):
        p, _ = QFileDialog.getOpenFileName(self, title, "", "Text (*.txt);;All files (*.*)")
        if p:
            line.setText(p)

    def _refresh(self):
        m = self.mode
        self._pw_wrap.setVisible(m == "direct")
        self._pf_wrap.setVisible(m == "file")
        self._wl_wrap.setVisible(m == "list")

    @property
    def mode(self) -> str:
        return self._mode.currentData()

    @property
    def password_file_path(self) -> str:
        return self._pf.text().strip()

    @property
    def wordlist_path(self) -> str | None:
        return self._wl.text().strip() or None

    def get_password(self) -> str | None:
        m = self.mode
        if m == "direct":
            return self._pw.text() or None
        if m == "file":
            return _load_password_file(self._pf.text().strip())
        return None  # none / list handled elsewhere


# ─────────────────────────────────────────────────────────────────────────────
# Output panel
# ─────────────────────────────────────────────────────────────────────────────
class OutputPanel(QGroupBox):
    def __init__(self):
        super().__init__("Output")
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 18, 12, 12)
        lay.setSpacing(8)

        self._mode = QComboBox()
        self._mode.addItem("Output directory", "dir")
        self._mode.addItem("In-place (overwrite original)", "inplace")
        self._mode.addItem("Explicit path (single file)", "file")
        self._mode.currentIndexChanged.connect(self._refresh)
        lay.addWidget(self._mode)

        # output dir
        self._dir = QLineEdit()
        self._dir.setPlaceholderText("blank → unlocked_<name>.ext beside original")
        self._dir.setMinimumWidth(60)
        self._dir_wrap = self._file_row(self._dir, directory=True)
        lay.addWidget(self._dir_wrap)

        # explicit file
        self._file = QLineEdit()
        self._file.setPlaceholderText("e.g.  ~/unlocked_report.pdf")
        self._file.setMinimumWidth(60)
        self._file_wrap = self._file_row(self._file, directory=False, save=True)
        lay.addWidget(self._file_wrap)

        # inplace note
        self._ip_note = QLabel("Original is overwritten. Enable backup below to keep a .bak copy.")
        self._ip_note.setObjectName("Warn")
        self._ip_note.setWordWrap(True)
        lay.addWidget(self._ip_note)

        # safety
        self._backup = QCheckBox("Backup (.bak) before overwriting")
        self._no_over = QCheckBox("Skip if output already exists")
        lay.addWidget(self._backup)
        lay.addWidget(self._no_over)

        self._refresh()

    def _file_row(self, line: QLineEdit, directory: bool, save: bool = False) -> QWidget:
        w = QWidget()
        row = QHBoxLayout(w)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(line, 1)
        btn = QPushButton("…")
        btn.setFixedWidth(34)
        btn.clicked.connect(lambda: self._pick(line, directory, save))
        row.addWidget(btn)
        return w

    def _pick(self, line: QLineEdit, directory: bool, save: bool):
        if directory:
            p = QFileDialog.getExistingDirectory(self, "Output directory")
        elif save:
            p, _ = QFileDialog.getSaveFileName(self, "Output file")
        else:
            p, _ = QFileDialog.getOpenFileName(self, "Select")
        if p:
            line.setText(p)

    def _refresh(self):
        m = self.mode
        self._dir_wrap.setVisible(m == "dir")
        self._file_wrap.setVisible(m == "file")
        self._ip_note.setVisible(m == "inplace")
        self._backup.setEnabled(m == "inplace")

    @property
    def mode(self) -> str:
        return self._mode.currentData()

    @property
    def output_dir(self) -> str | None:
        return self._dir.text().strip() or None

    @property
    def output_file(self) -> str | None:
        return self._file.text().strip() or None

    @property
    def in_place(self) -> bool:
        return self.mode == "inplace"

    @property
    def backup(self) -> bool:
        return self._backup.isChecked() and self.in_place

    @property
    def no_overwrite(self) -> bool:
        return self._no_over.isChecked()


# ─────────────────────────────────────────────────────────────────────────────
# Advanced panel
# ─────────────────────────────────────────────────────────────────────────────
class AdvancedPanel(QGroupBox):
    def __init__(self):
        super().__init__("Advanced")
        self._build()

    def _build(self):
        lay = QFormLayout(self)
        lay.setContentsMargins(12, 18, 12, 12)
        lay.setSpacing(10)
        lay.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        self._verbosity = QComboBox()
        self._verbosity.addItem("Normal", "normal")
        self._verbosity.addItem("Quiet (errors only)", "quiet")
        self._verbosity.addItem("Verbose (debug)", "verbose")
        lay.addRow(self._small("Log level"), self._verbosity)

        self._jobs = QSpinBox()
        self._jobs.setRange(1, 16)
        self._jobs.setValue(1)
        self._jobs.setToolTip("Values > 1 process files in parallel worker processes.")
        lay.addRow(self._small("Workers"), self._jobs)

        self._fail_fast = QCheckBox("Fail fast — stop on first error")
        lay.addRow(self._fail_fast)

        self._convert = QCheckBox("Convert legacy formats (.doc/.xls/.ppt → OpenXML)")
        lay.addRow(self._convert)
        conv_note = QLabel("Needs LibreOffice on PATH. Macros may be lost.")
        conv_note.setObjectName("Hint")
        conv_note.setWordWrap(True)
        lay.addRow(conv_note)

        self._json = QCheckBox("Emit JSON results  ·  Inspect mode only")
        lay.addRow(self._json)

    def _small(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("Hint")
        return lbl

    @property
    def verbosity(self) -> str:
        return self._verbosity.currentData()

    @property
    def jobs(self) -> int:
        return self._jobs.value()

    @property
    def fail_fast(self) -> bool:
        return self._fail_fast.isChecked()

    @property
    def convert(self) -> bool:
        return self._convert.isChecked()

    @property
    def json_export(self) -> bool:
        return self._json.isChecked()


# ─────────────────────────────────────────────────────────────────────────────
# Summary chips + log
# ─────────────────────────────────────────────────────────────────────────────
class StatsChips(QWidget):
    ORDER = [("unlocked", "unlocked"), ("open", "open"), ("protected", "protected"),
             ("skipped", "skipped"), ("failed", "failed"), ("pending", "queued")]

    def __init__(self):
        super().__init__()
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        self._chips: dict[str, QLabel] = {}
        for key, label in self.ORDER:
            chip = QLabel(f"0 {label}")
            chip.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self._chips[key] = chip
            lay.addWidget(chip)
        lay.addStretch(1)
        self.update_counts({})

    def update_counts(self, counts: dict[str, int]):
        for key, label in self.ORDER:
            n = counts.get(key, 0)
            c = STATUS_COLOR.get(key, FAINT)
            chip = self._chips[key]
            chip.setText(f"{n} {label}")
            on = n > 0
            chip.setStyleSheet(
                f"color: {c if on else FAINT};"
                f" background-color: {ELEVATED if on else 'transparent'};"
                f" border: 1px solid {c if on else BORDER};"
                f" border-radius: 11px; padding: 3px 10px; font-size: 11px;"
                f" font-weight: {'700' if on else '400'};"
            )


class LogPanel(QFrame):
    LEVEL_COLOR = {"debug": MUTED, "info": TEXT, "ok": SUCCESS,
                   "warn": WARN, "err": ERROR, "json": INFO}

    def __init__(self):
        super().__init__()
        self.setObjectName("ActionBar")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 12)
        lay.setSpacing(6)

        top = QHBoxLayout()
        title = QLabel("Output log")
        title.setObjectName("Subtitle")
        top.addWidget(title)
        top.addStretch(1)
        copy = QPushButton("Copy")
        copy.setObjectName("Ghost")
        copy.clicked.connect(self._copy)
        clear = QPushButton("Clear")
        clear.setObjectName("Ghost")
        clear.clicked.connect(self.clear)
        top.addWidget(copy)
        top.addWidget(clear)
        lay.addLayout(top)

        self._text = QPlainTextEdit()
        self._text.setReadOnly(True)
        self._text.setMinimumHeight(140)
        mono = QFont("Menlo")
        mono.setStyleHint(QFont.StyleHint.Monospace)
        mono.setPointSize(10)
        self._text.setFont(mono)
        lay.addWidget(self._text)

    def append(self, text: str, level: str = "info"):
        from html import escape
        import time
        color = self.LEVEL_COLOR.get(level, TEXT)
        ts = time.strftime("%H:%M:%S")
        self._text.appendHtml(
            f'<span style="color:{FAINT}">[{ts}]</span> '
            f'<span style="color:{color}">{escape(text)}</span>'
        )
        sb = self._text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def clear(self):
        self._text.clear()

    def _copy(self):
        QGuiApplication.clipboard().setText(self._text.toPlainText())


# ─────────────────────────────────────────────────────────────────────────────
# Workers (run in a QThread; communicate via signals)
# ─────────────────────────────────────────────────────────────────────────────
class _BaseWorker(QObject):
    log = pyqtSignal(str, str)           # text, level
    row = pyqtSignal(str, str, str)      # path, status, message
    progress = pyqtSignal(float, str)    # fraction, "n/total"
    finished = pyqtSignal()

    def __init__(self, paths: list[str]):
        super().__init__()
        self._paths = paths


class CheckWorker(_BaseWorker):
    json_ready = pyqtSignal(str)

    def __init__(self, paths, do_json: bool, fail_fast: bool):
        super().__init__(paths)
        self._do_json = do_json
        self._fail_fast = fail_fast

    def run(self):
        total = len(self._paths)
        acc: list[dict] = []
        for i, path in enumerate(self._paths):
            self.progress.emit(i / total, f"{i}/{total}")
            self.row.emit(path, "running", "inspecting…")
            try:
                result = check_file(path)
                self.row.emit(path, result.status, result.message)
                level = {"open": "ok", "protected": "warn",
                         "failed": "err"}.get(result.status, "info")
                self.log.emit(f"{Path(path).name}: [{result.status}] {result.message}", level)
                if self._do_json:
                    acc.append({"file": path, "status": result.status,
                                "layers": result.layers, "message": result.message})
            except Exception as e:  # noqa: BLE001
                self.row.emit(path, "failed", str(e))
                self.log.emit(f"{Path(path).name}: {e}", "err")
                if self._fail_fast:
                    self.log.emit("Fail-fast: stopping after first error.", "warn")
                    break
        if self._do_json and acc:
            blob = json.dumps(acc, indent=2)
            self.log.emit("JSON output:", "info")
            for line in blob.splitlines():
                self.log.emit(line, "json")
            self.json_ready.emit(blob)
        self.progress.emit(1.0, f"{total}/{total}")
        self.finished.emit()


class UnprotectWorker(_BaseWorker):
    def __init__(self, paths, *, password, pw_mode, wordlist_path,
                 out_mode, output_file, output_dir, in_place, backup,
                 no_overwrite, convert, fail_fast, jobs):
        super().__init__(paths)
        self._password = password
        self._pw_mode = pw_mode
        self._wordlist = wordlist_path
        self._out_mode = out_mode
        self._output_file = output_file
        self._output_dir = output_dir
        self._in_place = in_place
        self._backup = backup
        self._no_overwrite = no_overwrite
        self._convert = convert
        self._fail_fast = fail_fast
        self._jobs = max(1, jobs)

    def _call_for(self, path: str):
        common = dict(
            output_path=self._output_file if self._out_mode == "file" else None,
            in_place=self._in_place,
            output_dir=self._output_dir if self._out_mode == "dir" else None,
            backup=self._backup,
            no_overwrite=self._no_overwrite,
            convert=self._convert,
        )
        if self._pw_mode == "list":
            return _try_password_list, dict(input_path=path, wordlist_path=self._wordlist, **common)
        return unprotect_file, dict(input_path=path, password=self._password, **common)

    def run(self):
        total = len(self._paths)
        counts = {"unlocked": 0, "open": 0, "skipped": 0, "failed": 0}

        def record(path, result):
            self.row.emit(path, result.status, result.message)
            level = {"unlocked": "ok", "skipped": "warn",
                     "failed": "err", "open": "info"}.get(result.status, "info")
            self.log.emit(f"{Path(path).name}: {result.message}", level)
            counts[result.status if result.status in counts else "failed"] += 1

        def record_error(path, exc, unexpected=False):
            self.row.emit(path, "failed", str(exc))
            prefix = "unexpected error — " if unexpected else ""
            self.log.emit(f"{Path(path).name}: {prefix}{exc}", "err")
            counts["failed"] += 1

        if self._jobs > 1 and total > 1:
            self.log.emit(f"Processing {total} files with {self._jobs} workers…", "info")
            for p in self._paths:
                self.row.emit(p, "running", "queued…")
            done = 0
            stop = False
            with ProcessPoolExecutor(max_workers=self._jobs) as ex:
                fut_to_path = {}
                for path in self._paths:
                    fn, kwargs = self._call_for(path)
                    fut_to_path[ex.submit(fn, **kwargs)] = path
                for fut in as_completed(fut_to_path):
                    path = fut_to_path[fut]
                    try:
                        record(path, fut.result())
                    except UnprotectError as e:
                        record_error(path, e)
                        stop = self._fail_fast
                    except Exception as e:  # noqa: BLE001
                        record_error(path, e, unexpected=True)
                        stop = self._fail_fast
                    done += 1
                    self.progress.emit(done / total, f"{done}/{total}")
                    if stop:
                        self.log.emit("Fail-fast: cancelling remaining files.", "warn")
                        for f in fut_to_path:
                            f.cancel()
                        break
        else:
            for i, path in enumerate(self._paths):
                self.progress.emit(i / total, f"{i}/{total}")
                self.row.emit(path, "running", "processing…")
                try:
                    fn, kwargs = self._call_for(path)
                    record(path, fn(**kwargs))
                except UnprotectError as e:
                    record_error(path, e)
                    if self._fail_fast:
                        self.log.emit("Fail-fast: stopping after first error.", "warn")
                        break
                except Exception as e:  # noqa: BLE001
                    record_error(path, e, unexpected=True)
                    if self._fail_fast:
                        self.log.emit("Fail-fast: stopping after first error.", "warn")
                        break

        if total > 1:
            parts = []
            for k, lbl in [("unlocked", "unlocked"), ("open", "already open"),
                           ("skipped", "skipped"), ("failed", "failed")]:
                if counts[k]:
                    parts.append(f"{counts[k]} {lbl}")
            self.log.emit("Done: " + (", ".join(parts) if parts else "nothing processed"), "info")

        self.progress.emit(1.0, f"{total}/{total}")
        self.finished.emit()


# ─────────────────────────────────────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────────────────────────────────────
class UnprotectApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Unprotect")
        self.resize(1040, 760)
        self.setMinimumSize(860, 620)

        self._thread: QThread | None = None
        self._worker: _BaseWorker | None = None
        self._running = False

        self._build()

        if not _HAVE_LIB:
            self.log("unprotect.py not found alongside this file.", "warn")
            self.log("Place unprotect.py in the same directory and restart.", "warn")

    # ── Layout ───────────────────────────────────────────────────────────────
    def _build(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        root.addWidget(self._header())

        split = QSplitter(Qt.Orientation.Horizontal)
        split.setChildrenCollapsible(False)
        split.setHandleWidth(12)

        self._input_tab = InputPanel()
        self._input_tab.changed.connect(self._refresh_stats)
        split.addWidget(self._input_tab)

        split.addWidget(self._settings_column())
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setSizes([580, 430])
        root.addWidget(split, 1)

        root.addWidget(self._action_bar())
        self._log_panel = LogPanel()
        root.addWidget(self._log_panel)

        self.setStyleSheet(STYLESHEET)
        self._refresh_stats()

    def _header(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("Header")
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(18, 12, 18, 12)
        title = QLabel("Unprotect")
        title.setObjectName("Title")
        sub = QLabel("PDF · XLSX · DOCX · PPTX")
        sub.setObjectName("Subtitle")
        lay.addWidget(title)
        lay.addSpacing(10)
        lay.addWidget(sub)
        lay.addStretch(1)
        self._count = QLabel("0 files")
        self._count.setStyleSheet(
            f"color: {INFO}; background-color: {ELEVATED}; border: 1px solid {BORDER};"
            f" border-radius: 12px; padding: 4px 12px; font-weight: 600;")
        lay.addWidget(self._count)
        return bar

    def _settings_column(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        host = QWidget()
        lay = QVBoxLayout(host)
        lay.setContentsMargins(0, 0, 4, 0)
        lay.setSpacing(10)

        self._password_tab = PasswordPanel()
        self._output_tab = OutputPanel()
        self._advanced_tab = AdvancedPanel()
        for w in (self._password_tab, self._output_tab, self._advanced_tab):
            w.setMinimumWidth(0)
            lay.addWidget(w)
        lay.addStretch(1)

        # Let combo boxes shrink below their longest item so the column never
        # forces horizontal overflow; the dropdown still shows full labels.
        for combo in host.findChildren(QComboBox):
            combo.setSizeAdjustPolicy(
                QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(8)
            combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        scroll.setWidget(host)
        scroll.setMinimumWidth(320)
        return scroll

    def _action_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("ActionBar")
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(14, 10, 14, 10)
        lay.setSpacing(10)

        self._check_btn = QPushButton("Inspect")
        self._check_btn.clicked.connect(self._run_check)
        lay.addWidget(self._check_btn)

        self._run_btn = QPushButton("Unprotect")
        self._run_btn.setObjectName("Primary")
        self._run_btn.clicked.connect(self._run_unprotect)
        lay.addWidget(self._run_btn)

        self._progress = QProgressBar()
        self._progress.setFixedWidth(170)
        self._progress.setRange(0, 1000)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        lay.addWidget(self._progress)

        self._prog_lbl = QLabel("")
        self._prog_lbl.setObjectName("Hint")
        self._prog_lbl.setFixedWidth(60)
        lay.addWidget(self._prog_lbl)

        lay.addStretch(1)
        self._stats = StatsChips()
        lay.addWidget(self._stats, 1)
        return bar

    # ── Logging / stats ──────────────────────────────────────────────────────
    def log(self, text: str, level: str = "info"):
        verbosity = self._advanced_tab.verbosity if hasattr(self, "_advanced_tab") else "normal"
        if verbosity == "quiet" and level not in ("err", "warn"):
            return
        if verbosity != "verbose" and level == "debug":
            return
        self._log_panel.append(text, level)

    def _refresh_stats(self):
        rows = self._input_tab.file_rows
        n = len(rows)
        self._count.setText(f"{n} file" + ("" if n == 1 else "s"))
        counts: dict[str, int] = {}
        for r in rows:
            counts[r.status] = counts.get(r.status, 0) + 1
        self._stats.update_counts(counts)

    def _row_by_path(self, path: str) -> FileRow | None:
        for r in self._input_tab.file_rows:
            if r.path == path:
                return r
        return None

    # ── Busy state ───────────────────────────────────────────────────────────
    def _set_busy(self, busy: bool):
        self._running = busy
        self._run_btn.setEnabled(not busy)
        self._check_btn.setEnabled(not busy)
        if not busy:
            self._progress.setValue(0)
            self._prog_lbl.setText("")

    # ── Validation ───────────────────────────────────────────────────────────
    def _validate(self, check_only: bool) -> bool:
        if not _HAVE_LIB:
            QMessageBox.critical(self, "Missing library", "unprotect.py not found. See log.")
            return False
        rows = self._input_tab.file_rows
        if not rows:
            QMessageBox.warning(self, "No files", "Add at least one file first.")
            return False

        out, pw, adv = self._output_tab, self._password_tab, self._advanced_tab

        if out.mode == "file" and len(rows) > 1:
            QMessageBox.warning(self, "Config error",
                                "Explicit output path is single-file mode only.")
            return False
        if pw.mode == "file" and not pw.password_file_path:
            QMessageBox.warning(self, "Config error", "Password-file mode selected but no path given.")
            return False
        if pw.mode == "list" and not pw.wordlist_path:
            QMessageBox.warning(self, "Config error", "Wordlist mode selected but no wordlist chosen.")
            return False
        if adv.json_export and not check_only:
            QMessageBox.warning(self, "Config error", "JSON export is only available with Inspect.")
            return False
        if not check_only and adv.convert and not shutil.which("libreoffice"):
            resp = QMessageBox.question(
                self, "LibreOffice missing",
                "Conversion needs LibreOffice on PATH; without it, conversion will fail.\n\n"
                "Continue anyway?")
            if resp != QMessageBox.StandardButton.Yes:
                return False
        return True

    # ── Run: Check ───────────────────────────────────────────────────────────
    def _run_check(self):
        if self._running or not self._validate(check_only=True):
            return
        paths = [r.path for r in self._input_tab.file_rows]
        for r in self._input_tab.file_rows:
            r.set_status("running", "inspecting…")
        self._refresh_stats()
        worker = CheckWorker(paths, self._advanced_tab.json_export, self._advanced_tab.fail_fast)
        worker.json_ready.connect(self._offer_json_save)
        self._start(worker)

    # ── Run: Unprotect ───────────────────────────────────────────────────────
    def _run_unprotect(self):
        if self._running or not self._validate(check_only=False):
            return
        pw = self._password_tab
        password = None
        if pw.mode not in ("none", "list"):
            try:
                password = pw.get_password()
            except (UnprotectError, OSError) as e:
                QMessageBox.critical(self, "Password error", str(e))
                return
        paths = [r.path for r in self._input_tab.file_rows]
        for r in self._input_tab.file_rows:
            r.set_status("pending", "queued…")
        self._refresh_stats()
        out, adv = self._output_tab, self._advanced_tab
        worker = UnprotectWorker(
            paths, password=password, pw_mode=pw.mode, wordlist_path=pw.wordlist_path,
            out_mode=out.mode, output_file=out.output_file, output_dir=out.output_dir,
            in_place=out.in_place, backup=out.backup, no_overwrite=out.no_overwrite,
            convert=adv.convert, fail_fast=adv.fail_fast, jobs=adv.jobs)
        self._start(worker)

    # ── Thread plumbing ──────────────────────────────────────────────────────
    def _start(self, worker: _BaseWorker):
        self._set_busy(True)
        thread = QThread()
        worker.moveToThread(thread)
        worker.log.connect(self.log)
        worker.row.connect(self._on_row)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_finished)
        thread.started.connect(worker.run)
        self._thread, self._worker = thread, worker
        thread.start()

    def _on_row(self, path: str, status: str, message: str):
        row = self._row_by_path(path)
        if row is not None:
            row.set_status(status, message)
            self._refresh_stats()

    def _on_progress(self, fraction: float, label: str):
        self._progress.setValue(int(fraction * 1000))
        self._prog_lbl.setText(label)

    def _on_finished(self):
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
        self._thread = self._worker = None
        self._set_busy(False)

    def _offer_json_save(self, blob: str):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save JSON results", "results.json", "JSON (*.json);;All files (*.*)")
        if path:
            with open(path, "w") as f:
                f.write(blob)
            self.log(f"JSON saved → {path}", "ok")


# ─────────────────────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Unprotect")
    win = UnprotectApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
