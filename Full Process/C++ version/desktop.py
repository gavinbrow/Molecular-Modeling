#!/usr/bin/env python3
"""
desktop.py - PySide6 desktop application for the ORCA workflow manager.

Provides a native desktop UI to:
  - Enter SMILES codes for molecules with per-molecule settings
  - Configure ORCA calculation settings (with preset support)
  - Browse for ORCA executable and choose a project directory
  - Dynamically queue molecules while a calculation is running
  - Abort running calculations
  - Run the full pipeline (geometry -> input -> ORCA -> report)
  - View results with molecule structure images
  - Open the generated Excel report

Settings persist across sessions via config.json.

Start with:  python desktop.py
Requires:    pip install PySide6
"""

import sys
import os
import re
import json
import time
import copy
import threading
from pathlib import Path

try:
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QGroupBox, QTableWidget, QTableWidgetItem, QPushButton, QLabel,
        QComboBox, QSpinBox, QLineEdit, QPlainTextEdit, QProgressBar,
        QHeaderView, QAbstractItemView, QScrollArea, QDialog,
        QDialogButtonBox, QMessageBox, QSizePolicy, QFrame,
        QGridLayout, QStyleFactory, QFileDialog, QInputDialog,
    )
    from PySide6.QtCore import Qt, QThread, Signal, QTimer, QUrl
    from PySide6.QtGui import QPixmap, QFont, QColor, QDesktopServices, QIcon
except ImportError:
    print("ERROR: PySide6 is required.  Install with:  pip install PySide6")
    sys.exit(1)

from pipeline import (
    ORCA_EXE, INP_DIR, OUT_DIR, MID_DIR,
    validate_smiles, smiles_to_xyz, generate_inp,
    make_run_stamp, run_orca_job, parse_out_file, build_report, fmt_hhmmss,
)


# ── Windows sleep prevention ──────────────────────────────────────────

if sys.platform == "win32":
    import ctypes
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001

    def _prevent_sleep():
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        )

    def _allow_sleep():
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
else:
    def _prevent_sleep():
        pass

    def _allow_sleep():
        pass


# ── Configuration persistence ───────────────────────────────────────────

ROOT = Path(__file__).resolve().parent


def _find_app_icon() -> str | None:
    """Locate the app icon for window decoration."""
    candidates = []
    if getattr(sys, "frozen", False):
        # PyInstaller --onedir: icon next to the exe
        candidates.append(Path(sys.executable).parent / "icon.ico")
        # PyInstaller internal data dir
        candidates.append(Path(getattr(sys, "_MEIPASS", "")) / "icon.ico")
    candidates.append(ROOT / "icon.ico")
    for p in candidates:
        if p.exists():
            return str(p)
    return None
CONFIG_PATH = ROOT / "config.json"

DEFAULT_CONFIG = {
    "orca_path": str(ORCA_EXE),
    "project_dir": str(ROOT),
    "preset": "Custom",
    "functional": "B3LYP",
    "basis_set": "def2-SVP",
    "calc_type": "OPT FREQ",
    "charge": 0,
    "multiplicity": 1,
    "ram_mb": 4000,
    "cpus": 4,
    "extra_keywords": "",
    "extra_blocks": "",
    "custom_presets": {},
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = dict(DEFAULT_CONFIG)
            merged.update(cfg)
            return merged
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ── Presets ──────────────────────────────────────────────────────────────

BUILTIN_PRESETS = {
    "Standard Opt": {
        "functional": "B3LYP",
        "basis_set": "def2-SVP",
        "calc_type": "OPT",
        "extra_keywords": "RIJCOSX TightSCF NumFreq NoSym CPCM(THF)",
        "charge": 0,
        "multiplicity": 1,
        "cpus": 24,
        "ram_mb": 1000,
        "extra_blocks": "%geom MaxIter 300 end",
        "nprocs_group": 24,
    },
    "High-Level Refinement": {
        "functional": "wB97X-V",
        "basis_set": "ma-def2-TZVPP",
        "calc_type": "OPT",
        "extra_keywords": "RIJCOSX TightSCF NumFreq NoSym CPCM(THF)",
        "charge": 0,
        "multiplicity": 1,
        "cpus": 24,
        "ram_mb": 1000,
        "extra_blocks": "%geom MaxIter 300 end",
        "nprocs_group": 24,
    },
}

# Mutable dict rebuilt at startup from builtins + user custom presets
PRESETS: dict[str, dict | None] = {"Custom": None, **BUILTIN_PRESETS}

# Default settings dict used for molecules with no custom settings
DEFAULT_SETTINGS = {
    "functional": "B3LYP",
    "basis_set": "def2-SVP",
    "calc_type": "OPT FREQ",
    "charge": 0,
    "multiplicity": 1,
    "ram_mb": 4000,
    "cpus": 4,
    "extra_keywords": "",
    "extra_blocks": "",
}


# ── Helpers ──────────────────────────────────────────────────────────────

def _sanitize_name(name: str) -> str:
    safe = re.sub(r"[^\w\-]", "_", name.strip())
    return (safe or "mol")[:50]


def _form_label(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(
        "font-size: 11px; font-weight: 600; color: #94a3b8; background: transparent;"
    )
    return lbl


def _pixmap_from_bytes(data: bytes, w: int = 120, h: int = 85) -> QPixmap:
    pm = QPixmap()
    pm.loadFromData(data)
    return pm.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)


# ── Dark-mode stylesheet ────────────────────────────────────────────────

STYLE = """
/* ── Base ──────────────────────────────────────────── */
QMainWindow        { background: #0f172a; }
QScrollArea,
QWidget#scrollContent { background: #0f172a; }

/* ── Header banner ─────────────────────────────────── */
QFrame#header {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #1e293b, stop:1 #3b82f6);
    border-radius: 10px;
    padding: 20px;
}
QLabel#title     { color: #ffffff; font-size: 20px; font-weight: bold; }
QLabel#subtitle  { color: #93c5fd; font-size: 12px; }
QLabel#copyright { color: #93c5fd; font-size: 10px; margin-top: 2px; }

/* ── Cards (QGroupBox) ─────────────────────────────── */
QGroupBox {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px;
    padding-top: 26px;
    font-weight: bold;
    font-size: 12px;
    color: #e2e8f0;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px; top: 3px;
    padding: 0 6px;
    color: #93c5fd;
    background: #1e293b;
}

/* ── Tables ────────────────────────────────────────── */
QTableWidget {
    border: 1px solid #334155;
    border-radius: 4px;
    gridline-color: #1e293b;
    background: #0f172a;
    alternate-background-color: #1e293b;
    color: #e2e8f0;
    selection-background-color: #1e3a5f;
    font-size: 12px;
}
QTableWidget::item { padding: 3px 6px; }
QHeaderView::section {
    background: #1e293b;
    border: none;
    border-bottom: 2px solid #334155;
    padding: 5px 8px;
    font-weight: bold;
    font-size: 11px;
    color: #94a3b8;
}

/* ── Run button ────────────────────────────────────── */
QPushButton#runBtn {
    background: #3b82f6; color: white; border: none;
    border-radius: 8px; padding: 12px 48px;
    font-size: 14px; font-weight: bold;
}
QPushButton#runBtn:hover   { background: #2563eb; }
QPushButton#runBtn:pressed { background: #1d4ed8; }
QPushButton#runBtn:disabled { background: #475569; color: #64748b; }

/* ── Stop button ──────────────────────────────────── */
QPushButton#stopBtn {
    background: #ef4444; color: white; border: none;
    border-radius: 8px; padding: 12px 48px;
    font-size: 14px; font-weight: bold;
}
QPushButton#stopBtn:hover   { background: #dc2626; }
QPushButton#stopBtn:pressed { background: #b91c1c; }

/* ── Secondary / action buttons ────────────────────── */
QPushButton#addBtn, QPushButton#bulkBtn, QPushButton#applyAllBtn {
    background: transparent; color: #3b82f6;
    border: 1px solid #3b82f6; border-radius: 4px;
    padding: 4px 14px; font-size: 12px; font-weight: 600;
}
QPushButton#addBtn:hover, QPushButton#bulkBtn:hover, QPushButton#applyAllBtn:hover {
    background: #3b82f6; color: white;
}
QPushButton#queueBtn {
    background: #16a34a; color: white;
    border: none; border-radius: 4px;
    padding: 4px 14px; font-size: 12px; font-weight: 600;
}
QPushButton#queueBtn:hover { background: #15803d; }

QPushButton#removeBtn {
    background: transparent; color: #64748b; border: none;
    font-size: 15px; font-weight: bold;
    min-width: 26px; max-width: 26px;
}
QPushButton#removeBtn:hover { color: #ef4444; }

QPushButton#settingsBtn {
    background: transparent; border: none;
    font-size: 14px; font-weight: bold;
    min-width: 30px; max-width: 30px;
    padding: 0;
}

QPushButton#reportBtn, QPushButton#folderBtn {
    color: white; border: none; border-radius: 8px;
    padding: 10px 28px; font-size: 13px; font-weight: bold;
}
QPushButton#reportBtn       { background: #22c55e; }
QPushButton#reportBtn:hover { background: #16a34a; }
QPushButton#folderBtn       { background: #475569; }
QPushButton#folderBtn:hover { background: #334155; }

QPushButton#browseBtn {
    background: #334155; color: #e2e8f0; border: none;
    border-radius: 4px; padding: 5px 12px; font-size: 11px;
}
QPushButton#browseBtn:hover { background: #475569; }

/* ── Progress bar (status-aware, taller for text) ──── */
QProgressBar {
    border: none; border-radius: 6px;
    background: #334155; min-height: 22px; max-height: 22px;
    color: #e2e8f0; font-size: 10px; font-weight: bold;
    text-align: center;
}
QProgressBar::chunk { background: #3b82f6; border-radius: 6px; }

/* ── Inputs ────────────────────────────────────────── */
QLineEdit, QComboBox, QSpinBox {
    padding: 5px 8px; border: 1px solid #334155;
    border-radius: 4px; font-size: 12px;
    background: #0f172a; color: #e2e8f0;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
    border-color: #3b82f6;
}
QComboBox::drop-down { border: none; }
QComboBox QAbstractItemView {
    background: #1e293b; color: #e2e8f0;
    border: 1px solid #334155;
    selection-background-color: #3b82f6;
}
QSpinBox::up-button, QSpinBox::down-button {
    background: #334155; border: none; width: 16px;
}
QSpinBox::up-arrow   { image: none; border-left: 4px solid transparent;
    border-right: 4px solid transparent; border-bottom: 5px solid #94a3b8; }
QSpinBox::down-arrow { image: none; border-left: 4px solid transparent;
    border-right: 4px solid transparent; border-top: 5px solid #94a3b8; }

QPlainTextEdit {
    border: 1px solid #334155; border-radius: 4px;
    font-family: "Cascadia Code", Consolas, monospace;
    font-size: 12px; background: #0f172a; color: #e2e8f0;
}
QPlainTextEdit:focus { border-color: #3b82f6; }

/* ── Labels (default) ──────────────────────────────── */
QLabel { color: #e2e8f0; }

/* ── Dialogs & message boxes ───────────────────────── */
QDialog { background: #1e293b; color: #e2e8f0; }
QMessageBox { background: #1e293b; color: #e2e8f0; }
QMessageBox QLabel { color: #e2e8f0; }
QDialogButtonBox QPushButton {
    background: #334155; color: #e2e8f0; border: 1px solid #475569;
    border-radius: 4px; padding: 6px 16px; font-size: 12px;
}
QDialogButtonBox QPushButton:hover { background: #475569; }

/* ── Separator line ────────────────────────────────── */
QFrame#separator {
    background: #334155; max-height: 1px; min-height: 1px;
}

/* ── Tooltip ───────────────────────────────────────── */
QToolTip {
    background: #1e293b; color: #e2e8f0;
    border: 1px solid #475569; padding: 4px;
}

/* ── Input dialog ─────────────────────────────────── */
QInputDialog { background: #1e293b; color: #e2e8f0; }
QInputDialog QLabel { color: #e2e8f0; }
QInputDialog QLineEdit {
    background: #0f172a; color: #e2e8f0;
    border: 1px solid #334155; border-radius: 4px; padding: 5px 8px;
}
"""


# ── Bulk Import Dialog ───────────────────────────────────────────────────

class BulkImportDialog(QDialog):
    """Dialog for pasting multiple molecules at once."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Paste Molecules")
        self.setMinimumSize(480, 360)

        layout = QVBoxLayout(self)

        hint = QLabel(
            "One per line.  Format: <b>name, SMILES</b>"
            " &mdash; or just SMILES (names auto-generated)."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #94a3b8; font-size: 12px;")
        layout.addWidget(hint)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("water, O\nethanol, CCO\nc1ccccc1")
        layout.addWidget(self.text_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_molecules(self) -> list[tuple[str, str]]:
        results: list[tuple[str, str]] = []
        idx = 0
        for line in self.text_edit.toPlainText().split("\n"):
            line = line.strip()
            if not line:
                continue
            idx += 1
            sep = "," if "," in line else "\t" if "\t" in line else None
            if sep:
                parts = line.split(sep, 1)
                name = parts[0].strip()
                smiles = parts[1].strip() if len(parts) > 1 else ""
            else:
                name = f"mol_{idx}"
                smiles = line
            results.append((name, smiles))
        return results


# ── Per-Molecule Settings Dialog ─────────────────────────────────────────

class MoleculeSettingsDialog(QDialog):
    """Dialog for editing per-molecule calculation settings."""

    def __init__(self, mol_name: str, settings: dict | None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Settings for {mol_name}")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)

        hint = QLabel(
            "Configure settings for this specific molecule. "
            "Leave blank fields to use global defaults."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #94a3b8; font-size: 12px;")
        layout.addWidget(hint)

        g = QGridLayout()
        g.setSpacing(8)
        row = 0

        # Functional
        g.addWidget(_form_label("Functional"), row, 0)
        self.func_combo = QComboBox()
        self.func_combo.setEditable(True)
        self.func_combo.addItems([
            "B3LYP", "PBE", "PBE0", "M06-2X", "wB97X-D3",
            "wB97M-D3BJ", "wB97X-V", "TPSS", "BP86", "HF", "MP2",
            "RI-MP2", "DLPNO-CCSD(T)",
        ])
        g.addWidget(self.func_combo, row, 1, 1, 3)

        # Basis Set
        row += 1
        g.addWidget(_form_label("Basis Set"), row, 0)
        self.basis_combo = QComboBox()
        self.basis_combo.setEditable(True)
        self.basis_combo.addItems([
            "def2-SVP", "def2-TZVP", "def2-TZVPP", "def2-QZVPP",
            "ma-def2-TZVPP", "6-31G*", "6-31+G(d,p)", "6-311+G(d,p)",
            "cc-pVDZ", "cc-pVTZ", "aug-cc-pVDZ", "aug-cc-pVTZ",
        ])
        g.addWidget(self.basis_combo, row, 1, 1, 3)

        # Calculation Type
        row += 1
        g.addWidget(_form_label("Calculation Type"), row, 0)
        self.calc_combo = QComboBox()
        for label, value in [
            ("Optimisation + Frequency", "OPT FREQ"),
            ("Geometry Optimisation", "OPT"),
            ("Frequency Analysis", "FREQ"),
            ("Single Point Energy", "SP"),
            ("TS Optimisation + Frequency", "OPTTS FREQ"),
            ("TS Optimisation", "OPTTS"),
        ]:
            self.calc_combo.addItem(label, value)
        g.addWidget(self.calc_combo, row, 1, 1, 3)

        # Charge / Multiplicity
        row += 1
        g.addWidget(_form_label("Charge"), row, 0)
        self.charge_spin = QSpinBox()
        self.charge_spin.setRange(-10, 10)
        g.addWidget(self.charge_spin, row, 1)

        g.addWidget(_form_label("Multiplicity"), row, 2)
        self.mult_spin = QSpinBox()
        self.mult_spin.setRange(1, 10)
        g.addWidget(self.mult_spin, row, 3)

        # RAM / CPUs
        row += 1
        g.addWidget(_form_label("RAM / Core (MB)"), row, 0)
        self.ram_spin = QSpinBox()
        self.ram_spin.setRange(256, 128000)
        self.ram_spin.setSingleStep(256)
        g.addWidget(self.ram_spin, row, 1)

        g.addWidget(_form_label("CPU Cores"), row, 2)
        self.cpus_spin = QSpinBox()
        self.cpus_spin.setRange(1, 128)
        g.addWidget(self.cpus_spin, row, 3)

        # Extra Keywords
        row += 1
        g.addWidget(_form_label("Extra Keywords"), row, 0)
        self.extra_kw = QLineEdit()
        self.extra_kw.setPlaceholderText("e.g. D3BJ CPCM(water) TightSCF")
        g.addWidget(self.extra_kw, row, 1, 1, 3)

        # Extra Blocks
        row += 1
        g.addWidget(_form_label("Extra Blocks"), row, 0, Qt.AlignTop)
        self.extra_blocks = QPlainTextEdit()
        self.extra_blocks.setPlaceholderText("%scf\n  MaxIter 300\nend")
        self.extra_blocks.setMaximumHeight(80)
        g.addWidget(self.extra_blocks, row, 1, 1, 3)

        layout.addLayout(g)

        # Buttons
        btn_row = QHBoxLayout()
        clear_btn = QPushButton("Clear (Use Global)")
        clear_btn.setObjectName("browseBtn")
        clear_btn.clicked.connect(self._clear_settings)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        btn_row.addWidget(buttons)
        layout.addLayout(btn_row)

        self._cleared = False

        # Populate from existing settings
        if settings:
            self._populate(settings)

    def _populate(self, s: dict):
        self.func_combo.setCurrentText(s.get("functional", "B3LYP"))
        self.basis_combo.setCurrentText(s.get("basis_set", "def2-SVP"))
        idx = self.calc_combo.findData(s.get("calc_type", "OPT FREQ"))
        if idx >= 0:
            self.calc_combo.setCurrentIndex(idx)
        self.charge_spin.setValue(s.get("charge", 0))
        self.mult_spin.setValue(s.get("multiplicity", 1))
        self.ram_spin.setValue(s.get("ram_mb", 4000))
        self.cpus_spin.setValue(s.get("cpus", 4))
        self.extra_kw.setText(s.get("extra_keywords", ""))
        self.extra_blocks.setPlainText(s.get("extra_blocks", ""))

    def _clear_settings(self):
        self._cleared = True
        self.accept()

    def get_settings(self) -> dict | None:
        """Return settings dict, or None if cleared."""
        if self._cleared:
            return None
        return {
            "functional": self.func_combo.currentText().strip(),
            "basis_set": self.basis_combo.currentText().strip(),
            "calc_type": self.calc_combo.currentData() or "OPT FREQ",
            "charge": self.charge_spin.value(),
            "multiplicity": self.mult_spin.value(),
            "ram_mb": self.ram_spin.value(),
            "cpus": self.cpus_spin.value(),
            "extra_keywords": self.extra_kw.text().strip(),
            "extra_blocks": self.extra_blocks.toPlainText().strip(),
        }


# ── Pipeline Worker (background thread) ──────────────────────────────────

class PipelineWorker(QThread):
    """Runs the full pipeline in a background thread.

    Supports abort via *abort_event* and dynamic queue via
    appending to *molecules* list while running.
    """

    finished = Signal()

    def __init__(
        self,
        molecules: list[dict],
        global_settings: dict,
        mol_settings: dict,
        job: dict,
        orca_path: Path,
        project_dir: Path,
        abort_event: threading.Event,
    ):
        super().__init__()
        self.molecules = molecules
        self.global_settings = global_settings
        self.mol_settings = mol_settings  # {mol_name: settings_dict}
        self.job = job
        self.orca_path = Path(orca_path)
        self.project_dir = Path(project_dir)
        self.abort_event = abort_event
        self.png_data: dict[str, bytes] = {}
        self._wall_start: float = 0.0

    def run(self):
        job = self.job
        self._wall_start = time.monotonic()

        inp_dir = self.project_dir / "INP"
        out_dir = self.project_dir / "OUT"
        mid_dir = self.project_dir / "MID"

        try:
            stamp = make_run_stamp()
            job["stamp"] = stamp
            out_run = out_dir / stamp
            mid_run = mid_dir / stamp
            out_run.mkdir(parents=True, exist_ok=True)
            mid_run.mkdir(parents=True, exist_ok=True)
            inp_dir.mkdir(parents=True, exist_ok=True)

            image_map: dict[str, str] = {}
            smiles_map: dict[str, str] = {}
            inp_paths: dict[str, Path] = {}

            processed = 0

            # Keep processing until all molecules (including dynamically added) are done
            while processed < len(self.molecules):
                if self.abort_event.is_set():
                    job["status"] = "aborted"
                    job["phase"] = "aborted"
                    break

                mol = self.molecules[processed]
                name = mol["name"]
                smiles = mol["smiles"]

                # Use per-molecule settings if set, else global
                settings = self.mol_settings.get(name, self.global_settings)

                # Ensure job["molecules"] is long enough for dynamically added
                while len(job["molecules"]) <= processed:
                    job["molecules"].append({
                        "name": mol["name"],
                        "smiles": mol["smiles"],
                        "status": "pending",
                        "gibbs": None,
                        "electronic_energy": None,
                        "error": None,
                    })

                job["total"] = len(self.molecules)
                job["current"] = processed + 1
                job["current_name"] = name
                job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)

                # ── Phase: geometry ──
                job["phase"] = "geometry"
                job["molecules"][processed]["status"] = "generating"

                try:
                    xyz_block, png_bytes, _ = smiles_to_xyz(smiles)
                    inp_content = generate_inp(name, xyz_block, settings)

                    inp_path = inp_dir / f"{name}.inp"
                    inp_path.write_text(inp_content, encoding="utf-8")
                    inp_paths[name] = inp_path

                    img_dir = out_run / name
                    img_dir.mkdir(parents=True, exist_ok=True)
                    img_path = img_dir / f"{name}.png"
                    img_path.write_bytes(png_bytes)
                    image_map[name] = str(img_path)
                    smiles_map[name] = smiles
                    self.png_data[name] = png_bytes

                    job["molecules"][processed]["status"] = "generated"
                except Exception as exc:
                    job["molecules"][processed]["status"] = "error"
                    job["molecules"][processed]["error"] = str(exc)
                    processed += 1
                    continue

                if self.abort_event.is_set():
                    job["status"] = "aborted"
                    job["phase"] = "aborted"
                    break

                # ── Phase: ORCA ──
                job["phase"] = "orca"
                job["current_name"] = name
                job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)
                job["molecules"][processed]["status"] = "running"

                env = os.environ.copy()
                rc = run_orca_job(
                    name, inp_paths[name], mid_run, out_run,
                    env, status_dict=job, pipeline_start=self._wall_start,
                    orca_exe=self.orca_path,
                    abort_event=self.abort_event,
                )
                job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)

                if self.abort_event.is_set() or rc == -9:
                    job["molecules"][processed]["status"] = "aborted"
                    job["molecules"][processed]["error"] = "Aborted by user"
                    job["status"] = "aborted"
                    job["phase"] = "aborted"
                    break

                if rc == 0:
                    out_file = out_run / name / f"{name}.out"
                    if out_file.exists():
                        parsed = parse_out_file(out_file)
                        job["molecules"][processed]["gibbs"] = parsed.get("gibbs_eh")
                        job["molecules"][processed]["electronic_energy"] = parsed.get(
                            "electronic_energy_eh"
                        )
                        job["molecules"][processed]["status"] = (
                            "completed" if parsed["normal_term"] else "warning"
                        )
                    else:
                        job["molecules"][processed]["status"] = "completed"
                else:
                    job["molecules"][processed]["status"] = "error"
                    job["molecules"][processed]["error"] = f"ORCA exited with code {rc}"

                processed += 1

            # ── Phase: report ──
            if job.get("status") != "aborted":
                job["phase"] = "report"
                job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)

                out_files = []
                for mol in self.molecules[:processed]:
                    f = out_run / mol["name"] / f"{mol['name']}.out"
                    if f.exists():
                        out_files.append(f)

                records = [parse_out_file(f) for f in out_files]
                xlsx_path = out_run / "orca_report.xlsx"

                if records:
                    build_report(records, image_map, smiles_map, xlsx_path)
                    job["report_path"] = str(xlsx_path)

                job["status"] = "completed"
                job["phase"] = "done"

            job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)
            job["out_dir"] = str(out_run)

        except Exception as exc:
            job["status"] = "failed"
            job["error"] = str(exc)
            job["elapsed"] = fmt_hhmmss(time.monotonic() - self._wall_start)
        finally:
            self.finished.emit()


# ── Main Window ──────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ORCA Workflow Manager")
        self.setMinimumSize(1050, 700)
        self.resize(1120, 820)

        icon_path = _find_app_icon()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))

        self._updating = False
        self._applying_preset = False
        self._preset_nprocs_group: int | None = None
        self.job_status: dict | None = None
        self.worker: PipelineWorker | None = None
        self.abort_event: threading.Event | None = None
        self.report_path: str | None = None
        self.output_folder: str | None = None
        self._queued_count: int = 0  # how many rows have been sent to the worker

        # Per-molecule settings: {row_index: settings_dict or None}
        self._mol_settings: dict[int, dict | None] = {}

        self.config = load_config()

        # Load user custom presets into runtime dict
        for name, data in self.config.get("custom_presets", {}).items():
            if name not in ("Custom",) and isinstance(data, dict):
                PRESETS[name] = data

        self.poll_timer = QTimer()
        self.poll_timer.timeout.connect(self._poll_progress)

        self._build_ui()
        self._load_settings_from_config()
        self._update_orca_badge()
        self._update_delete_btn()

    # ── UI construction ───────────────────────────────────────────

    def _build_ui(self):
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(self.scroll_area)

        content = QWidget()
        content.setObjectName("scrollContent")
        self.scroll_area.setWidget(content)

        root = QVBoxLayout(content)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(16)

        root.addWidget(self._build_header())

        panels = QHBoxLayout()
        panels.setSpacing(16)
        panels.addWidget(self._build_molecules_panel(), stretch=3)
        panels.addWidget(self._build_settings_panel(), stretch=2)
        root.addLayout(panels)

        # Run / Stop buttons row
        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignCenter)
        btn_row.setSpacing(12)

        self.run_btn = QPushButton("Run Calculations")
        self.run_btn.setObjectName("runBtn")
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(self._on_run_clicked)
        btn_row.addWidget(self.run_btn)

        self.stop_btn = QPushButton("Stop Calculation")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.clicked.connect(self._on_stop_clicked)
        self.stop_btn.setVisible(False)
        btn_row.addWidget(self.stop_btn)

        root.addLayout(btn_row)

        self.progress_group = self._build_progress_section()
        self.progress_group.setVisible(False)
        root.addWidget(self.progress_group)

        self.results_group = self._build_results_section()
        self.results_group.setVisible(False)
        root.addWidget(self.results_group)

        root.addStretch()

    # -- Header --

    def _build_header(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("header")
        lay = QVBoxLayout(frame)

        title = QLabel("ORCA Workflow Manager")
        title.setObjectName("title")
        lay.addWidget(title)

        sub = QLabel(
            "SMILES \u2192 Force-Field Optimisation "
            "\u2192 ORCA Calculation \u2192 Excel Report"
        )
        sub.setObjectName("subtitle")
        lay.addWidget(sub)

        self.orca_badge = QLabel()
        self.orca_badge.setObjectName("orcaBadge")
        lay.addWidget(self.orca_badge)

        copy_label = QLabel("\u00a9 Gavin Brown")
        copy_label.setObjectName("copyright")
        lay.addWidget(copy_label)

        return frame

    def _update_orca_badge(self):
        orca = Path(self.orca_path_edit.text().strip())
        if orca.exists():
            self.orca_badge.setText(f"\u2713  ORCA found: {orca}")
            self.orca_badge.setStyleSheet(
                "background:#166534; color:#dcfce7; font-size:11px; "
                "padding:3px 12px; border-radius:12px;"
            )
        else:
            self.orca_badge.setText(f"\u2717  ORCA not found: {orca}")
            self.orca_badge.setStyleSheet(
                "background:#991b1b; color:#fee2e2; font-size:11px; "
                "padding:3px 12px; border-radius:12px;"
            )

    # -- Molecules panel --

    def _build_molecules_panel(self) -> QGroupBox:
        grp = QGroupBox("Molecules")
        lay = QVBoxLayout(grp)

        # Columns: Name, SMILES, Settings, Remove
        self.mol_table = QTableWidget(0, 4)
        self.mol_table.setHorizontalHeaderLabels(["Name", "SMILES", "Settings", ""])
        hdr = self.mol_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.Fixed)
        hdr.setSectionResizeMode(3, QHeaderView.Fixed)
        self.mol_table.setColumnWidth(2, 60)
        self.mol_table.setColumnWidth(3, 36)
        self.mol_table.verticalHeader().setVisible(False)
        self.mol_table.verticalHeader().setDefaultSectionSize(34)
        self.mol_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.mol_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
        )
        self.mol_table.cellChanged.connect(self._on_cell_changed)
        lay.addWidget(self.mol_table)

        self._add_mol_row()

        btn_row = QHBoxLayout()
        for text, name, slot in [
            ("+ Add Molecule", "addBtn", self._on_add_clicked),
            ("Paste Bulk", "bulkBtn", self._on_bulk_clicked),
        ]:
            b = QPushButton(text)
            b.setObjectName(name)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(slot)
            btn_row.addWidget(b)

        self.queue_btn = QPushButton("\u25b6 Add to Queue")
        self.queue_btn.setObjectName("queueBtn")
        self.queue_btn.setCursor(Qt.PointingHandCursor)
        self.queue_btn.setToolTip(
            "Submit new molecules to the running calculation"
        )
        self.queue_btn.clicked.connect(self._on_add_to_queue_clicked)
        self.queue_btn.setVisible(False)
        btn_row.addWidget(self.queue_btn)

        btn_row.addStretch()
        lay.addLayout(btn_row)

        return grp

    def _add_mol_row(self, name: str = "", smiles: str = ""):
        self._updating = True
        row = self.mol_table.rowCount()
        self.mol_table.insertRow(row)

        self.mol_table.setItem(row, 0, QTableWidgetItem(name))
        self.mol_table.setItem(row, 1, QTableWidgetItem(smiles))

        # Settings button — shows red X (no custom settings) by default
        settings_btn = QPushButton("\u2717")
        settings_btn.setObjectName("settingsBtn")
        settings_btn.setStyleSheet("color: #ef4444;")
        settings_btn.setCursor(Qt.PointingHandCursor)
        settings_btn.setToolTip("Click to set per-molecule settings")
        settings_btn.clicked.connect(self._on_settings_clicked)
        self.mol_table.setCellWidget(row, 2, settings_btn)

        btn = QPushButton("\u00d7")
        btn.setObjectName("removeBtn")
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._on_remove_clicked)
        self.mol_table.setCellWidget(row, 3, btn)

        self._updating = False

    def _update_settings_indicator(self, row: int):
        """Update the settings button icon for a given row."""
        btn = self.mol_table.cellWidget(row, 2)
        if not btn:
            return
        if self._mol_settings.get(row) is not None:
            btn.setText("\u2713")
            btn.setStyleSheet("color: #22c55e;")
            btn.setToolTip("Custom settings configured (click to edit)")
        else:
            btn.setText("\u2717")
            btn.setStyleSheet("color: #ef4444;")
            btn.setToolTip("Using global settings (click to customize)")

    # -- Molecule event handlers --

    def _on_cell_changed(self, row: int, col: int):
        if self._updating:
            return

    def _on_add_clicked(self):
        self._add_mol_row()
        last = self.mol_table.rowCount() - 1
        self.mol_table.setCurrentCell(last, 0)
        self.mol_table.editItem(self.mol_table.item(last, 0))

    def _on_remove_clicked(self):
        if self.mol_table.rowCount() <= 1:
            return
        btn = self.sender()
        for row in range(self.mol_table.rowCount()):
            if self.mol_table.cellWidget(row, 3) is btn:
                self._mol_settings.pop(row, None)
                self.mol_table.removeRow(row)
                # Re-index mol_settings
                new = {}
                for k, v in self._mol_settings.items():
                    if k > row:
                        new[k - 1] = v
                    elif k < row:
                        new[k] = v
                self._mol_settings = new
                return

    def _on_bulk_clicked(self):
        dlg = BulkImportDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        molecules = dlg.get_molecules()
        if not molecules:
            return
        if self.mol_table.rowCount() == 1:
            i0 = self.mol_table.item(0, 0)
            i1 = self.mol_table.item(0, 1)
            if (not (i0 and i0.text().strip()) and
                    not (i1 and i1.text().strip())):
                self.mol_table.removeRow(0)
                self._mol_settings.clear()
        for name, smiles in molecules:
            self._add_mol_row(name, smiles)

    def _on_add_to_queue_clicked(self):
        """Submit new (un-queued) molecule rows to the running worker."""
        if not self.worker or not self.worker.isRunning():
            return

        new_mols = []
        new_settings: dict[str, dict] = {}
        errors = []

        for row in range(self._queued_count, self.mol_table.rowCount()):
            i0 = self.mol_table.item(row, 0)
            i1 = self.mol_table.item(row, 1)
            name_text = (i0.text() if i0 else "").strip()
            smiles = (i1.text() if i1 else "").strip()
            if not smiles:
                continue
            if not validate_smiles(smiles):
                errors.append(f"Row {row + 1}: invalid SMILES \"{smiles}\"")
                continue
            name = _sanitize_name(name_text or f"mol_{row + 1}")
            new_mols.append({"name": name, "smiles": smiles})

            per_mol = self._mol_settings.get(row)
            if per_mol is not None:
                new_settings[name] = per_mol

        if errors:
            QMessageBox.warning(
                self, "Invalid SMILES",
                "Some molecules have invalid SMILES and were skipped:\n\n"
                + "\n".join(errors),
            )

        if not new_mols:
            if not errors:
                QMessageBox.information(
                    self, "Nothing to Add",
                    "No new molecules to add. Enter molecules in new rows first.",
                )
            return

        # Append to the running worker's shared lists
        self.worker.molecules.extend(new_mols)
        self.worker.mol_settings.update(new_settings)
        self._queued_count = self.mol_table.rowCount()

        # Immediately update the shared job_status so the UI reflects the new total
        job = self.job_status
        if job is not None:
            for m in new_mols:
                job["molecules"].append({
                    "name": m["name"],
                    "smiles": m["smiles"],
                    "status": "pending",
                    "gibbs": None,
                    "electronic_energy": None,
                    "error": None,
                })
            job["total"] = len(self.worker.molecules)

        QMessageBox.information(
            self, "Added to Queue",
            f"{len(new_mols)} molecule(s) added to the running calculation.",
        )

    def _on_settings_clicked(self):
        """Open per-molecule settings dialog."""
        btn = self.sender()
        for row in range(self.mol_table.rowCount()):
            if self.mol_table.cellWidget(row, 2) is btn:
                name_item = self.mol_table.item(row, 0)
                mol_name = (name_item.text() if name_item else "").strip() or f"mol_{row+1}"

                existing = self._mol_settings.get(row)
                # If no custom settings, pre-populate with current global settings
                if existing is None:
                    existing = self._collect_global_settings()

                dlg = MoleculeSettingsDialog(mol_name, existing, self)
                if dlg.exec() == QDialog.Accepted:
                    self._mol_settings[row] = dlg.get_settings()
                    self._update_settings_indicator(row)
                return

    # -- Settings panel --

    def _build_settings_panel(self) -> QGroupBox:
        grp = QGroupBox("Global Settings")
        g = QGridLayout(grp)
        g.setSpacing(8)
        row = 0

        # ── Preset ──
        g.addWidget(_form_label("Preset"), row, 0)
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(list(PRESETS.keys()))
        self.preset_combo.currentTextChanged.connect(self._on_preset_changed)
        g.addWidget(self.preset_combo, row, 1)

        save_preset_btn = QPushButton("Save")
        save_preset_btn.setObjectName("browseBtn")
        save_preset_btn.setCursor(Qt.PointingHandCursor)
        save_preset_btn.setToolTip("Save current settings as a preset")
        save_preset_btn.clicked.connect(self._save_preset)
        g.addWidget(save_preset_btn, row, 2)

        self.del_preset_btn = QPushButton("Delete")
        self.del_preset_btn.setObjectName("browseBtn")
        self.del_preset_btn.setCursor(Qt.PointingHandCursor)
        self.del_preset_btn.setToolTip("Delete the selected preset")
        self.del_preset_btn.clicked.connect(self._delete_preset)
        g.addWidget(self.del_preset_btn, row, 3)

        # ── ORCA Path ──
        row += 1
        g.addWidget(_form_label("ORCA Path"), row, 0)
        self.orca_path_edit = QLineEdit()
        self.orca_path_edit.setPlaceholderText("Path to orca executable")
        self.orca_path_edit.editingFinished.connect(self._update_orca_badge)
        g.addWidget(self.orca_path_edit, row, 1, 1, 2)
        orca_browse = QPushButton("Browse")
        orca_browse.setObjectName("browseBtn")
        orca_browse.setCursor(Qt.PointingHandCursor)
        orca_browse.clicked.connect(self._browse_orca)
        g.addWidget(orca_browse, row, 3)

        # ── Project Directory ──
        row += 1
        g.addWidget(_form_label("Project Dir"), row, 0)
        self.project_dir_edit = QLineEdit()
        self.project_dir_edit.setPlaceholderText("Directory for INP/OUT/MID")
        g.addWidget(self.project_dir_edit, row, 1, 1, 2)
        dir_browse = QPushButton("Browse")
        dir_browse.setObjectName("browseBtn")
        dir_browse.setCursor(Qt.PointingHandCursor)
        dir_browse.clicked.connect(self._browse_project_dir)
        g.addWidget(dir_browse, row, 3)

        # ── Separator ──
        row += 1
        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.HLine)
        g.addWidget(sep, row, 0, 1, 4)

        # ── Functional ──
        row += 1
        g.addWidget(_form_label("Functional"), row, 0)
        self.func_combo = QComboBox()
        self.func_combo.setEditable(True)
        self.func_combo.addItems([
            "B3LYP", "PBE", "PBE0", "M06-2X", "wB97X-D3",
            "wB97M-D3BJ", "wB97X-V", "TPSS", "BP86", "HF", "MP2",
            "RI-MP2", "DLPNO-CCSD(T)",
        ])
        self.func_combo.currentTextChanged.connect(self._on_setting_changed)
        g.addWidget(self.func_combo, row, 1, 1, 3)

        # ── Basis Set ──
        row += 1
        g.addWidget(_form_label("Basis Set"), row, 0)
        self.basis_combo = QComboBox()
        self.basis_combo.setEditable(True)
        self.basis_combo.addItems([
            "def2-SVP", "def2-TZVP", "def2-TZVPP", "def2-QZVPP",
            "ma-def2-TZVPP", "6-31G*", "6-31+G(d,p)", "6-311+G(d,p)",
            "cc-pVDZ", "cc-pVTZ", "aug-cc-pVDZ", "aug-cc-pVTZ",
        ])
        self.basis_combo.currentTextChanged.connect(self._on_setting_changed)
        g.addWidget(self.basis_combo, row, 1, 1, 3)

        # ── Calculation Type ──
        row += 1
        g.addWidget(_form_label("Calculation Type"), row, 0)
        self.calc_combo = QComboBox()
        for label, value in [
            ("Optimisation + Frequency", "OPT FREQ"),
            ("Geometry Optimisation", "OPT"),
            ("Frequency Analysis", "FREQ"),
            ("Single Point Energy", "SP"),
            ("TS Optimisation + Frequency", "OPTTS FREQ"),
            ("TS Optimisation", "OPTTS"),
        ]:
            self.calc_combo.addItem(label, value)
        self.calc_combo.currentIndexChanged.connect(self._on_setting_changed)
        g.addWidget(self.calc_combo, row, 1, 1, 3)

        # ── Charge / Multiplicity ──
        row += 1
        g.addWidget(_form_label("Charge"), row, 0)
        self.charge_spin = QSpinBox()
        self.charge_spin.setRange(-10, 10)
        self.charge_spin.setValue(0)
        self.charge_spin.valueChanged.connect(self._on_setting_changed)
        g.addWidget(self.charge_spin, row, 1)

        g.addWidget(_form_label("Multiplicity"), row, 2)
        self.mult_spin = QSpinBox()
        self.mult_spin.setRange(1, 10)
        self.mult_spin.setValue(1)
        self.mult_spin.valueChanged.connect(self._on_setting_changed)
        g.addWidget(self.mult_spin, row, 3)

        # ── RAM / CPUs ──
        row += 1
        g.addWidget(_form_label("RAM / Core (MB)"), row, 0)
        self.ram_spin = QSpinBox()
        self.ram_spin.setRange(256, 128000)
        self.ram_spin.setSingleStep(256)
        self.ram_spin.setValue(4000)
        self.ram_spin.valueChanged.connect(self._on_setting_changed)
        g.addWidget(self.ram_spin, row, 1)

        g.addWidget(_form_label("CPU Cores"), row, 2)
        self.cpus_spin = QSpinBox()
        self.cpus_spin.setRange(1, 128)
        self.cpus_spin.setValue(4)
        self.cpus_spin.valueChanged.connect(self._on_setting_changed)
        g.addWidget(self.cpus_spin, row, 3)

        # ── Extra Keywords ──
        row += 1
        g.addWidget(_form_label("Extra Keywords"), row, 0)
        self.extra_kw = QLineEdit()
        self.extra_kw.setPlaceholderText("e.g. D3BJ CPCM(water) TightSCF")
        self.extra_kw.textChanged.connect(self._on_setting_changed)
        g.addWidget(self.extra_kw, row, 1, 1, 3)

        # ── Extra Blocks ──
        row += 1
        g.addWidget(_form_label("Extra Blocks"), row, 0, Qt.AlignTop)
        self.extra_blocks = QPlainTextEdit()
        self.extra_blocks.setPlaceholderText("%scf\n  MaxIter 300\nend")
        self.extra_blocks.setMaximumHeight(80)
        self.extra_blocks.textChanged.connect(self._on_setting_changed)
        g.addWidget(self.extra_blocks, row, 1, 1, 3)

        # ── Apply to All button ──
        row += 1
        apply_all_btn = QPushButton("Apply Settings to All Tests")
        apply_all_btn.setObjectName("applyAllBtn")
        apply_all_btn.setCursor(Qt.PointingHandCursor)
        apply_all_btn.setToolTip(
            "Overwrite all per-molecule settings with the current global settings"
        )
        apply_all_btn.clicked.connect(self._apply_settings_to_all)
        g.addWidget(apply_all_btn, row, 0, 1, 4)

        return grp

    # -- Preset handling --

    def _on_preset_changed(self, name: str):
        self._update_delete_btn()
        preset = PRESETS.get(name)
        if preset is None:
            self._preset_nprocs_group = None
            return

        self._applying_preset = True
        try:
            self.func_combo.setCurrentText(preset.get("functional", "B3LYP"))
            self.basis_combo.setCurrentText(preset.get("basis_set", "def2-SVP"))
            idx = self.calc_combo.findData(preset.get("calc_type", "OPT FREQ"))
            if idx >= 0:
                self.calc_combo.setCurrentIndex(idx)
            self.charge_spin.setValue(preset.get("charge", 0))
            self.mult_spin.setValue(preset.get("multiplicity", 1))
            self.ram_spin.setValue(preset.get("ram_mb", 4000))
            self.cpus_spin.setValue(preset.get("cpus", 4))
            self.extra_kw.setText(preset.get("extra_keywords", ""))
            self.extra_blocks.setPlainText(preset.get("extra_blocks", ""))
            self._preset_nprocs_group = preset.get("nprocs_group")
        finally:
            self._applying_preset = False

    def _on_setting_changed(self):
        if self._applying_preset:
            return
        self._preset_nprocs_group = None
        self.preset_combo.blockSignals(True)
        self.preset_combo.setCurrentText("Custom")
        self.preset_combo.blockSignals(False)
        self._update_delete_btn()

    def _update_delete_btn(self):
        """Enable Delete for any non-Custom preset (including built-ins)."""
        name = self.preset_combo.currentText()
        self.del_preset_btn.setEnabled(name != "Custom")

    def _save_preset(self):
        """Save current settings as a named custom preset."""
        name, ok = QInputDialog.getText(
            self, "Save Preset", "Preset name:",
        )
        if not ok or not name.strip():
            return
        name = name.strip()

        if name == "Custom":
            QMessageBox.warning(self, "Reserved Name",
                                '"Custom" is reserved and cannot be used.')
            return

        # If overwriting an existing preset, confirm
        if name in PRESETS and PRESETS[name] is not None:
            reply = QMessageBox.question(
                self, "Overwrite Preset",
                f'A preset named "{name}" already exists.\nOverwrite it?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return

        preset_data = {
            "functional": self.func_combo.currentText().strip(),
            "basis_set": self.basis_combo.currentText().strip(),
            "calc_type": self.calc_combo.currentData() or "OPT FREQ",
            "extra_keywords": self.extra_kw.text().strip(),
            "charge": self.charge_spin.value(),
            "multiplicity": self.mult_spin.value(),
            "cpus": self.cpus_spin.value(),
            "ram_mb": self.ram_spin.value(),
            "extra_blocks": self.extra_blocks.toPlainText().strip(),
        }
        if self._preset_nprocs_group is not None:
            preset_data["nprocs_group"] = self._preset_nprocs_group

        PRESETS[name] = preset_data

        # Persist to config
        cfg = load_config()
        custom = cfg.get("custom_presets", {})
        custom[name] = preset_data
        cfg["custom_presets"] = custom
        save_config(cfg)

        if self.preset_combo.findText(name) < 0:
            self.preset_combo.addItem(name)

        self.preset_combo.blockSignals(True)
        self.preset_combo.setCurrentText(name)
        self.preset_combo.blockSignals(False)
        self._update_delete_btn()

    def _delete_preset(self):
        """Delete the currently selected preset (built-in or custom)."""
        name = self.preset_combo.currentText()
        if name == "Custom":
            return

        reply = QMessageBox.question(
            self, "Delete Preset",
            f'Delete preset "{name}"?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        PRESETS.pop(name, None)

        # Remove from persisted custom presets
        cfg = load_config()
        custom = cfg.get("custom_presets", {})
        custom.pop(name, None)
        cfg["custom_presets"] = custom
        save_config(cfg)

        idx = self.preset_combo.findText(name)
        if idx >= 0:
            self.preset_combo.blockSignals(True)
            self.preset_combo.removeItem(idx)
            self.preset_combo.setCurrentText("Custom")
            self.preset_combo.blockSignals(False)
        self._update_delete_btn()

    def _apply_settings_to_all(self):
        """Overwrite all per-molecule settings with current global settings."""
        settings = self._collect_global_settings()
        for row in range(self.mol_table.rowCount()):
            self._mol_settings[row] = copy.deepcopy(settings)
            self._update_settings_indicator(row)

    # -- Path browsers --

    def _browse_orca(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Locate ORCA Executable",
            self.orca_path_edit.text() or str(ROOT),
            "Executables (orca orca.exe);;All Files (*)",
        )
        if path:
            self.orca_path_edit.setText(path)
            self._update_orca_badge()
            self._save_config()

    def _browse_project_dir(self):
        path = QFileDialog.getExistingDirectory(
            self, "Choose Project Directory",
            self.project_dir_edit.text() or str(ROOT),
        )
        if path:
            self.project_dir_edit.setText(path)
            self._save_config()

    # -- Progress section --

    def _build_progress_section(self) -> QGroupBox:
        grp = QGroupBox("Progress")
        lay = QVBoxLayout(grp)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Starting...")
        lay.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setTextFormat(Qt.RichText)
        self.progress_label.setStyleSheet("color: #94a3b8; font-size: 12px;")
        lay.addWidget(self.progress_label)

        self.progress_molecules = QLabel()
        self.progress_molecules.setTextFormat(Qt.RichText)
        self.progress_molecules.setWordWrap(True)
        lay.addWidget(self.progress_molecules)

        return grp

    # -- Results section --

    def _build_results_section(self) -> QGroupBox:
        grp = QGroupBox("Results")
        lay = QVBoxLayout(grp)

        self.results_table = QTableWidget(0, 6)
        self.results_table.setHorizontalHeaderLabels([
            "Structure", "Name", "SMILES", "Status",
            "Gibbs Free Energy", "Errors",
        ])
        hdr = self.results_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        self.results_table.setColumnWidth(0, 140)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.Stretch)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        lay.addWidget(self.results_table)

        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignCenter)
        btn_row.setSpacing(12)

        self.report_btn = QPushButton("Open Excel Report")
        self.report_btn.setObjectName("reportBtn")
        self.report_btn.setCursor(Qt.PointingHandCursor)
        self.report_btn.clicked.connect(self._open_report)
        self.report_btn.setVisible(False)
        btn_row.addWidget(self.report_btn)

        self.folder_btn = QPushButton("Open Output Folder")
        self.folder_btn.setObjectName("folderBtn")
        self.folder_btn.setCursor(Qt.PointingHandCursor)
        self.folder_btn.clicked.connect(self._open_folder)
        self.folder_btn.setVisible(False)
        btn_row.addWidget(self.folder_btn)

        lay.addLayout(btn_row)
        return grp

    # ── Config load / save ───────────────────────────────────────

    def _load_settings_from_config(self):
        cfg = self.config
        self._applying_preset = True
        try:
            self.orca_path_edit.setText(cfg.get("orca_path", ""))
            self.project_dir_edit.setText(cfg.get("project_dir", str(ROOT)))

            preset_name = cfg.get("preset", "Custom")
            idx = self.preset_combo.findText(preset_name)
            if idx >= 0:
                self.preset_combo.setCurrentIndex(idx)

            self.func_combo.setCurrentText(cfg.get("functional", "B3LYP"))
            self.basis_combo.setCurrentText(cfg.get("basis_set", "def2-SVP"))

            ct_idx = self.calc_combo.findData(cfg.get("calc_type", "OPT FREQ"))
            if ct_idx >= 0:
                self.calc_combo.setCurrentIndex(ct_idx)

            self.charge_spin.setValue(cfg.get("charge", 0))
            self.mult_spin.setValue(cfg.get("multiplicity", 1))
            self.ram_spin.setValue(cfg.get("ram_mb", 4000))
            self.cpus_spin.setValue(cfg.get("cpus", 4))
            self.extra_kw.setText(cfg.get("extra_keywords", ""))
            self.extra_blocks.setPlainText(cfg.get("extra_blocks", ""))

            if preset_name != "Custom" and preset_name in PRESETS:
                p = PRESETS[preset_name]
                if p:
                    self._preset_nprocs_group = p.get("nprocs_group")
            self._update_delete_btn()
        finally:
            self._applying_preset = False

    def _save_config(self):
        self.config = {
            "orca_path": self.orca_path_edit.text().strip(),
            "project_dir": self.project_dir_edit.text().strip(),
            "preset": self.preset_combo.currentText(),
            "functional": self.func_combo.currentText().strip(),
            "basis_set": self.basis_combo.currentText().strip(),
            "calc_type": self.calc_combo.currentData() or "OPT FREQ",
            "charge": self.charge_spin.value(),
            "multiplicity": self.mult_spin.value(),
            "ram_mb": self.ram_spin.value(),
            "cpus": self.cpus_spin.value(),
            "extra_keywords": self.extra_kw.text().strip(),
            "extra_blocks": self.extra_blocks.toPlainText().strip(),
            "custom_presets": load_config().get("custom_presets", {}),
        }
        save_config(self.config)

    # ── Collect form data ────────────────────────────────────────

    def _collect_global_settings(self) -> dict:
        """Return the current global settings from the panel."""
        settings = {
            "functional": self.func_combo.currentText().strip(),
            "basis_set": self.basis_combo.currentText().strip(),
            "calc_type": self.calc_combo.currentData() or "OPT FREQ",
            "charge": self.charge_spin.value(),
            "multiplicity": self.mult_spin.value(),
            "ram_mb": self.ram_spin.value(),
            "cpus": self.cpus_spin.value(),
            "extra_keywords": self.extra_kw.text().strip(),
            "extra_blocks": self.extra_blocks.toPlainText().strip(),
        }
        if self._preset_nprocs_group is not None:
            settings["nprocs_group"] = self._preset_nprocs_group
        return settings

    def _collect_data(self) -> tuple[list[dict], dict, dict]:
        """Return (molecules, global_settings, per_mol_settings)."""
        molecules = []
        mol_settings_by_name: dict[str, dict] = {}

        for row in range(self.mol_table.rowCount()):
            i0 = self.mol_table.item(row, 0)
            i1 = self.mol_table.item(row, 1)
            name_text = (i0.text() if i0 else "").strip()
            smiles = (i1.text() if i1 else "").strip()
            if smiles:
                name = _sanitize_name(name_text or f"mol_{row + 1}")
                molecules.append({"name": name, "smiles": smiles})

                per_mol = self._mol_settings.get(row)
                if per_mol is not None:
                    mol_settings_by_name[name] = per_mol

        global_settings = self._collect_global_settings()
        return molecules, global_settings, mol_settings_by_name

    # ── Run pipeline ─────────────────────────────────────────────

    def _on_run_clicked(self):
        molecules, global_settings, mol_settings = self._collect_data()

        if not molecules:
            QMessageBox.warning(
                self, "No Molecules",
                "Add at least one molecule with a SMILES string.",
            )
            return

        for mol in molecules:
            if not validate_smiles(mol["smiles"]):
                QMessageBox.warning(
                    self, "Invalid SMILES",
                    f"Invalid SMILES string: {mol['smiles']}",
                )
                return

        orca_path = Path(self.orca_path_edit.text().strip())
        if not orca_path.exists():
            reply = QMessageBox.question(
                self, "ORCA Not Found",
                f"ORCA executable not found at:\n{orca_path}\n\nRun anyway?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return

        project_dir = Path(self.project_dir_edit.text().strip() or str(ROOT))

        # Save settings before running
        self._save_config()

        # Lock UI (but keep molecule table editable for dynamic queueing)
        self._queued_count = self.mol_table.rowCount()
        self.run_btn.setEnabled(False)
        self.run_btn.setText("Running\u2026")
        self.stop_btn.setVisible(True)
        self.queue_btn.setVisible(True)
        self.results_group.setVisible(False)
        self.progress_group.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("")
        self.progress_bar.setFormat("Initializing...")
        self.progress_label.setText("")
        self.progress_molecules.setText("")

        self.abort_event = threading.Event()

        self.job_status = {
            "status": "running",
            "phase": "starting",
            "current": 0,
            "total": len(molecules),
            "current_name": "",
            "elapsed": "00:00:00",
            "molecules": [
                {
                    "name": m["name"],
                    "smiles": m["smiles"],
                    "status": "pending",
                    "gibbs": None,
                    "electronic_energy": None,
                    "error": None,
                }
                for m in molecules
            ],
            "report_path": None,
            "stamp": None,
            "out_dir": None,
            "error": None,
        }

        # Prevent system sleep
        _prevent_sleep()

        self.worker = PipelineWorker(
            molecules, global_settings, mol_settings, self.job_status,
            orca_path=orca_path,
            project_dir=project_dir,
            abort_event=self.abort_event,
        )
        self.worker.finished.connect(self._on_pipeline_finished)
        self.worker.start()
        self.poll_timer.start(1000)

        self.scroll_area.ensureWidgetVisible(self.progress_group)

    def _on_stop_clicked(self):
        """Abort the running pipeline."""
        if self.abort_event:
            reply = QMessageBox.question(
                self, "Stop Calculation",
                "Are you sure you want to abort the running calculation?\n"
                "The current ORCA process will be killed.",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.abort_event.set()
                self.stop_btn.setEnabled(False)
                self.stop_btn.setText("Stopping\u2026")

    # ── Progress polling ─────────────────────────────────────────

    def _poll_progress(self):
        job = self.job_status
        if not job:
            return

        total = job.get("total", 1) or 1
        current = job.get("current", 0)
        pct = int(current / total * 100)
        self.progress_bar.setValue(pct)

        phase = job.get("phase", "")
        name = job.get("current_name", "")
        elapsed = job.get("elapsed", "00:00:00")

        # Status-aware progress-bar text
        phase_formats = {
            "starting":  "Initializing...",
            "geometry":  f"Generating 3D Coordinates for {name}...  {pct}%",
            "orca":      f"Executing ORCA Calculation for {name}...  {pct}%",
            "report":    "Building Excel Report...",
            "done":      "Complete!",
            "aborted":   "Aborted",
        }
        self.progress_bar.setFormat(
            phase_formats.get(phase, f"{phase}  {pct}%")
        )

        # Detail label below the bar
        self.progress_label.setText(
            f"{current}/{total} molecules &mdash; Elapsed: {elapsed}"
        )

        # Per-molecule status list
        status_colors = {
            "pending": "#64748b", "generating": "#3b82f6",
            "generated": "#22c55e", "running": "#3b82f6",
            "completed": "#22c55e", "warning": "#f59e0b",
            "error": "#ef4444", "aborted": "#f59e0b",
        }

        lines = []
        for m in job.get("molecules", []):
            c = status_colors.get(m["status"], "#64748b")
            parts = [
                f'<span style="color:{c}; font-size:15px;">\u25cf</span> ',
                f'<b>{m["name"]}</b> ',
                f'<span style="color:#64748b; font-family:Consolas,monospace;'
                f' font-size:11px;">{m["smiles"]}</span>',
            ]
            if m.get("gibbs") is not None:
                parts.append(
                    f' <span style="color:#22c55e;">'
                    f'(G = {m["gibbs"]:.6f} Eh)</span>'
                )
            if m.get("error"):
                parts.append(
                    f' <span style="color:#ef4444;"> &mdash; '
                    f'{m["error"]}</span>'
                )
            elif m["status"] == "running":
                parts.append(
                    ' <span style="color:#3b82f6;">Executing ORCA\u2026</span>'
                )
            elif m["status"] == "generating":
                parts.append(
                    ' <span style="color:#3b82f6;">RDKit Force Field\u2026</span>'
                )
            lines.append("".join(parts))

        self.progress_molecules.setText("<br>".join(lines))
        self.setWindowTitle(
            f"ORCA Workflow Manager \u2014 {current}/{total} [{elapsed}]"
        )

    # ── Pipeline finished ────────────────────────────────────────

    def _on_pipeline_finished(self):
        self.poll_timer.stop()
        self._poll_progress()

        # Allow system sleep again
        _allow_sleep()

        self.run_btn.setEnabled(True)
        self.run_btn.setText("Run Calculations")
        self.stop_btn.setVisible(False)
        self.stop_btn.setEnabled(True)
        self.stop_btn.setText("Stop Calculation")
        self.queue_btn.setVisible(False)
        self.setWindowTitle("ORCA Workflow Manager")

        job = self.job_status
        if not job:
            return

        if job["status"] == "completed":
            self.progress_bar.setValue(100)
            self.progress_bar.setStyleSheet(
                "QProgressBar::chunk { background: #22c55e; border-radius: 6px; }"
            )
            self.progress_bar.setFormat(
                f"Complete \u2014 {job.get('elapsed', '')}"
            )
            self._show_results()
            QApplication.alert(self, 0)
        elif job["status"] == "aborted":
            self.progress_bar.setStyleSheet(
                "QProgressBar::chunk { background: #f59e0b; border-radius: 6px; }"
            )
            self.progress_bar.setFormat(
                f"Aborted \u2014 {job.get('elapsed', '')}"
            )
            self._show_results()
        elif job["status"] == "failed":
            self.progress_bar.setStyleSheet(
                "QProgressBar::chunk { background: #ef4444; border-radius: 6px; }"
            )
            self.progress_bar.setFormat("Pipeline Failed")
            self.progress_label.setText(
                f'<span style="color:#ef4444;">Failed: '
                f'{job.get("error", "unknown error")}</span>'
            )
            QMessageBox.critical(
                self, "Pipeline Failed",
                f"The pipeline failed:\n{job.get('error', 'Unknown error')}",
            )

    # ── Show results ─────────────────────────────────────────────

    def _show_results(self):
        job = self.job_status
        if not job:
            return

        self.results_table.setRowCount(0)

        for m in job.get("molecules", []):
            row = self.results_table.rowCount()
            self.results_table.insertRow(row)
            self.results_table.setRowHeight(row, 95)

            # Structure image
            png = self.worker.png_data.get(m["name"]) if self.worker else None
            if png:
                lbl = QLabel()
                lbl.setPixmap(_pixmap_from_bytes(png))
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet("background: transparent;")
                self.results_table.setCellWidget(row, 0, lbl)
            else:
                self.results_table.setItem(row, 0, QTableWidgetItem(""))

            # Name
            self.results_table.setItem(row, 1, QTableWidgetItem(m["name"]))

            # SMILES
            si = QTableWidgetItem(m["smiles"])
            si.setForeground(QColor("#94a3b8"))
            si.setFont(QFont("Consolas", 9))
            self.results_table.setItem(row, 2, si)

            # Status
            status_text = m["status"].upper()
            st = QTableWidgetItem(status_text)
            color_map = {
                "completed": "#22c55e", "warning": "#f59e0b",
                "error": "#ef4444", "generated": "#94a3b8",
                "pending": "#64748b", "aborted": "#f59e0b",
            }
            st.setForeground(QColor(color_map.get(m["status"], "#e2e8f0")))
            st.setFont(QFont("Segoe UI", 10, QFont.Bold))
            self.results_table.setItem(row, 3, st)

            # Gibbs free energy
            g_text = f"{m['gibbs']:.8f} Eh" if m.get("gibbs") is not None else "\u2014"
            gi = QTableWidgetItem(g_text)
            gi.setFont(QFont("Consolas", 9))
            self.results_table.setItem(row, 4, gi)

            # Errors
            ei = QTableWidgetItem(m.get("error") or "")
            ei.setForeground(QColor("#ef4444"))
            self.results_table.setItem(row, 5, ei)

        # Show action buttons
        self.report_path = job.get("report_path")
        self.output_folder = job.get("out_dir")

        self.report_btn.setVisible(
            bool(self.report_path and Path(self.report_path).exists())
        )
        self.folder_btn.setVisible(
            bool(self.output_folder and Path(self.output_folder).exists())
        )

        self.results_group.setVisible(True)
        self.scroll_area.ensureWidgetVisible(self.results_group)

    # ── Actions ──────────────────────────────────────────────────

    def _open_report(self):
        if self.report_path and Path(self.report_path).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.report_path))

    def _open_folder(self):
        if self.output_folder and Path(self.output_folder).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_folder))

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "Job Running",
                "A calculation is still running. Are you sure you want to quit?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
            # Kill the running job
            if self.abort_event:
                self.abort_event.set()
            _allow_sleep()
        self._save_config()
        event.accept()


# ── Entry point ──────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setStyleSheet(STYLE)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
