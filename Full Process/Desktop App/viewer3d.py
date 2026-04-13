"""
viewer3d.py - Interactive 3D molecular viewer widget.

Provides an Avogadro-style 3D viewer for inspecting and manipulating
molecular geometries before ORCA submission.  Embedded inside the
desktop PySide6 application.

Supports multi-fragment SMILES (e.g. "CC.OO") for transition-state
calculations.

Requires:  pyqtgraph, PyOpenGL, numpy
"""

import numpy as np

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QSizePolicy,
    QDialog, QFormLayout, QSpinBox, QDoubleSpinBox, QCheckBox, QMessageBox,
    QDialogButtonBox,
)
from PySide6.QtCore import Signal, Qt, QRect, QPoint, QTimer, QThread
from PySide6.QtGui import QVector3D, QPainter, QPen, QColor, QBrush

try:
    import pyqtgraph as pg
    import pyqtgraph.opengl as gl
    HAS_3D_VIEWER = True
except ImportError:
    HAS_3D_VIEWER = False

from rdkit import Chem
from rdkit.Chem import AllChem

from pipeline import mol_to_xyz_block, global_conformer_search

# ── CPK-style element colours (R, G, B, A) floats 0-1 ────────────────
ELEMENT_COLORS = {
    "H":  (1.00, 1.00, 1.00, 1.0),
    "C":  (0.56, 0.56, 0.56, 1.0),
    "N":  (0.19, 0.31, 0.97, 1.0),
    "O":  (1.00, 0.05, 0.05, 1.0),
    "F":  (0.56, 0.88, 0.31, 1.0),
    "Cl": (0.12, 0.94, 0.12, 1.0),
    "Br": (0.65, 0.16, 0.16, 1.0),
    "I":  (0.58, 0.00, 0.58, 1.0),
    "S":  (1.00, 1.00, 0.19, 1.0),
    "P":  (1.00, 0.50, 0.00, 1.0),
    "B":  (1.00, 0.71, 0.71, 1.0),
    "Si": (0.94, 0.78, 0.63, 1.0),
    "Li": (0.80, 0.50, 1.00, 1.0),
    "Na": (0.67, 0.36, 0.95, 1.0),
    "K":  (0.56, 0.25, 0.83, 1.0),
}
DEFAULT_COLOR = (0.70, 0.50, 1.00, 1.0)

# Ball-and-stick atom radii (Angstroms, world-space) —
# roughly 0.3× van der Waals radius for a nice Avogadro look.
ELEMENT_RADII = {
    "H":  0.30,  "C":  0.50,  "N":  0.48,  "O":  0.45,  "F":  0.42,
    "Cl": 0.55,  "Br": 0.60,  "I":  0.68,  "S":  0.55,  "P":  0.55,
    "B":  0.50,  "Si": 0.58,  "Li": 0.50,  "Na": 0.55,  "K":  0.70,
}
DEFAULT_RADIUS = 0.50

# Double-bond offset (Å) — half-gap between the two parallel lines
DOUBLE_BOND_OFFSET = 0.12
BOND_LINE_WIDTH = 3.0

PICK_RADIUS_PX = 30  # click proximity in pixels


# ── Custom GLViewWidget that intercepts right-click for atom editing ──

if HAS_3D_VIEWER:

    class MolGLViewWidget(gl.GLViewWidget):
        """GLViewWidget subclass that redirects right-button events to
        the parent Molecule3DViewer for atom picking / dragging instead
        of the default pyqtgraph zoom behaviour.

        Left-drag  : orbit (default pyqtgraph)
        Right-click: atom pick / drag / box-select (handled by viewer)
        Middle-drag: pan   (default pyqtgraph)
        Scroll     : zoom  (default pyqtgraph)
        """

        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            # Rubber-band box selection overlay
            self._box_rect: QRect | None = None

        def mousePressEvent(self, ev):
            if ev.button() == Qt.RightButton:
                viewer = self._mol_viewer()
                if viewer is not None:
                    viewer._handle_right_press(ev)
                    ev.accept()
                    return
            super().mousePressEvent(ev)

        def mouseMoveEvent(self, ev):
            viewer = self._mol_viewer()
            if viewer is not None and (viewer._dragging or viewer._box_selecting or viewer._rotating):
                viewer._handle_right_move(ev)
                ev.accept()
                return
            super().mouseMoveEvent(ev)

        def mouseReleaseEvent(self, ev):
            if ev.button() == Qt.RightButton:
                viewer = self._mol_viewer()
                if viewer is not None:
                    viewer._handle_right_release(ev)
                    ev.accept()
                    return
            super().mouseReleaseEvent(ev)

        def keyPressEvent(self, ev):
            viewer = self._mol_viewer()
            if viewer is not None:
                if ev.key() == Qt.Key_A and ev.modifiers() & Qt.ControlModifier:
                    viewer._select_all()
                    ev.accept()
                    return
                if ev.key() == Qt.Key_Escape:
                    viewer._deselect_all()
                    ev.accept()
                    return
            super().keyPressEvent(ev)

        def paintEvent(self, ev):
            super().paintEvent(ev)
            # Draw rubber-band rectangle on top of GL scene
            if self._box_rect is not None:
                painter = QPainter(self)
                painter.setRenderHint(QPainter.Antialiasing, False)
                pen = QPen(QColor(147, 197, 253, 200))  # #93c5fd
                pen.setWidth(1)
                pen.setStyle(Qt.DashLine)
                painter.setPen(pen)
                painter.setBrush(QBrush(QColor(147, 197, 253, 35)))
                painter.drawRect(self._box_rect)
                painter.end()

        def _mol_viewer(self):
            """Walk up the parent chain to find the Molecule3DViewer."""
            p = self.parent()
            while p is not None:
                if isinstance(p, Molecule3DViewer):
                    return p
                p = p.parent()
            return None


class ConformerSearchDialog(QDialog):
    """Settings dialog for the global conformer search."""

    def __init__(self, n_rotatable: int = 0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Global Conformer Search Settings")
        self.setMinimumWidth(340)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        auto_n = max(50, min(500, 50 * n_rotatable))

        self.auto_check = QCheckBox(f"Auto ({auto_n} based on {n_rotatable} rotatable bonds)")
        self.auto_check.setChecked(True)
        self.auto_check.toggled.connect(self._on_auto_toggled)
        form.addRow("Conformers:", self.auto_check)

        self.n_confs_spin = QSpinBox()
        self.n_confs_spin.setRange(10, 2000)
        self.n_confs_spin.setValue(auto_n)
        self.n_confs_spin.setEnabled(False)
        form.addRow("  Custom count:", self.n_confs_spin)

        self.rmsd_spin = QDoubleSpinBox()
        self.rmsd_spin.setRange(0.1, 5.0)
        self.rmsd_spin.setSingleStep(0.1)
        self.rmsd_spin.setDecimals(2)
        self.rmsd_spin.setValue(0.50)
        self.rmsd_spin.setSuffix(" \u00c5")
        form.addRow("RMSD threshold:", self.rmsd_spin)

        self.repro_check = QCheckBox("Reproducible (seed = 42)")
        self.repro_check.setChecked(False)
        form.addRow("Random seed:", self.repro_check)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_auto_toggled(self, checked):
        self.n_confs_spin.setEnabled(not checked)

    def get_params(self) -> dict:
        return {
            "n_conformers": None if self.auto_check.isChecked() else self.n_confs_spin.value(),
            "rmsd_threshold": self.rmsd_spin.value(),
            "random_seed": 42 if self.repro_check.isChecked() else -1,
        }


class ConformerSearchWorker(QThread):
    """Background thread for global conformer search."""

    # (mol_3d, energy_kcal, n_unique, n_cpus, ff_name)
    finished = Signal(object, float, int, int, str)
    failed = Signal(str)

    def __init__(self, mol_3d, params: dict, parent=None):
        super().__init__(parent)
        self._mol = Chem.RWMol(mol_3d)  # work on a copy
        self._params = params

    def run(self):
        try:
            mol, energy, n_unique, n_cpus, ff_name = global_conformer_search(
                self._mol,
                n_conformers=self._params.get("n_conformers"),
                rmsd_threshold=self._params.get("rmsd_threshold", 0.5),
                random_seed=self._params.get("random_seed", -1),
            )
            self.finished.emit(mol, energy, n_unique, n_cpus, ff_name)
        except Exception as exc:
            self.failed.emit(str(exc))


class Molecule3DViewer(QWidget):
    """Interactive 3D molecular viewer with multi-atom selection and dragging.

    Features
    --------
    - Left-drag rotates the view (pyqtgraph default).
    - Right-click an atom to select it; right-drag to move selected atoms.
    - Shift+right-click to toggle atoms into / out of selection.
    - Right-click empty space to deselect all.
    - Ctrl+A selects all atoms.  Escape deselects all.
    - Scroll to zoom, middle-drag to pan.
    - Supports multi-fragment SMILES for transition-state calculations.

    Signals
    -------
    viewerClosed()
        Emitted when the user presses *Done*.
    preOptimized()
        Emitted after a successful force-field optimisation.
    conformerSearchDone()
        Emitted after a successful global conformer search.
    """

    viewerClosed = Signal()
    preOptimized = Signal()
    conformerSearchDone = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.mol_3d = None
        self.smiles = ""
        self.mol_name = ""
        self.atom_positions: np.ndarray | None = None
        self.atom_elements: list[str] = []
        self.bonds: list[tuple[int, int]] = []

        # Selection state — set of selected atom indices
        self._selected_atoms: set[int] = set()
        self._dragging: bool = False
        self._drag_plane_normal: np.ndarray | None = None
        self._drag_plane_point: np.ndarray | None = None
        self._drag_anchor: np.ndarray | None = None
        self._drag_offsets: dict[int, np.ndarray] = {}

        # Box-selection state
        self._box_selecting: bool = False
        self._box_start: tuple[float, float] = (0, 0)
        self._box_shift: bool = False  # was Shift held when box started?
        self._box_pre_selection: set[int] = set()  # selection before this box

        # Bond-rotation state (Shift+right-drag on atom)
        self._rotating: bool = False
        self._rotate_axis_origin: np.ndarray | None = None  # pivot point
        self._rotate_axis_dir: np.ndarray | None = None      # unit axis vector
        self._rotate_atoms: set[int] = set()                  # atoms to rotate
        self._rotate_initial_pos: dict[int, np.ndarray] = {}  # original positions
        self._rotate_start_mx: float = 0.0                    # mouse x at start

        # GL items
        self._atom_meshes: list = []      # one GLMeshItem per atom
        self._highlight_meshes: list = [] # glow meshes for selected atoms
        self._bond_items: list = []       # GLLinePlotItem list (flat)
        self._bond_render_info: list = [] # (atom_i, atom_j, offset_frac) per line

        if HAS_3D_VIEWER:
            self._build_ui()

    # ── UI construction ──────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)

        # Header
        hdr = QHBoxLayout()
        title = QLabel("3D Molecular Editor")
        title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #93c5fd; background: transparent;"
        )
        hdr.addWidget(title)
        hdr.addStretch()
        self.name_label = QLabel("")
        self.name_label.setStyleSheet(
            "color: #94a3b8; font-size: 12px; background: transparent;"
        )
        hdr.addWidget(self.name_label)
        layout.addLayout(hdr)

        # GL view — use our custom subclass that redirects right-click
        self.gl_view = MolGLViewWidget()
        self.gl_view.setBackgroundColor(15, 23, 42)  # #0f172a
        self.gl_view.setCameraPosition(distance=25)
        self.gl_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.gl_view.setMinimumHeight(300)
        self.gl_view.setFocusPolicy(Qt.StrongFocus)
        layout.addWidget(self.gl_view, stretch=1)

        # Instructions
        info = QLabel(
            "Left-drag: rotate view  |  Right-click atom: select & drag  |  "
            "Right-drag empty: box select\n"
            "Shift+right on atom: rotate bond  |  Shift+right empty: add to selection  |  "
            "Ctrl+A: select all  |  Escape: deselect"
        )
        info.setStyleSheet(
            "color: #64748b; font-size: 10px; background: transparent;"
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        # Status
        self.status_label = QLabel("")
        self.status_label.setStyleSheet(
            "color: #94a3b8; font-size: 11px; background: transparent;"
        )
        layout.addWidget(self.status_label)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.ff_btn = QPushButton("Force Field Optimize")
        self.ff_btn.setObjectName("applyAllBtn")
        self.ff_btn.setCursor(Qt.PointingHandCursor)
        self.ff_btn.setToolTip(
            "Run MMFF/UFF force-field optimisation on the current geometry"
        )
        self.ff_btn.clicked.connect(self._run_force_field)
        btn_row.addWidget(self.ff_btn)

        self.conf_search_btn = QPushButton("Global Conformer Search")
        self.conf_search_btn.setObjectName("confSearchBtn")
        self.conf_search_btn.setCursor(Qt.PointingHandCursor)
        self.conf_search_btn.setToolTip(
            "Runs a stochastic global conformer search. Generates many random\n"
            "starting geometries, optimises all with MMFF, deduplicates by RMSD,\n"
            "and returns the lowest-energy unique conformer.\n\n"
            "Use this instead of Force Field Optimize when the molecule has\n"
            "flexible bonds and you want to avoid local minima."
        )
        self.conf_search_btn.clicked.connect(self._run_conformer_search)
        btn_row.addWidget(self.conf_search_btn)

        btn_row.addStretch()

        self.done_btn = QPushButton("Done")
        self.done_btn.setObjectName("browseBtn")
        self.done_btn.setFixedWidth(80)
        self.done_btn.setCursor(Qt.PointingHandCursor)
        self.done_btn.clicked.connect(self._on_done)
        btn_row.addWidget(self.done_btn)

        layout.addLayout(btn_row)

    # ── Public API ───────────────────────────────────────────────────

    def load_molecule(self, mol_3d, smiles: str = "", name: str = ""):
        """Load an RDKit mol (with 3D conformer) into the viewer."""
        self.mol_3d = mol_3d
        self.smiles = smiles
        self.mol_name = name
        self._selected_atoms.clear()
        self.name_label.setText(f"{name}  ({smiles})" if name else smiles)

        n_frags = len(smiles.split(".")) if smiles else 1
        if n_frags > 1:
            self.status_label.setText(
                f"Loaded {n_frags} fragments (transition-state mode) — not yet optimised"
            )
        else:
            self.status_label.setText("Structure loaded — not yet optimised")

        self._extract_data()
        self._render(reset_camera=True)

    def get_xyz_block(self) -> str:
        """Return the current XYZ coordinate block."""
        if self.mol_3d is None:
            return ""
        return mol_to_xyz_block(self.mol_3d)

    # ── Data extraction ──────────────────────────────────────────────

    def _extract_data(self):
        if not self.mol_3d or self.mol_3d.GetNumConformers() == 0:
            self.atom_positions = None
            self.atom_elements = []
            self.bonds = []
            return

        conf = self.mol_3d.GetConformer()
        n = self.mol_3d.GetNumAtoms()

        self.atom_positions = np.zeros((n, 3), dtype=np.float64)
        self.atom_elements = []
        for i in range(n):
            pos = conf.GetAtomPosition(i)
            self.atom_positions[i] = [pos.x, pos.y, pos.z]
            self.atom_elements.append(self.mol_3d.GetAtomWithIdx(i).GetSymbol())

        self.bonds = []
        for bond in self.mol_3d.GetBonds():
            order = bond.GetBondTypeAsDouble()  # 1.0, 1.5, 2.0, 3.0
            self.bonds.append((bond.GetBeginAtomIdx(), bond.GetEndAtomIdx(), order))

    # ── Camera helpers ───────────────────────────────────────────────

    def _save_camera(self) -> dict:
        opts = self.gl_view.opts
        return {
            "center": pg.Vector(opts["center"]),
            "distance": opts["distance"],
            "elevation": opts["elevation"],
            "azimuth": opts["azimuth"],
            "fov": opts.get("fov", 60),
        }

    def _restore_camera(self, cam: dict):
        self.gl_view.opts["center"] = cam["center"]
        self.gl_view.opts["distance"] = cam["distance"]
        self.gl_view.opts["elevation"] = cam["elevation"]
        self.gl_view.opts["azimuth"] = cam["azimuth"]
        if "fov" in cam:
            self.gl_view.opts["fov"] = cam["fov"]

    def _get_camera_position(self) -> np.ndarray:
        cam_opts = self.gl_view.opts
        center = cam_opts["center"]
        dist = cam_opts["distance"]
        elev = np.radians(cam_opts["elevation"])
        azim = np.radians(cam_opts["azimuth"])

        cam_x = center.x() + dist * np.cos(elev) * np.cos(azim)
        cam_y = center.y() + dist * np.cos(elev) * np.sin(azim)
        cam_z = center.z() + dist * np.sin(elev)
        return np.array([cam_x, cam_y, cam_z])

    # ── Rendering ────────────────────────────────────────────────────

    @staticmethod
    def _bond_perp(p1, p2):
        """Return a unit vector perpendicular to the bond p1→p2."""
        d = p2 - p1
        n = np.linalg.norm(d)
        if n < 1e-12:
            return np.array([1.0, 0.0, 0.0])
        d /= n
        ref = np.array([0.0, 1.0, 0.0]) if abs(d[1]) < 0.9 else np.array([1.0, 0.0, 0.0])
        perp = np.cross(d, ref)
        pn = np.linalg.norm(perp)
        if pn < 1e-12:
            return np.array([1.0, 0.0, 0.0])
        return perp / pn

    def _render(self, reset_camera: bool = False):
        """(Re-)render the molecule.  Preserves camera unless reset_camera=True."""
        cam = None
        if not reset_camera:
            try:
                cam = self._save_camera()
            except Exception:
                pass

        for item in list(self.gl_view.items):
            self.gl_view.removeItem(item)
        self._atom_meshes.clear()
        self._highlight_meshes.clear()
        self._bond_items.clear()
        self._bond_render_info.clear()

        if self.atom_positions is None or len(self.atom_positions) == 0:
            return

        # Shared sphere mesh template (used for all atoms)
        sphere_md = gl.MeshData.sphere(rows=14, cols=14)

        # ── Atoms as 3D shaded spheres ──────────────────────────────
        for i, (pos, elem) in enumerate(
            zip(self.atom_positions, self.atom_elements)
        ):
            rgba = ELEMENT_COLORS.get(elem, DEFAULT_COLOR)
            radius = ELEMENT_RADII.get(elem, DEFAULT_RADIUS)

            is_selected = i in self._selected_atoms
            if is_selected:
                # Brighten colour for selected atoms
                color = (
                    min(rgba[0] + 0.35, 1.0),
                    min(rgba[1] + 0.35, 1.0),
                    min(rgba[2] + 0.10, 1.0),
                    rgba[3],
                )
            else:
                color = rgba

            mesh = gl.GLMeshItem(
                meshdata=sphere_md,
                smooth=True,
                shader='shaded',
                color=color,
                glOptions='opaque',
            )
            mesh.resetTransform()
            mesh.scale(radius, radius, radius)
            mesh.translate(float(pos[0]), float(pos[1]), float(pos[2]))
            self.gl_view.addItem(mesh)
            self._atom_meshes.append(mesh)

        # ── Selection glow (slightly larger translucent sphere) ─────
        if self._selected_atoms:
            for idx in sorted(self._selected_atoms):
                pos = self.atom_positions[idx]
                elem = self.atom_elements[idx]
                radius = ELEMENT_RADII.get(elem, DEFAULT_RADIUS) * 1.5

                glow = gl.GLMeshItem(
                    meshdata=sphere_md,
                    smooth=True,
                    shader='shaded',
                    color=(1.0, 1.0, 0.2, 0.30),
                    glOptions='translucent',
                )
                glow.resetTransform()
                glow.scale(radius, radius, radius)
                glow.translate(float(pos[0]), float(pos[1]), float(pos[2]))
                self.gl_view.addItem(glow)
                self._highlight_meshes.append(glow)

        # ── Bonds (single / double / triple / aromatic) ─────────────
        bond_color = (0.50, 0.50, 0.55, 1.0)

        for i, j, order in self.bonds:
            p1 = self.atom_positions[i]
            p2 = self.atom_positions[j]

            if order >= 2.0:
                # Double bond → two parallel lines
                perp = self._bond_perp(p1, p2) * DOUBLE_BOND_OFFSET
                for frac in (-1.0, 1.0):
                    off = perp * frac
                    pts = np.array([p1 + off, p2 + off], dtype=np.float32)
                    line = gl.GLLinePlotItem(
                        pos=pts, color=bond_color,
                        width=BOND_LINE_WIDTH, antialias=True,
                        glOptions='opaque',
                    )
                    self.gl_view.addItem(line)
                    self._bond_items.append(line)
                    self._bond_render_info.append((i, j, frac))

                if order >= 3.0:
                    # Triple bond → add center line
                    pts = np.array([p1, p2], dtype=np.float32)
                    line = gl.GLLinePlotItem(
                        pos=pts, color=bond_color,
                        width=BOND_LINE_WIDTH, antialias=True,
                        glOptions='opaque',
                    )
                    self.gl_view.addItem(line)
                    self._bond_items.append(line)
                    self._bond_render_info.append((i, j, 0.0))

            elif order == 1.5:
                # Aromatic — solid + dashed
                pts = np.array([p1, p2], dtype=np.float32)
                line = gl.GLLinePlotItem(
                    pos=pts, color=bond_color,
                    width=BOND_LINE_WIDTH, antialias=True,
                    glOptions='opaque',
                )
                self.gl_view.addItem(line)
                self._bond_items.append(line)
                self._bond_render_info.append((i, j, 0.0))

                # Second dashed-like thinner line offset
                perp = self._bond_perp(p1, p2) * DOUBLE_BOND_OFFSET
                pts2 = np.array([p1 + perp, p2 + perp], dtype=np.float32)
                line2 = gl.GLLinePlotItem(
                    pos=pts2, color=(0.50, 0.50, 0.55, 0.50),
                    width=BOND_LINE_WIDTH * 0.6, antialias=True,
                    glOptions='translucent',
                )
                self.gl_view.addItem(line2)
                self._bond_items.append(line2)
                self._bond_render_info.append((i, j, 1.0))

            else:
                # Single bond
                pts = np.array([p1, p2], dtype=np.float32)
                line = gl.GLLinePlotItem(
                    pos=pts, color=bond_color,
                    width=BOND_LINE_WIDTH, antialias=True,
                    glOptions='opaque',
                )
                self.gl_view.addItem(line)
                self._bond_items.append(line)
                self._bond_render_info.append((i, j, 0.0))

        if reset_camera:
            center = self.atom_positions.mean(axis=0)
            extent = np.max(np.linalg.norm(self.atom_positions - center, axis=1))
            self.gl_view.opts["center"] = pg.Vector(
                float(center[0]), float(center[1]), float(center[2])
            )
            self.gl_view.opts["distance"] = max(float(extent) * 3.5, 10.0)
        elif cam is not None:
            self._restore_camera(cam)

        self.gl_view.update()

    def _update_positions(self):
        """Fast path: push updated positions without full re-render."""
        if self.atom_positions is None:
            return

        # Move atom spheres
        for idx, mesh in enumerate(self._atom_meshes):
            pos = self.atom_positions[idx]
            elem = self.atom_elements[idx]
            radius = ELEMENT_RADII.get(elem, DEFAULT_RADIUS)
            mesh.resetTransform()
            mesh.scale(radius, radius, radius)
            mesh.translate(float(pos[0]), float(pos[1]), float(pos[2]))

        # Move highlight glows
        sel_list = sorted(self._selected_atoms) if self._selected_atoms else []
        for hi, idx in enumerate(sel_list):
            if hi < len(self._highlight_meshes):
                pos = self.atom_positions[idx]
                elem = self.atom_elements[idx]
                radius = ELEMENT_RADII.get(elem, DEFAULT_RADIUS) * 1.5
                glow = self._highlight_meshes[hi]
                glow.resetTransform()
                glow.scale(radius, radius, radius)
                glow.translate(float(pos[0]), float(pos[1]), float(pos[2]))

        # Move bond lines (recompute perpendicular offsets for double bonds)
        for li, (ai, aj, frac) in enumerate(self._bond_render_info):
            if li >= len(self._bond_items):
                break
            p1 = self.atom_positions[ai]
            p2 = self.atom_positions[aj]
            if frac != 0.0:
                perp = self._bond_perp(p1, p2) * DOUBLE_BOND_OFFSET * frac
                pts = np.array([p1 + perp, p2 + perp], dtype=np.float32)
            else:
                pts = np.array([p1, p2], dtype=np.float32)
            self._bond_items[li].setData(pos=pts)

    # ── Projection helper ───────────────────────────────────────────

    def _get_mvp(self):
        """Return the model-view-projection QMatrix4x4."""
        w = self.gl_view.width()
        h = self.gl_view.height()
        region = (0, 0, w, h)
        viewport = (0, 0, w, h)
        view = self.gl_view.viewMatrix()
        proj = self.gl_view.projectionMatrix(region, viewport)
        return proj * view

    # ── Atom picking ─────────────────────────────────────────────────

    def _pick_atom(self, mouse_x: float, mouse_y: float) -> int:
        """Return index of atom nearest to screen point, or -1."""
        if self.atom_positions is None or len(self.atom_positions) == 0:
            return -1

        w = self.gl_view.width()
        h = self.gl_view.height()
        if w == 0 or h == 0:
            return -1

        mvp = self._get_mvp()

        best_idx = -1
        best_dist = PICK_RADIUS_PX

        for i, pos in enumerate(self.atom_positions):
            v = QVector3D(float(pos[0]), float(pos[1]), float(pos[2]))
            ndc = mvp.map(v)
            sx = (ndc.x() + 1.0) * 0.5 * w
            sy = (1.0 - ndc.y()) * 0.5 * h
            d = ((sx - mouse_x) ** 2 + (sy - mouse_y) ** 2) ** 0.5
            if d < best_dist:
                best_dist = d
                best_idx = i

        return best_idx

    def _screen_to_ray(self, mouse_x: float, mouse_y: float):
        """Return (origin, direction) of the ray through the given screen pixel."""
        w = self.gl_view.width()
        h = self.gl_view.height()

        ndc_x = 2.0 * mouse_x / w - 1.0
        ndc_y = 1.0 - 2.0 * mouse_y / h

        mvp = self._get_mvp()
        inv_mvp, ok = mvp.inverted()
        if not ok:
            return np.zeros(3), np.array([0, 0, -1.0])

        near = inv_mvp.map(QVector3D(ndc_x, ndc_y, -1.0))
        far = inv_mvp.map(QVector3D(ndc_x, ndc_y, 1.0))

        origin = np.array([near.x(), near.y(), near.z()])
        direction = np.array(
            [far.x() - near.x(), far.y() - near.y(), far.z() - near.z()]
        )
        norm = np.linalg.norm(direction)
        if norm > 1e-12:
            direction /= norm
        return origin, direction

    def _ray_plane_intersect(self, ray_origin, ray_dir, plane_point, plane_normal):
        denom = np.dot(ray_dir, plane_normal)
        if abs(denom) < 1e-12:
            return plane_point.copy()
        t = np.dot(plane_point - ray_origin, plane_normal) / denom
        return ray_origin + t * ray_dir

    # ── Screen-space projection for box selection ──────────────────

    def _project_all_atoms(self) -> list[tuple[float, float]]:
        """Project every atom to screen (x, y) coordinates."""
        w = self.gl_view.width()
        h = self.gl_view.height()
        if w == 0 or h == 0 or self.atom_positions is None:
            return []

        mvp = self._get_mvp()
        screen_pts = []
        for pos in self.atom_positions:
            v = QVector3D(float(pos[0]), float(pos[1]), float(pos[2]))
            ndc = mvp.map(v)
            sx = (ndc.x() + 1.0) * 0.5 * w
            sy = (1.0 - ndc.y()) * 0.5 * h
            screen_pts.append((sx, sy))
        return screen_pts

    def _atoms_in_rect(self, x1: float, y1: float,
                       x2: float, y2: float) -> set[int]:
        """Return indices of atoms whose screen projections fall inside
        the rectangle from (x1, y1) to (x2, y2)."""
        lo_x, hi_x = min(x1, x2), max(x1, x2)
        lo_y, hi_y = min(y1, y2), max(y1, y2)
        result = set()
        for i, (sx, sy) in enumerate(self._project_all_atoms()):
            if lo_x <= sx <= hi_x and lo_y <= sy <= hi_y:
                result.add(i)
        return result

    # ── Bond rotation helpers ──────────────────────────────────────────

    def _build_adjacency(self) -> dict[int, set[int]]:
        adj: dict[int, set[int]] = {}
        for i, j, *_ in self.bonds:
            adj.setdefault(i, set()).add(j)
            adj.setdefault(j, set()).add(i)
        return adj

    def _find_rotation_group(self, atom_idx: int):
        """Determine the bond to rotate around and the fragment to move.

        Returns (pivot_idx, rotate_atoms) where:
        - pivot_idx is the neighbour anchoring the rotation axis
        - rotate_atoms is the set of atom indices on atom_idx's side
        Returns (None, None) if rotation is not possible (e.g. ring bond,
        isolated atom).
        """
        adj = self._build_adjacency()
        neighbors = list(adj.get(atom_idx, []))
        if not neighbors:
            return None, None

        n_total = len(self.atom_positions)

        # Pick the neighbour whose "other side" is largest (= main body).
        # That neighbour becomes the fixed pivot; we rotate atom_idx's side.
        best_pivot = None
        best_frag = None

        for nb in neighbors:
            # BFS from atom_idx, skipping the edge atom_idx↔nb
            visited = {atom_idx}
            queue = [atom_idx]
            while queue:
                node = queue.pop(0)
                for nxt in adj.get(node, []):
                    if nxt not in visited:
                        if node == atom_idx and nxt == nb:
                            continue  # skip this one edge
                        visited.add(nxt)
                        queue.append(nxt)

            # If visited includes nb (reachable via another path), the bond
            # is in a ring → rotation would break geometry.  Skip it.
            if nb in visited:
                continue

            # visited = atoms on atom_idx's side when cutting this bond
            if best_frag is None or len(visited) < len(best_frag):
                best_frag = visited
                best_pivot = nb

        if best_pivot is None:
            return None, None

        return best_pivot, best_frag

    @staticmethod
    def _rodrigues_rotate(points: np.ndarray, origin: np.ndarray,
                          axis: np.ndarray, angle_rad: float) -> np.ndarray:
        """Rotate *points* around *axis* through *origin* (Rodrigues)."""
        k = axis
        cos_a = np.cos(angle_rad)
        sin_a = np.sin(angle_rad)
        shifted = points - origin
        # vectorised Rodrigues: v' = v cosθ + (k×v) sinθ + k(k·v)(1-cosθ)
        dot = (shifted * k).sum(axis=1, keepdims=True)
        cross = np.cross(k, shifted)
        rotated = shifted * cos_a + cross * sin_a + k * dot * (1 - cos_a)
        return rotated + origin

    # ── Right-button handlers (called by MolGLViewWidget) ────────────

    def _handle_right_press(self, ev):
        """Right-click: pick atom → drag, or start box selection on empty space."""
        mx, my = ev.position().x(), ev.position().y()
        idx = self._pick_atom(mx, my)
        shift = bool(ev.modifiers() & Qt.ShiftModifier)

        if idx >= 0:
            if shift:
                # ── Shift+right-click on atom: bond rotation mode ──
                pivot, frag = self._find_rotation_group(idx)
                if pivot is None:
                    self.status_label.setText(
                        "Cannot rotate — atom is in a ring or isolated"
                    )
                    return

                self._rotating = True
                self._dragging = False
                self._box_selecting = False

                pivot_pos = self.atom_positions[pivot]
                atom_pos = self.atom_positions[idx]
                axis = atom_pos - pivot_pos
                norm = np.linalg.norm(axis)
                if norm < 1e-12:
                    self._rotating = False
                    return
                self._rotate_axis_dir = axis / norm
                self._rotate_axis_origin = pivot_pos.copy()
                self._rotate_atoms = frag
                self._rotate_initial_pos = {
                    i: self.atom_positions[i].copy() for i in frag
                }
                self._rotate_start_mx = mx

                # Highlight the atoms that will rotate
                self._selected_atoms = set(frag)
                self._render()
                self.status_label.setText(
                    f"Rotating {len(frag)} atoms around bond "
                    f"{pivot}-{idx} — drag left/right"
                )
                return

            # ── Plain right-click on atom: select + start drag ──
            if idx not in self._selected_atoms:
                self._selected_atoms = {idx}

            if not self._selected_atoms:
                self._dragging = False
                self._render()
                self.status_label.setText("")
                return

            # Prepare drag for all selected atoms
            self._dragging = True
            self._box_selecting = False
            self._rotating = False
            cam_pos = self._get_camera_position()
            sel_list = sorted(self._selected_atoms)
            centroid = self.atom_positions[sel_list].mean(axis=0)

            view_dir = centroid - cam_pos
            n = np.linalg.norm(view_dir)
            if n > 1e-12:
                view_dir /= n
            else:
                view_dir = np.array([0.0, 0.0, -1.0])

            self._drag_plane_normal = view_dir
            self._drag_plane_point = centroid.copy()

            ray_o, ray_d = self._screen_to_ray(mx, my)
            self._drag_anchor = self._ray_plane_intersect(
                ray_o, ray_d, self._drag_plane_point, self._drag_plane_normal
            )

            self._drag_offsets = {}
            for si in sel_list:
                self._drag_offsets[si] = (
                    self.atom_positions[si].copy() - self._drag_anchor
                )

            self._render()
            n_sel = len(self._selected_atoms)
            if n_sel == 1:
                elem = self.atom_elements[idx] if idx < len(self.atom_elements) else "?"
                self.status_label.setText(
                    f"Selected atom {idx} ({elem}) — drag to move"
                )
            else:
                self.status_label.setText(f"Selected {n_sel} atoms — drag to move")
        else:
            # ── Clicked empty space: begin rubber-band box selection ──
            self._dragging = False
            self._box_selecting = True
            self._box_start = (mx, my)
            self._box_shift = shift
            # Remember current selection so Shift adds to it
            self._box_pre_selection = (
                set(self._selected_atoms) if shift else set()
            )
            if not shift:
                self._selected_atoms.clear()
                self._render()
            self.status_label.setText("Drag to select region...")

    def _handle_right_move(self, ev):
        """Right-drag: move selected atoms, update box selection, or rotate."""
        mx, my = ev.position().x(), ev.position().y()

        if self._rotating:
            # Horizontal mouse delta → rotation angle (0.5 deg per pixel)
            delta_px = mx - self._rotate_start_mx
            angle_rad = np.radians(delta_px * 0.5)

            # Rotate from the original saved positions (avoids drift)
            indices = sorted(self._rotate_atoms)
            orig = np.array([self._rotate_initial_pos[i] for i in indices])
            rotated = self._rodrigues_rotate(
                orig, self._rotate_axis_origin, self._rotate_axis_dir, angle_rad
            )

            conf = self.mol_3d.GetConformer()
            for k, si in enumerate(indices):
                self.atom_positions[si] = rotated[k]
                conf.SetAtomPosition(
                    si, (float(rotated[k][0]), float(rotated[k][1]),
                         float(rotated[k][2]))
                )

            self._update_positions()
            deg = delta_px * 0.5
            self.status_label.setText(f"Rotation: {deg:+.1f}\u00b0")
            return

        if self._box_selecting:
            # Update rubber-band rectangle
            x1, y1 = self._box_start
            rect = QRect(
                QPoint(int(min(x1, mx)), int(min(y1, my))),
                QPoint(int(max(x1, mx)), int(max(y1, my))),
            )
            self.gl_view._box_rect = rect
            self.gl_view.update()  # trigger paintEvent for overlay

            # Live-update selection as user drags
            box_atoms = self._atoms_in_rect(x1, y1, mx, my)
            self._selected_atoms = self._box_pre_selection | box_atoms
            self._render()
            self.status_label.setText(
                f"Box: {len(box_atoms)} atoms  |  "
                f"Total selected: {len(self._selected_atoms)}"
            )
            return

        # Atom-drag mode
        if not self._selected_atoms or self._drag_plane_normal is None:
            return

        ray_o, ray_d = self._screen_to_ray(mx, my)
        hit = self._ray_plane_intersect(
            ray_o, ray_d, self._drag_plane_point, self._drag_plane_normal
        )

        conf = self.mol_3d.GetConformer()
        for si, offset in self._drag_offsets.items():
            new_pos = hit + offset
            self.atom_positions[si] = new_pos
            conf.SetAtomPosition(
                si, (float(new_pos[0]), float(new_pos[1]), float(new_pos[2]))
            )

        self._update_positions()

    def _handle_right_release(self, ev):
        """Right-button release: finalise box selection, rotation, or drag."""
        # Clear rubber-band overlay
        self.gl_view._box_rect = None
        self.gl_view.update()

        if self._rotating:
            self._rotating = False
            n = len(self._rotate_atoms)
            self._rotate_atoms = set()
            self._rotate_initial_pos.clear()
            self._render()
            self.status_label.setText(f"Rotated {n} atoms")
            return

        if self._box_selecting:
            self._box_selecting = False
            mx, my = ev.position().x(), ev.position().y()
            x1, y1 = self._box_start
            box_atoms = self._atoms_in_rect(x1, y1, mx, my)
            self._selected_atoms = self._box_pre_selection | box_atoms
            self._render()
            n_sel = len(self._selected_atoms)
            if n_sel:
                self.status_label.setText(
                    f"Selected {n_sel} atoms"
                    + (" (Shift to add more)" if not self._box_shift else "")
                )
            else:
                self.status_label.setText("")
            return

        if self._dragging:
            self._dragging = False
            n_sel = len(self._selected_atoms)
            if n_sel == 1:
                idx = next(iter(self._selected_atoms))
                elem = (
                    self.atom_elements[idx]
                    if idx < len(self.atom_elements)
                    else "?"
                )
                self.status_label.setText(f"Atom {idx} ({elem}) moved")
            elif n_sel > 1:
                self.status_label.setText(f"Moved {n_sel} atoms")

    # ── Keyboard handlers (called by MolGLViewWidget) ────────────────

    def _select_all(self):
        if self.atom_positions is not None:
            self._selected_atoms = set(range(len(self.atom_positions)))
            self._render()
            self.status_label.setText(
                f"Selected all {len(self._selected_atoms)} atoms"
            )

    def _deselect_all(self):
        if self._selected_atoms:
            self._selected_atoms.clear()
            self._render()
            self.status_label.setText("Selection cleared")

    # ── Global conformer search ────────────────────────────────────────

    def _run_conformer_search(self):
        """Open the settings dialog and launch the conformer search."""
        if self.mol_3d is None:
            return

        # Warn if the user edited the geometry by hand
        reply = QMessageBox.question(
            self,
            "Global Conformer Search",
            "This will replace the current geometry with the lowest-energy\n"
            "conformer found by a stochastic search. Any manual edits to\n"
            "atom positions will be lost.\n\nProceed?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply != QMessageBox.Yes:
            return

        from rdkit.Chem import Descriptors as Desc
        n_rot = Desc.NumRotatableBonds(self.mol_3d)

        dlg = ConformerSearchDialog(n_rotatable=n_rot, parent=self)
        if dlg.exec() != QDialog.Accepted:
            return

        params = dlg.get_params()

        # Disable buttons during search
        self.ff_btn.setEnabled(False)
        self.conf_search_btn.setEnabled(False)
        self.done_btn.setEnabled(False)
        import os as _os
        n_cpus = _os.cpu_count() or 4
        self._selected_atoms.clear()
        self.status_label.setText(
            f"Global conformer search running on {n_cpus} threads ..."
        )

        self._conf_worker = ConformerSearchWorker(self.mol_3d, params, parent=self)
        self._conf_worker.finished.connect(self._on_conformer_search_done)
        self._conf_worker.failed.connect(self._on_conformer_search_failed)
        self._conf_worker.start()

    def _on_conformer_search_done(self, mol_result, energy_kcal, n_unique, n_cpus, ff_name):
        """Conformer search completed successfully."""
        # Replace current molecule with the result
        self.mol_3d = mol_result
        self._extract_data()
        self._render(reset_camera=False)

        self.ff_btn.setEnabled(True)
        self.conf_search_btn.setEnabled(True)
        self.done_btn.setEnabled(True)

        self.status_label.setText(
            f"Conformer search complete — {n_unique} unique conformer(s), "
            f"best {ff_name} energy: {energy_kcal:.1f} kcal/mol  "
            f"[{n_cpus} threads]"
        )
        self.conformerSearchDone.emit()

    def _on_conformer_search_failed(self, error_msg):
        """Conformer search failed."""
        self.ff_btn.setEnabled(True)
        self.conf_search_btn.setEnabled(True)
        self.done_btn.setEnabled(True)

        self.status_label.setText(f"Conformer search failed: {error_msg}")
        QMessageBox.warning(self, "Conformer Search Failed", error_msg)

    # ── Force-field optimisation (animated) ───────────────────────────

    def _add_ionic_constraints(self, ff):
        """Add distance constraints between oppositely-charged fragments.

        UFF/MMFF often lack electrostatic terms for metal ions, so counter-
        ions like K+ just sit still.  This finds the closest cation–anion
        atom pair across fragments and adds a distance constraint so the
        FF actively pulls them together.
        """
        mol = self.mol_3d
        frags = Chem.GetMolFrags(mol, asMols=False)
        if len(frags) <= 1:
            return

        # Classify each fragment by its net formal charge
        pos_frags = []   # (frag_idx_list, net_charge)
        neg_frags = []
        for frag_idxs in frags:
            net_q = sum(mol.GetAtomWithIdx(i).GetFormalCharge() for i in frag_idxs)
            if net_q > 0:
                pos_frags.append(frag_idxs)
            elif net_q < 0:
                neg_frags.append(frag_idxs)

        if not pos_frags or not neg_frags:
            return

        conf = mol.GetConformer()

        # For each positive fragment, constrain it to the nearest negative
        for pfrag in pos_frags:
            for nfrag in neg_frags:
                # Find the charged atoms specifically
                cat_atoms = [i for i in pfrag
                             if mol.GetAtomWithIdx(i).GetFormalCharge() > 0]
                an_atoms = [i for i in nfrag
                            if mol.GetAtomWithIdx(i).GetFormalCharge() < 0]
                if not cat_atoms:
                    cat_atoms = list(pfrag)
                if not an_atoms:
                    an_atoms = list(nfrag)

                # Find closest cation–anion pair
                best_dist = 1e9
                best_pair = (cat_atoms[0], an_atoms[0])
                for ci in cat_atoms:
                    cp = conf.GetAtomPosition(ci)
                    for ai in an_atoms:
                        ap = conf.GetAtomPosition(ai)
                        d = cp.Distance(ap)
                        if d < best_dist:
                            best_dist = d
                            best_pair = (ci, ai)

                # Typical ionic contact distance (K–O ≈ 2.7 Å)
                target_dist = 2.7
                # Force constant high enough to pull them together
                force_k = 200.0
                ff.AddDistanceConstraint(
                    best_pair[0], best_pair[1],
                    target_dist - 0.3, target_dist + 0.3,
                    force_k,
                )

    def _run_force_field(self):
        if self.mol_3d is None:
            return

        self.ff_btn.setEnabled(False)
        self.done_btn.setEnabled(False)
        self._selected_atoms.clear()

        # Obtain a force-field object for step-wise minimisation
        ff = None
        ff_name = ""
        try:
            ff = AllChem.MMFFGetMoleculeForceField(
                self.mol_3d,
                AllChem.MMFFGetMoleculeProperties(self.mol_3d),
            )
            ff_name = "MMFF"
        except Exception:
            ff = None

        if ff is None:
            try:
                ff = AllChem.UFFGetMoleculeForceField(self.mol_3d)
                ff_name = "UFF"
            except Exception:
                ff = None

        if ff is None:
            self.status_label.setText("Could not set up force field")
            self.ff_btn.setEnabled(True)
            self.done_btn.setEnabled(True)
            return

        # Add distance constraints between oppositely-charged fragments
        # so counter-ions (e.g. K+ / O-) are pulled together by the FF.
        self._add_ionic_constraints(ff)

        ff.Initialize()

        # Animation state
        self._ff_obj = ff
        self._ff_step = 0
        self._ff_max_steps = 80        # number of animation frames
        self._ff_iters_per_step = 50   # FF iterations per frame
        self._ff_name = ff_name

        self.status_label.setText(f"Optimising ({ff_name}) ...")
        self._ff_timer = QTimer(self)
        self._ff_timer.setInterval(30)  # ~33 fps
        self._ff_timer.timeout.connect(self._ff_animation_step)
        self._ff_timer.start()

    def _ff_animation_step(self):
        """Run a small batch of FF iterations, update the view, repeat."""
        converged = self._ff_obj.Minimize(
            maxIts=self._ff_iters_per_step,
        )
        self._ff_step += 1

        # Pull updated coordinates from the FF back into the conformer
        conf = self.mol_3d.GetConformer()
        n = self.mol_3d.GetNumAtoms()
        pos = self._ff_obj.Positions()
        for i in range(n):
            x, y, z = pos[i * 3], pos[i * 3 + 1], pos[i * 3 + 2]
            conf.SetAtomPosition(i, (x, y, z))
            self.atom_positions[i] = [x, y, z]

        # Update visuals without full re-render (keeps camera)
        self._update_positions()

        total_done = self._ff_step * self._ff_iters_per_step
        self.status_label.setText(
            f"Optimising ({self._ff_name}) ... {total_done} iterations"
        )

        # Stop when converged or max steps reached
        if converged == 0 or self._ff_step >= self._ff_max_steps:
            self._ff_timer.stop()
            self._ff_timer.deleteLater()
            self._ff_timer = None
            self._ff_obj = None
            self._render(reset_camera=False)
            self.ff_btn.setEnabled(True)
            self.done_btn.setEnabled(True)
            result = "converged" if converged == 0 else "reached max iterations"
            self.status_label.setText(
                f"Force-field optimisation complete ({result}, {total_done} iters)"
            )
            self.preOptimized.emit()

    # ── Done ─────────────────────────────────────────────────────────

    def _on_done(self):
        self._selected_atoms.clear()
        self.viewerClosed.emit()
