#!/usr/bin/env python3
"""
build_exe.py — Build a standalone executable using PyInstaller.

Uses --onedir mode so the app starts instantly (no temp-extraction delay).
Distribute the entire  dist/ORCAWorkflowManager/  folder.

App Icon
--------
To set a custom icon, place an .ico file and pass its path:

    python build_exe.py --icon myicon.ico

Icon specs:
  - Format : .ico (Windows) — must contain multiple sizes embedded
  - Sizes  : 16x16, 32x32, 48x48, 64x64, 128x128, 256x256 px
  - Depth  : 32-bit (RGBA, with transparency)
  - Tool   : Use https://www.icoconverter.com or GIMP / ImageMagick
             to convert a PNG to a multi-resolution .ico file.
  - Place  : Anywhere; pass the path as an argument.
             Recommended: put it in the project root as "icon.ico".

Prerequisites:
    pip install pyinstaller

Usage:
    python build_exe.py                    # no icon
    python build_exe.py --icon icon.ico    # with icon
"""

import argparse
import subprocess
import sys
import platform
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Build ORCA Workflow Manager executable")
    parser.add_argument(
        "--icon",
        type=str,
        default=None,
        help="Path to an .ico file for the application icon "
             "(recommended sizes: 16/32/48/64/128/256 px, 32-bit RGBA)",
    )
    args = parser.parse_args()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", "ORCAWorkflowManager",
        # --onedir: files are pre-extracted -> near-instant startup
        # (--onefile packs everything into one exe that must decompress
        #  into a temp folder each launch, causing ~60 s delays for large apps)
        "--onedir",
        "--windowed",
        "--noconfirm",
        "--clean",
        # RDKit ships compiled extensions + data files — collect everything
        "--collect-all", "rdkit",
        # PySide6 plugins and resources
        "--collect-all", "PySide6",
        # pyqtgraph + OpenGL for 3D molecular viewer
        "--collect-all", "pyqtgraph",
        "--collect-all", "OpenGL",
        # Hidden imports that PyInstaller's analysis may miss
        "--hidden-import", "openpyxl",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "pyqtgraph.opengl",
        "--hidden-import", "OpenGL.platform.win32",
        "--hidden-import", "OpenGL.GL",
        "--hidden-import", "OpenGL.GLU",
        "--hidden-import", "numpy",
    ]

    # Exclude heavy packages pulled in transitively but never used
    for pkg in [
        "torch", "torchvision", "torchaudio",
        "cv2", "opencv-python",
        "imageio_ffmpeg", "av",
        "pyarrow",
        "scipy",
        "transformers",
        "onnxruntime",
        "botocore", "boto3", "s3transfer",
        "sklearn", "scikit-learn",
        "grpc", "grpcio",
        "pandas",
        "matplotlib",
        "IPython", "jupyter", "notebook",
        "tensorflow",
    ]:
        cmd += ["--exclude-module", pkg]

    # Resolve icon: explicit --icon flag, or auto-detect icon.ico next to this script
    icon_path = None
    if args.icon:
        icon_path = Path(args.icon).resolve()
    else:
        auto = Path(__file__).resolve().parent / "icon.ico"
        if auto.exists():
            icon_path = auto
            print(f"  Auto-detected icon: {icon_path}")

    if icon_path is not None:
        if not icon_path.exists():
            print(f"ERROR: Icon file not found: {icon_path}")
            sys.exit(1)
        if icon_path.suffix.lower() != ".ico":
            print(f"WARNING: Icon file should be .ico format, got: {icon_path.suffix}")
        cmd += ["--icon", str(icon_path)]
        # Also bundle the icon so the Qt window can use it at runtime
        sep = ";" if sys.platform == "win32" else ":"
        cmd += ["--add-data", f"{icon_path}{sep}."]
        print(f"  Icon: {icon_path}")

    # Windows version metadata (copyright, description, etc.)
    version_file = Path(__file__).resolve().parent / "version_info.txt"
    if version_file.exists():
        cmd += ["--version-file", str(version_file)]

    # Entry point
    cmd.append("desktop.py")

    print("Building standalone executable (--onedir mode)...")
    print(f"  Command: {' '.join(cmd)}\n")
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with exit code {e.returncode}.")
        print("Troubleshooting tips:")
        print("  - Make sure PyInstaller is installed: pip install pyinstaller")
        print("  - Check that all dependencies are installed in the active environment")
        sys.exit(1)

    print("\nBuild complete!")
    print("Distribute the whole folder: dist/ORCAWorkflowManager/")
    print("Launch with:  dist/ORCAWorkflowManager/ORCAWorkflowManager.exe")


if __name__ == "__main__":
    main()
