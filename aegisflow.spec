# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for AegisFlow Desktop Application
===================================================
Build command:
    pyinstaller aegisflow.spec

Or use the helper script:
    build_exe.bat
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── Gather customtkinter assets (themes, images) ──────────────────────────
ctk_datas = collect_data_files("customtkinter")

# ── Hidden imports ─────────────────────────────────────────────────────────
hidden = [
    "pandas",
    "openpyxl",
    "openpyxl.cell.rich_text",
    "openpyxl.cell.text",
    "openpyxl.styles",
    "openpyxl.utils",
    "numpy",
    "customtkinter",
    "PIL",
    "PIL._tkinter_finder",
    "tkinter",
    "tkinter.filedialog",
    "mosaic_convert",
    "fill_tlf_template",
    "fill_tlf_status",
    "extract_programs",
    "generate_batch_xml",
]
hidden += collect_submodules("customtkinter")

# ── Analysis ───────────────────────────────────────────────────────────────
a = Analysis(
    ["aegisflow_app.py"],
    pathex=["."],
    binaries=[],
    datas=ctk_datas + [
        ("01_template", "01_template"),   # bundled template files
    ],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "scipy", "IPython", "notebook",
        "pytest", "setuptools", "distutils",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="AegisFlow",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="assets/aegisflow.ico",   # uncomment if you add an icon file
    # NOTE: passing binaries+datas to EXE (not COLLECT) creates a single-file EXE
)
