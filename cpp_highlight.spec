# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

# Get the project directory
project_dir = Path(SPECPATH)

a = Analysis(
    ['cpp_highlight.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include theme.json in the package
        ('theme.json', '.'),
        # Include the entire cpp_highlight package
        ('cpp_highlight', 'cpp_highlight'),
    ],
    hiddenimports=[
        'openpyxl',
        'pygments',
        'pygments.lexers',
        'pygments.lexers.cpp',
        'pygments.formatters.html',
        'pygments.styles.default',
        # Package modules
        'cpp_highlight',
        'cpp_highlight.cli',
        'cpp_highlight.processor',
        'cpp_highlight.config',
        'cpp_highlight.config.settings',
        'cpp_highlight.config.theme',
        'cpp_highlight.core',
        'cpp_highlight.core.detection',
        'cpp_highlight.core.highlighter',
        'cpp_highlight.models',
        'cpp_highlight.models.text_block',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='cpp_highlight',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
