# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['gui_certificate_generator.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Base', 'Base'),  # Include Base folder with templates
    ],
    hiddenimports=[
        'win32com.client',
        'win32com.gen_py',
        'openpyxl',
        'pandas',
        'tkinter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CertificateGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI only)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path if you have one: icon='icon.ico'
)
