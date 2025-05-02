# timer_sent.spec

from PyInstaller.utils.hooks import collect_submodules

a = Analysis(
    ['timer_sent.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config/icono.ico', 'config'),
    ],
    hiddenimports=collect_submodules("win32com"),
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='timer_sent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='config/icono.ico',
    single_file=True
)