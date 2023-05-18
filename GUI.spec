# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    [
    'TestReport\\App.py',
    'TestReport\\Copy_Issue_To_Report.py',
    'TestReport\\Copy_Test_Case.py',
    'TestReport\\Find_Issue_Number.py',
    'TestReport\\GUI.py',
    'TestReport\\Handle_Issue_Table.py',
    'TestReport\\Judge_New_Issue.py',
    'TestReport\\Merge_Report_Table.py',
    'TestReport\\Refresh_All_Table.py',
    ],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name='单体报告自动化工具',
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
