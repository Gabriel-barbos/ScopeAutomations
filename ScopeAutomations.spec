# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['menu_principal.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('AdicionarGrupo.xlsx', '.'),
        ('RemoverGrupo.xlsx', '.'),
        ('ID_billing.xlsx', '.'),
        ('QTGO_ID.xlsx', '.'),
        ('qtgo_automation.py', '.'),
    ],
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.common',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.wait',
        'selenium.common.exceptions',
        'openpyxl',
        'pandas',
        'xlrd',
        'xlsxwriter',
        'add_automation',
        'remove_automation',
        'billing_automation',
        'qtgo_automation'
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
    name='ScopeAutomations',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)
