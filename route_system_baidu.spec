# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置文件

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'numpy.random',
        'numpy.random.bit_generator',
        'numpy.random._generator',
        'numpy.random._bounded_integers',
        'openpyxl',
        'flask',
        'werkzeug',
        'jinja2',
        'requests',
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.edge',
        'selenium.webdriver.edge.service',
        'selenium.webdriver.edge.options',
        'selenium.webdriver.common.by',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'tkinter',
        'PyQt5',
        'PyQt6',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'unittest',
        'sphinx',
        'pydoc',
    ],
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
    name='route_system_baidu',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # Windows上strip工具可能不可用，设为False避免错误
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为 False 表示无控制台窗口，True 表示显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以在这里指定图标文件路径，例如: icon='icon.ico'
)
