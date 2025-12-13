# -*- mode: python ; coding: utf-8 -*-
# PyInstaller配置文件

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
        # Flask相关
        'flask',
        'werkzeug',
        'jinja2',
        'markupsafe',
        # Pandas相关
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.skiplist',
        'pandas.io.excel._openpyxl',
        'pandas.io.common',
        'pandas.io.parsers',
        # NumPy相关（pandas依赖）
        'numpy',
        'numpy.random',
        'numpy.random.bit_generator',
        'numpy.random._generator',
        'numpy.random._bounded_integers',
        # Excel处理
        'openpyxl',
        # 网络请求
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 只排除明确不需要的模块
        'matplotlib',
        'scipy',
        'tkinter',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
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
    strip=False,
    upx=False,  # 关闭UPX压缩，避免兼容性问题
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False隐藏控制台窗口，True则显示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以在这里指定图标文件路径，例如: 'icon.ico'
)
