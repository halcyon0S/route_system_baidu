# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates'), ('static', 'static')],
    hiddenimports=['pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.nattype', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.skiplist', 'numpy.random', 'numpy.random.bit_generator', 'numpy.random._generator', 'numpy.random._bounded_integers', 'openpyxl', 'flask', 'werkzeug', 'jinja2', 'requests'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'tkinter', 'PyQt5', 'PyQt6', 'IPython', 'jupyter', 'notebook', 'pytest', 'unittest', 'sphinx', 'pydoc'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='route_system_baidu',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
