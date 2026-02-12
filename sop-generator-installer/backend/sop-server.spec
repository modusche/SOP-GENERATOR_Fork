# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\camu-new\\camunda-modeler\\client\\src\\plugins\\sop-generator-installer\\backend\\sop_server.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates'), ('final_master_template_2.docx', '.'), ('sabah_template.docx', '.'), ('sana_template.docx', '.'), ('tarabut_template.docx', '.'), ('window_world_template.docx', '.')],
    hiddenimports=['waitress', 'docxtpl', 'lxml', 'flask'],
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
    name='sop-server',
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
