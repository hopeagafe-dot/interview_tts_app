# -*- mode: python ; coding: utf-8 -*-
# collect_all() 은 datas / binaries / hiddenimports 를 한번에 수집합니다.
# python-docx 가 의존하는 lxml 은 C 확장(.dll) 을 포함하므로 반드시 collect_all 필요.
from PyInstaller.utils.hooks import collect_all

docx_datas,  docx_binaries,  docx_hidden  = collect_all('docx')
lxml_datas,  lxml_binaries,  lxml_hidden  = collect_all('lxml')

a = Analysis(
    ['interview_tts_generator.py'],
    pathex=['C:\\Users\\MCE\\AppData\\Roaming\\Python\\Python314\\site-packages'],
    binaries=docx_binaries + lxml_binaries,
    datas=[
        ('MCE_logo.png', '.'),
        ('MCE_logo.ico', '.'),
    ] + docx_datas + lxml_datas,
    hiddenimports=docx_hidden + lxml_hidden + ['edge_tts'],
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
    name='Interview_MP3_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['MCE_logo.ico'],
)
