# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 실행 환경: Python 3.10
# (C:\Users\MCE\AppData\Local\Programs\Python\Python310)
# 패키지도 동일 환경에 설치되어야 함:
#   python.exe -m pip install python-docx edge-tts
import os, glob
from PyInstaller.utils.hooks import collect_all

docx_datas, docx_binaries, docx_hidden = collect_all('docx')
lxml_datas, lxml_binaries, lxml_hidden = collect_all('lxml')

# lxml C 확장(.pyd) 수동 수집
import lxml as _lxml_pkg
_lxml_dir = os.path.dirname(_lxml_pkg.__file__)
_manual_lxml_bins = [
    (fp, 'lxml')
    for fp in glob.glob(os.path.join(_lxml_dir, '*.pyd'))
]

a = Analysis(
    ['interview_tts_generator.py'],
    pathex=[],          # 동일 Python 환경이므로 추가 경로 불필요
    binaries=docx_binaries + lxml_binaries + _manual_lxml_bins,
    datas=[
        ('MCE_logo.png', '.'),
        ('MCE_logo.ico', '.'),
    ] + docx_datas + lxml_datas,
    hiddenimports=docx_hidden + lxml_hidden + [
        'edge_tts',
        'lxml.etree',
        'lxml._elementpath',
        'lxml.html',
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
    name='Interview_MP3_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
