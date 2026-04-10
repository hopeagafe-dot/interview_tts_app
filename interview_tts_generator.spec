# -*- mode: python ; coding: utf-8 -*-
import os, sys, glob

# 사용자 site-packages 경로를 sys.path에 추가해 spec 안에서도 import 가능하게 함
USER_SITE = r'C:\Users\MCE\AppData\Roaming\Python\Python314\site-packages'
if USER_SITE not in sys.path:
    sys.path.insert(0, USER_SITE)

from PyInstaller.utils.hooks import collect_all

# ── docx 전체 수집 ────────────────────────────────────────────────────────────
docx_datas, docx_binaries, docx_hidden = collect_all('docx')

# ── lxml 전체 수집 + .pyd 파일 수동 추가 ─────────────────────────────────────
# collect_all 이 Python 3.14 에서 .pyd 를 누락하는 경우를 대비해 직접 glob
lxml_datas, lxml_binaries, lxml_hidden = collect_all('lxml')

import lxml as _lxml_pkg
_lxml_dir = os.path.dirname(_lxml_pkg.__file__)
_manual_lxml_bins = [
    (fp, 'lxml')
    for fp in glob.glob(os.path.join(_lxml_dir, '*.pyd'))
]

all_binaries = docx_binaries + lxml_binaries + _manual_lxml_bins

a = Analysis(
    ['interview_tts_generator.py'],
    pathex=[USER_SITE],
    binaries=all_binaries,
    datas=[
        ('MCE_logo.png', '.'),
        ('MCE_logo.ico', '.'),
    ] + docx_datas + lxml_datas,
    hiddenimports=docx_hidden + lxml_hidden + [
        'edge_tts',
        'lxml.etree',
        'lxml._elementpath',
        'lxml.html',
        'lxml.objectify',
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
    upx=False,          # lxml .pyd 압축 시 오작동 방지
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
