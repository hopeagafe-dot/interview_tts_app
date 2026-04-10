# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['interview_tts_generator.py'],
    # user-site-packages 경로 추가 (pip install 이 user 설치로 진행된 경우)
    pathex=['C:\\Users\\MCE\\AppData\\Roaming\\Python\\Python314\\site-packages'],
    binaries=[],
    # 로고/아이콘 파일을 exe 번들에 포함 (목적지: '.' = exe와 같은 위치)
    datas=[
        ('MCE_logo.png', '.'),
        ('MCE_logo.ico', '.'),
    ],
    # PyInstaller 자동 탐지 실패 패키지 명시
    hiddenimports=[
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'edge_tts',
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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI 앱 — 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['MCE_logo.ico'],
)
