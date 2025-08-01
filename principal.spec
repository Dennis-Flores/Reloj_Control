# -*- mode: python ; coding: utf-8 -*-
import os

a = Analysis(
    ['principal.py'],
    pathex=[],
    binaries=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,

    datas=[
    (os.path.abspath('reloj_control.db'), '.'),
    (os.path.abspath('rostros'), 'rostros'),
    (os.path.abspath('face_recognition_models/models/shape_predictor_68_face_landmarks.dat'),
     'face_recognition_models/models'),
    (os.path.abspath('face_recognition_models/models/shape_predictor_5_face_landmarks.dat'),
     'face_recognition_models/models'),
    (os.path.abspath('face_recognition_models/models/mmod_human_face_detector.dat'),
     'face_recognition_models/models'),
    (os.path.abspath('face_recognition_models/models/dlib_face_recognition_resnet_model_v1.dat'),
     'face_recognition_models/models'),
],

)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BioAccess',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
)
