# -*- mode: python ; coding: utf-8 -*-
# גרסה קלה: בלי דיבור (faster_whisper/ctranslate2) ובלי תרגום אופליין (argostranslate).
# Build: pyinstaller AutoLangLight.spec
# התוצאה: dist\AutoLangLight.exe — קטן יותר, נטען מהר יותר.
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('wordfreq')
# אין faster_whisper, ctranslate2, sentencepiece

binaries = []

a = Analysis(
    ['auto_lang_ui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'auto_lang', 'keyboard', 'keyboard._winkeyboard', 'mouse', 'mouse._winmouse',
        'PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui',
        'keyboard_maps', 'wordfreq', 'translator', 'spell_module', 'grammar_module', 'spellchecker',
        'numpy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'stanza', 'transformers', 'sacremoses', 'torchaudio', 'torchvision',
              'tensorflow', 'tensorboard', 'scipy', 'matplotlib', 'IPython', 'notebook', 'jupyter',
              'faster_whisper', 'ctranslate2', 'argostranslate', 'sentencepiece', 'pyaudio',
              'speech_module'],
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
    name='AutoLangLight',
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
    uac_admin=False,
)
