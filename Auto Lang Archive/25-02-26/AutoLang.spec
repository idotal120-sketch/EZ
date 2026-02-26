# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = []
datas += collect_data_files('wordfreq')
datas += collect_data_files('ctranslate2')
datas += collect_data_files('sentencepiece')
datas += collect_data_files('faster_whisper')

binaries = []
binaries += collect_dynamic_libs('ctranslate2')


a = Analysis(
    ['auto_lang_ui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=['auto_lang', 'keyboard', 'keyboard._winkeyboard', 'mouse', 'mouse._winmouse', 'PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui', 'keyboard_maps', 'wordfreq', 'translator', 'speech_module', 'spell_module', 'grammar_module', 'spellchecker', 'pyaudio', 'faster_whisper', 'numpy', 'argostranslate', 'argostranslate.apis', 'argostranslate.apply_bpe', 'argostranslate.models', 'argostranslate.networking', 'argostranslate.package', 'argostranslate.sbd', 'argostranslate.settings', 'argostranslate.tags', 'argostranslate.tokenizer', 'argostranslate.translate', 'argostranslate.utils', 'ctranslate2', 'sentencepiece', 'minisbd'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'stanza', 'transformers', 'sacremoses', 'torchaudio', 'torchvision', 'tensorflow', 'tensorboard', 'scipy', 'matplotlib', 'IPython', 'notebook', 'jupyter'],
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
    name='AutoLang',
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
