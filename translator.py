"""
Translation module: Google Translate (online) with Argos Translate offline fallback.

Strategy:
  1. Try Google Translate via free web API (no key needed).
  2. If network is unavailable, fall back to Argos Translate (offline).
  Argos packages are loaded lazily in a background thread on first use.
"""

import json
import threading
import traceback
import urllib.parse
import urllib.request
from pathlib import Path

_LOG_FILE = Path.home() / 'translator_debug.log'


def _log(msg: str):
    try:
        with open(_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        pass
    print(msg)


# ── State ──────────────────────────────────────────────

_ready = False          # Argos offline models loaded
_failed = False         # Argos loading failed
_online = True          # last online attempt succeeded
_done_event = threading.Event()
_started = False
_start_lock = threading.Lock()

PAIRS = [('en', 'he'), ('he', 'en')]

_log('[Translator] Module loaded')


# ── Helpers ────────────────────────────────────────────

def _is_hebrew(text: str) -> bool:
    """Check if the majority of alphabetic chars are Hebrew."""
    heb = 0
    other = 0
    for ch in text:
        cp = ord(ch)
        if 0x0590 <= cp <= 0x05FF:
            heb += 1
        elif ch.isalpha():
            other += 1
    return heb > other


def detect_direction(text: str) -> str:
    """Return 'en-he' or 'he-en' based on text content."""
    return 'he-en' if _is_hebrew(text) else 'en-he'


# ── Google Translate (free, no API key) ────────────────

def _google_translate(text: str, from_code: str, to_code: str) -> str | None:
    """Translate using Google Translate's free web endpoint.

    Returns translated text or None on failure.
    """
    global _online
    try:
        # Google Translate single-string API endpoint
        url = 'https://translate.googleapis.com/translate_a/single'
        params = {
            'client': 'gtx',
            'sl': from_code,
            'tl': to_code,
            'dt': 't',
            'q': text,
        }
        qs = urllib.parse.urlencode(params)
        req = urllib.request.Request(
            f'{url}?{qs}',
            headers={
                'User-Agent': 'AutoLang/3 (Windows; Translator)',
                'Accept': 'application/json',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            raw = resp.read().decode('utf-8', errors='replace')

        # Response is a nested JSON array: [[["translated","original",...],...],...]
        data = json.loads(raw)
        translated_parts = []
        if isinstance(data, list) and len(data) > 0 and isinstance(data[0], list):
            for segment in data[0]:
                if isinstance(segment, list) and len(segment) > 0 and segment[0]:
                    translated_parts.append(segment[0])
        result = ''.join(translated_parts)
        if result:
            _online = True
            _log(f'[Translator] Google OK: {text[:40]!r} -> {result[:40]!r}')
            return result
        _log('[Translator] Google returned empty result')
        return None
    except Exception as e:
        _online = False
        _log(f'[Translator] Google failed (offline?): {e}')
        return None


# ── Argos Translate (offline fallback) ─────────────────

def _patch_stanza():
    """Inject a fake stanza module so argostranslate.sbd can import
    without pulling in torch."""
    import sys
    import types
    if 'stanza' not in sys.modules:
        fake = types.ModuleType('stanza')
        fake.__path__ = []
        fake.__file__ = 'fake'
        fake_pipeline = types.ModuleType('stanza.pipeline')
        fake_pipeline.__path__ = []
        fake_core = types.ModuleType('stanza.pipeline.core')

        class _DummyPipeline:
            def __init__(self, *a, **kw): pass
            def __call__(self, text):
                class _Doc:
                    sentences = [type('S', (), {'text': text})()]
                return _Doc()
        fake_core.Pipeline = _DummyPipeline
        fake.Pipeline = _DummyPipeline
        sys.modules['stanza'] = fake
        sys.modules['stanza.pipeline'] = fake_pipeline
        sys.modules['stanza.pipeline.core'] = fake_core
        for sub in [
            'stanza.models', 'stanza.models.common',
            'stanza.models.common.doc', 'stanza.models.common.utils',
            'stanza.resources', 'stanza.resources.common',
            'stanza.utils', 'stanza.utils.conll',
        ]:
            if sub not in sys.modules:
                m = types.ModuleType(sub)
                m.__path__ = []
                sys.modules[sub] = m
        _log('[Translator] Stanza stubbed out successfully')


def _ensure_packages() -> None:
    """Download and install Argos language packages if not already installed."""
    global _ready, _failed
    try:
        _log('[Translator] Argos: Patching stanza...')
        _patch_stanza()
        _log('[Translator] Argos: Importing argostranslate...')
        import argostranslate.package
        import argostranslate.translate
        _log('[Translator] Argos: argostranslate imported OK')

        installed = argostranslate.package.get_installed_packages()
        installed_pairs = {(p.from_code, p.to_code) for p in installed}
        _log(f'[Translator] Argos: Installed pairs: {installed_pairs}')

        missing = [p for p in PAIRS if p not in installed_pairs]
        if missing:
            _log(f'[Translator] Argos: Missing pairs: {missing}, updating index...')
            argostranslate.package.update_package_index()
            available = argostranslate.package.get_available_packages()

            for from_code, to_code in missing:
                pkg = next(
                    (p for p in available
                     if p.from_code == from_code and p.to_code == to_code),
                    None,
                )
                if pkg:
                    _log(f'[Translator] Argos: Installing {from_code}->{to_code}...')
                    argostranslate.package.install_from_path(pkg.download())
                    _log(f'[Translator] Argos: {from_code}->{to_code} ready.')
                else:
                    _log(f'[Translator] Argos: Package {from_code}->{to_code} not found!')
        else:
            _log('[Translator] Argos: Packages already installed.')

        _ready = True
        _log('[Translator] Argos: Ready.')
    except Exception as e:
        _failed = True
        _log(f'[Translator] Argos: Setup failed: {e}')
        _log(traceback.format_exc())
    finally:
        _done_event.set()


def _argos_translate(text: str, from_code: str, to_code: str) -> str | None:
    """Translate using Argos Translate (offline). Blocks until models are loaded."""
    _start_loading()
    if not _ready:
        _log('[Translator] Argos: Waiting for loading...')
        _done_event.wait(timeout=60)
    if not _ready:
        _log('[Translator] Argos: Still not ready after waiting!')
        return None
    try:
        import argostranslate.translate
        result = argostranslate.translate.translate(text, from_code, to_code)
        _log(f'[Translator] Argos OK: {text[:40]!r} -> {result[:40]!r}')
        return result
    except Exception as e:
        _log(f'[Translator] Argos failed: {e}')
        _log(traceback.format_exc())
        return None


# ── Loading / Readiness ────────────────────────────────

def _start_loading():
    """Start loading Argos models once in a background thread."""
    global _started
    with _start_lock:
        if _started:
            return
        _started = True
    _log('[Translator] Starting Argos background loading thread')
    threading.Thread(target=_ensure_packages, daemon=True).start()


def ensure_models_loaded(callback=None):
    """Start loading Argos packages in the background (for offline fallback).

    Also probes Google Translate availability.
    If callback is provided, it is called (no args) once ready.
    """
    _log(f'[Translator] ensure_models_loaded called, ready={_ready}')
    _start_loading()

    if callback:
        def _wait_and_callback():
            # Online is always "ready" -- call back immediately
            # but also wait for Argos in case we need it later
            try:
                callback()
            except Exception:
                pass
            _done_event.wait(timeout=120)
        threading.Thread(target=_wait_and_callback, daemon=True).start()


def is_ready(pair: str = None) -> bool:
    # Online is always ready; offline readiness depends on Argos
    return True


def translate(text: str, pair: str = None) -> str | None:
    """Translate text. Auto-detects direction if pair is None.

    Strategy: try Google Translate first, fall back to Argos if offline.
    Returns translated text, or None on failure.
    """
    _log(f'[Translator] translate() called, text={text[:50]!r}')
    if not text or not text.strip():
        return ''
    if pair is None:
        pair = detect_direction(text)

    parts = pair.split('-')
    if len(parts) != 2:
        return None
    from_code, to_code = parts

    # 1) Try Google Translate (online)
    result = _google_translate(text, from_code, to_code)
    if result:
        return result

    # 2) Fallback to Argos Translate (offline)
    _log('[Translator] Falling back to Argos (offline)...')
    return _argos_translate(text, from_code, to_code)
