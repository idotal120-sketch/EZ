"""
Speech-to-Text module for AutoLang.
===================================
Uses SpeechRecognition + Google Web Speech API.
Auto-detects Hebrew / English by trying both and picking the better result.
"""

import threading
import os

# Debug log
_LOG_PATH = os.path.expanduser('~/speech_debug.log')

def _log(msg):
    try:
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            from datetime import datetime
            f.write(f'[{datetime.now():%H:%M:%S}] {msg}\n')
    except Exception:
        pass

# ── State ──────────────────────────────────────────────
_is_recording = False
_stop_event = threading.Event()
_lock = threading.Lock()

# Callback signature: callback(text: str, lang_code: str)
# Error callback:     on_error(msg: str)
# State callback:     on_state(state: str)  -> 'listening', 'processing', 'idle'


def is_recording() -> bool:
    return _is_recording


def start_recording(callback, on_error=None, on_state=None):
    """
    Start continuous speech recording from the default microphone.
    Auto-detects Hebrew or English.

    callback(text, lang_code): called on the recording thread when text is recognized.
    on_error(msg): called on any error.
    on_state(state): called with 'listening', 'processing', 'idle'.
    """
    global _is_recording
    with _lock:
        if _is_recording:
            return
        _is_recording = True
        _stop_event.clear()

    t = threading.Thread(
        target=_record_loop,
        args=(callback, on_error, on_state),
        daemon=True,
    )
    t.start()
    _log('Recording started')


def stop_recording():
    """Stop the current recording session."""
    global _is_recording
    with _lock:
        _is_recording = False
        _stop_event.set()
    _log('Recording stop requested')


def _record_loop(callback, on_error, on_state):
    """Main recording loop — runs in background thread."""
    global _is_recording
    try:
        import speech_recognition as sr
    except ImportError as e:
        _log(f'speech_recognition import failed: {e}')
        if on_error:
            on_error('חסרה חבילת SpeechRecognition')
        _is_recording = False
        if on_state:
            on_state('idle')
        return

    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    recognizer.energy_threshold = 300
    recognizer.pause_threshold = 0.8

    try:
        mic = sr.Microphone()
    except Exception as e:
        _log(f'Microphone init failed: {e}')
        if on_error:
            on_error('לא נמצא מיקרופון')
        _is_recording = False
        if on_state:
            on_state('idle')
        return

    try:
        with mic as source:
            # Calibrate for ambient noise
            if on_state:
                on_state('calibrating')
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            _log('Ambient noise calibrated')

            while not _stop_event.is_set():
                if on_state:
                    on_state('listening')
                try:
                    audio = recognizer.listen(
                        source,
                        timeout=3,
                        phrase_time_limit=10,
                    )
                except sr.WaitTimeoutError:
                    continue  # No speech detected, keep listening

                # Always process captured audio, even if stop was requested
                # (don't discard audio that was already recorded)
                if on_state:
                    on_state('processing')

                # Try both languages in parallel
                text, lang = _recognize_auto(recognizer, audio)
                if text:
                    _log(f'Recognized [{lang}]: {text}')
                    try:
                        callback(text, lang)
                        _log(f'Callback returned OK for: {text[:50]!r}')
                    except Exception as cb_err:
                        _log(f'Callback FAILED: {cb_err}')
                else:
                    _log('No speech recognized in either language')

    except Exception as e:
        _log(f'Recording error: {e}')
        if on_error:
            on_error(f'שגיאת הקלטה: {e}')
    finally:
        _is_recording = False
        if on_state:
            on_state('idle')
        _log('Recording loop ended')


def _recognize_auto(recognizer, audio):
    """
    Try recognizing audio in both Hebrew and English.
    Returns (text, lang_code) or (None, None).
    """
    results = {}
    errors = {}
    done = threading.Event()
    lock = threading.Lock()

    def _try_lang(lang_tag, lang_code):
        try:
            import speech_recognition as sr
            # show_all=True returns the full API response with alternatives
            resp = recognizer.recognize_google(
                audio, language=lang_tag, show_all=True)
            if resp and isinstance(resp, dict):
                alts = resp.get('alternative', [])
                if alts:
                    text = alts[0].get('transcript', '')
                    confidence = alts[0].get('confidence', 0.0)
                    with lock:
                        results[lang_code] = (text, confidence)
            elif resp and isinstance(resp, list):
                # Some versions return list
                if resp:
                    text = resp[0] if isinstance(resp[0], str) else str(resp[0])
                    with lock:
                        results[lang_code] = (text, 0.5)
        except Exception as e:
            with lock:
                errors[lang_code] = str(e)

    # Run both in parallel
    t_he = threading.Thread(target=_try_lang, args=('he-IL', 'he'), daemon=True)
    t_en = threading.Thread(target=_try_lang, args=('en-US', 'en'), daemon=True)
    t_he.start()
    t_en.start()
    t_he.join(timeout=10)
    t_en.join(timeout=10)

    if not results:
        # Maybe try simple recognize (not show_all) as fallback
        try:
            text = recognizer.recognize_google(audio, language='he-IL')
            return (text, 'he')
        except Exception:
            pass
        try:
            text = recognizer.recognize_google(audio, language='en-US')
            return (text, 'en')
        except Exception:
            pass
        return (None, None)

    # Pick the result with highest confidence
    best_lang = max(results, key=lambda k: results[k][1])
    text, conf = results[best_lang]
    _log(f'Auto-detect results: {[(k, v[1]) for k, v in results.items()]} → picked {best_lang} (conf={conf})')

    if text:
        return (text, best_lang)
    return (None, None)
