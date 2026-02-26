"""
Speech-to-Text module for AutoLang — **Whisper (offline)**.
============================================================
Uses faster-whisper (CTranslate2-based) for local speech recognition.
Auto-detects Hebrew / English. No internet required.
Model: "small" (~461 MB) — good balance of speed vs accuracy for he+en.
"""

import threading
import os
import numpy as np

# Debug log
_LOG_PATH = os.path.expanduser('~/speech_debug.log')


def _log(msg):
    try:
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            from datetime import datetime
            f.write(f'[{datetime.now():%H:%M:%S}] {msg}\n')
    except Exception:
        pass


# ── Whisper model management ───────────────────────────
_model = None
_model_lock = threading.Lock()
_model_loading = False
_model_ready = threading.Event()
_model_failed = False
_MODEL_SIZE = 'small'          # ~461 MB, good for he+en
_DEVICE = 'cpu'                # safe default; change to 'cuda' if GPU available
_COMPUTE_TYPE = 'int8'         # fast on CPU


def _ensure_model():
    """Load the Whisper model (lazy, thread-safe, once)."""
    global _model, _model_loading, _model_failed
    with _model_lock:
        if _model is not None:
            return
        if _model_loading:
            return
        _model_loading = True

    _log(f'Loading faster-whisper model "{_MODEL_SIZE}" (device={_DEVICE}, compute={_COMPUTE_TYPE})...')
    try:
        from faster_whisper import WhisperModel
        _model = WhisperModel(
            _MODEL_SIZE,
            device=_DEVICE,
            compute_type=_COMPUTE_TYPE,
        )
        _model_ready.set()
        _log('Whisper model loaded OK')
    except Exception as e:
        _model_failed = True
        _model_ready.set()  # unblock waiters
        _log(f'Whisper model load FAILED: {e}')


def preload_model():
    """Start loading the model in background (called at app startup)."""
    threading.Thread(target=_ensure_model, daemon=True).start()


# ── Recording state ───────────────────────────────────
_is_recording = False
_stop_event = threading.Event()
_state_lock = threading.Lock()
_session_id = 0               # incremented on each start; stale threads check this

SAMPLE_RATE = 16000   # Whisper expects 16 kHz mono
CHUNK = 1024
SILENCE_THRESHOLD = 500       # RMS below this = silence
SILENCE_DURATION = 1.2        # seconds of silence to end a phrase
MIN_PHRASE_DURATION = 0.3     # ignore phrases shorter than this


def is_recording() -> bool:
    return _is_recording


def start_recording(callback, on_error=None, on_state=None):
    """
    Start continuous speech recording.
    callback(text, lang_code): called when text is recognized.
    on_error(msg): called on error.
    on_state(state): 'loading', 'listening', 'processing', 'idle'.
    """
    global _is_recording, _session_id
    with _state_lock:
        if _is_recording:
            return
        _is_recording = True
        _session_id += 1
        my_session = _session_id
        _stop_event.clear()

    t = threading.Thread(
        target=_record_loop,
        args=(callback, on_error, on_state, my_session),
        daemon=True,
    )
    t.start()
    _log(f'Recording started (session {my_session})')


def stop_recording():
    """Stop the current recording session."""
    global _is_recording
    with _state_lock:
        _is_recording = False
        _stop_event.set()
    _log('Recording stop requested')


def _record_loop(callback, on_error, on_state, my_session):
    """Main recording loop — captures audio, detects phrases, transcribes."""
    global _is_recording

    # ── Import PyAudio ──
    try:
        import pyaudio
    except ImportError:
        _log('pyaudio import failed')
        if on_error:
            on_error('חסרה חבילת PyAudio')
        _is_recording = False
        if on_state:
            on_state('idle')
        return

    # ── Ensure Whisper model is loaded ──
    if on_state:
        on_state('loading')
    _ensure_model()
    _model_ready.wait(timeout=180)

    # ── Check if we were cancelled while waiting for the model ──
    if _stop_event.is_set() or _session_id != my_session:
        _log(f'Session {my_session} cancelled during model load (current={_session_id})')
        _is_recording = False
        if on_state:
            on_state('idle')
        return

    if _model_failed or _model is None:
        _log('Model not available')
        if on_error:
            on_error('טעינת מודל Whisper נכשלה')
        _is_recording = False
        if on_state:
            on_state('idle')
        return

    _log(f'Model ready, opening microphone... (session {my_session})')

    # ── Open mic stream ──
    pa = pyaudio.PyAudio()
    stream = None
    try:
        stream = pa.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=SAMPLE_RATE,
            input=True,
            frames_per_buffer=CHUNK,
        )
        _log('Microphone opened')

        silence_chunks = int(SILENCE_DURATION * SAMPLE_RATE / CHUNK)
        min_chunks = int(MIN_PHRASE_DURATION * SAMPLE_RATE / CHUNK)

        phrase_chunks = []      # accumulated audio chunks for current phrase
        silent_count = 0        # consecutive silent chunks
        speaking = False        # are we inside a phrase?

        while not _stop_event.is_set():
            if on_state and not speaking:
                on_state('listening')

            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
            except Exception as e:
                _log(f'Stream read error: {e}')
                continue

            # Convert to numpy for RMS
            samples = np.frombuffer(data, dtype=np.int16).astype(np.float32)
            rms = np.sqrt(np.mean(samples ** 2)) if len(samples) > 0 else 0

            if rms > SILENCE_THRESHOLD:
                # Sound detected
                phrase_chunks.append(data)
                silent_count = 0
                speaking = True
            else:
                if speaking:
                    phrase_chunks.append(data)  # include trailing silence
                    silent_count += 1
                    if silent_count >= silence_chunks:
                        # End of phrase — transcribe
                        if len(phrase_chunks) >= min_chunks:
                            _transcribe_phrase(
                                phrase_chunks, callback, on_state)
                        phrase_chunks = []
                        silent_count = 0
                        speaking = False

        # Process any remaining audio
        if phrase_chunks and len(phrase_chunks) >= min_chunks:
            _transcribe_phrase(phrase_chunks, callback, on_state)

    except Exception as e:
        _log(f'Recording error: {e}')
        if on_error:
            on_error(f'שגיאת הקלטה: {e}')
    finally:
        if stream:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
        pa.terminate()
        _is_recording = False
        if on_state:
            on_state('idle')
        _log('Recording loop ended')


def _transcribe_phrase(chunks, callback, on_state):
    """Transcribe a collected phrase using Whisper."""
    if _model is None:
        return

    if on_state:
        on_state('processing')

    # Combine chunks into a single numpy array (float32, normalized)
    raw = b''.join(chunks)
    audio = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0

    _log(f'Transcribing {len(audio)/SAMPLE_RATE:.1f}s of audio...')

    try:
        segments, info = _model.transcribe(
            audio,
            beam_size=5,
            language=None,          # auto-detect language
            vad_filter=True,        # filter out non-speech
            vad_parameters=dict(
                min_silence_duration_ms=500,
            ),
        )

        detected_lang = info.language
        lang_prob = info.language_probability
        _log(f'Detected language: {detected_lang} (prob={lang_prob:.2f})')

        full_text = ''
        for segment in segments:
            full_text += segment.text

        full_text = full_text.strip()
        if full_text:
            _log(f'Transcribed [{detected_lang}]: {full_text}')
            try:
                callback(full_text, detected_lang)
                _log(f'Callback returned OK for: {full_text[:50]!r}')
            except Exception as cb_err:
                _log(f'Callback FAILED: {cb_err}')
        else:
            _log('Empty transcription result')

    except Exception as e:
        _log(f'Transcription error: {e}')
