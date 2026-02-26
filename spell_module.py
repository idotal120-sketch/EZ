"""
Spell checking module for AutoLang.

Uses pyspellchecker for English and wordfreq-powered Hebrew checking.
Supports three modes:
  - 'tooltip'  : show suggestion balloon, user presses Tab to accept (default)
  - 'auto'     : silently replace misspelled word with best candidate
  - 'visual'   : red underline in floating panel only (no correction)
"""

import threading
from typing import Optional

# ── Lazy imports (keep startup fast) ──────────────────────
_en_spell = None           # SpellChecker instance for English
_he_spell = None           # SpellChecker instance for Hebrew (custom dict)
_he_initialized = False
_he_init_lock = threading.Lock()

SPELL_AVAILABLE = False
try:
    from spellchecker import SpellChecker
    SPELL_AVAILABLE = True
except ImportError:
    pass

try:
    from wordfreq import zipf_frequency, top_n_list
    _WORDFREQ_AVAILABLE = True
except ImportError:
    _WORDFREQ_AVAILABLE = False

# ── Constants ─────────────────────────────────────────────
_HEBREW_RANGE = range(0x0590, 0x05FF)
_HE_PREFIXES = 'בהוכלמש'
_HE_PREFIX_COMBOS = (
    'ב', 'ה', 'ו', 'כ', 'ל', 'מ', 'ש',
    'בה', 'וה', 'כה', 'לה', 'מה', 'שה',
    'של', 'וב', 'וכ', 'ול', 'ומ', 'וש',
    'שב', 'שכ', 'של', 'שמ',
)
_MIN_WORD_LEN = 2
_HE_ZIPF_THRESHOLD = 2.0   # below this → unknown word
_HE_DICT_SIZE = 50000       # top N Hebrew words from wordfreq

# ── Public config (set by engine) ─────────────────────────
SPELL_ENABLED = True
SPELL_MODE = 'tooltip'     # 'tooltip' | 'auto' | 'visual'
SPELL_CALLBACK = None      # type: callable | None  — (word, suggestions, mode) -> None


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Lazy initialization                                          ║
# ╚═══════════════════════════════════════════════════════════════╝

def _ensure_english():
    """Lazy-init the English spellchecker."""
    global _en_spell
    if _en_spell is not None or not SPELL_AVAILABLE:
        return
    _en_spell = SpellChecker(language='en', distance=1)


def _ensure_hebrew():
    """Lazy-init the Hebrew spellchecker from wordfreq top words."""
    global _he_spell, _he_initialized
    if _he_initialized or not SPELL_AVAILABLE or not _WORDFREQ_AVAILABLE:
        return
    with _he_init_lock:
        if _he_initialized:
            return
        _he_spell = SpellChecker(language=None, distance=1)
        words = top_n_list('he', _HE_DICT_SIZE)
        freq_dict = {}
        for w in words:
            z = zipf_frequency(w, 'he')
            freq_dict[w] = max(1, int(10 ** z))
        _he_spell.word_frequency.load_dictionary(freq_dict)
        _he_initialized = True


def init_background():
    """Pre-load dictionaries in a background thread (called at startup)."""
    def _load():
        _ensure_english()
        _ensure_hebrew()
    threading.Thread(target=_load, name='spell-init', daemon=True).start()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Core spell check                                             ║
# ╚═══════════════════════════════════════════════════════════════╝

def is_hebrew_word(word: str) -> bool:
    """True if word contains at least one Hebrew letter."""
    return any(ord(c) in _HEBREW_RANGE for c in word)


def check_word(word: str) -> Optional[dict]:
    """Check a single word and return spell result.

    Returns None if the word is correct (or too short / not checkable).
    Returns {'original': str, 'suggestions': list[str]} if misspelled.
    """
    if not SPELL_ENABLED or not SPELL_AVAILABLE:
        return None
    if not word or len(word) < _MIN_WORD_LEN:
        return None

    clean = word.strip()
    if not clean:
        return None

    if is_hebrew_word(clean):
        return _check_hebrew(clean)
    elif clean.isalpha():
        return _check_english(clean)
    return None


def _check_english(word: str) -> Optional[dict]:
    _ensure_english()
    if _en_spell is None:
        return None

    lower = word.lower()
    misspelled = _en_spell.unknown([lower])
    if not misspelled:
        return None

    candidates = _en_spell.candidates(lower)
    if not candidates:
        return {'original': word, 'suggestions': []}

    # Sort by frequency (best first)
    if _WORDFREQ_AVAILABLE:
        suggestions = sorted(candidates, key=lambda w: zipf_frequency(w, 'en'), reverse=True)
    else:
        suggestions = list(candidates)

    # Preserve original case for first suggestion
    if suggestions and word[0].isupper():
        suggestions = [s.capitalize() if i == 0 else s for i, s in enumerate(suggestions)]

    return {'original': word, 'suggestions': suggestions[:5]}


def _check_hebrew(word: str) -> Optional[dict]:
    if not _WORDFREQ_AVAILABLE:
        return None

    # Quick check: is the word known in wordfreq?
    if zipf_frequency(word, 'he') >= _HE_ZIPF_THRESHOLD:
        return None

    # Check with common prefix removal
    for prefix in _HE_PREFIX_COMBOS:
        if word.startswith(prefix) and len(word) > len(prefix) + 1:
            root = word[len(prefix):]
            if zipf_frequency(root, 'he') >= _HE_ZIPF_THRESHOLD:
                return None  # word = prefix + valid root → correct

    # Word seems misspelled — get candidates from custom dictionary
    _ensure_hebrew()
    if _he_spell is None:
        return None

    misspelled = _he_spell.unknown([word])
    if not misspelled:
        return None  # known in our custom dict

    candidates = _he_spell.candidates(word)
    if not candidates:
        return {'original': word, 'suggestions': []}

    suggestions = sorted(candidates, key=lambda w: zipf_frequency(w, 'he'), reverse=True)
    return {'original': word, 'suggestions': suggestions[:5]}


# ╔═══════════════════════════════════════════════════════════════╗
# ║ High-level: check and fire callback                          ║
# ╚═══════════════════════════════════════════════════════════════╝

def check_and_notify(word: str):
    """Check word and fire SPELL_CALLBACK if misspelled.

    Called from the keyboard hook thread after a word boundary.
    """
    if not SPELL_ENABLED or not SPELL_CALLBACK:
        return

    result = check_word(word)
    if result and result['suggestions']:
        try:
            SPELL_CALLBACK(result['original'], result['suggestions'], SPELL_MODE)
        except Exception:
            pass
