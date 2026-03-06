"""
keyboard_maps.py — Multi-language keyboard layout profiles
==========================================================
Each LanguageProfile maps physical QWERTY key positions between English
and a non-English keyboard layout.

The module auto-discovers installed keyboard layouts via Windows API
and provides the translation tables needed by the engine.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Set, Tuple


@dataclass
class LanguageProfile:
    """A single non-English keyboard layout profile."""
    code: str               # ISO 639-1 code, e.g. 'he', 'ru'
    name: str               # Display name, e.g. 'עברית'
    flag: str               # Emoji flag
    to_english: Dict[str, str]          # native_char -> en_char
    from_english: Dict[str, str]        # en_char -> native_char
    common_short_words: Dict[str, str]  # en_typed -> native_word (1-3 char shortcuts)
    unicode_ranges: List[Tuple[int, int]]  # list of (start, end) codepoint ranges for script detection
    win_lang_ids: Set[int]  # Windows LANGID values, e.g. {0x040D} for Hebrew

    def contains_script(self, text: str) -> bool:
        """Check if text contains any character from this language's script."""
        for ch in text:
            cp = ord(ch)
            for start, end in self.unicode_ranges:
                if start <= cp <= end:
                    return True
        return False

    def is_plausible(self, word: str) -> bool:
        """Quick structural check: does the word look like it belongs to this language?"""
        if not word:
            return False
        cp = ord(word[0])
        return any(s <= cp <= e for s, e in self.unicode_ranges)


# ═══════════════════════════════════════════════════════════════
# Language Profiles
# ═══════════════════════════════════════════════════════════════

# ── Hebrew ──────────────────────────────────────────────────────

_HE_TO_EN = {
    # Main keys (lowercase row by row)
    'ק': 'e', 'ר': 'r', 'א': 't', 'ט': 'y', 'ו': 'u', 'ן': 'i', 'ם': 'o', 'פ': 'p',
    'ש': 'a', 'ד': 's', 'ג': 'd', 'כ': 'f', 'ע': 'g', 'י': 'h', 'ח': 'j', 'ל': 'k', 'ך': 'l', 'ף': ';',
    'ז': 'z', 'ס': 'x', 'ב': 'c', 'ה': 'v', 'נ': 'b', 'מ': 'n', 'צ': 'm', 'ת': ',', 'ץ': '.',
    "'": 'w',
    '/': 'q',
    # Shift variants (Hebrew doesn't have uppercase, but some keys produce different chars with Shift)
}

_EN_TO_HE = {v: k for k, v in _HE_TO_EN.items()}
# Add uppercase English -> same Hebrew (Hebrew has no uppercase)
_EN_TO_HE.update({v.upper(): k for k, v in _HE_TO_EN.items() if v.isalpha()})

# Common Hebrew words -> English (for correction: use meaning instead of key mapping)
HE_WORD_TO_EN: Dict[str, str] = {
    'אוהב': 'love', 'אוהבת': 'love', 'שלום': 'hello', 'תודה': 'thanks', 'בבקשה': 'please',
    'כן': 'yes', 'לא': 'no', 'טוב': 'good', 'יפה': 'nice', 'אני': 'I', 'אתה': 'you',
    'הוא': 'he', 'היא': 'she', 'זה': 'this', 'זאת': 'this', 'מה': 'what', 'איך': 'how',
    'למה': 'why', 'איפה': 'where', 'מתי': 'when', 'אם': 'if', 'גם': 'also', 'רק': 'only',
    'עוד': 'more', 'כבר': 'already', 'עכשיו': 'now', 'פה': 'here', 'שם': 'there',
    'היום': 'today', 'מחר': 'tomorrow', 'אתמול': 'yesterday', 'אחד': 'one', 'שתיים': 'two',
    'בננה': 'banana', 'בננות': 'bananas', 'לאכול': 'eat', 'לשתות': 'drink', 'מים': 'water',
    'לחם': 'bread', 'חלב': 'milk', 'תפוח': 'apple', 'תפוחים': 'apples',
}

# English key sequences that mean a Hebrew word (wrong layout) -> correct Hebrew.
EN_WORD_TO_HE: Dict[str, str] = {
    'vrcv': 'הרבה',
    'hu,r': 'הרבה',
    'hfr': 'הרבה',
}

_HE_SHORT_WORDS = {
    # Common 1-2 letter Hebrew words typed with English layout
    'tk': 'את',
    'ka': 'שא',
    'gk': 'עך',
    'hk': 'יך',
    'fo': 'כם',
    'tm': 'אצ',
    'dv': 'גה',
    'nt': 'בא',
    'ct': 'הא',  # correction: 'v' maps to 'ה', not 'c'
    'vt': 'הא',
    'kv': 'לה',
    'cv': 'בה',
    'ka': 'שת',
    # Single letter
    't': 'א',
    'v': 'ה',
    'u': 'ו',
    'h': 'י',
    'k': 'ל',
    'n': 'מ',
    'a': 'ש',
}

HEBREW_PROFILE = LanguageProfile(
    code='he',
    name='עברית',
    flag='🇮🇱',
    to_english=_HE_TO_EN,
    from_english=_EN_TO_HE,
    common_short_words=_HE_SHORT_WORDS,
    unicode_ranges=[(0x0590, 0x05FF), (0xFB1D, 0xFB4F)],  # Hebrew block + presentation forms
    win_lang_ids={0x040D},
)

# ── Russian ─────────────────────────────────────────────────────

_RU_TO_EN = {
    'й': 'q', 'ц': 'w', 'у': 'e', 'к': 'r', 'е': 't', 'н': 'y', 'г': 'u', 'ш': 'i', 'щ': 'o', 'з': 'p',
    'х': '[', 'ъ': ']',
    'ф': 'a', 'ы': 's', 'в': 'd', 'а': 'f', 'п': 'g', 'р': 'h', 'о': 'j', 'л': 'k', 'д': 'l', 'ж': ';', 'э': "'",
    'я': 'z', 'ч': 'x', 'с': 'c', 'м': 'v', 'и': 'b', 'т': 'n', 'ь': 'm', 'б': ',', 'ю': '.',
    # Uppercase
    'Й': 'Q', 'Ц': 'W', 'У': 'E', 'К': 'R', 'Е': 'T', 'Н': 'Y', 'Г': 'U', 'Ш': 'I', 'Щ': 'O', 'З': 'P',
    'Х': '{', 'Ъ': '}',
    'Ф': 'A', 'Ы': 'S', 'В': 'D', 'А': 'F', 'П': 'G', 'Р': 'H', 'О': 'J', 'Л': 'K', 'Д': 'L', 'Ж': ':', 'Э': '"',
    'Я': 'Z', 'Ч': 'X', 'С': 'C', 'М': 'V', 'И': 'B', 'Т': 'N', 'Ь': 'M', 'Б': '<', 'Ю': '>',
    'ё': '`', 'Ё': '~',
}

_EN_TO_RU = {v: k for k, v in _RU_TO_EN.items()}

_RU_SHORT_WORDS = {
    'f': 'а',    # Russian 'a'  (и, а, в, я, о, к, с)
    'd': 'в',    # Russian 'v' (в)
    'j': 'о',    # Russian 'o'
    'z': 'я',    # Russian 'ya'
    'c': 'с',    # Russian 's'
    'r': 'к',    # Russian 'k'
    'yf': 'на',
    'gj': 'по',
    'bp': 'из',
    'lf': 'да',
    'yt': 'не',
    'nj': 'то',
    'pf': 'за',
    'jy': 'он',
}

RUSSIAN_PROFILE = LanguageProfile(
    code='ru',
    name='Русский',
    flag='🇷🇺',
    to_english=_RU_TO_EN,
    from_english=_EN_TO_RU,
    common_short_words=_RU_SHORT_WORDS,
    unicode_ranges=[(0x0400, 0x04FF), (0x0500, 0x052F)],  # Cyrillic + Cyrillic supplement
    win_lang_ids={0x0419},
)

# ── Arabic ──────────────────────────────────────────────────────

_AR_TO_EN = {
    'ض': 'q', 'ص': 'w', 'ث': 'e', 'ق': 'r', 'ف': 't', 'غ': 'y', 'ع': 'u', 'ه': 'i', 'خ': 'o', 'ح': 'p',
    'ج': '[', 'د': ']',
    'ش': 'a', 'س': 's', 'ي': 'd', 'ب': 'f', 'ل': 'g', 'ا': 'h', 'ت': 'j', 'ن': 'k', 'م': 'l', 'ك': ';', 'ط': "'",
    'ئ': 'z', 'ء': 'x', 'ؤ': 'c', 'ر': 'v', 'لا': 'b', 'ى': 'n', 'ة': 'm', 'و': ',', 'ز': '.',
    'ظ': '/',
}

_EN_TO_AR = {v: k for k, v in _AR_TO_EN.items()}

_AR_SHORT_WORDS = {
    'd': 'ي',
    'h': 'ا',
    'k': 'ن',
    'l': 'م',
    'td': 'تي',
    'hk': 'ان',
}

ARABIC_PROFILE = LanguageProfile(
    code='ar',
    name='العربية',
    flag='🇸🇦',
    to_english=_AR_TO_EN,
    from_english=_EN_TO_AR,
    common_short_words=_AR_SHORT_WORDS,
    unicode_ranges=[(0x0600, 0x06FF), (0x0750, 0x077F), (0x08A0, 0x08FF), (0xFB50, 0xFDFF), (0xFE70, 0xFEFF)],
    win_lang_ids={0x0401, 0x0801, 0x0C01, 0x1001, 0x1401, 0x1801, 0x1C01, 0x2001, 0x2401, 0x2801, 0x2C01, 0x3001, 0x3401, 0x3801, 0x3C01, 0x4001},
)

# ── Ukrainian ───────────────────────────────────────────────────

_UK_TO_EN = {
    'й': 'q', 'ц': 'w', 'у': 'e', 'к': 'r', 'е': 't', 'н': 'y', 'г': 'u', 'ш': 'i', 'щ': 'o', 'з': 'p',
    'х': '[', 'ї': ']',
    'ф': 'a', 'і': 's', 'в': 'd', 'а': 'f', 'п': 'g', 'р': 'h', 'о': 'j', 'л': 'k', 'д': 'l', 'ж': ';', 'є': "'",
    'я': 'z', 'ч': 'x', 'с': 'c', 'м': 'v', 'и': 'b', 'т': 'n', 'ь': 'm', 'б': ',', 'ю': '.',
    'ґ': '`',
}

_EN_TO_UK = {v: k for k, v in _UK_TO_EN.items()}

UKRAINIAN_PROFILE = LanguageProfile(
    code='uk',
    name='Українська',
    flag='🇺🇦',
    to_english=_UK_TO_EN,
    from_english=_EN_TO_UK,
    common_short_words={
        'f': 'а',
        'd': 'в',
        'j': 'о',
        'z': 'я',
        'c': 'с',
        'yf': 'на',
        'lf': 'да',
        'yt': 'не',
    },
    unicode_ranges=[(0x0400, 0x04FF), (0x0500, 0x052F)],
    win_lang_ids={0x0422},
)

# ── French ──────────────────────────────────────────────────────
# AZERTY layout

_FR_TO_EN = {
    'a': 'q', 'z': 'w', 'q': 'a', 'm': ';',
    'w': 'z',
    # Most letters are the same in AZERTY, the layout differs mainly in number/symbol row
    # and a few letters. For language detection, we mainly need the swapped ones.
}

_EN_TO_FR = {v: k for k, v in _FR_TO_EN.items()}

FRENCH_PROFILE = LanguageProfile(
    code='fr',
    name='Français',
    flag='🇫🇷',
    to_english=_FR_TO_EN,
    from_english=_EN_TO_FR,
    common_short_words={},
    unicode_ranges=[(0x00C0, 0x00FF)],  # Latin Extended (accented chars)
    win_lang_ids={0x040C, 0x080C, 0x0C0C, 0x100C, 0x140C},
)

# ── German ──────────────────────────────────────────────────────
# QWERTZ layout

_DE_TO_EN = {
    'z': 'y', 'y': 'z',
    'ö': ';', 'ä': "'", 'ü': '[',
    'ß': '-',
}

_EN_TO_DE = {v: k for k, v in _DE_TO_EN.items()}

GERMAN_PROFILE = LanguageProfile(
    code='de',
    name='Deutsch',
    flag='🇩🇪',
    to_english=_DE_TO_EN,
    from_english=_EN_TO_DE,
    common_short_words={},
    unicode_ranges=[(0x00C0, 0x00FF)],
    win_lang_ids={0x0407, 0x0807, 0x0C07, 0x1007, 0x1407},
)

# ── Spanish ─────────────────────────────────────────────────────

_ES_TO_EN = {
    'ñ': ';',
}

_EN_TO_ES = {v: k for k, v in _ES_TO_EN.items()}

SPANISH_PROFILE = LanguageProfile(
    code='es',
    name='Español',
    flag='🇪🇸',
    to_english=_ES_TO_EN,
    from_english=_EN_TO_ES,
    common_short_words={},
    unicode_ranges=[(0x00C0, 0x00FF)],
    win_lang_ids={0x0C0A, 0x040A, 0x080A, 0x100A, 0x140A, 0x180A, 0x1C0A, 0x200A, 0x240A, 0x280A, 0x2C0A, 0x300A, 0x340A, 0x380A, 0x3C0A, 0x400A, 0x440A, 0x480A, 0x4C0A, 0x500A},
)

# ── Greek ───────────────────────────────────────────────────────

_EL_TO_EN = {
    ';': 'q', 'ς': 'w', 'ε': 'e', 'ρ': 'r', 'τ': 't', 'υ': 'y', 'θ': 'u', 'ι': 'i', 'ο': 'o', 'π': 'p',
    'α': 'a', 'σ': 's', 'δ': 'd', 'φ': 'f', 'γ': 'g', 'η': 'h', 'ξ': 'j', 'κ': 'k', 'λ': 'l',
    'ζ': 'z', 'χ': 'x', 'ψ': 'c', 'ω': 'v', 'β': 'b', 'ν': 'n', 'μ': 'm',
}

_EN_TO_EL = {v: k for k, v in _EL_TO_EN.items()}

GREEK_PROFILE = LanguageProfile(
    code='el',
    name='Ελληνικά',
    flag='🇬🇷',
    to_english=_EL_TO_EN,
    from_english=_EN_TO_EL,
    common_short_words={},
    unicode_ranges=[(0x0370, 0x03FF), (0x1F00, 0x1FFF)],
    win_lang_ids={0x0408},
)

# ── Persian (Farsi) ─────────────────────────────────────────────

_FA_TO_EN = {
    'ض': 'q', 'ص': 'w', 'ث': 'e', 'ق': 'r', 'ف': 't', 'غ': 'y', 'ع': 'u', 'ه': 'i', 'خ': 'o', 'ح': 'p',
    'ج': '[', 'چ': ']',
    'ش': 'a', 'س': 's', 'ی': 'd', 'ب': 'f', 'ل': 'g', 'ا': 'h', 'ت': 'j', 'ن': 'k', 'م': 'l', 'ک': ';', 'گ': "'",
    'ظ': 'z', 'ط': 'x', 'ز': 'c', 'ر': 'v', 'ذ': 'b', 'د': 'n', 'پ': 'm', 'و': ',', 'ژ': '.',
}

_EN_TO_FA = {v: k for k, v in _FA_TO_EN.items()}

PERSIAN_PROFILE = LanguageProfile(
    code='fa',
    name='فارسی',
    flag='🇮🇷',
    to_english=_FA_TO_EN,
    from_english=_EN_TO_FA,
    common_short_words={},
    unicode_ranges=[(0x0600, 0x06FF), (0x0750, 0x077F), (0xFB50, 0xFDFF), (0xFE70, 0xFEFF)],
    win_lang_ids={0x0429},
)

# ── Turkish ─────────────────────────────────────────────────────
# Turkish QWERTY (mostly same as English but with ı, İ, ö, ü, ş, ç, ğ)

_TR_TO_EN = {
    'ı': 'x',  # approximate physical mapping
    'ö': ',',
    'ü': '[',
    'ş': ';',
    'ç': '.',
    'ğ': '[',
}

_EN_TO_TR = {v: k for k, v in _TR_TO_EN.items()}

TURKISH_PROFILE = LanguageProfile(
    code='tr',
    name='Türkçe',
    flag='🇹🇷',
    to_english=_TR_TO_EN,
    from_english=_EN_TO_TR,
    common_short_words={},
    unicode_ranges=[(0x00C0, 0x00FF), (0x011E, 0x011F), (0x0130, 0x0131), (0x015E, 0x015F)],
    win_lang_ids={0x041F},
)

# ── Polish ──────────────────────────────────────────────────────

POLISH_PROFILE = LanguageProfile(
    code='pl',
    name='Polski',
    flag='🇵🇱',
    to_english={},   # Polish programmer layout is QWERTY — detection is via Unicode chars
    from_english={},
    common_short_words={},
    unicode_ranges=[(0x0100, 0x017F)],  # Latin Extended-A (ą, ć, ę, ł, ń, ó, ś, ź, ż)
    win_lang_ids={0x0415},
)


# ═══════════════════════════════════════════════════════════════
# Registry of all known profiles
# ═══════════════════════════════════════════════════════════════

ALL_PROFILES: Dict[str, LanguageProfile] = {
    'he': HEBREW_PROFILE,
    'ru': RUSSIAN_PROFILE,
    'ar': ARABIC_PROFILE,
    'uk': UKRAINIAN_PROFILE,
    'fr': FRENCH_PROFILE,
    'de': GERMAN_PROFILE,
    'es': SPANISH_PROFILE,
    'el': GREEK_PROFILE,
    'fa': PERSIAN_PROFILE,
    'tr': TURKISH_PROFILE,
    'pl': POLISH_PROFILE,
}

# English language IDs (common variants)
ENGLISH_LANG_IDS: Set[int] = {
    0x0409,  # en-US
    0x0809,  # en-GB
    0x0C09,  # en-AU
    0x1009,  # en-CA
    0x1409,  # en-NZ
    0x1809,  # en-IE
    0x1C09,  # en-ZA
    0x2009,  # en-JM
    0x2409,  # en-029
    0x2809,  # en-BZ
    0x2C09,  # en-TT
    0x3009,  # en-ZW
    0x3409,  # en-PH
    0x3809,  # en-ID
    0x3C09,  # en-HK
    0x4009,  # en-IN
    0x4409,  # en-MY
    0x4809,  # en-SG
}


def is_english_lang_id(lang_id: int) -> bool:
    """Check if a LANGID corresponds to any English variant."""
    # English primary language ID is 0x09
    return (lang_id & 0x00FF) == 0x09


def lang_id_to_profile(lang_id: int) -> LanguageProfile | None:
    """Look up which LanguageProfile corresponds to a Windows LANGID."""
    for profile in ALL_PROFILES.values():
        if lang_id in profile.win_lang_ids:
            return profile
    return None


def detect_script(text: str) -> str | None:
    """Detect which non-English script is present in text. Returns language code or None."""
    for code, profile in ALL_PROFILES.items():
        if profile.contains_script(text):
            return code
    return None
