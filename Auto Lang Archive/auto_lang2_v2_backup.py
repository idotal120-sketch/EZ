import ctypes
import json
import os
import sys

# When running as --noconsole EXE (PyInstaller), sys.stdout/stderr are None.
# Redirect to devnull so print() calls don't crash.
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

"""
Auto Language Layout Fixer for Windows
========================================
סקריפט Python מתקדם לזיהוי אוטומטי ותיקון טעויות פריסת מקלדת (עברית/אנגלית) בזמן אמת על Windows.
תכונות עיקריות:
----------------
1. **זיהוי ותיקון אוטומטי של טעויות הקלדה**:
    - מזהה כאשר המשתמש הקליד טקסט בפריסה הלא נכונה (עברית במקום אנגלית או להיפך)
    - מתקן אוטומטית את המילה האחרונה כשמזוהה גבול מילה (רווח, סימני פיסוק)
    - דוגמאות: 'יקךךם' → 'hello', 'akuo' → 'שלום'
2. **תמיכה מיוחדת ב-WhatsApp Desktop**:
    - שיטות תיקון מותאמות לסביבת Electron (שבה פועלת WhatsApp Desktop)
    - תיקון מבוסס clipboard כאשר אפשרי
    - fallback למצב keybuffer כאשר clipboard חסום
    - תיקון עמיד יותר באמצעות Ctrl+Backspace או מחיקה מדויקת
    - טיפול מיוחד בטקסט RTL (Right-to-Left)
3. **החלפת פריסה אינטליגנטית**:
    - מחליפה אוטומטית את פריסת המקלדת במערכת אחרי מספר תיקונים רצופים לאותה שפה
    - תומכת במספר קיצורי מקשים נפוצים להחלפת שפה (Alt+Shift, Win+Space, Ctrl+Shift)
4. **פריסת ברירת מחדל לפי אפליקציה**:
    - מגדירה אוטומטית את השפה המועדפת כאשר עוברים בין אפליקציות
    - ניתן להגדיר שפת ברירת מחדל לפי שם exe או כותרת חלון
    - דוגמאות: VS Code → אנגלית, WhatsApp → עברית
5. **היוריסטיקות חכמות**:
    - מזהה מילים אנגליות נפוצות ומונע תיקון מיותר
    - בודקת הגיון לשוני (תנועות באנגלית, צפיפות עיצורים)
    - טיפול במילים קצרות (2-3 אותיות) באופן שמרני יותר
    - תמיכה בראשי תיבות ללא תנועות (idk, lol, וכו')
ארכיטקטורה טכנית:
-------------------
- **Win32 API Integration**: שימוש ב-ctypes לשליטה ישירה במקלדת וב-clipboard
- **Unicode Input**: הזרקת תווים Unicode ישירות (ללא תלות בפריסה נוכחית)
- **Keyboard Hook**: ניטור כל הקשות מקלדת ברמת מערכת באמצעות ספריית keyboard
- **Threading**: עיבוד אסינכרוני של תיקונים למניעת lag
- **Process Detection**: זיהוי החלון הפעיל וה-exe שלו לקביעת התנהגות
מיפויי פריסה:
-------------
- מיפוי דו-כיווני מלא בין תווים עבריים ואנגליים
- תמיכה באותיות גדולות וקטנות
- מיפוי סימני פיסוק ותווים מיוחדים
תצורה:
-------
הסקריפט מאפשר התאמה אישית נרחבת דרך משתני תצורה בראש הקובץ:
- EXIT_HOTKEY: קיצור דרך ליציאה (ברירת מחדל: Ctrl+Alt+Q)
- USE_CLIPBOARD_CORRECTION_ON_SPACE: שימוש בתיקון מבוסס clipboard ב-WhatsApp
- APP_DEFAULT_LANG_BY_EXE: מיפוי שפות ברירת מחדל לפי אפליקציות
- AUTO_SWITCH_AFTER_CONSECUTIVE: מספר תיקונים רצופים להחלפת פריסה אוטומטית
דרישות מערכת:
--------------
- Windows (שימוש נרחב ב-Win32 API)
- Python 3.9+
- ספריית keyboard (pip install keyboard)
- הרשאות מנהל (לניטור keyboard ברמת מערכת)
שימוש:
-------
הפעל את הסקריפט עם הרשאות מנהל:
     python auto_lang2.py
לחץ Ctrl+Alt+Q ליציאה.
הערות חשובות:
--------------
- הסקריפט מיועד לשימוש אישי ולא למערכות ייצור
- מומלץ לבדוק את ההתנהגות בכל אפליקציה לפני הסתמכות מלאה
- חלק מהאפליקציות (במיוחד Electron-based) עשויות לדרוש התאמות ספציפיות
- שימוש ב-clipboard עשוי להתנגש עם אפליקציות אחרות שמשתמשות בו במקביל
"""
from ctypes import wintypes
import threading
import time

import keyboard

# NLP word-frequency validation (replaces heuristic word lists)
from wordfreq import zipf_frequency

# Multi-language keyboard maps
try:
    from keyboard_maps import (
        ALL_PROFILES, ENGLISH_LANG_IDS, LanguageProfile,
        lang_id_to_profile, is_english_lang_id, detect_script,
    )
except ImportError:
    # Fallback: will be available when bundled with PyInstaller
    pass


# ----------------------------
# Single-instance guard (Windows mutex)
# ----------------------------

_AUTO_LANG2_MUTEX_NAME = 'Global\\AutoLang2PyMutex'


def _ensure_single_instance() -> None:
    """Exit early if another instance is already running.

    Without this, users often end up launching multiple elevated instances, which leads to
    duplicated hook handling and confusing behavior.
    """
    try:
        k32 = ctypes.WinDLL('kernel32', use_last_error=True)
        k32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
        k32.CreateMutexW.restype = wintypes.HANDLE

        handle = k32.CreateMutexW(None, False, _AUTO_LANG2_MUTEX_NAME)
        if not handle:
            return
        # ERROR_ALREADY_EXISTS (183)
        if ctypes.get_last_error() == 183:
            print('auto_lang2.py: another instance is already running; exiting this one.')
            raise SystemExit(0)
    except SystemExit:
        raise
    except Exception:
        # If mutex fails for any reason, continue (best-effort).
        return


# ----------------------------
# Configuration
# ----------------------------

# קיצור דרך ליציאה (גם כשאין פוקוס על הטרמינל)
EXIT_HOTKEY = 'ctrl+alt+q'

# דיבאג מהיר: מדפיס למסך exe+title+lang_id של החלון שבפוקוס
INFO_HOTKEY = 'ctrl+alt+i'

# מצב תאימות ל-WhatsApp Desktop (Electron): מתקנים על בסיס Clipboard במקום לסמוך על תווים מה-hook
USE_CLIPBOARD_CORRECTION_ON_SPACE = True

# באילו אפליקציות נאפשר תיקון על רווח גם כש-USE_CLIPBOARD_CORRECTION_ON_SPACE פעיל.
# (בברירת מחדל אנחנו נמנעים מלגעת בשדות אחרים כדי לא "להדביק"/למחוק בטרמינל וכו')
SPACE_CORRECTION_ALLOWED_EXE = {
    'whatsapp.root.exe',
    'whatsapp.exe',
    'teams.exe',
    'ms-teams.exe',
    'code.exe',
    'pycharm64.exe',
    'winword.exe',
    'excel.exe',
    'powerpnt.exe',
    'outlook.exe',
    'notepad.exe',
    'notepad++.exe',
}

# For "safe" apps (VS Code/PyCharm/Teams, etc.) we avoid clipboard paste because the clipboard
# can be locked and we might accidentally paste stale content (e.g., logs). Instead, we do a
# key-up based correction on space and replace via backspace + keyboard.write.
SAFE_APP_NO_CLIPBOARD_REPLACE = True

# אם WhatsApp Desktop חוסם Ctrl+C, נשתמש ב-buffer של מקשים (לא תלוי ב-clipboard)
USE_KEYBUFFER_FALLBACK_FOR_WHATSAPP = True

# ב-WhatsApp Desktop לעיתים SendInput מתעלם; ננסה להזריק קיצורים דרך הספרייה keyboard
WHATSAPP_REPLACE_USE_KEYBOARD_SEND = True

# ברירת מחדל לווטסאפ: מחיקה דרך Ctrl+Backspace ואז Paste (עמיד יותר מ-Shift+Left ב-RTL)
WHATSAPP_REPLACE_USE_CTRL_BACKSPACE = True

# ב-WhatsApp: עבור קיצורים עם פיסוק (למשל "ac,") מחיקה לפי ספירת Backspace יכולה לפעמים לפספס תו אחד,
# ולהשאיר אות מובילה (למשל "aשבת"). עבור מקרים כאלה, selection+paste יציב יותר.
WHATSAPP_REPLACE_PUNCTUATION_BY_SELECTION = True

# ב-WhatsApp: סימני פיסוק (כמו ',') לעתים הם בעצם אותיות (למשל ',' -> 'ת').
# לכן נתייחס אליהם כחלק מהמילה ונבצע תיקון רק על רווח.
WHATSAPP_PUNCTUATION_AS_TEXT = True

DEBUG = True

# ─── NLP word-frequency thresholds ───
# Words with zipf_frequency >= this threshold are considered "valid" in a language.
# Tested: 3.0 perfectly separates real words from gibberish for both English and Hebrew.
NLP_ZIPF_THRESHOLD = 3.0

# מילים/קיצורים לא להמיר (UI מעדכן את הרשימה הזו)
EXCLUDE_WORDS: set[str] = set()

# UI toggle - allows disabling corrections without stopping the hook
ENGINE_ENABLED = True

# החלפת פריסה אוטומטית אחרי N תיקונים רצופים לאותה שפה (best-effort)
AUTO_SWITCH_LAYOUT = True
AUTO_SWITCH_AFTER_CONSECUTIVE = 2

# איך לנסות להחליף שפה במערכת (בסדר ניסיון). אפשר לשנות לפי ההגדרות שלך ב-Windows.
LAYOUT_TOGGLE_HOTKEYS = ('alt+shift', 'win+space', 'ctrl+shift')

# ----------------------------
# Default layout per focused app
# ----------------------------

ENABLE_APP_DEFAULT_LAYOUT = True
APP_DEFAULT_POLL_INTERVAL_SEC = 0.35

# לא להכריח החלפת שפה מיד אחרי תיקון מילה (כדי לא "להילחם" עם AUTO_SWITCH_LAYOUT)
APP_DEFAULT_COOLDOWN_AFTER_CORRECTION_SEC = 0.7

# ניסיונות חוזרים אם האפליקציה "בלעה" את ההחלפה (שכיח במיוחד ב-Teams/WebView)
APP_DEFAULT_RETRY_IF_MISMATCH = True
APP_DEFAULT_RETRY_INTERVAL_SEC = 1.0
APP_DEFAULT_RETRY_EXE = {
    'teams.exe',
    'ms-teams.exe',
}

# ברירת מחדל לפי exe (מומלץ). אפשר להוסיף/להסיר לפי מה שיש אצלך.
APP_DEFAULT_LANG_BY_EXE: dict[str, str] = {
    # Hebrew-first apps
    'whatsapp.root.exe': 'he',
    'whatsapp.exe': 'he',
    'telegram.exe': 'he',
    'winword.exe': 'he',
    'teams.exe': 'he',
    'ms-teams.exe': 'he',

    # English-first apps
    'code.exe': 'en',
    'pycharm64.exe': 'en',
    'windowsterminal.exe': 'en',
    'powershell.exe': 'en',
    'cmd.exe': 'en',
    'slack.exe': 'en',
    'excel.exe': 'en',
    'outlook.exe': 'en',
}

# ברירת מחדל לפי "צ'אט"/מסך בתוך אותה אפליקציה לפי כותרת חלון.
# עובד טוב בעיקר כשהכותרת כוללת את שם הצ'אט/איש קשר (למשל: "Maria - WhatsApp").
# ההשוואה היא substring (לא regex).
APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING: dict[str, dict[str, str]] = {
    'whatsapp.root.exe': {
        'יעל': 'he',
        'yael': 'he',
        'maria': 'en',
    },
    'whatsapp.exe': {
        'יעל': 'he',
        'yael': 'he',
        'maria': 'en',
    },
    'teams.exe': {
        'יעל': 'he',
        'yael': 'he',
    },
    'ms-teams.exe': {
        'יעל': 'he',
        'yael': 'he',
    },
}

# באילו אפליקציות נעקוב גם אחרי שינוי בכותרת (כי מעבר צ'אט לא משנה hwnd)
# נבנה אוטומטית מ-APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING + רשימה ידנית נוספת
_EXTRA_WATCH_TITLE_EXE: set[str] = {              # אפליקציות נוספות לעקוב - גם בלי chat defaults
    'chrome.exe',
    'msedge.exe',
    'outlook.exe',
    'olk.exe',
}
WATCH_TITLE_CHANGES_EXE: set[str] = set()          # ייבנה דינמית


def _rebuild_watch_title_set() -> None:
    """Rebuild WATCH_TITLE_CHANGES_EXE from chat-defaults keys + extras."""
    WATCH_TITLE_CHANGES_EXE.clear()
    WATCH_TITLE_CHANGES_EXE.update(APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING.keys())
    WATCH_TITLE_CHANGES_EXE.update(_EXTRA_WATCH_TITLE_EXE)


# Build initial set from defaults above
_rebuild_watch_title_set()

# ----------------------------
# Auto-learning language preference per chat
# ----------------------------

# Enable automatic language learning based on what you actually type in each chat.
ENABLE_AUTO_LEARN_CHAT_LANG = True

# Minimum characters typed before making a language decision (to avoid false positives).
AUTO_LEARN_MIN_CHARS = 50

# Threshold: if >60% Hebrew chars, prefer Hebrew; otherwise prefer English.
AUTO_LEARN_HEBREW_THRESHOLD = 0.60

# File to persist learned preferences between script runs.
AUTO_LEARN_CACHE_FILE = os.path.expanduser('~/.auto_lang2_learned.json')

# In-memory cache: {(exe, clean_title): {'he': count, 'en': count, 'learned': 'he'|'en'|None}}
_chat_lang_stats: dict[tuple[str, str], dict] = {}

# fallback לפי כותרת חלון (כמו בקובץ auto_lang.py). שימושי כששם exe לא יציב.
ENGLISH_APPS_TITLE = [
    'Visual Studio Code', 'Terminal', 'PyCharm', 'Command Prompt', 'Slack'
]
HEBREW_APPS_TITLE = [
    'WhatsApp', 'Telegram', 'Word', 'Team'
]

# גבולות מילה: רווח וסימני פיסוק (הסקריפט יחליף את המילה ואז ידפיס גם את הסימן)
WORD_BOUNDARIES = {'space', ',', '.', '!', '?', ';', ':'}

# מיפוי שמות מקשים ל"תו גבול" בפועל (כי keyboard לפעמים מחזיר 'comma' ולא ',')
BOUNDARY_KEY_TO_TEXT = {
    'space': ' ',
    ',': ',',
    'comma': ',',
    '.': '.',
    'dot': '.',
    ';': ';',
    'semicolon': ';',
    ':': ':',
    '!': '!',
    'exclamation': '!',
    '?': '?',
    'question': '?',
}

# ─── Multi-Language Dynamic Maps ─────────────────────────────────
# Built from keyboard_maps.py profiles based on installed keyboard layouts.
# These are populated at startup by _detect_and_load_languages().

# Active language profiles (only those installed on this machine)
ACTIVE_PROFILES: dict[str, 'LanguageProfile'] = {}

# Legacy compat aliases (will point to Hebrew profile if installed)
HEBREW_TO_ENGLISH: dict[str, str] = {}
ENGLISH_TO_HEBREW: dict[str, str] = {}
COMMON_SHORT_HEBREW_WORDS: dict[str, str] = {}

# Per-profile maps: lang_code -> {native_char: en_char}
PROFILE_TO_ENGLISH: dict[str, dict[str, str]] = {}
# Per-profile reverse maps: lang_code -> {en_char: native_char}
PROFILE_FROM_ENGLISH: dict[str, dict[str, str]] = {}
# Per-profile short words: lang_code -> {en_typed: native_word}
PROFILE_SHORT_WORDS: dict[str, dict[str, str]] = {}



def _detect_installed_keyboard_layouts() -> list[int]:
    """Query Windows for all installed keyboard layout language IDs."""
    try:
        count = user32.GetKeyboardLayoutList(0, None)
        if count <= 0:
            return []
        arr = (wintypes.HKL * count)()
        user32.GetKeyboardLayoutList(count, arr)
        return [int(h) & 0xFFFF for h in arr]
    except Exception:
        return []


def _detect_and_load_languages() -> None:
    """Auto-detect installed keyboard layouts and load matching profiles."""
    global ACTIVE_PROFILES, PROFILE_TO_ENGLISH, PROFILE_FROM_ENGLISH, PROFILE_SHORT_WORDS
    global HEBREW_TO_ENGLISH, ENGLISH_TO_HEBREW, COMMON_SHORT_HEBREW_WORDS

    installed_ids = _detect_installed_keyboard_layouts()
    if DEBUG:
        print(f'Installed keyboard layout IDs: {[f"0x{lid:04x}" for lid in installed_ids]}')

    loaded = []
    for lid in installed_ids:
        if is_english_lang_id(lid):
            continue  # English is the base, not a "profile"
        profile = lang_id_to_profile(lid)
        if profile and profile.code not in ACTIVE_PROFILES:
            ACTIVE_PROFILES[profile.code] = profile
            PROFILE_TO_ENGLISH[profile.code] = profile.to_english
            PROFILE_FROM_ENGLISH[profile.code] = profile.from_english
            PROFILE_SHORT_WORDS[profile.code] = profile.common_short_words
            loaded.append(f'{profile.flag} {profile.name} ({profile.code})')

    # Legacy Hebrew compat
    if 'he' in ACTIVE_PROFILES:
        he = ACTIVE_PROFILES['he']
        HEBREW_TO_ENGLISH.update(he.to_english)
        ENGLISH_TO_HEBREW.update(he.from_english)
        COMMON_SHORT_HEBREW_WORDS.update(he.common_short_words)

    if loaded:
        print(f'Loaded {len(loaded)} language profiles: {", ".join(loaded)}')
    else:
        # Fallback: load Hebrew if nothing detected
        if 'he' in ALL_PROFILES:
            he = ALL_PROFILES['he']
            ACTIVE_PROFILES['he'] = he
            PROFILE_TO_ENGLISH['he'] = he.to_english
            PROFILE_FROM_ENGLISH['he'] = he.from_english
            PROFILE_SHORT_WORDS['he'] = he.common_short_words
            HEBREW_TO_ENGLISH.update(he.to_english)
            ENGLISH_TO_HEBREW.update(he.from_english)
            COMMON_SHORT_HEBREW_WORDS.update(he.common_short_words)
            print('No layouts detected; loaded Hebrew as fallback')


# ----------------------------
# Win32 Unicode typing (layout independent)
# ----------------------------

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

psapi = ctypes.windll.psapi

# Some Python builds don't expose wintypes.ULONG_PTR; define our own.
ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong

# Fix ctypes default signatures (avoid pointer truncation on 64-bit)
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
user32.GetKeyboardLayout.restype = wintypes.HKL
user32.MapVirtualKeyExW.argtypes = [wintypes.UINT, wintypes.UINT, wintypes.HKL]
user32.MapVirtualKeyExW.restype = wintypes.UINT
user32.GetKeyboardState.argtypes = [ctypes.POINTER(ctypes.c_ubyte)]
user32.GetKeyboardState.restype = wintypes.BOOL
user32.ToUnicodeEx.argtypes = [
    wintypes.UINT,  # wVirtKey
    wintypes.UINT,  # wScanCode
    ctypes.POINTER(ctypes.c_ubyte),  # lpKeyState
    wintypes.LPWSTR,  # pwszBuff
    ctypes.c_int,  # cchBuff
    wintypes.UINT,  # wFlags
    wintypes.HKL,  # dwhkl
]
user32.ToUnicodeEx.restype = ctypes.c_int
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = wintypes.INT
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, wintypes.INT]
user32.GetWindowTextW.restype = wintypes.INT
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL
user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
    wintypes.UINT,
    wintypes.UINT,
    ctypes.POINTER(ULONG_PTR),
]
user32.SendMessageTimeoutW.restype = wintypes.LPARAM
user32.LoadKeyboardLayoutW.argtypes = [wintypes.LPCWSTR, wintypes.UINT]
user32.LoadKeyboardLayoutW.restype = wintypes.HKL
user32.ActivateKeyboardLayout.argtypes = [wintypes.HKL, wintypes.UINT]
user32.ActivateKeyboardLayout.restype = wintypes.HKL
user32.GetKeyboardLayoutList.argtypes = [ctypes.c_int, ctypes.POINTER(wintypes.HKL)]
user32.GetKeyboardLayoutList.restype = ctypes.c_int
user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, wintypes.LPVOID]
user32.GetGUIThreadInfo.restype = wintypes.BOOL
user32.EnumChildWindows.argtypes = [wintypes.HWND, ctypes.c_void_p, wintypes.LPARAM]
user32.EnumChildWindows.restype = wintypes.BOOL
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wintypes.BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wintypes.BOOL
user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
user32.GetClipboardData.argtypes = [wintypes.UINT]
user32.GetClipboardData.restype = wintypes.HANDLE
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE

kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = wintypes.LPVOID
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL
kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalFree.restype = wintypes.HGLOBAL
kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalSize.restype = ctypes.c_size_t

kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL

kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL

psapi.GetModuleBaseNameW.argtypes = [wintypes.HANDLE, wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
psapi.GetModuleBaseNameW.restype = wintypes.DWORD

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004

VK_BACK = 0x08
VK_RETURN = 0x0D
VK_SPACE = 0x20
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_LEFT = 0x25
VK_RIGHT = 0x27
VK_C = 0x43
VK_V = 0x56
VK_MENU = 0x12
VK_LWIN = 0x5B

MAPVK_VSC_TO_VK_EX = 3

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

WM_INPUTLANGCHANGEREQUEST = 0x0050
KLF_ACTIVATE = 0x00000001
KLF_SETFORPROCESS = 0x00000100

SMTO_ABORTIFHUNG = 0x0002


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk', ctypes.c_ushort),
        ('wScan', ctypes.c_ushort),
        ('dwFlags', ctypes.c_ulong),
        ('time', ctypes.c_ulong),
        ('dwExtraInfo', ULONG_PTR),
    ]


class INPUT_I(ctypes.Union):
    _fields_ = [('ki', KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [('type', ctypes.c_ulong), ('ii', INPUT_I)]


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('flags', wintypes.DWORD),
        ('hwndActive', wintypes.HWND),
        ('hwndFocus', wintypes.HWND),
        ('hwndCapture', wintypes.HWND),
        ('hwndMenuOwner', wintypes.HWND),
        ('hwndMoveSize', wintypes.HWND),
        ('hwndCaret', wintypes.HWND),
        ('rcCaret', wintypes.RECT),
    ]


# Now that GUITHREADINFO is defined, tighten the signature.
user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
user32.GetGUIThreadInfo.restype = wintypes.BOOL


def _get_foreground_focus_hwnd() -> wintypes.HWND:
    """Best-effort: return focused control HWND for the foreground thread (helps Teams)."""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return hwnd
    pid = wintypes.DWORD(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not tid:
        return hwnd
    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    if user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
        if info.hwndFocus:
            return info.hwndFocus
    return hwnd


def _send_vk_down(vk: int) -> None:
    down = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))


def _send_vk_up(vk: int) -> None:
    up = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def _send_vk_combo(*vks: int, inter_delay: float = 0.0) -> None:
    """Press all keys down in order, then release in reverse."""
    for vk in vks:
        _send_vk_down(vk)
        if inter_delay:
            time.sleep(inter_delay)
    for vk in reversed(vks):
        _send_vk_up(vk)
        if inter_delay:
            time.sleep(inter_delay)


def _send_vk(vk: int) -> None:
    down = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0)))
    up = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def _send_unicode_char(ch: str) -> None:
    code = ord(ch)
    down = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE, time=0, dwExtraInfo=0)))
    up = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def send_unicode_text(text: str, interval: float = 0.0) -> None:
    for ch in text:
        _send_unicode_char(ch)
        if interval:
            time.sleep(interval)


# ----------------------------
# Clipboard helpers (Unicode)
# ----------------------------

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002


def _get_clipboard_text() -> str | None:
    if not user32.OpenClipboard(None):
        return None
    try:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ''

        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ''

        size_bytes = kernel32.GlobalSize(handle)
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ''
        try:
            # size_bytes includes the null terminator; cap read to avoid AV.
            max_wchars = max(0, (size_bytes // ctypes.sizeof(ctypes.c_wchar)) - 1)
            return ctypes.wstring_at(ptr, max_wchars)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def _set_clipboard_text(text: str | None) -> bool:
    if text is None:
        text = ''
    data = text + '\x00'
    size_bytes = len(data) * ctypes.sizeof(ctypes.c_wchar)

    if not user32.OpenClipboard(None):
        return False
    hglob = None
    try:
        user32.EmptyClipboard()
        hglob = kernel32.GlobalAlloc(GMEM_MOVEABLE, size_bytes)
        if not hglob:
            return False
        ptr = kernel32.GlobalLock(hglob)
        if not ptr:
            return False
        try:
            ctypes.memmove(ptr, ctypes.create_unicode_buffer(data), size_bytes)
        finally:
            kernel32.GlobalUnlock(hglob)

        # If SetClipboardData succeeds, the system owns hglob.
        if not user32.SetClipboardData(CF_UNICODETEXT, hglob):
            kernel32.GlobalFree(hglob)
            hglob = None
            return False
        hglob = None
        return True
    finally:
        if hglob:
            kernel32.GlobalFree(hglob)
        user32.CloseClipboard()


def _get_foreground_hkl():
    # Prefer the focused control's thread (important for Teams/WebView), fall back to foreground.
    hwnd = _get_foreground_focus_hwnd() or user32.GetForegroundWindow()
    if not hwnd:
        return user32.GetKeyboardLayout(0)
    pid = wintypes.DWORD(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    hkl = user32.GetKeyboardLayout(tid)
    if not hkl:
        hkl = user32.GetKeyboardLayout(0)
    return hkl


def _pid_to_exe(pid_value: int) -> str:
    """Return lowercase exe basename for a PID, or '' if unknown."""
    hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid_value)
    if not hproc:
        return ''
    try:
        buf_len = wintypes.DWORD(260)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if kernel32.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(buf_len)):
            return buf.value.rsplit('\\', 1)[-1].lower()
        name_buf = ctypes.create_unicode_buffer(260)
        if psapi.GetModuleBaseNameW(hproc, None, name_buf, 260):
            return name_buf.value.lower()
        return ''
    finally:
        kernel32.CloseHandle(hproc)


# Callback type for EnumChildWindows
_ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _get_uwp_real_exe(parent_hwnd) -> str:
    """
    For UWP apps (ApplicationFrameHost.exe), the real app runs in a child window
    with a different PID. Enumerate children to find it.
    """
    parent_pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(parent_hwnd, ctypes.byref(parent_pid))
    result = ['']

    def _enum_cb(child_hwnd, _lparam):
        child_pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(child_hwnd, ctypes.byref(child_pid))
        if child_pid.value and child_pid.value != parent_pid.value:
            exe = _pid_to_exe(child_pid.value)
            if exe and exe != 'applicationframehost.exe':
                result[0] = exe
                return False  # stop enumeration
        return True  # continue

    cb = _ENUM_CHILD_PROC(_enum_cb)
    try:
        user32.EnumChildWindows(parent_hwnd, cb, 0)
    except Exception:
        pass
    return result[0]


def _get_foreground_exe_name() -> str:
    """Return lowercase exe basename of the foreground window process, or '' if unknown."""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ''

    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ''

    exe = _pid_to_exe(pid.value)

    # UWP apps: the foreground window belongs to ApplicationFrameHost.exe,
    # but the real app (WhatsApp, Calculator, etc.) is a child with a different PID.
    if exe == 'applicationframehost.exe':
        real = _get_uwp_real_exe(hwnd)
        if real:
            return real

    return exe


def _get_foreground_window_title() -> str:
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ''
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ''
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _strip_invisible_controls(text: str) -> str:
    if not text:
        return ''
    # Remove common bidi / zero-width / BOM characters that can break substring matching.
    bad = {
        '\u200e', '\u200f', '\u202a', '\u202b', '\u202c', '\u202d', '\u202e', '\ufeff',
        '\u200b', '\u200c', '\u200d',
    }
    return ''.join(ch for ch in text if ch not in bad)


def _klid_from_lang_id(lang_id: int) -> str:
    # KLID format: 8 hex digits, e.g. 00000409 / 0000040D
    return f'{lang_id:08x}'.upper()


def _set_foreground_input_language(lang_id: int) -> bool:
    """Explicitly request input language for the current foreground window (best-effort)."""
    fg_hwnd = user32.GetForegroundWindow()
    hwnd = _get_foreground_focus_hwnd() or fg_hwnd
    if not hwnd:
        return False

    klid = _klid_from_lang_id(lang_id)
    try:
        # KLF_SETFORPROCESS helps WebView/Electron apps that keep their own thread defaults.
        hkl = user32.LoadKeyboardLayoutW(klid, KLF_ACTIVATE | KLF_SETFORPROCESS)
    except Exception:
        return False

    if not hkl:
        return False

    # Activate for current thread/process (best-effort) and ask the focused control to change.
    try:
        user32.ActivateKeyboardLayout(hkl, KLF_SETFORPROCESS)
    except Exception:
        pass

    def _send_change(to_hwnd: wintypes.HWND) -> bool:
        if not to_hwnd:
            return False
        # Prefer SendMessageTimeout (more reliable than PostMessage in some apps)
        try:
            result = wintypes.ULONG_PTR(0)
            rc = user32.SendMessageTimeoutW(
                to_hwnd,
                WM_INPUTLANGCHANGEREQUEST,
                0,
                hkl,
                SMTO_ABORTIFHUNG,
                80,
                ctypes.byref(result),
            )
            return bool(rc)
        except Exception:
            try:
                return bool(user32.PostMessageW(to_hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl))
            except Exception:
                return False

    # Try multiple candidate hwnds from GUI thread info (Teams often needs hwndFocus/caret).
    sent_any = False
    try:
        if fg_hwnd:
            pid = wintypes.DWORD(0)
            tid = user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(pid))
        else:
            tid = 0

        if tid:
            info = GUITHREADINFO()
            info.cbSize = ctypes.sizeof(GUITHREADINFO)
            if user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
                for candidate in (info.hwndFocus, info.hwndCaret, info.hwndActive, hwnd, fg_hwnd):
                    if candidate and _send_change(candidate):
                        sent_any = True
        else:
            for candidate in (hwnd, fg_hwnd):
                if candidate and _send_change(candidate):
                    sent_any = True
    except Exception:
        sent_any = False

    return sent_any


def _pick_default_lang_for_foreground(exe_name: str, title: str) -> str | None:
    if not exe_name and not title:
        return None

    exe = (exe_name or '').lower()
    raw_title = title or ''
    clean_title = _strip_invisible_controls(raw_title)
    title_l = clean_title.lower()

    # 0) Auto-learned preference (highest priority)
    if ENABLE_AUTO_LEARN_CHAT_LANG:
        key = (exe, clean_title)
        stats = _chat_lang_stats.get(key)
        if stats and stats.get('learned'):
            return stats['learned']

    # 1) Chat/screen specific rules (by window title)
    per_title = APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING.get(exe)
    if per_title:
        # check both lowercase and original title to support Hebrew substrings
        for needle, lang in per_title.items():
            if not needle:
                continue
            if (needle.lower() in title_l) or (needle in clean_title):
                if lang == 'en' or lang in ACTIVE_PROFILES:
                    return lang

    direct = APP_DEFAULT_LANG_BY_EXE.get(exe)
    if direct == 'en' or direct in ACTIVE_PROFILES:
        return direct

    if any(tok.lower() in title_l for tok in ENGLISH_APPS_TITLE):
        return 'en'
    if any(tok.lower() in title_l for tok in HEBREW_APPS_TITLE):
        return 'he'

    return None


def _layout_matches_target(lang_id: int, target: str) -> bool:
    if target == 'en':
        return is_english_lang_id(lang_id)
    # For any non-English target, check if it matches the specific profile
    profile = ACTIVE_PROFILES.get(target)
    if profile:
        return lang_id in profile.win_lang_ids
    # Legacy fallback
    if target == 'he':
        return _is_hebrew_lang(lang_id)
    return False


def _apply_app_default_layout_if_needed() -> bool:
    if not ENABLE_APP_DEFAULT_LAYOUT:
        return False

    # Don't apply app default if we're paused after auto-switching layout.
    # Let the user continue in the language they're typing until sentence boundary.
    if _auto_switched_wait_for_boundary:
        return False

    # Don't apply app default if user actively chose a language via auto-switch.
    # Only revert on focus change.
    if _user_chose_language:
        return False

    # Avoid fighting right after a correction/layout auto-switch.
    if (time.time() - last_correction_time) < APP_DEFAULT_COOLDOWN_AFTER_CORRECTION_SEC:
        return False

    exe = _get_foreground_exe_name()
    title = _get_foreground_window_title()
    target = _pick_default_lang_for_foreground(exe, title)
    if not target:
        return False

    lang_id = _foreground_lang_id()
    if _layout_matches_target(lang_id, target):
        return True

    if target == 'en':
        desired = LANG_ENGLISH_US
    elif target in ACTIVE_PROFILES:
        desired = next(iter(ACTIVE_PROFILES[target].win_lang_ids))
    else:
        desired = LANG_ENGLISH_US

    if DEBUG:
        print(f'App default layout: exe={exe!r} title={title!r} -> {target!r}')

    # First attempt: WM_INPUTLANGCHANGEREQUEST — a window message, not a keyboard
    # event, so we do NOT need the 'injecting' guard here.  Keeping the guard off
    # lets the keyboard hook keep buffering real user keystrokes instead of
    # swallowing them for 1+ second.
    ok = _set_foreground_input_language(desired)
    time.sleep(0.06)

    # Verify it actually changed; some apps ignore WM_INPUTLANGCHANGEREQUEST.
    lang_after = _foreground_lang_id()
    good = _layout_matches_target(lang_after, target)

    if (not good):
        if DEBUG:
            print(
                f'App default layout: request_ok={ok} but lang_id stayed 0x{lang_after:04x}; '
                'falling back to toggle hotkeys'
            )
        # Fall back to OS hotkeys.  We set 'injecting' ONLY around each
        # synthetic key send (~20 ms) so that the hook ignores the fake
        # Alt+Shift / Win+Space events, but still buffers real user
        # keystrokes during the sleep gaps between retries.
        is_teams = exe in APP_DEFAULT_RETRY_EXE
        max_attempts = 14 if (is_teams and target == 'he') else 8
        sleep_after = 0.08 if is_teams else 0.05
        for _ in range(max_attempts):
            lang_now = _foreground_lang_id()
            if _layout_matches_target(lang_now, target):
                break
            injecting.set()
            try:
                _toggle_layout_once()
            finally:
                injecting.clear()
            time.sleep(sleep_after)

    return _layout_matches_target(_foreground_lang_id(), target)


def _app_default_layout_watcher() -> None:
    global _auto_switched_wait_for_boundary, en_streak, he_streak, _user_chose_language
    global words_in_sentence, sentence_lang, _pending_words
    global buffer_chars, buffer_keys, buffer_was_hebrew
    last_hwnd = None
    last_title = ''
    while not stop_event.is_set():
        try:
            if injecting.is_set():
                time.sleep(APP_DEFAULT_POLL_INTERVAL_SEC)
                continue

            hwnd = user32.GetForegroundWindow()
            if not hwnd:
                time.sleep(APP_DEFAULT_POLL_INTERVAL_SEC)
                continue

            if hwnd != last_hwnd:
                # Focus changed to different window - reset everything and apply app default
                last_hwnd = hwnd
                last_title = _get_foreground_window_title()
                _auto_switched_wait_for_boundary = False
                _user_chose_language = False
                en_streak = 0
                he_streak = 0
                words_in_sentence = 0
                sentence_lang = None
                _pending_words = []
                buffer_chars = ''
                buffer_keys = ''
                buffer_was_hebrew = False
                if DEBUG:
                    exe = _get_foreground_exe_name()
                    print(f'Focus changed - resetting to app default | exe={exe!r} title={last_title!r}')
                _apply_app_default_layout_if_needed()
            else:
                exe = _get_foreground_exe_name()
                if exe in WATCH_TITLE_CHANGES_EXE:
                    title = _get_foreground_window_title()
                    if title != last_title:
                        last_title = title
                        words_in_sentence = 0
                        sentence_lang = None
                        _pending_words = []
                        buffer_chars = ''
                        buffer_keys = ''
                        buffer_was_hebrew = False
                        _auto_switched_wait_for_boundary = False
                        if DEBUG:
                            print(f'Chat/title changed: exe={exe!r} title={last_title!r}')
                        _apply_app_default_layout_if_needed()

                # Retry-on-mismatch: some apps (Teams) may ignore a single attempt.
                if APP_DEFAULT_RETRY_IF_MISMATCH and exe in APP_DEFAULT_RETRY_EXE:
                    title = last_title or _get_foreground_window_title()
                    target = _pick_default_lang_for_foreground(exe, title)
                    if target:
                        if not _layout_matches_target(_foreground_lang_id(), target):
                            key = (int(hwnd), target)
                            now = time.time()
                            last = _last_app_default_attempt.get(key, 0.0)
                            if (now - last) >= APP_DEFAULT_RETRY_INTERVAL_SEC:
                                _last_app_default_attempt[key] = now
                                _apply_app_default_layout_if_needed()
        except Exception:
            pass

        time.sleep(APP_DEFAULT_POLL_INTERVAL_SEC)


def _event_to_char(event: keyboard.KeyboardEvent) -> str | None:
    """Best-effort: get the actual character that would be typed for this keydown."""
    try:
        scan = int(getattr(event, 'scan_code', 0) or 0)
    except Exception:
        return None
    if not scan:
        return None

    hkl = _get_foreground_hkl()
    vk = user32.MapVirtualKeyExW(scan, MAPVK_VSC_TO_VK_EX, hkl)
    if not vk:
        return None

    state = (ctypes.c_ubyte * 256)()
    if not user32.GetKeyboardState(state):
        return None

    buf = ctypes.create_unicode_buffer(8)
    rc = user32.ToUnicodeEx(vk, scan, state, buf, len(buf) - 1, 0, hkl)
    if rc > 0:
        ch = buf.value[:rc]
        if ch and not any(ord(c) < 32 for c in ch):
            return ch
        # Got control character — likely stale Ctrl modifier from Electron/WebView apps
        # (WhatsApp, Teams, etc.).  Clear Ctrl bits and retry.
        state[0x11] &= 0x7F   # VK_CONTROL
        state[0xA2] &= 0x7F   # VK_LCONTROL
        state[0xA3] &= 0x7F   # VK_RCONTROL
        rc = user32.ToUnicodeEx(vk, scan, state, buf, len(buf) - 1, 0, hkl)
        if rc > 0:
            ch = buf.value[:rc]
            if ch and not any(ord(c) < 32 for c in ch):
                return ch
    return None


# ----------------------------
# Core logic
# ----------------------------

buffer_chars = ''
buffer_keys = ''
buffer_was_hebrew = False  # Track if buffer was typed in Hebrew layout
words_in_sentence = 0  # Track word count in current sentence (reset on ENTER/period)
sentence_lang = None  # Track language determined by first word ('en', 'he', or None)

en_streak = 0
he_streak = 0

# Track characters that the user types while a replacement is in progress.
# These characters reach the app but on_key ignores them (injecting is set),
# so we stash them here so _replace_word can account for them.
_inject_overflow = ''

# אחרי החלפת שפה אוטומטית, נפסיק לבדוק תיקונים עד ENTER או נקודה
_auto_switched_wait_for_boundary = False
# האם המשתמש בחר שפה באופן אקטיבי (דרך auto-switch)? אם כן, לא לחזור לברירת מחדל עד שינוי פוקוס

# ─── NLP pending-word buffer ───
# When both language versions of a word are valid (AMBIG), we buffer the word
# and wait for a disambiguating word.  Once found, all buffered words are
# corrected backwards in a single batch.
# Each entry: {'original': str, 'en_version': str, 'native_version': str,
#              'native_lang': str, 'boundary': str, 'typed_is_native': bool}
_pending_words: list[dict] = []
_user_chose_language = False


def _load_learned_prefs() -> None:
    """Load previously learned chat language preferences from disk."""
    global _chat_lang_stats
    if not ENABLE_AUTO_LEARN_CHAT_LANG:
        return
    try:
        if os.path.exists(AUTO_LEARN_CACHE_FILE):
            with open(AUTO_LEARN_CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Convert JSON keys "exe|title" back to tuples
                for key_str, stats in data.items():
                    if '|' in key_str:
                        exe, title = key_str.split('|', 1)
                        _chat_lang_stats[(exe, title)] = stats
    except Exception as e:
        if DEBUG:
            print(f'Failed to load learned prefs: {e}')


def _save_learned_prefs() -> None:
    """Persist learned chat language preferences to disk."""
    if not ENABLE_AUTO_LEARN_CHAT_LANG:
        return
    try:
        # Convert tuple keys to "exe|title" strings for JSON
        data = {}
        for (exe, title), stats in _chat_lang_stats.items():
            if stats.get('learned'):
                data[f'{exe}|{title}'] = stats
        with open(AUTO_LEARN_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        if DEBUG:
            print(f'Failed to save learned prefs: {e}')


def _update_chat_lang_stats(exe: str, title: str, char: str) -> None:
    """Update language statistics for current chat based on typed character."""
    if not ENABLE_AUTO_LEARN_CHAT_LANG:
        return
    if not exe or not title:
        return
    
    clean_title = _strip_invisible_controls(title)
    key = (exe, clean_title)
    
    if key not in _chat_lang_stats:
        _chat_lang_stats[key] = {'he': 0, 'en': 0, 'learned': None}
    
    stats = _chat_lang_stats[key]
    
    # Count characters by language
    detected = contains_non_english(char)
    if detected:
        stats[detected] = stats.get(detected, 0) + 1
    elif char.isascii() and char.isalpha():
        stats['en'] = stats.get('en', 0) + 1
    else:
        return  # Ignore punctuation/digits for language stats
    
    total = sum(v for k, v in stats.items() if k not in ('learned',))
    if total < AUTO_LEARN_MIN_CHARS:
        return  # Not enough data yet
    
    # Decide learned preference: language with highest ratio
    lang_counts = {k: v for k, v in stats.items() if k not in ('learned',)}
    best_lang = max(lang_counts, key=lang_counts.get)
    best_ratio = lang_counts[best_lang] / total if total > 0 else 0
    old_learned = stats.get('learned')
    new_learned = best_lang if best_ratio >= AUTO_LEARN_HEBREW_THRESHOLD else 'en'
    
    if new_learned != old_learned:
        stats['learned'] = new_learned
        if DEBUG:
            print(
                f'Auto-learned: {exe}:{clean_title[:40]} -> {new_learned!r} '
                f'(best={best_lang} ratio={best_ratio:.2f})'
            )
        _save_learned_prefs()

_toggle_method_index = 0

injecting = threading.Event()  # מונע לופ: לא מאחסנים אירועים שאנחנו עצמנו שולחים
stop_event = threading.Event()

last_correction_time = 0.0

# Throttle retries per hwnd/target to avoid spamming.
_last_app_default_attempt: dict[tuple[int, str], float] = {}

correction_lock = threading.Lock()


def _sanitize_copied_word(text: str) -> str:
    if not text:
        return ''
    # מסירים תווי כיווניות/בקרה נפוצים כדי שהיורסיטיקות יעבדו
    bad = {
        '\u200e', '\u200f', '\u202a', '\u202b', '\u202c', '\u202d', '\u202e', '\ufeff'
    }
    return ''.join(ch for ch in text if ch not in bad).strip('\r\n')


LANG_ENGLISH_US = 0x0409
LANG_ENGLISH_UK = 0x0809
LANG_HEBREW = 0x040D


def _foreground_lang_id() -> int:
    try:
        hkl = _get_foreground_hkl()
        lang_id = int(hkl) & 0xFFFF
        if lang_id == 0:
            # Some windows/threads can report 0 transiently; fall back to current thread.
            lang_id = int(user32.GetKeyboardLayout(0)) & 0xFFFF
        return lang_id
    except Exception:
        return 0


def _foreground_profile() -> 'LanguageProfile | None':
    """Get the LanguageProfile for the current foreground keyboard layout, or None if English."""
    lid = _foreground_lang_id()
    return lang_id_to_profile(lid)


def _is_hebrew_lang(lang_id: int) -> bool:
    return lang_id == LANG_HEBREW


def _is_english_lang(lang_id: int) -> bool:
    return is_english_lang_id(lang_id)


def _is_non_english_lang(lang_id: int) -> bool:
    """Check if lang_id belongs to any active non-English profile."""
    for p in ACTIVE_PROFILES.values():
        if lang_id in p.win_lang_ids:
            return True
    return False


def _lang_code_for_id(lang_id: int) -> str | None:
    """Return the language code ('he','ru','ar',...) for a lang_id, or 'en', or None."""
    if is_english_lang_id(lang_id):
        return 'en'
    p = lang_id_to_profile(lang_id)
    return p.code if p else None


def _toggle_layout_once() -> None:
    """Best-effort: attempt one toggle method (rotating across calls).

    Rationale: some apps ignore a specific injection route; rotating avoids getting stuck
    retrying the same ineffective method.
    """
    global _toggle_method_index

    methods: list[str] = list(LAYOUT_TOGGLE_HOTKEYS) + [
        '__sendinput_alt_shift',
        '__sendinput_win_space',
        '__sendinput_ctrl_shift',
    ]
    if not methods:
        return

    pick = methods[_toggle_method_index % len(methods)]
    _toggle_method_index += 1

    try:
        if pick == '__sendinput_alt_shift':
            _send_vk_combo(VK_MENU, VK_SHIFT, inter_delay=0.02)
            return
        if pick == '__sendinput_win_space':
            _send_vk_combo(VK_LWIN, VK_SPACE, inter_delay=0.02)
            return
        if pick == '__sendinput_ctrl_shift':
            _send_vk_combo(VK_CONTROL, VK_SHIFT, inter_delay=0.02)
            return

        # keyboard library route
        keyboard.send(pick)
    except Exception:
        # ignore and let the caller retry
        return


def _switch_layout_to(target: str) -> bool:
    """Best-effort switch input language for the foreground window.
    target can be 'en', 'he', 'ru', 'ar', or any language code in ACTIVE_PROFILES.
    """
    # Validate target
    if target != 'en' and target not in ACTIVE_PROFILES:
        return False

    # Try explicit WM_INPUTLANGCHANGEREQUEST first
    if target == 'en':
        desired_id = LANG_ENGLISH_US
    elif target in ACTIVE_PROFILES:
        p = ACTIVE_PROFILES[target]
        desired_id = next(iter(p.win_lang_ids))
    else:
        return False

    injecting.set()
    try:
        # First try explicit language set (most reliable)
        _set_foreground_input_language(desired_id)
        time.sleep(0.05)
        lang_id = _foreground_lang_id()
        if _layout_matches_target(lang_id, target):
            return True

        # Fall back to toggling
        for _ in range(8):
            lang_id = _foreground_lang_id()
            if DEBUG:
                print(f'Layout check: lang_id=0x{lang_id:04x} target={target!r}')
            if _layout_matches_target(lang_id, target):
                return True
            _toggle_layout_once()
            time.sleep(0.05)
        return False
    finally:
        time.sleep(0.01)
        injecting.clear()


# Per-language correction streak counters: lang_code -> count
_correction_streaks: dict[str, int] = {}


def _note_correction_target(target: str | None) -> None:
    """Update consecutive correction counters and maybe auto-switch layout."""
    global en_streak, he_streak, last_correction_time, _auto_switched_wait_for_boundary, _user_chose_language

    # Update per-language streaks
    if target:
        for k in list(_correction_streaks.keys()):
            if k != target:
                _correction_streaks[k] = 0
        _correction_streaks[target] = _correction_streaks.get(target, 0) + 1

    # Legacy compat
    if target == 'en':
        en_streak += 1
        he_streak = 0
    elif target == 'he':
        he_streak += 1
        en_streak = 0
    else:
        en_streak = 0
        he_streak = 0

    if target:
        last_correction_time = time.time()

    if not AUTO_SWITCH_LAYOUT:
        return

    # Check if any language has enough consecutive corrections to trigger auto-switch
    for lang, count in _correction_streaks.items():
        if count >= AUTO_SWITCH_AFTER_CONSECUTIVE:
            if DEBUG:
                print(f'Auto-switching layout -> {lang}')
            if _switch_layout_to(lang):
                _correction_streaks[lang] = 0
                _auto_switched_wait_for_boundary = True
                _user_chose_language = True
                if DEBUG:
                    print('Pausing corrections until ENTER or period')
                break


def _replace_by_delete_len(delete_len: int, corrected: str, boundary_text: str) -> None:
    global buffer_chars, buffer_keys
    injecting.set()
    try:
        # תן לרווח להיכנס לפני מחיקה (Electron לפעמים מאחר)
        time.sleep(0.06)

        # מוחקים קודם את ה-boundary שכבר הודפס, ואז את המילה
        _send_vk(VK_BACK)
        time.sleep(0.005)

        for _ in range(delete_len):
            _send_vk(VK_BACK)
            time.sleep(0.003)

        send_unicode_text(corrected)
        if boundary_text == ' ':
            _send_vk(VK_SPACE)
        else:
            send_unicode_text(boundary_text)

        if DEBUG:
            print(f'Keybuffer replaced: delete_len={delete_len} corrected={corrected!r}')
    finally:
        buffer_chars = ''
        buffer_keys = ''
        time.sleep(0.01)
        injecting.clear()


def _replace_by_select_and_paste(select_chars: int, replacement: str) -> None:
    """Replace last characters by selecting them (Shift+Left) and pasting replacement via clipboard."""
    injecting.set()
    original_clip = None
    clipboard_restored = False
    try:
        original_clip = _get_clipboard_text()

        # Let WhatsApp commit the user's last key (space) before we start selecting.
        time.sleep(0.06)

        # Ensure we're still in WhatsApp right before injecting.
        fg = _get_foreground_exe_name()
        if not (fg and ('whatsapp' in fg)):
            if DEBUG:
                print(f'Abort replace: foreground changed to {fg!r}')
            return

        # Select N chars to the left
        for _ in range(select_chars):
            if WHATSAPP_REPLACE_USE_KEYBOARD_SEND:
                keyboard.send('shift+left')
                time.sleep(0.002)
            else:
                _send_vk_combo(VK_SHIFT, VK_LEFT, inter_delay=0.001)

        _set_clipboard_text(replacement)
        time.sleep(0.03)
        if WHATSAPP_REPLACE_USE_KEYBOARD_SEND:
            keyboard.send('ctrl+v')
        else:
            _send_vk_combo(VK_CONTROL, VK_V, inter_delay=0.001)
        time.sleep(0.18)

        if WHATSAPP_REPLACE_USE_KEYBOARD_SEND:
            keyboard.send('right')
        else:
            _send_vk(VK_RIGHT)

        if original_clip is not None:
            _set_clipboard_text(original_clip)
            clipboard_restored = True

        if DEBUG:
            print(f'Keybuffer pasted replacement: {replacement!r} (selected {select_chars})')
    except Exception as e:
        print(f'Keybuffer paste failed: {e!r}')
    finally:
        if (not clipboard_restored) and (original_clip is not None):
            try:
                _set_clipboard_text(original_clip)
            except Exception:
                pass
        time.sleep(0.01)
        injecting.clear()


def _replace_prev_word_by_ctrl_backspace_paste(replacement: str) -> bool:
    """WhatsApp/Electron-friendly replace: delete last space + previous word, then paste replacement."""
    injecting.set()
    original_clip = None
    clipboard_restored = False
    try:
        # Ensure we're still in WhatsApp right before injecting.
        fg = _get_foreground_exe_name()
        if not (fg and ('whatsapp' in fg)):
            if DEBUG:
                print(f'Abort replace: foreground changed to {fg!r}')
            return False

        original_clip = _get_clipboard_text()

        # Let WhatsApp commit the user's last key (space) before we start deleting.
        time.sleep(0.06)

        # Delete the just-typed space, then the previous word.
        # This avoids relying on selection behavior in RTL.
        keyboard.send('backspace')
        time.sleep(0.01)
        keyboard.send('ctrl+backspace')
        time.sleep(0.02)

        _set_clipboard_text(replacement)
        time.sleep(0.03)
        keyboard.send('ctrl+v')
        time.sleep(0.12)

        if original_clip is not None:
            _set_clipboard_text(original_clip)
            clipboard_restored = True

        if DEBUG:
            print(f'Keybuffer ctrl+backspace replaced with: {replacement!r}')
        return True
    except Exception as e:
        print(f'Keybuffer ctrl+backspace replace failed: {e!r}')
        return False
    finally:
        if (not clipboard_restored) and (original_clip is not None):
            try:
                _set_clipboard_text(original_clip)
            except Exception:
                pass
        time.sleep(0.01)
        injecting.clear()


def _replace_by_backspace_count_paste(backspaces: int, replacement: str) -> bool:
    """Delete exact number of chars with Backspace, then paste replacement."""
    if backspaces <= 0:
        return False

    injecting.set()
    original_clip = None
    clipboard_restored = False
    try:
        fg = _get_foreground_exe_name()
        if not (fg and ('whatsapp' in fg)):
            if DEBUG:
                print(f'Abort replace: foreground changed to {fg!r}')
            return False

        original_clip = _get_clipboard_text()
        time.sleep(0.08)

        for _ in range(backspaces):
            keyboard.send('backspace')
            time.sleep(0.01)

        _set_clipboard_text(replacement)
        time.sleep(0.03)
        keyboard.send('ctrl+v')
        time.sleep(0.12)

        if original_clip is not None:
            _set_clipboard_text(original_clip)
            clipboard_restored = True

        if DEBUG:
            print(f'Keybuffer backspace-count replaced with: {replacement!r} (backspaces={backspaces})')
        return True
    except Exception as e:
        print(f'Keybuffer backspace-count replace failed: {e!r}')
        return False
    finally:
        if (not clipboard_restored) and (original_clip is not None):
            try:
                _set_clipboard_text(original_clip)
            except Exception:
                pass
        time.sleep(0.01)
        injecting.clear()


def contains_hebrew(text: str) -> bool:
    """Legacy compat: check if text contains Hebrew characters."""
    return any('\u05D0' <= c <= '\u05EA' for c in text)


def contains_non_english(text: str) -> str | None:
    """Detect which active non-English language is present in text.
    Returns the language code (e.g. 'he', 'ru') or None.
    """
    for code, profile in ACTIVE_PROFILES.items():
        if profile.contains_script(text):
            return code
    return None


def translate(text: str, mapping: dict[str, str]) -> str:
    return ''.join(mapping.get(c, c) for c in text)


def english_plausible(candidate: str) -> bool:
    """DEPRECATED — kept for backward compat. Use _nlp_is_valid() instead."""
    return _nlp_is_valid(candidate, 'en')

# Legacy alias
english_plausible_for_hebrew_to_english = english_plausible


def native_plausible(word: str, lang_code: str) -> bool:
    """Check if word is plausible in the given language using its profile."""
    profile = ACTIVE_PROFILES.get(lang_code)
    if not profile:
        return False
    return profile.is_plausible(word)


def hebrew_plausible(word: str) -> bool:
    """Legacy compat: check if word looks like Hebrew."""
    return native_plausible(word, 'he')


# ─── NLP word-frequency validation ───────────────────────────────────────────

def _nlp_is_valid(word: str, lang: str) -> bool:
    """Check if *word* is a real word in *lang* using wordfreq + structural checks.

    Combines:
    1. Structural plausibility — first char must belong to the script.
    2. Frequency plausibility — zipf_frequency ≥ NLP_ZIPF_THRESHOLD (3.0).
    """
    if not word or len(word) < 2:
        return False

    w = word.strip().lower()
    if not w:
        return False

    # ── structural check ──
    if lang == 'en':
        # English: first char must be ASCII letter
        if not (w[0].isascii() and w[0].isalpha()):
            return False
    else:
        profile = ACTIVE_PROFILES.get(lang)
        if profile:
            cp = ord(w[0])
            if not any(s <= cp <= e for s, e in profile.unicode_ranges):
                return False

    # ── frequency check ──
    freq = zipf_frequency(w, lang)
    return freq >= NLP_ZIPF_THRESHOLD


def _nlp_decide_word(original: str, boundary_text: str) -> None:
    """Core NLP decision logic shared by decide_and_correct and the snapshot path.

    For every word that arrives:
    1. Generate both versions (English ↔ native).
    2. If it's a known short word → force-convert (bypass NLP).
    3. Run _nlp_is_valid on both versions.
    4. Decision matrix:
       en_valid & !he_valid  → winner = 'en'
       he_valid & !en_valid  → winner = native lang
       both valid            → AMBIG  → buffer word for later
       neither valid         → NEITHER → buffer word (don't lock sentence), correct later if disambiguated
    5. When winner found and there are pending words → batch-correct backwards.
    """
    global sentence_lang, _pending_words

    # ── Determine typed language & generate both versions ──
    typed_lang_code = contains_non_english(original)

    if typed_lang_code:
        # Text contains non-English script (e.g., user typed Hebrew chars)
        native_version = original
        native_lang = typed_lang_code
        to_en_map = PROFILE_TO_ENGLISH.get(typed_lang_code, {})
        en_version = translate(original, to_en_map)
        typed_is_native = True
    else:
        # Text is English/ASCII chars
        en_version = original
        typed_is_native = False
        # Pick first active profile for the native translation
        native_version = None
        native_lang = None
        for lang_code in ACTIVE_PROFILES:
            from_en_map = PROFILE_FROM_ENGLISH.get(lang_code, {})
            native_version = translate(original, from_en_map)
            native_lang = lang_code
            break
        if not native_version or not native_lang:
            return  # No active profiles

    lw = original.lower()

    # ── Short-words bypass (1-2 char words) ──
    short_words = PROFILE_SHORT_WORDS.get(native_lang, {})
    if len(lw) <= 2:
        if lw in short_words:
            if not typed_is_native:
                # User typed English shortcut → convert to native
                _do_single_correction(original, short_words[lw], boundary_text, native_lang)
            # else: already native, no correction
        # Short words not in dict → ignore
        return

    # Known short-word match for longer words (e.g., 'tbh' → 'אני' if in dict)
    if lw in short_words:
        if not typed_is_native:
            _do_single_correction(original, short_words[lw], boundary_text, native_lang)
        return

    # ── NLP validation ──
    en_valid = _nlp_is_valid(en_version, 'en')
    he_valid = _nlp_is_valid(native_version, native_lang)

    if DEBUG:
        print(
            f'nlp: original={original!r} en={en_version!r}(valid={en_valid}) '
            f'{native_lang}={native_version!r}(valid={he_valid}) '
            f'pending={len(_pending_words)} typed_native={typed_is_native}'
        )

    if en_valid and not he_valid:
        # ── Winner: English ──
        winner_lang = 'en'
        if typed_is_native:
            # User typed native chars but it's English → correct
            _flush_and_correct(winner_lang, original, en_version, boundary_text, typed_is_native)
        else:
            # Already English, no correction needed
            _flush_and_set_lang(winner_lang, boundary_text)

    elif he_valid and not en_valid:
        # ── Winner: native language ──
        winner_lang = native_lang
        if not typed_is_native:
            # User typed English but it's native → correct
            _flush_and_correct(winner_lang, original, native_version, boundary_text, typed_is_native)
        else:
            # Already native, no correction needed
            _flush_and_set_lang(winner_lang, boundary_text)

    elif en_valid and he_valid:
        # ── AMBIGUOUS — both versions are valid words ──
        _pending_words.append({
            'original': original,
            'en_version': en_version,
            'native_version': native_version,
            'native_lang': native_lang,
            'boundary': boundary_text,
            'typed_is_native': typed_is_native,
        })
        if DEBUG:
            print(f'nlp: AMBIG — buffered word #{len(_pending_words)}: {original!r}')

    else:
        # ── NEITHER valid — buffer it like AMBIG ──
        # The word might be a typo in the wrong layout.  If a later word
        # disambiguates the language, we'll batch-correct this one too
        # (the translated version, even if it's not a dictionary word).
        _pending_words.append({
            'original': original,
            'en_version': en_version,
            'native_version': native_version,
            'native_lang': native_lang,
            'boundary': boundary_text,
            'typed_is_native': typed_is_native,
        })
        if DEBUG:
            print(f'nlp: NEITHER valid — buffered word #{len(_pending_words)}: {original!r} '
                  f'(sentence_lang stays {sentence_lang})')


def _flush_and_set_lang(winner_lang: str, boundary_text: str) -> None:
    """Winner found and current word is already correct — flush pending words."""
    global sentence_lang, _pending_words

    sentence_lang = winner_lang

    if not _pending_words:
        return  # No pending words to correct

    # Check if any pending word actually needs correction
    corrections_needed = []
    for pw in _pending_words:
        if winner_lang == 'en' and pw['typed_is_native']:
            corrections_needed.append(pw)
        elif winner_lang != 'en' and not pw['typed_is_native']:
            corrections_needed.append(pw)

    if not corrections_needed:
        _pending_words = []
        return

    # There ARE pending words that need correction but the current word is fine.
    # We need to batch-correct the pending words AND retype the current word
    # (since we must backspace through it to reach the pending words).
    # This is handled in _replace_word_batch.
    # NOTE: The current word just committed its boundary char, so we include it.
    threading.Thread(
        target=_replace_word_batch,
        args=(winner_lang, _pending_words.copy(), None, None, boundary_text),
        daemon=True,
    ).start()
    _pending_words = []


def _flush_and_correct(winner_lang: str, current_original: str, current_corrected: str,
                       boundary_text: str, current_typed_is_native: bool) -> None:
    """Winner found and current word needs correction — flush pending + correct current."""
    global sentence_lang, _pending_words

    sentence_lang = winner_lang

    if not _pending_words:
        # No pending words — just correct the current word (common fast path)
        _do_single_correction(current_original, current_corrected, boundary_text, winner_lang)
        _pending_words = []
        return

    # Batch correct: pending words + current word
    threading.Thread(
        target=_replace_word_batch,
        args=(winner_lang, _pending_words.copy(), current_original, current_corrected, boundary_text),
        daemon=True,
    ).start()
    _pending_words = []


def _do_single_correction(original: str, corrected: str, boundary_text: str, target_lang: str) -> None:
    """Correct a single word (no pending buffer). Wrapper around _replace_word."""
    if DEBUG:
        print(f'nlp: CORRECT {original!r} → {corrected!r} (lang={target_lang})')
    threading.Thread(
        target=_replace_word,
        args=(original, corrected, boundary_text),
        daemon=True,
    ).start()


def _get_current_app_default() -> str | None:
    """Get the default language for the current foreground app/chat."""
    exe = _get_foreground_exe_name()
    title = _get_foreground_window_title()
    return _pick_default_lang_for_foreground(exe, title)


def decide_and_correct(boundary_text: str) -> None:
    """NLP-based decision: translate both ways, check both with wordfreq, correct if needed."""
    global buffer_chars, buffer_was_hebrew, words_in_sentence, sentence_lang
    if not buffer_chars:
        return

    if not ENGINE_ENABLED:
        buffer_chars = ''
        buffer_was_hebrew = False
        return

    original = buffer_chars

    # Skip excluded words
    if EXCLUDE_WORDS and original.lower() in EXCLUDE_WORDS:
        if DEBUG:
            print(f'decide: skipping excluded word: {original!r}')
        buffer_chars = ''
        buffer_was_hebrew = False
        return

    words_in_sentence += 1

    # If sentence language already determined, auto-correct subsequent words
    # that are typed in the wrong layout (no NLP needed — we know the language).
    if sentence_lang is not None:
        typed_lang_code = contains_non_english(original)
        typed_is_native = bool(typed_lang_code)
        wrong_layout = False

        if sentence_lang == 'en' and typed_is_native:
            # Sentence is English but user typed native chars → correct to English
            to_en_map = PROFILE_TO_ENGLISH.get(typed_lang_code, {})
            corrected = translate(original, to_en_map)
            wrong_layout = True
        elif sentence_lang != 'en' and not typed_is_native:
            # Sentence is native but user typed English chars → correct to native
            from_en_map = PROFILE_FROM_ENGLISH.get(sentence_lang, {})
            corrected = translate(original, from_en_map)
            wrong_layout = True

        if wrong_layout:
            if DEBUG:
                print(f'decide: sentence_lang={sentence_lang} → auto-correct {original!r} → {corrected!r}')
            _do_single_correction(original, corrected, boundary_text, sentence_lang)
        else:
            if DEBUG:
                print(f'decide: sentence_lang={sentence_lang} already set, {original!r} matches layout')
        buffer_chars = ''
        buffer_was_hebrew = False
        return

    # Delegate to the unified NLP decision tree
    _nlp_decide_word(original, boundary_text)

    # Always clear buffers
    buffer_chars = ''
    buffer_was_hebrew = False


def _decide_and_correct_for_word(original: str, boundary_text: str) -> None:
    """Decide and replace for an explicit word snapshot (used for delayed space handling).

    Uses the same unified NLP decision tree as decide_and_correct.
    Runs synchronously (called from a thread already).
    """
    global words_in_sentence, sentence_lang
    if not original:
        return

    if not ENGINE_ENABLED:
        return

    # Skip excluded words
    if EXCLUDE_WORDS and original.lower() in EXCLUDE_WORDS:
        if DEBUG:
            print(f'decide(snap): skipping excluded word: {original!r}')
        return

    words_in_sentence += 1

    # If sentence language already determined, auto-correct wrong-layout words
    if sentence_lang is not None:
        typed_lang_code = contains_non_english(original)
        typed_is_native = bool(typed_lang_code)
        wrong_layout = False
        corrected = original

        if sentence_lang == 'en' and typed_is_native:
            to_en_map = PROFILE_TO_ENGLISH.get(typed_lang_code, {})
            corrected = translate(original, to_en_map)
            wrong_layout = True
        elif sentence_lang != 'en' and not typed_is_native:
            from_en_map = PROFILE_FROM_ENGLISH.get(sentence_lang, {})
            corrected = translate(original, from_en_map)
            wrong_layout = True

        if wrong_layout:
            if DEBUG:
                print(f'decide(snap): sentence_lang={sentence_lang} → auto-correct {original!r} → {corrected!r}')
            _do_single_correction(original, corrected, boundary_text, sentence_lang)
        else:
            if DEBUG:
                print(f'decide(snap): sentence_lang={sentence_lang} already set, {original!r} matches layout')
        return

    # Use the same NLP decision tree
    # NOTE: _nlp_decide_word spawns threads for _replace_word / _replace_word_batch,
    #       but since _decide_and_correct_for_word is already on a thread, that's fine.
    _nlp_decide_word(original, boundary_text)


def _replace_word_batch(winner_lang: str, pending: list[dict],
                        current_original: str | None, current_corrected: str | None,
                        current_boundary: str) -> None:
    """Batch-replace pending ambiguous words (and optionally the current word).

    Backspaces through ALL the text from the first pending word to the cursor,
    then retypes everything with corrected versions.

    *current_original* / *current_corrected* may be None when the current word
    itself is already correct but pending words need correction.  In that case
    we still need to backspace through the current word's boundary to reach the
    pending words, then retype it verbatim.
    """
    global buffer_chars, _inject_overflow
    injecting.set()
    try:
        time.sleep(0.01)

        # Build the ordered list: [pending_0, pending_1, ..., current_word]
        all_words: list[dict] = []
        for pw in pending:
            if winner_lang == 'en':
                corrected = pw['en_version'] if pw['typed_is_native'] else pw['original']
            else:
                corrected = pw['native_version'] if not pw['typed_is_native'] else pw['original']
            all_words.append({
                'original': pw['original'],
                'corrected': corrected,
                'boundary': pw['boundary'],
            })

        # Append the current word (if any) — it may or may not need correction
        if current_original is not None:
            all_words.append({
                'original': current_original,
                'corrected': current_corrected or current_original,
                'boundary': current_boundary,
            })

        # Check if any word actually needs correction
        needs_correction = any(w['original'] != w['corrected'] for w in all_words)
        if not needs_correction:
            if DEBUG:
                print('batch: no corrections needed in pending words')
            return

        # Account for extra chars typed during the delay
        extra_pre = len(buffer_chars)
        extra_inj = len(_inject_overflow)
        total_extra = extra_pre + extra_inj

        # Total chars to delete: all words + their boundaries + extra
        total_delete = sum(len(w['original']) + len(w['boundary']) for w in all_words) + total_extra

        if DEBUG:
            originals = [w['original'] for w in all_words]
            correcteds = [w['corrected'] for w in all_words]
            print(f'batch: {originals} → {correcteds} delete={total_delete} extra={total_extra}')

        # Determine the target language so we can ensure correct layout
        fg = _get_foreground_exe_name()
        desired = contains_non_english(all_words[0]['corrected']) or 'en'

        # For safe apps, ensure layout matches what we'll type
        if fg in SPACE_CORRECTION_ALLOWED_EXE and SAFE_APP_NO_CLIPBOARD_REPLACE:
            for _ in range(8):
                lang_now = _foreground_lang_id()
                if _layout_matches_target(lang_now, desired):
                    break
                _toggle_layout_once()
                time.sleep(0.03)

        # Delete everything
        for _ in range(total_delete):
            keyboard.send('backspace')
            time.sleep(0.004)

        # Retype everything with corrections
        for w in all_words:
            corrected_text = w['corrected']
            bnd = w['boundary']

            # Use send_unicode_text for layout-independent output
            send_unicode_text(corrected_text)

            if bnd == ' ':
                _send_vk(VK_SPACE)
            else:
                send_unicode_text(bnd)

        # Note the correction target for auto-switch tracking
        correction_target = contains_non_english(all_words[-1]['corrected']) or 'en'
        _note_correction_target(correction_target)

    finally:
        _inject_overflow = ''
        buffer_chars = ''
        time.sleep(0.02)
        injecting.clear()


def _replace_word(original: str, corrected: str, boundary_text: str) -> None:
    global buffer_chars, _inject_overflow
    injecting.set()
    try:
        # Small delay so the target app finishes committing the last keystroke.
        time.sleep(0.01)

        fg = _get_foreground_exe_name()
        is_whatsapp = bool(fg) and ('whatsapp' in fg)

        # In allowlisted non-WhatsApp apps, do a no-clipboard replace on SPACE.
        if (
            SAFE_APP_NO_CLIPBOARD_REPLACE
            and boundary_text == ' '
            and (fg in SPACE_CORRECTION_ALLOWED_EXE)
            and (not is_whatsapp)
        ):
            # Safe-app (VS Code, etc.) no-clipboard replace.
            # Delete: space + wrong word, then type corrected + space.

            desired = contains_non_english(corrected) or 'en'

            # Ensure the active layout matches what we're about to type, because keyboard.write
            # is layout-dependent (and VS Code may ignore KEYEVENTF_UNICODE injections).
            for _ in range(8):
                lang_now = _foreground_lang_id()
                if _layout_matches_target(lang_now, desired):
                    break
                _toggle_layout_once()
                time.sleep(0.03)

            # Snapshot default target (if any) so we can restore after correction.
            title_now = _get_foreground_window_title()
            default_target = _pick_default_lang_for_foreground(fg, title_now)

            # Account for characters the user typed during the correction delay.
            # These characters are in the app (user's physical keystrokes always reach it)
            # but were tracked in buffer_chars / _inject_overflow by on_key.
            extra_pre = len(buffer_chars)   # typed during the 30ms delay
            extra_inj = len(_inject_overflow)  # typed after injecting.set()
            total_extra = extra_pre + extra_inj
            total_to_delete = len(original) + 1 + total_extra
            if DEBUG and total_extra > 0:
                print(f'replace: accounting for {total_extra} extra chars (pre={extra_pre} inj={extra_inj})')

            for _ in range(total_to_delete):
                keyboard.send('backspace')
                time.sleep(0.004)

            keyboard.write(corrected)
            keyboard.send('space')

            # Note the correction target for auto-switch tracking
            _note_correction_target(desired)

            # Restore app default layout if we temporarily switched away (e.g., user typed Hebrew
            # in VS Code but default is English) - BUT ONLY if we haven't auto-switched layout yet.
            # If we already auto-switched, don't fight with the user's intended language.
            if (
                default_target in {'en', 'he'}
                and default_target != desired
                and not _auto_switched_wait_for_boundary
            ):
                def _restore_default():
                    try:
                        time.sleep(0.12)
                        _switch_layout_to(default_target)
                    except Exception:
                        return

                threading.Thread(target=_restore_default, daemon=True).start()
            if DEBUG:
                print(
                    f'replace: safe-app keys fg={fg!r} original={original!r} corrected={corrected!r} '
                    f'desired={desired!r} default={default_target!r} paused={_auto_switched_wait_for_boundary}'
                )
            return

        # מוחקים קודם את ה-boundary שכבר הודפס (רווח/פיסוק), ואז את המילה
        _send_vk(VK_BACK)
        time.sleep(0.005)

        # מוחקים את המילה כפי שנכתבה
        for _ in range(len(original)):
            _send_vk(VK_BACK)
            time.sleep(0.005)

        # מדפיסים טקסט מתוקן באופן בלתי תלוי בפריסה
        send_unicode_text(corrected)
        if boundary_text == ' ':
            _send_vk(VK_SPACE)
        else:
            send_unicode_text(boundary_text)

        # Note the correction target for auto-switch tracking
        correction_target = contains_non_english(corrected) or 'en'
        _note_correction_target(correction_target)

        if DEBUG:
            print(f'replace: unicode original={original!r} corrected={corrected!r} boundary={boundary_text!r}')
    finally:
        _inject_overflow = ''
        buffer_chars = ''
        time.sleep(0.02)
        injecting.clear()


def on_key(event: keyboard.KeyboardEvent):
    global buffer_chars, buffer_keys, _auto_switched_wait_for_boundary, buffer_was_hebrew, words_in_sentence, sentence_lang, _inject_overflow, _pending_words

    if event.event_type == 'up':
        return

    # אם אנחנו באמצע הקלדה "סינתטית" – לא מאחסנים כלום ולא חוסמים
    if injecting.is_set():
        # Track user physical keystrokes so _replace_word can account for them
        if isinstance(event.name, str) and len(event.name) == 1 and event.name.isprintable():
            _inject_overflow += event.name
        elif event.name in BOUNDARY_KEY_TO_TEXT:
            _inject_overflow += BOUNDARY_KEY_TO_TEXT[event.name]
        return

    # Resume corrections on ENTER or period, but stay in current language
    if event.name in {'enter', '.', 'dot'}:
        if _auto_switched_wait_for_boundary:
            _auto_switched_wait_for_boundary = False
            if DEBUG:
                print('Resuming corrections (staying in current language)')
        # Reset word counter, sentence language, and pending buffer - start of new sentence
        words_in_sentence = 0
        sentence_lang = None
        _pending_words = []

    boundary_text = BOUNDARY_KEY_TO_TEXT.get(event.name)
    if boundary_text is None and event.name in WORD_BOUNDARIES:
        boundary_text = ' ' if event.name == 'space' else str(event.name)

    if boundary_text is not None:
        # אם החלפנו שפה אוטומטית, לא נבדוק תיקונים עד גבול משפט
        if _auto_switched_wait_for_boundary:
            buffer_chars = ''
            buffer_keys = ''
            return

        # ב-WhatsApp Desktop מתקנים בצורה אמינה רק על רווח, באמצעות Clipboard
        if USE_CLIPBOARD_CORRECTION_ON_SPACE and boundary_text == ' ':
            fg = _get_foreground_exe_name()
            if DEBUG:
                print(f'Foreground exe: {fg!r}')
            is_whatsapp = bool(fg) and ('whatsapp' in fg)
            if not is_whatsapp:
                # Outside WhatsApp: allow space-correction only for a small allowlist (e.g., Teams).
                if fg in SPACE_CORRECTION_ALLOWED_EXE:
                    # For allowlisted safe apps, schedule a delayed correction so the just-typed
                    # space is already committed to the editor, without using clipboard.
                    if SAFE_APP_NO_CLIPBOARD_REPLACE:
                        snapshot = buffer_chars
                        buffer_chars = ''
                        buffer_keys = ''

                        def _delayed():
                            try:
                                # Give the target app time to commit the space.
                                time.sleep(0.03)
                                # Ensure we are still in the same kind of safe app.
                                now_fg = _get_foreground_exe_name()
                                if now_fg != fg:
                                    return
                                _decide_and_correct_for_word(snapshot, ' ')
                            except Exception:
                                return

                        threading.Thread(target=_delayed, daemon=True).start()
                        return

                    decide_and_correct(boundary_text)
                else:
                    # לא לתקן מחוץ לווטסאפ (כדי לא להעתיק/להדביק בטרמינל)
                    if DEBUG:
                        print('Skipping correction (foreground is not WhatsApp)')
                buffer_chars = ''
                buffer_keys = ''
                return

            # WhatsApp: use standard correction
            decide_and_correct(boundary_text)

            # Ensure buffers don't leak across words
            buffer_chars = ''
            buffer_keys = ''
            buffer_was_hebrew = False
            return

        # ב-WhatsApp: אם זה פיסוק, נשמור אותו כחלק מהמילה (כדי ש-ac, יהפוך לשבת)
        if USE_CLIPBOARD_CORRECTION_ON_SPACE and WHATSAPP_PUNCTUATION_AS_TEXT and boundary_text != ' ':
            buffer_chars += boundary_text
            buffer_keys += boundary_text
            return

        decide_and_correct(boundary_text)
        return

    if event.name == 'enter':
        buffer_chars = ''
        buffer_keys = ''
        buffer_was_hebrew = False
        _pending_words = []
        # Don't reset _auto_switched_wait_for_boundary here - it's already handled above
        return

    if event.name == 'backspace':
        buffer_chars = buffer_chars[:-1]
        buffer_keys = buffer_keys[:-1]
        return

    if isinstance(event.name, str) and len(event.name) == 1:
        ch = _event_to_char(event) or event.name
        was_hebrew_key = False
        # לפעמים נקבל יותר מתו אחד (נדיר); נכניס את כולם
        
        # Filter out control characters - if we got them, something went wrong
        # In Electron/WebView apps, _event_to_char sometimes returns control chars
        # because GetKeyboardState reports stale Ctrl modifier.
        if ch and any(ord(c) < 32 for c in ch):
            if DEBUG:
                print(f'Ignoring control character: {ch!r} from event.name={event.name!r}')
            
            # Determine intended character based on current keyboard layout.
            # event.name is always the English physical key name (e.g. 'a').
            # If layout is non-English, map physical key -> native character
            # using PROFILE_FROM_ENGLISH (en->native), NOT PROFILE_TO_ENGLISH.
            mapped = False
            lang_now = _foreground_lang_id()
            if not is_english_lang_id(lang_now):
                for from_en_map in PROFILE_FROM_ENGLISH.values():
                    if event.name in from_en_map:
                        ch = from_en_map[event.name]
                        was_hebrew_key = True
                        mapped = True
                        break
            if not mapped:
                if event.name.isprintable() and event.name.isascii():
                    ch = event.name
                else:
                    return  # Skip this key entirely
        
        buffer_chars += ch
        buffer_keys += event.name
        # Track if any key in the buffer was typed in non-English layout
        if not was_hebrew_key and len(buffer_chars) >= 1:
            for to_en_map in PROFILE_TO_ENGLISH.values():
                if event.name in to_en_map:
                    was_hebrew_key = True
                    break
        if was_hebrew_key:
            buffer_was_hebrew = True
        
        if DEBUG and len(buffer_chars) > 0 and len(buffer_chars) % 5 == 0:
            print(f'Buffer: {buffer_chars!r}')
        
        # Update auto-learning stats for current chat
        if ENABLE_AUTO_LEARN_CHAT_LANG:
            fg = _get_foreground_exe_name()
            if fg in WATCH_TITLE_CHANGES_EXE:
                title = _get_foreground_window_title()
                _update_chat_lang_stats(fg, title, ch)
        
        return

    return


def _request_stop():
    stop_event.set()


def _print_foreground_info() -> None:
    try:
        exe = _get_foreground_exe_name()
        title = _get_foreground_window_title()
        clean = _strip_invisible_controls(title)
        focus_hwnd = _get_foreground_focus_hwnd()

        lang_id = _foreground_lang_id()
        # Also report the raw HKL/lang_id for the top-level foreground window thread.
        fg_hwnd = user32.GetForegroundWindow()
        fg_lang_id = 0
        try:
            if fg_hwnd:
                pid = wintypes.DWORD(0)
                tid = user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(pid))
                fg_lang_id = int(user32.GetKeyboardLayout(tid)) & 0xFFFF
        except Exception:
            fg_lang_id = 0
        print(
            f'FG info: exe={exe!r} lang_id=0x{lang_id:04x} '
            f'fg_lang_id=0x{fg_lang_id:04x} focus_hwnd=0x{int(focus_hwnd):X} '
            f'title={title!r} clean_title={clean!r}'
        )
    except Exception as e:
        print(f'FG info failed: {e!r}')


def _load_ui_config() -> None:
    """Load config from UI config file only when launched from the UI."""
    # Only load UI config when explicitly launched from UI (AUTO_LANG2_UI_MODE=1)
    if os.environ.get('AUTO_LANG2_UI_MODE') != '1':
        return

    config_path = os.environ.get('AUTO_LANG2_CONFIG', '')
    if not config_path:
        # Also check default path
        default_path = os.path.expanduser('~/.auto_lang2_config.json')
        if os.path.exists(default_path):
            config_path = default_path
    if not config_path or not os.path.exists(config_path):
        return

    global DEBUG, EXCLUDE_WORDS, AUTO_SWITCH_LAYOUT, AUTO_SWITCH_AFTER_CONSECUTIVE
    global APP_DEFAULT_LANG_BY_EXE, APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)

        # App defaults
        if 'app_defaults' in cfg:
            APP_DEFAULT_LANG_BY_EXE.clear()
            APP_DEFAULT_LANG_BY_EXE.update(cfg['app_defaults'])

        # Chat defaults
        if 'chat_defaults' in cfg:
            APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING.clear()
            for exe, chats in cfg['chat_defaults'].items():
                if chats:
                    APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING[exe] = chats

        # Extra watch-title exes (from config, for apps without chat defaults)
        if 'watch_title_exes' in cfg:
            _EXTRA_WATCH_TITLE_EXE.clear()
            _EXTRA_WATCH_TITLE_EXE.update(cfg['watch_title_exes'])

        # Rebuild the unified set
        _rebuild_watch_title_set()

        # Exclude words
        if 'exclude_words' in cfg:
            EXCLUDE_WORDS = set(w.lower() for w in cfg['exclude_words'])

        # General settings
        if 'debug' in cfg:
            DEBUG = cfg['debug']
        if 'auto_switch' in cfg:
            AUTO_SWITCH_LAYOUT = cfg['auto_switch']
        if 'auto_switch_count' in cfg:
            AUTO_SWITCH_AFTER_CONSECUTIVE = cfg['auto_switch_count']

        if DEBUG:
            print(f'Loaded UI config from: {config_path}')
            print(f'  - app_defaults: {len(APP_DEFAULT_LANG_BY_EXE)} apps')
            print(f'  - watch_title_exes: {WATCH_TITLE_CHANGES_EXE}')
            print(f'  - exclude_words: {EXCLUDE_WORDS or "none"}')
    except Exception as e:
        print(f'Failed to load UI config: {e}')


def _check_expiration() -> None:
    """Check if the trial period has expired."""
    from datetime import date
    expiration = date(2026, 3, 1)
    if date.today() > expiration:
        print('Trial period has expired. Please contact the developer.')
        import sys
        sys.exit(0)


def main():
    _check_expiration()
    _ensure_single_instance()

    # Auto-detect installed keyboard layouts and load language profiles
    _detect_and_load_languages()

    # Load UI config if available
    _load_ui_config()

    # Load previously learned chat preferences
    _load_learned_prefs()

    keyboard.add_hotkey(EXIT_HOTKEY, _request_stop)
    keyboard.add_hotkey(INFO_HOTKEY, _print_foreground_info)
    keyboard.on_press(on_key, suppress=False)

    if ENABLE_APP_DEFAULT_LAYOUT:
        threading.Thread(target=_app_default_layout_watcher, name='app_default_layout', daemon=True).start()

    print('Auto layout fixer running')
    print(f'- Exit hotkey: {EXIT_HOTKEY}')
    langs = ', '.join(f'{p.flag} {p.name}' for p in ACTIVE_PROFILES.values())
    print(f'- Active languages: {langs or "none (Hebrew fallback)"}')
    print('- Examples:')
    for code, p in ACTIVE_PROFILES.items():
        # Show one example per language
        first_native = list(p.to_english.keys())[:4]
        first_en = [p.to_english[c] for c in first_native]
        print(f'  - {"".join(first_native)} -> {"".join(first_en)} ({p.name})')
    if ENABLE_AUTO_LEARN_CHAT_LANG:
        print(f'- Auto-learning enabled (tracking language per chat)')

    stop_event.wait()
    keyboard.unhook_all()

    # Save learned preferences on exit
    _save_learned_prefs()


if __name__ == '__main__':
    main()
