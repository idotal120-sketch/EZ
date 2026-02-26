import ctypes
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
- COMMON_ENGLISH_WORDS: רשימת מילים אנגליות שלא לתקן
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


# ----------------------------
# Configuration
# ----------------------------

# קיצור דרך ליציאה (גם כשאין פוקוס על הטרמינל)
EXIT_HOTKEY = 'ctrl+alt+q'

# מצב תאימות ל-WhatsApp Desktop (Electron): מתקנים על בסיס Clipboard במקום לסמוך על תווים מה-hook
USE_CLIPBOARD_CORRECTION_ON_SPACE = True

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

# ברירת מחדל לפי exe (מומלץ). אפשר להוסיף/להסיר לפי מה שיש אצלך.
APP_DEFAULT_LANG_BY_EXE: dict[str, str] = {
    # Hebrew-first apps
    'whatsapp.root.exe': 'he',
    'whatsapp.exe': 'he',
    'telegram.exe': 'he',
    'winword.exe': 'he',
    'teams.exe': 'he',

    # English-first apps
    'code.exe': 'en',
    'pycharm64.exe': 'en',
    'windowsterminal.exe': 'en',
    'powershell.exe': 'en',
    'cmd.exe': 'en',
    'slack.exe': 'en',
    'excel.exe': 'en',
}

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

# מיפוי מקלדת עברית→אנגלית (מקליד אנגלית כשהפריסה על עברית)
HEBREW_TO_ENGLISH = {
    'ק': 'e', 'ר': 'r', 'א': 't', 'ט': 'y', 'ו': 'u', 'ן': 'i', 'ם': 'o', 'פ': 'p',
    'ש': 'a', 'ד': 's', 'ג': 'd', 'כ': 'f', 'ע': 'g', 'י': 'h', 'ח': 'j', 'ל': 'k', 'ך': 'l', 'ף': ';',
    'ז': 'z', 'ס': 'x', 'ב': 'c', 'ה': 'v', 'נ': 'b', 'מ': 'n', 'צ': 'm', 'ת': ',', 'ץ': '.',
    "'": 'w',
}

# מיפוי מקלדת אנגלית→עברית (מקליד עברית כשהפריסה על אנגלית)
ENGLISH_TO_HEBREW = {v: k for k, v in HEBREW_TO_ENGLISH.items()}

# מרחיבים גם אותיות גדולות
ENGLISH_TO_HEBREW.update({k.upper(): v for k, v in ENGLISH_TO_HEBREW.items() if len(k) == 1 and k.isalpha()})

# מילים אנגליות נפוצות שלא נרצה להפוך לעברית בטעות
COMMON_ENGLISH_WORDS = {
    'how', 'are', 'you', 'the', 'is', 'at', 'it', 'in', 'on', 'to', 'of', 'for', 'and', 'or', 'not',
    'be', 'he', 'she', 'we', 'me', 'my', 'your', 'his', 'her', 'our', 'was', 'were', 'have', 'has',
    'had', 'do', 'does', 'did', 'can', 'could', 'will', 'would', 'should', 'may', 'might', 'must'
}

# מילים אנגליות קצרות (2-3 אותיות) שנרצה להשאיר באנגלית ולא להפוך לעברית.
SHORT_ENGLISH_WORDS = {
    'as', 'an', 'am', 'if', 'ok', 'no', 'so', 'up', 'us', 'we', 'me', 'he', 'be', 'do', 'go', 'hi',
}

# ראשי תיבות באנגלית ללא תנועות (ברירת מחדל: ריק).
# אם אתה רוצה שהמרות כמו "אני" -> "tbh" כן יעבדו, הוסף כאן למשל: {'tbh', 'idk', 'lol'}
ALLOW_VOWELLESS_ENGLISH = set()


# ----------------------------
# Win32 Unicode typing (layout independent)
# ----------------------------

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

psapi = ctypes.windll.psapi

# Fix ctypes default signatures (avoid pointer truncation on 64-bit)
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = wintypes.INT
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, wintypes.INT]
user32.GetWindowTextW.restype = wintypes.INT
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL
user32.LoadKeyboardLayoutW.argtypes = [wintypes.LPCWSTR, wintypes.UINT]
user32.LoadKeyboardLayoutW.restype = wintypes.HKL
user32.ActivateKeyboardLayout.argtypes = [wintypes.HKL, wintypes.UINT]
user32.ActivateKeyboardLayout.restype = wintypes.HKL
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

MAPVK_VSC_TO_VK_EX = 3

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

WM_INPUTLANGCHANGEREQUEST = 0x0050
KLF_ACTIVATE = 0x00000001


ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong


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
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return user32.GetKeyboardLayout(0)
    pid = ctypes.c_ulong(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return user32.GetKeyboardLayout(tid)


def _get_foreground_exe_name() -> str:
    """Return lowercase exe basename of the foreground window process, or '' if unknown."""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ''

    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ''

    hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not hproc:
        return ''
    try:
        buf_len = wintypes.DWORD(260)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if kernel32.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(buf_len)):
            full = buf.value
            base = full.rsplit('\\', 1)[-1]
            return base.lower()

        # fallback
        name_buf = ctypes.create_unicode_buffer(260)
        if psapi.GetModuleBaseNameW(hproc, None, name_buf, 260):
            return name_buf.value.lower()
        return ''
    finally:
        kernel32.CloseHandle(hproc)


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


def _klid_from_lang_id(lang_id: int) -> str:
    # KLID format: 8 hex digits, e.g. 00000409 / 0000040D
    return f'{lang_id:08x}'.upper()


def _set_foreground_input_language(lang_id: int) -> bool:
    """Explicitly request input language for the current foreground window (best-effort)."""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return False

    klid = _klid_from_lang_id(lang_id)
    try:
        hkl = user32.LoadKeyboardLayoutW(klid, KLF_ACTIVATE)
    except Exception:
        return False

    if not hkl:
        return False

    # Activate for current thread (best-effort) and ask the foreground window to change.
    try:
        user32.ActivateKeyboardLayout(hkl, 0)
    except Exception:
        pass

    try:
        user32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl)
    except Exception:
        return False

    return True


def _pick_default_lang_for_foreground(exe_name: str, title: str) -> str | None:
    if not exe_name and not title:
        return None

    exe = (exe_name or '').lower()
    title_l = (title or '').lower()

    direct = APP_DEFAULT_LANG_BY_EXE.get(exe)
    if direct in {'en', 'he'}:
        return direct

    if any(tok.lower() in title_l for tok in ENGLISH_APPS_TITLE):
        return 'en'
    if any(tok.lower() in title_l for tok in HEBREW_APPS_TITLE):
        return 'he'

    return None


def _apply_app_default_layout_if_needed() -> None:
    if not ENABLE_APP_DEFAULT_LAYOUT:
        return

    exe = _get_foreground_exe_name()
    title = _get_foreground_window_title()
    target = _pick_default_lang_for_foreground(exe, title)
    if not target:
        return

    lang_id = _foreground_lang_id()
    if target == 'he':
        desired = LANG_HEBREW
        if _is_hebrew_lang(lang_id):
            return
    else:
        desired = LANG_ENGLISH_US
        if (lang_id != 0) and (not _is_hebrew_lang(lang_id)):
            return

    if DEBUG:
        print(f'App default layout: exe={exe!r} title={title!r} -> {target!r}')

    injecting.set()
    try:
        _set_foreground_input_language(desired)
        time.sleep(0.03)
    finally:
        injecting.clear()


def _app_default_layout_watcher() -> None:
    last_hwnd = None
    while not stop_event.is_set():
        try:
            if injecting.is_set():
                time.sleep(APP_DEFAULT_POLL_INTERVAL_SEC)
                continue

            hwnd = user32.GetForegroundWindow()
            if hwnd and hwnd != last_hwnd:
                last_hwnd = hwnd
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
    if rc <= 0:
        return None

    ch = buf.value[:rc]
    if not ch:
        return None
    return ch


# ----------------------------
# Core logic
# ----------------------------

buffer_chars = ''
buffer_keys = ''

en_streak = 0
he_streak = 0

injecting = threading.Event()  # מונע לופ: לא מאחסנים אירועים שאנחנו עצמנו שולחים
stop_event = threading.Event()

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
        return int(hkl) & 0xFFFF
    except Exception:
        return 0


def _is_hebrew_lang(lang_id: int) -> bool:
    return lang_id == LANG_HEBREW


def _is_english_lang(lang_id: int) -> bool:
    return lang_id in {LANG_ENGLISH_US, LANG_ENGLISH_UK}


def _toggle_layout_once() -> None:
    # best-effort: try common OS hotkeys (some apps ignore SendInput modifiers)
    for hk in LAYOUT_TOGGLE_HOTKEYS:
        try:
            keyboard.send(hk)
            return
        except Exception:
            pass

    # last resort (may be ignored by some apps)
    _send_vk_combo(VK_MENU, VK_SHIFT, inter_delay=0.001)


def _switch_layout_to(target: str) -> bool:
    """Best-effort switch input language for the foreground window."""
    if target not in {'en', 'he'}:
        return False

    injecting.set()
    try:
        for _ in range(6):
            lang_id = _foreground_lang_id()
            if DEBUG:
                print(f'Layout check: lang_id=0x{lang_id:04x} target={target!r}')
            if target == 'he':
                if _is_hebrew_lang(lang_id):
                    return True
            else:
                # English target: accept any non-Hebrew as "good enough".
                if (lang_id != 0) and (not _is_hebrew_lang(lang_id)):
                    return True

            _toggle_layout_once()
            time.sleep(0.05)
        return False
    finally:
        time.sleep(0.01)
        injecting.clear()


def _note_correction_target(target: str | None) -> None:
    """Update consecutive correction counters and maybe auto-switch layout."""
    global en_streak, he_streak

    if target == 'en':
        en_streak += 1
        he_streak = 0
    elif target == 'he':
        he_streak += 1
        en_streak = 0
    else:
        en_streak = 0
        he_streak = 0

    if not AUTO_SWITCH_LAYOUT:
        return

    if en_streak >= AUTO_SWITCH_AFTER_CONSECUTIVE:
        if DEBUG:
            print('Auto-switching layout -> English')
        if _switch_layout_to('en'):
            en_streak = 0

    if he_streak >= AUTO_SWITCH_AFTER_CONSECUTIVE:
        if DEBUG:
            print('Auto-switching layout -> Hebrew')
        if _switch_layout_to('he'):
            he_streak = 0


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


def _correct_previous_word_from_keybuffer(boundary_text: str) -> bool:
    """Fallback for WhatsApp Desktop when clipboard copy is blocked."""
    global buffer_keys, buffer_chars
    if not buffer_keys and not buffer_chars:
        return False

    # שים לב: ב-WhatsApp Desktop, keyboard.event.name לרוב מחזיר את התו שהוקלד בפועל
    # (כלומר "ים'"), לא את המקש הפיזי ("h o w"). לכן מתקנים לפי התווים בפועל.
    # Prefer the actual typed characters (buffer_chars). buffer_keys is the physical key names.
    typed = buffer_chars or buffer_keys
    corrected = None
    delete_len = len(typed)
    correction_target = None

    lang_id = _foreground_lang_id()
    if DEBUG:
        print(f'Keybuffer: active lang_id=0x{lang_id:04x}')

    if contains_hebrew(typed):
        # הוקלד בעברית כשהכוונה לאנגלית
        candidate = translate(typed, HEBREW_TO_ENGLISH)
        if DEBUG:
            print(f'Keybuffer: typed(he)={typed!r} -> en={candidate!r}')
        # ברירת מחדל: מתקנים HE->EN כשפריסה עברית.
        # חריג: אם התוצאה היא מילה אנגלית נפוצה מאוד, נתקן גם אם זיהוי השפה הפעילה לא מדויק.
        strong_english = candidate.lower() in COMMON_ENGLISH_WORDS
        if (_is_hebrew_lang(lang_id) or strong_english) and english_plausible_for_hebrew_to_english(candidate):
            corrected = candidate
            correction_target = 'en'
    else:
        # הוקלד באנגלית כשהכוונה לעברית
        candidate = translate(typed, ENGLISH_TO_HEBREW)
        if DEBUG:
            print(f'Keybuffer: typed(en)={typed!r} -> he={candidate!r}')
        # מתקנים EN->HE בעיקר כשאנחנו בפריסה אנגלית והמילה נראית "לא אנגלית"
        if (not _is_hebrew_lang(lang_id)) and (typed.lower() not in COMMON_ENGLISH_WORDS) and hebrew_plausible(candidate) and english_unlikely_for_en_to_he(typed):
            corrected = candidate
            correction_target = 'he'

    if not corrected or delete_len is None:
        if DEBUG:
            print('Keybuffer: no correction decision')
        _note_correction_target(None)
        buffer_keys = ''
        buffer_chars = ''
        return False

    # WhatsApp Desktop: prefer deletion (Ctrl+Backspace) + paste.
    if boundary_text == ' ':
        replacement = corrected + ' '

        # Ctrl+Backspace is fast, but can fail around punctuation/RTL. Use it only for simple alnum tokens.
        is_ascii = typed.isascii()
        is_simple_alnum = is_ascii and typed.isalnum()
        has_ascii_punct = is_ascii and (not typed.isalnum())

        # For shortcuts like "ac,": selection+paste is more stable than backspace-count.
        if WHATSAPP_REPLACE_PUNCTUATION_BY_SELECTION and has_ascii_punct:
            select_chars = delete_len + 1
            _replace_by_select_and_paste(select_chars, replacement)
            _note_correction_target(correction_target)
            buffer_keys = ''
            buffer_chars = ''
            return True
        if WHATSAPP_REPLACE_USE_CTRL_BACKSPACE and is_simple_alnum:
            ok = _replace_prev_word_by_ctrl_backspace_paste(replacement)
            if ok:
                _note_correction_target(correction_target)
                buffer_keys = ''
                buffer_chars = ''
                return True

        # Reliable default: delete exact char count (word + trailing space), then paste.
        ok = _replace_by_backspace_count_paste(delete_len + 1, replacement)
        if ok:
            _note_correction_target(correction_target)
            buffer_keys = ''
            buffer_chars = ''
            return True

        # Last resort: selection+paste
        select_chars = delete_len + 1
        _replace_by_select_and_paste(select_chars, replacement)
        _note_correction_target(correction_target)
    else:
        # fallback to delete/inject for non-space boundaries
        _replace_by_delete_len(delete_len, corrected, boundary_text)
        _note_correction_target(correction_target)
    buffer_keys = ''
    buffer_chars = ''
    return True


def _correct_previous_word_via_clipboard() -> bool:
    """WhatsApp Desktop fallback: copy previous word from the input field and replace it if needed."""
    if not correction_lock.acquire(blocking=False):
        return False

    injecting.set()
    original_clip = None
    clipboard_restored = False
    try:
        original_clip = _get_clipboard_text()

        if DEBUG:
            print('Space detected -> attempting clipboard correction')

        sentinel = f'__AUTO_LANG_SENTINEL__{time.time_ns()}__'
        _set_clipboard_text(sentinel)
        time.sleep(0.01)

        def _try_copy_with_selection(select_space_too: bool) -> str:
            # caret is after the just-typed space.
            if select_space_too:
                # Select space then expand to previous word
                _send_vk_combo(VK_SHIFT, VK_LEFT, inter_delay=0.001)
                _send_vk_combo(VK_CONTROL, VK_SHIFT, VK_LEFT, inter_delay=0.001)
            else:
                # Select previous word only
                _send_vk_combo(VK_CONTROL, VK_SHIFT, VK_LEFT, inter_delay=0.001)

            _send_vk_combo(VK_CONTROL, VK_C, inter_delay=0.001)
            time.sleep(0.10)
            return _get_clipboard_text() or ''

        copied_raw = _try_copy_with_selection(select_space_too=False)
        if copied_raw == sentinel or copied_raw == '':
            copied_raw = _try_copy_with_selection(select_space_too=True)

        if copied_raw == sentinel or copied_raw == '':
            if DEBUG:
                print('Copy failed (clipboard unchanged)')
            # Clear selection and return caret to end
            _send_vk(VK_RIGHT)
            return False

        copied_raw = copied_raw.replace('\r\n', '\n')
        copied = _sanitize_copied_word(copied_raw)

        had_trailing_space = copied.endswith(' ')
        copied_word = copied[:-1] if had_trailing_space else copied

        if DEBUG:
            print(f'Copied selection: {copied!r} -> word={copied_word!r} trailing_space={had_trailing_space}')

        if not copied_word:
            _send_vk(VK_RIGHT)
            return False

        corrected = None
        if contains_hebrew(copied_word):
            candidate = translate(copied_word, HEBREW_TO_ENGLISH)
            if english_plausible(candidate):
                corrected = candidate
        else:
            if copied_word.lower() not in COMMON_ENGLISH_WORDS:
                candidate = translate(copied_word, ENGLISH_TO_HEBREW)
                if hebrew_plausible(candidate):
                    corrected = candidate

        if corrected and corrected != copied_word:
            replacement = corrected + (' ' if had_trailing_space else '')
            # Paste is more reliable in Electron apps than Unicode injection
            _set_clipboard_text(replacement)
            time.sleep(0.03)
            _send_vk_combo(VK_CONTROL, VK_V, inter_delay=0.001)
            time.sleep(0.15)

            # Move caret back to where the user expects (after the existing typed space)
            _send_vk(VK_RIGHT)

            # חשוב: ב-Electron לפעמים הקריאה ל-clipboard מתבצעת באיחור,
            # אז אסור לשחזר את ה-clipboard מוקדם מדי.
            if original_clip is not None:
                _set_clipboard_text(original_clip)
                clipboard_restored = True

            if DEBUG:
                print(f'Replaced with: {replacement!r}')
            return True
        else:
            if DEBUG:
                print('No correction needed')
            return False

        # תמיד מנקים בחירה ומחזירים את הסמן לסוף (אם לא עשינו כבר)
        _send_vk(VK_RIGHT)
    except Exception as e:
        # לא מפילים את הסקריפט אם ה-clipboard עסוק/ננעל ע"י מערכת אחרת
        print(f'Clipboard correction failed: {e!r}')
        return False
    finally:
        if (not clipboard_restored) and (original_clip is not None):
            try:
                _set_clipboard_text(original_clip)
            except Exception:
                pass
        time.sleep(0.01)
        injecting.clear()
        correction_lock.release()


def contains_hebrew(text: str) -> bool:
    return any('\u05D0' <= c <= '\u05EA' for c in text)


def translate(text: str, mapping: dict[str, str]) -> str:
    return ''.join(mapping.get(c, c) for c in text)


def english_plausible(word: str) -> bool:
    w = word.strip()
    if not w:
        return False
    # חייב להיות ASCII בלבד
    try:
        w.encode('ascii')
    except UnicodeEncodeError:
        return False
    # חייב להכיל לפחות אות אחת
    if not any(ch.isalpha() for ch in w):
        return False
    # לא מכיל תווים מוזרים
    allowed = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'`,-./;[]? !:")
    if any((ch not in allowed) for ch in w):
        return False
    # אם זו מילה אנגלית ממש נפוצה – חזקה מאוד
    if w.lower() in COMMON_ENGLISH_WORDS:
        return True
    # הימנעות ממחרוזות עם הרבה סימני פיסוק
    punct = sum(1 for ch in w if (not ch.isalnum()) and (ch != ' '))
    if punct > max(1, len(w) // 3):
        return False
    return True


def _has_english_vowel(word: str) -> bool:
    return any(ch in 'aeiouAEIOU' for ch in word)


def _vowel_density(word: str) -> float:
    letters = [ch for ch in word if ch.isalpha() and ch.isascii()]
    if not letters:
        return 0.0
    vowels = sum(1 for ch in letters if ch in 'aeiouAEIOU')
    return vowels / len(letters)


def _has_consonant_cluster(word: str, run: int = 4) -> bool:
    cluster = 0
    for ch in word:
        if ch.isalpha() and ch.isascii() and (ch not in 'aeiouAEIOU'):
            cluster += 1
            if cluster >= run:
                return True
        else:
            cluster = 0
    return False


def english_unlikely(word: str) -> bool:
    """Heuristic: word that doesn't look like real English (useful for deciding EN->HE conversion)."""
    w = word.strip()
    if not w:
        return False
    if not english_plausible(w):
        return True

    letters = [ch for ch in w if ch.isalpha() and ch.isascii()]
    if len(letters) >= 4:
        d = _vowel_density(w)
        if d < 0.20 or d > 0.60:
            return True
        if _has_consonant_cluster(w, run=4):
            return True
    return False


def english_unlikely_for_en_to_he(word: str) -> bool:
    """More permissive for EN->HE: converts short gibberish like 'ac,' but keeps common short words."""
    w = word.strip()
    if not w:
        return False

    if not english_plausible(w):
        return True

    lw = w.lower()
    if lw in COMMON_ENGLISH_WORDS or lw in SHORT_ENGLISH_WORDS:
        return False

    letters = [ch for ch in w if ch.isalpha() and ch.isascii()]
    if 1 <= len(letters) <= 3:
        return True

    return english_unlikely(w)


def english_plausible_for_hebrew_to_english(candidate: str) -> bool:
    """Stricter check for WhatsApp fallback to reduce false positives on real Hebrew words."""
    w = candidate.strip()
    if not english_plausible(w):
        return False
    lw = w.lower()
    if lw in COMMON_ENGLISH_WORDS:
        return True
    if lw in ALLOW_VOWELLESS_ENGLISH:
        return True
    return _has_english_vowel(w)


def hebrew_plausible(word: str) -> bool:
    w = word.strip()
    if not w:
        return False
    # חייב להכיל אות עברית אחת לפחות
    if not contains_hebrew(w):
        return False
    # לא יותר מדי לטיניות בתוך המילה
    latin_letters = sum(1 for ch in w if ch.isalpha() and not ('\u05D0' <= ch <= '\u05EA'))
    if latin_letters:
        return False
    return True


def decide_and_correct(boundary_text: str) -> None:
    global buffer_chars
    if not buffer_chars:
        return

    original = buffer_chars

    # 1) אם נכתב בעברית – כנראה התכוונו לאנגלית (אבל רק אם התרגום נראה באמת אנגלי)
    if contains_hebrew(original):
        candidate = translate(original, HEBREW_TO_ENGLISH)
        if english_plausible(candidate):
            threading.Thread(
                target=_replace_word,
                args=(original, candidate, boundary_text),
                daemon=True,
            ).start()
            return

        buffer_chars = ''
        return

    # 2) אם נכתב בלטינית – כנראה התכוונו לעברית (אבל לא אם זו מילה אנגלית מוכרת)
    if original.lower() not in COMMON_ENGLISH_WORDS:
        candidate = translate(original, ENGLISH_TO_HEBREW)
        if hebrew_plausible(candidate):
            threading.Thread(
                target=_replace_word,
                args=(original, candidate, boundary_text),
                daemon=True,
            ).start()
            return

    # לא זוהה כטעות
    buffer_chars = ''


def _replace_word(original: str, corrected: str, boundary_text: str) -> None:
    global buffer_chars
    injecting.set()
    try:
        # נותנים רגע קצר כדי שה-boundary (רווח/פיסוק) שנלחץ ייכנס לשדה הטקסט
        time.sleep(0.01)

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
    finally:
        buffer_chars = ''
        time.sleep(0.02)
        injecting.clear()


def on_key(event: keyboard.KeyboardEvent):
    global buffer_chars, buffer_keys

    if event.event_type == 'up':
        return

    # אם אנחנו באמצע הקלדה “סינתטית” – לא מאחסנים כלום ולא חוסמים
    if injecting.is_set():
        return

    boundary_text = BOUNDARY_KEY_TO_TEXT.get(event.name)
    if boundary_text is None and event.name in WORD_BOUNDARIES:
        boundary_text = ' ' if event.name == 'space' else str(event.name)

    if boundary_text is not None:
        # ב-WhatsApp Desktop מתקנים בצורה אמינה רק על רווח, באמצעות Clipboard
        if USE_CLIPBOARD_CORRECTION_ON_SPACE and boundary_text == ' ':
            fg = _get_foreground_exe_name()
            if DEBUG:
                print(f'Foreground exe: {fg!r}')
            is_whatsapp = bool(fg) and ('whatsapp' in fg)
            if not is_whatsapp:
                # לא לתקן מחוץ לווטסאפ (כדי לא להעתיק/להדביק בטרמינל)
                if DEBUG:
                    print('Skipping correction (foreground is not WhatsApp)')
                buffer_chars = ''
                buffer_keys = ''
                return

            # מנסים Clipboard; אם חסום (כמו אצלך), נופלים ל-keybuffer.
            ok = _correct_previous_word_via_clipboard()
            if (not ok) and USE_KEYBUFFER_FALLBACK_FOR_WHATSAPP:
                if DEBUG:
                    print('Falling back to keybuffer correction')
                ok2 = _correct_previous_word_from_keybuffer(boundary_text)
                if DEBUG and (not ok2):
                    print('Keybuffer fallback did not apply')

            # Ensure buffers don't leak across words
            buffer_chars = ''
            buffer_keys = ''
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
        return

    if event.name == 'backspace':
        buffer_chars = buffer_chars[:-1]
        buffer_keys = buffer_keys[:-1]
        return

    if isinstance(event.name, str) and len(event.name) == 1:
        ch = _event_to_char(event) or event.name
        # לפעמים נקבל יותר מתו אחד (נדיר); נכניס את כולם
        buffer_chars += ch
        buffer_keys += event.name
        return

    return


def _request_stop():
    stop_event.set()


keyboard.add_hotkey(EXIT_HOTKEY, _request_stop)
keyboard.on_press(on_key, suppress=False)

if ENABLE_APP_DEFAULT_LAYOUT:
    threading.Thread(target=_app_default_layout_watcher, name='app_default_layout', daemon=True).start()

print('Auto layout fixer running')
print(f'- Exit hotkey: {EXIT_HOTKEY}')
print('- Examples:')
print("  - 'יקךךם ' -> 'hello '")
print("  - 'akuo ' -> 'שלום '")

stop_event.wait()
keyboard.unhook_all()
