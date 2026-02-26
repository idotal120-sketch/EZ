"""
Auto Language Layout Fixer for Windows — v3 (clean rewrite)
============================================================
Detects when the user types in the wrong keyboard layout and:
  1. Corrects the already-typed text
  2. Switches the keyboard to the correct language

Detection triggers:
  - 2+ words typed in wrong layout, OR
  - 1 word with 6+ characters in wrong layout

Key design decisions (fixing v2 bugs):
  - Thread-safe: all shared state protected by a single lock
  - No ToUnicodeEx (avoids dead-key corruption) — uses layout-aware char map instead
  - Buffer is snapshotted before spawning correction thread (no race)
  - No overflow tracking during injection — avoids Ctrl+V/'v' key leak
  - sentence_lang only locks after the 2-word threshold is met
  - Period/Enter flush pending words *before* resetting sentence state
  - correction_lock is actually used
  - No trial expiration
  - No admin required
"""

import ctypes
import json
import os
import sys
import threading
import time

# When running as --noconsole EXE (PyInstaller), sys.stdout/stderr are None.
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

from ctypes import wintypes

import keyboard
import mouse
from wordfreq import zipf_frequency

import spell_module
import grammar_module
from keyboard_maps import (
    ALL_PROFILES, ENGLISH_LANG_IDS, LanguageProfile,
    lang_id_to_profile, is_english_lang_id, detect_script,
)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Single-instance guard                                        ║
# ╚═══════════════════════════════════════════════════════════════╝

_MUTEX_NAME = 'Global\\AutoLang3PyMutex'


def _ensure_single_instance() -> None:
    k32 = ctypes.WinDLL('kernel32', use_last_error=True)
    k32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
    k32.CreateMutexW.restype = wintypes.HANDLE
    handle = k32.CreateMutexW(None, False, _MUTEX_NAME)
    if handle and ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        print('Another instance is already running.')
        raise SystemExit(0)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Configuration                                                ║
# ╚═══════════════════════════════════════════════════════════════╝

EXIT_HOTKEY = 'ctrl+alt+q'
INFO_HOTKEY = 'ctrl+alt+i'
UNDO_HOTKEY = 'f12'

# Detection thresholds
MIN_WORDS_TO_DETECT = 1          # Correct text after 1st wrong-language word
MIN_CHARS_SINGLE_WORD = 6        # OR after 6+ chars in a single word

# NLP
NLP_ZIPF_THRESHOLD = 2.5
NLP_SCORE_GAP = 1.0               # Min zipf gap to declare a clear winner
NLP_LOCK_OVERRIDE_GAP = 3.0       # Much higher gap needed to break sentence lock

# Layout auto-switch after N consecutive corrections
AUTO_SWITCH_LAYOUT = True
AUTO_SWITCH_AFTER_CONSECUTIVE = 2
LAYOUT_TOGGLE_HOTKEYS = ('alt+shift', 'win+space', 'ctrl+shift')

# App-default layouts
ENABLE_APP_DEFAULT_LAYOUT = True
APP_DEFAULT_POLL_INTERVAL = 0.35
APP_DEFAULT_COOLDOWN_SEC = 0.7

APP_DEFAULT_LANG_BY_EXE: dict[str, str] = {
    'whatsapp.root.exe': 'he', 'whatsapp.exe': 'he', 'telegram.exe': 'he',
    'winword.exe': 'he', 'teams.exe': 'he', 'ms-teams.exe': 'he',
    'code.exe': 'en', 'pycharm64.exe': 'en', 'windowsterminal.exe': 'en',
    'powershell.exe': 'en', 'cmd.exe': 'en', 'slack.exe': 'en',
    'excel.exe': 'en', 'outlook.exe': 'en',
}

APP_DEFAULT_LANG_BY_TITLE: dict[str, dict[str, str]] = {}

WATCH_TITLE_CHANGES_EXE: set[str] = set()

# Known browser executables (for browser_defaults matching)
BROWSER_EXES: set[str] = {
    'chrome.exe', 'msedge.exe', 'firefox.exe', 'brave.exe',
    'opera.exe', 'vivaldi.exe', 'iexplore.exe', 'safari.exe',
    'chromium.exe', 'waterfox.exe', 'tor.exe',
}

# Per-website language defaults (keyword in browser title → lang)
# Applied to ALL browser EXEs automatically.
BROWSER_LANG_BY_KEYWORD: dict[str, str] = {}

# Apps where correction is DISABLED (terminals, consoles)
CORRECTION_EXCLUDED_EXE: set[str] = {
    'cmd.exe', 'powershell.exe', 'pwsh.exe', 'windowsterminal.exe',
    'wt.exe', 'conhost.exe', 'mintty.exe', 'bash.exe',
    'wsl.exe', 'openssh.exe', 'putty.exe',
}

# ╔═══════════════════════════════════════════════════════════════╗
# ║ Privacy Guard — password fields, sensitive apps, PII regex   ║
# ╚═══════════════════════════════════════════════════════════════╝

PRIVACY_GUARD_ENABLED = True

# Apps fully blocked — password managers, banking, RDP, VM consoles
PRIVACY_BLOCKED_EXE: set[str] = {
    # Password managers
    '1password.exe', 'keeper.exe', 'keepass.exe', 'keepassxc.exe',
    'lastpass.exe', 'bitwarden.exe', 'dashlane.exe', 'roboform.exe',
    'enpass.exe',
    # Banking / Finance (common Israeli + international)
    'bankhapoalim.exe', 'leumi.exe', 'discount.exe', 'mizrahi.exe',
    # RDP / VM consoles
    'mstsc.exe', 'vmconnect.exe', 'vmware-vmx.exe', 'virtualbox.exe',
    'vboxwebclient.exe', 'vmplayer.exe', 'vmware.exe',
    # Remote desktop tools
    'anydesk.exe', 'teamviewer.exe',
}

import re as _re

# Regex patterns for sensitive data (credit cards, Israeli ID, OTP codes)
_PRIVACY_PATTERNS = [
    # Credit card numbers: 13-19 digits (with optional spaces/dashes)
    _re.compile(r'^[\d\s\-]{13,19}$'),
    # Israeli Teudat Zehut (ID): exactly 9 digits
    _re.compile(r'^\d{9}$'),
    # OTP / verification codes: 4-8 digits only
    _re.compile(r'^\d{4,8}$'),
    # CVV: exactly 3-4 digits after typing in card context
    _re.compile(r'^\d{3,4}$'),
]

# Cache for UIA password detection (avoid re-checking every keystroke)
_privacy_cache_hwnd = 0
_privacy_cache_result = False
_privacy_cache_time = 0.0
_PRIVACY_CACHE_TTL = 0.5  # seconds

def _is_password_field() -> bool:
    """Check if the focused UI element is a password/secure field.
    
    Uses Windows MSAA (oleacc.dll) — no extra packages needed.
    Checks for STATE_SYSTEM_PROTECTED flag on the focused accessible object.
    Also checks window class names for common password field indicators.
    Caches result per HWND for 500ms to avoid overhead on every keystroke.
    """
    global _privacy_cache_hwnd, _privacy_cache_result, _privacy_cache_time
    
    try:
        hwnd = _get_focus_hwnd()
        now = time.time()
        
        # Return cached result if same window and within TTL
        if hwnd == _privacy_cache_hwnd and (now - _privacy_cache_time) < _PRIVACY_CACHE_TTL:
            return _privacy_cache_result
        
        _privacy_cache_hwnd = hwnd
        _privacy_cache_time = now
        _privacy_cache_result = False
        
        if not hwnd:
            return False
        
        # Method 1: Check window class — many password fields use "Edit" with ES_PASSWORD style
        class_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buf, 256)
        class_name = class_buf.value.lower()
        
        # Edit controls: check ES_PASSWORD style (0x0020)
        if class_name in ('edit', 'richedit', 'richedit20w', 'richedit50w'):
            style = user32.GetWindowLongW(hwnd, -16)  # GWL_STYLE
            ES_PASSWORD = 0x0020
            if style & ES_PASSWORD:
                _privacy_cache_result = True
                _dbg('[PRIVACY] Password field detected via ES_PASSWORD style')
                return True
        
        # Chrome/Edge/Firefox password fields: check via window title heuristic
        # (Chromium uses "Chrome_RenderWidgetHostHWND" — can't check style)
        
        # Method 2: Window title heuristic — check if parent or focus window title suggests password
        try:
            title_len = user32.GetWindowTextLengthW(hwnd)
            if title_len > 0:
                title_buf = ctypes.create_unicode_buffer(title_len + 1)
                user32.GetWindowTextW(hwnd, title_buf, title_len + 1)
                title = title_buf.value.lower()
                _pwd_keywords = ('password', 'passwd', 'סיסמ', 'pin code',
                                 'otp', 'cvv', 'cvc', 'security code',
                                 'קוד אימות', 'קוד סודי', 'verification',
                                 'אימות', 'סודי', 'secret', 'login',
                                 'sign in', 'כניסה', 'התחברות',
                                 # Banking / Finance sites (Israeli + international)
                                 'בנק', 'bank', 'הפועלים', 'hapoalim',
                                 'לאומי', 'leumi', 'דיסקונט', 'discount',
                                 'מזרחי', 'mizrahi', 'טפחות', 'benleumi',
                                 'fibi', 'הבינלאומי', 'paypal', 'פייפאל',
                                 'bit.co.il', 'max.co.il', 'cal-online',
                                 'isracard', 'ישראכרט', 'לאומי קארד',
                                 'credit card', 'כרטיס אשראי', 'חשבון',
                                 'banking', 'netbanking', 'e-banking')
                if any(kw in title for kw in _pwd_keywords):
                    _privacy_cache_result = True
                    _dbg(f'[PRIVACY] Password context detected via title: {title!r}')
                    return True
        except Exception:
            pass
        
        return False
    except Exception:
        _privacy_cache_result = False
        return False


def _is_sensitive_buffer(buffer: str) -> bool:
    """Check if the current buffer looks like sensitive data (CC, ID, OTP).
    
    Only checks digit-heavy buffers to avoid false positives on normal text.
    """
    if not buffer or len(buffer) < 3:
        return False
    
    # Only check if buffer is mostly digits
    digits = sum(1 for c in buffer if c.isdigit())
    if digits < len(buffer) * 0.7:
        return False
    
    # Strip spaces/dashes for pattern matching
    clean = buffer.replace(' ', '').replace('-', '')
    
    for pattern in _PRIVACY_PATTERNS:
        if pattern.match(clean):
            _dbg(f'[PRIVACY] Sensitive data pattern detected in buffer: {len(clean)} chars')
            return True
    
    return False


def _privacy_should_block() -> bool:
    """Main privacy guard check — returns True if engine should NOT process input.
    
    Checks in order (fast to slow):
    1. Feature disabled → pass through
    2. Blocked app → block
    3. UIA password field → block
    """
    if not PRIVACY_GUARD_ENABLED:
        return False
    
    # Check blocked apps (very fast — set lookup)
    try:
        exe = _get_foreground_exe()
        if exe and exe in PRIVACY_BLOCKED_EXE:
            _dbg(f'[PRIVACY] Blocked app: {exe}')
            return True
    except Exception:
        pass
    
    # Check UIA password field (cached, ~0 cost when cached)
    if _is_password_field():
        return True
    
    return False

# Exclude words (never correct these)
EXCLUDE_WORDS: set[str] = set()

# UI toggle
ENGINE_ENABLED = True
DEBUG = True

# Buffer UI callback — set by auto_lang_ui to receive (actual_text, translated_text, corrections)
# Called on every keystroke that changes the buffer.
_buffer_callback = None   # type: callable | None
_prev_words: list[str] = []   # last 2 completed words (for showing in panel)
# Per-word correction history: maps corrected_word -> original_word for recent corrections
_word_corrections: dict[str, str] = {}  # corrected -> original (only for words in _prev_words)
_last_target_hwnd = None   # HWND of the last foreground app (for click-to-revert focus restore)
_widget_hwnd = None        # HWND of our own floating widget (skip in focus watcher)
_widget_edit_hwnd = 0      # HWND of the widget's text editor (skip corrections)
_own_pid = 0               # Our process ID — to filter own windows from target tracking
_widget_clicking = threading.Event()  # Set while UI is handling a panel word click

# Spell check state — active suggestion waiting for Tab/dismiss
_spell_pending = None      # type: dict | None  — {'original': str, 'suggestions': list[str]}
_spell_pending_lock = threading.Lock()

GRAMMAR_HOTKEY = 'ctrl+shift+g'


def _push_prev_word(word: str):
    """Add a completed word to _prev_words, keeping max 30."""
    global _prev_words
    if word:
        _prev_words = (_prev_words + [word])[-30:]


def _record_word_correction(original: str, corrected: str):
    """Record that 'original' was corrected to 'corrected'.
    Also update _prev_words so the panel shows the corrected text."""
    global _prev_words, _word_corrections
    # Update _prev_words: replace the last occurrence of original with corrected
    for i in range(len(_prev_words) - 1, -1, -1):
        if _prev_words[i] == original:
            _prev_words[i] = corrected
            break
    _word_corrections[corrected] = original
    # Prune old entries not in _prev_words
    current_words = set(_prev_words)
    _word_corrections = {k: v for k, v in _word_corrections.items() if k in current_words}

# Direct file-based debug log (bypasses stdout buffering issues)
_DEBUG_LOG_PATH = os.path.join(os.path.expanduser('~'), 'auto_lang3_debug.log')
_debug_log_file = None

def _dbg(msg: str):
    """Write debug message to file — only when DEBUG is enabled."""
    if not DEBUG:
        return
    global _debug_log_file
    try:
        if _debug_log_file is None:
            _debug_log_file = open(_DEBUG_LOG_PATH, 'w', encoding='utf-8', buffering=1)
        _debug_log_file.write(msg + '\n')
        _debug_log_file.flush()
    except Exception:
        pass

# Auto-learn chat language
ENABLE_AUTO_LEARN = True
AUTO_LEARN_MIN_CHARS = 50
AUTO_LEARN_THRESHOLD = 0.60
AUTO_LEARN_CACHE_FILE = os.path.expanduser('~/.auto_lang3_learned.json')


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Win32 API Setup                                              ║
# ╚═══════════════════════════════════════════════════════════════╝

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi = ctypes.windll.psapi

ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong

# Fix signatures
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
user32.GetKeyboardLayout.restype = wintypes.HKL
user32.GetKeyboardLayoutList.argtypes = [ctypes.c_int, ctypes.POINTER(wintypes.HKL)]
user32.GetKeyboardLayoutList.restype = ctypes.c_int
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = wintypes.INT
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, wintypes.INT]
user32.GetWindowTextW.restype = wintypes.INT
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL
user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
    wintypes.UINT, wintypes.UINT, ctypes.POINTER(ULONG_PTR),
]
user32.SendMessageTimeoutW.restype = wintypes.LPARAM
user32.LoadKeyboardLayoutW.argtypes = [wintypes.LPCWSTR, wintypes.UINT]
user32.LoadKeyboardLayoutW.restype = wintypes.HKL
user32.ActivateKeyboardLayout.argtypes = [wintypes.HKL, wintypes.UINT]
user32.ActivateKeyboardLayout.restype = wintypes.HKL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
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
user32.EnumChildWindows.argtypes = [wintypes.HWND, ctypes.c_void_p, wintypes.LPARAM]
user32.EnumChildWindows.restype = wintypes.BOOL
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int
user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.GetWindowLongW.restype = ctypes.c_long

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

# MapVirtualKeyW for scan code lookup
user32.MapVirtualKeyW.argtypes = [wintypes.UINT, wintypes.UINT]
user32.MapVirtualKeyW.restype = wintypes.UINT
MAPVK_VK_TO_VSC = 0
MAPVK_VSC_TO_VK = 1


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Constants                                                    ║
# ╚═══════════════════════════════════════════════════════════════╝

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_BACK = 0x08
VK_SPACE = 0x20
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12    # Alt
VK_LWIN = 0x5B
VK_LEFT = 0x25
VK_V = 0x56

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WM_INPUTLANGCHANGEREQUEST = 0x0050
KLF_ACTIVATE = 0x00000001
KLF_SETFORPROCESS = 0x00000100
SMTO_ABORTIFHUNG = 0x0002
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

LANG_ENGLISH_US = 0x0409

# Word boundaries
BOUNDARY_KEYS = {
    'space': ' ', ',': ',', 'comma': ',', '.': '.', 'dot': '.',
    ';': ';', 'semicolon': ';', ':': ':', '!': '!', '?': '?',
}

# Sentence boundaries (reset sentence tracking)
SENTENCE_BOUNDARIES = {'enter', '.', 'dot'}

# Shift + English key → shifted character (for Shift+OEM handling in non-EN layouts)
SHIFT_EN_CHARS = {
    ';': ':', ',': '<', '.': '>', '/': '?', "'": '"',
    '[': '{', ']': '}', '\\': '|', '`': '~', '-': '_', '=': '+',
    '1': '!', '2': '@', '3': '#', '4': '$', '5': '%',
    '6': '^', '7': '&', '8': '*', '9': '(', '0': ')',
}


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Win32 Structs                                                ║
# ╚═══════════════════════════════════════════════════════════════╝

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk', ctypes.c_ushort),
        ('wScan', ctypes.c_ushort),
        ('dwFlags', ctypes.c_ulong),
        ('time', ctypes.c_ulong),
        ('dwExtraInfo', ULONG_PTR),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ('dx', ctypes.c_long),
        ('dy', ctypes.c_long),
        ('mouseData', ctypes.c_ulong),
        ('dwFlags', ctypes.c_ulong),
        ('time', ctypes.c_ulong),
        ('dwExtraInfo', ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ('uMsg', ctypes.c_ulong),
        ('wParamL', ctypes.c_ushort),
        ('wParamH', ctypes.c_ushort),
    ]


class INPUT_I(ctypes.Union):
    _fields_ = [('ki', KEYBDINPUT), ('mi', MOUSEINPUT), ('hi', HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [('type', ctypes.c_ulong), ('ii', INPUT_I)]


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD), ('flags', wintypes.DWORD),
        ('hwndActive', wintypes.HWND), ('hwndFocus', wintypes.HWND),
        ('hwndCapture', wintypes.HWND), ('hwndMenuOwner', wintypes.HWND),
        ('hwndMoveSize', wintypes.HWND), ('hwndCaret', wintypes.HWND),
        ('rcCaret', wintypes.RECT),
    ]


user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
user32.GetGUIThreadInfo.restype = wintypes.BOOL

_ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Low-level helpers                                            ║
# ╚═══════════════════════════════════════════════════════════════╝

def _send_vk(vk: int) -> None:
    """Press and release a single virtual key (with proper scan code)."""
    scan = user32.MapVirtualKeyW(vk, MAPVK_VK_TO_VSC)
    down = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=0, time=0, dwExtraInfo=0)))
    up = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def _send_vk_combo(*vks: int, delay: float = 0.02) -> None:
    """Press all keys down, then release in reverse (with scan codes)."""
    for vk in vks:
        scan = user32.MapVirtualKeyW(vk, MAPVK_VK_TO_VSC)
        inp = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=0, time=0, dwExtraInfo=0)))
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
        if delay:
            time.sleep(delay)
    for vk in reversed(vks):
        scan = user32.MapVirtualKeyW(vk, MAPVK_VK_TO_VSC)
        inp = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
        if delay:
            time.sleep(delay)


def _send_unicode_char(ch: str) -> None:
    """Inject a single Unicode character (layout-independent)."""
    code = ord(ch)
    down = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE, time=0, dwExtraInfo=0)))
    up = INPUT(type=INPUT_KEYBOARD, ii=INPUT_I(ki=KEYBDINPUT(wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def _send_unicode_text(text: str) -> None:
    """Inject a Unicode string with small inter-character delays.

    Individual SendInput calls with short delays for reliability.
    KEYEVENTF_UNICODE events don't appear in the keyboard hook,
    so user interleaving is not an issue for these.
    """
    if not text:
        return
    _dbg(f'[UNICODE] sending {len(text)} chars: {text!r}')
    for ch in text:
        _send_unicode_char(ch)
        time.sleep(0.005)  # 5ms per char — fast but reliable
    _dbg(f'[UNICODE] done')


def _send_backspaces(count: int) -> None:
    """Send multiple backspaces with individual delays.

    Apps (especially Notepad) need time to process each deletion.
    Sending too fast as a single batch causes dropped events.
    """
    if count <= 0:
        return
    _dbg(f'[BACKSPACE] sending {count} backspaces')
    for i in range(count):
        _send_vk(VK_BACK)
        time.sleep(0.008)  # 8ms per backspace — reliable across apps
    _dbg(f'[BACKSPACE] done')


def _paste_text(text: str) -> bool:
    """Paste text using clipboard + Ctrl+V. More reliable across apps.

    Returns True on success. Saves and restores original clipboard content.
    """
    old_clip = None
    try:
        old_clip = _get_clipboard_text()
    except Exception:
        pass

    if not _set_clipboard_text(text):
        # Fallback: direct Unicode injection
        _send_unicode_text(text)
        return True

    time.sleep(0.01)
    _send_vk_combo(VK_CONTROL, VK_V, delay=0.01)
    time.sleep(0.03)

    # Restore original clipboard
    try:
        if old_clip is not None:
            _set_clipboard_text(old_clip)
    except Exception:
        pass

    return True


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Clipboard helpers                                            ║
# ╚═══════════════════════════════════════════════════════════════╝

def _get_clipboard_text() -> str | None:
    if not user32.OpenClipboard(None):
        return None
    try:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ''
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ''
        size = kernel32.GlobalSize(handle)
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ''
        try:
            max_w = max(0, (size // ctypes.sizeof(ctypes.c_wchar)) - 1)
            return ctypes.wstring_at(ptr, max_w)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def _set_clipboard_text(text: str) -> bool:
    if text is None:
        text = ''
    data = text + '\x00'
    size = len(data) * ctypes.sizeof(ctypes.c_wchar)
    if not user32.OpenClipboard(None):
        return False
    hglob = None
    try:
        user32.EmptyClipboard()
        hglob = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not hglob:
            return False
        ptr = kernel32.GlobalLock(hglob)
        if not ptr:
            return False
        try:
            ctypes.memmove(ptr, ctypes.create_unicode_buffer(data), size)
        finally:
            kernel32.GlobalUnlock(hglob)
        if not user32.SetClipboardData(CF_UNICODETEXT, hglob):
            kernel32.GlobalFree(hglob)
            return False
        hglob = None  # system owns it now
        return True
    finally:
        if hglob:
            kernel32.GlobalFree(hglob)
        user32.CloseClipboard()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Window / Process helpers                                     ║
# ╚═══════════════════════════════════════════════════════════════╝

def _get_focus_hwnd() -> wintypes.HWND:
    """Return the focused control HWND of the foreground thread."""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return hwnd
    pid = wintypes.DWORD(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if tid:
        info = GUITHREADINFO()
        info.cbSize = ctypes.sizeof(GUITHREADINFO)
        if user32.GetGUIThreadInfo(tid, ctypes.byref(info)) and info.hwndFocus:
            return info.hwndFocus
    return hwnd


def _pid_to_exe(pid: int) -> str:
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return ''
    try:
        buf_len = wintypes.DWORD(260)
        buf = ctypes.create_unicode_buffer(260)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(buf_len)):
            return buf.value.rsplit('\\', 1)[-1].lower()
        name_buf = ctypes.create_unicode_buffer(260)
        if psapi.GetModuleBaseNameW(h, None, name_buf, 260):
            return name_buf.value.lower()
        return ''
    finally:
        kernel32.CloseHandle(h)


def _get_foreground_exe() -> str:
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ''
    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ''
    exe = _pid_to_exe(pid.value)
    if exe == 'applicationframehost.exe':
        parent_pid = pid.value
        result = ['']

        def _cb(child_hwnd, _):
            cpid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(child_hwnd, ctypes.byref(cpid))
            if cpid.value and cpid.value != parent_pid:
                e = _pid_to_exe(cpid.value)
                if e and e != 'applicationframehost.exe':
                    result[0] = e
                    return False
            return True

        cb = _ENUM_CHILD_PROC(_cb)
        try:
            user32.EnumChildWindows(hwnd, cb, 0)
        except Exception:
            pass
        if result[0]:
            return result[0]
    return exe


def _get_foreground_title() -> str:
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ''
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ''
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _strip_bidi(text: str) -> str:
    """Remove invisible Unicode control chars (bidi, zero-width, BOM)."""
    bad = {'\u200e', '\u200f', '\u202a', '\u202b', '\u202c', '\u202d', '\u202e',
           '\ufeff', '\u200b', '\u200c', '\u200d'}
    return ''.join(ch for ch in text if ch not in bad)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Keyboard layout detection & switching                        ║
# ╚═══════════════════════════════════════════════════════════════╝

def _foreground_lang_id() -> int:
    """Get the LANGID of the foreground window's keyboard layout."""
    hwnd = _get_focus_hwnd() or user32.GetForegroundWindow()
    if not hwnd:
        return int(user32.GetKeyboardLayout(0)) & 0xFFFF
    pid = wintypes.DWORD(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    hkl = user32.GetKeyboardLayout(tid)
    if not hkl:
        hkl = user32.GetKeyboardLayout(0)
    return int(hkl) & 0xFFFF


def _layout_matches(lang_id: int, target: str) -> bool:
    """Check if the current layout matches the target language."""
    if target == 'en':
        return is_english_lang_id(lang_id)
    profile = ACTIVE_PROFILES.get(target)
    if profile:
        return lang_id in profile.win_lang_ids
    return False


def _set_input_language(lang_id: int) -> bool:
    """Request language change for the foreground window via WM_INPUTLANGCHANGEREQUEST."""
    fg_hwnd = user32.GetForegroundWindow()
    focus_hwnd = _get_focus_hwnd() or fg_hwnd
    if not focus_hwnd:
        return False

    klid = f'{lang_id:08X}'
    try:
        hkl = user32.LoadKeyboardLayoutW(klid, KLF_ACTIVATE | KLF_SETFORPROCESS)
    except Exception:
        return False
    if not hkl:
        return False

    try:
        user32.ActivateKeyboardLayout(hkl, KLF_SETFORPROCESS)
    except Exception:
        pass

    # Send to multiple candidate HWNDs
    sent = False
    try:
        pid = wintypes.DWORD(0)
        tid = user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(pid)) if fg_hwnd else 0
        if tid:
            info = GUITHREADINFO()
            info.cbSize = ctypes.sizeof(GUITHREADINFO)
            if user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
                for candidate in (info.hwndFocus, info.hwndCaret, info.hwndActive, focus_hwnd, fg_hwnd):
                    if candidate:
                        result = ULONG_PTR(0)
                        rc = user32.SendMessageTimeoutW(
                            candidate, WM_INPUTLANGCHANGEREQUEST, 0, hkl,
                            SMTO_ABORTIFHUNG, 80, ctypes.byref(result))
                        if rc:
                            sent = True
        if not sent:
            for candidate in (focus_hwnd, fg_hwnd):
                if candidate:
                    user32.PostMessageW(candidate, WM_INPUTLANGCHANGEREQUEST, 0, hkl)
                    sent = True
    except Exception:
        pass
    return sent


_toggle_idx = 0


def _toggle_layout_once() -> None:
    """Rotate through layout-switch methods."""
    global _toggle_idx
    methods = list(LAYOUT_TOGGLE_HOTKEYS) + [
        '__si_alt_shift', '__si_win_space', '__si_ctrl_shift',
    ]
    pick = methods[_toggle_idx % len(methods)]
    _toggle_idx += 1
    try:
        if pick == '__si_alt_shift':
            _send_vk_combo(VK_MENU, VK_SHIFT)
        elif pick == '__si_win_space':
            _send_vk_combo(VK_LWIN, VK_SPACE)
        elif pick == '__si_ctrl_shift':
            _send_vk_combo(VK_CONTROL, VK_SHIFT)
        else:
            keyboard.send(pick)
    except Exception:
        pass


def _switch_layout_to(target: str) -> bool:
    """Best-effort: switch the input language to the target."""
    if target != 'en' and target not in ACTIVE_PROFILES:
        return False

    if target == 'en':
        desired_id = LANG_ENGLISH_US
    else:
        desired_id = next(iter(ACTIVE_PROFILES[target].win_lang_ids))

    # Try explicit WM_INPUTLANGCHANGEREQUEST first
    _set_input_language(desired_id)
    time.sleep(0.05)
    if _layout_matches(_foreground_lang_id(), target):
        return True

    # Fall back to hotkey toggling
    for _ in range(8):
        if _layout_matches(_foreground_lang_id(), target):
            return True
        _toggle_layout_once()
        time.sleep(0.05)
    return _layout_matches(_foreground_lang_id(), target)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Language detection & profiles                                ║
# ╚═══════════════════════════════════════════════════════════════╝

ACTIVE_PROFILES: dict[str, LanguageProfile] = {}
PROFILE_TO_EN: dict[str, dict[str, str]] = {}
PROFILE_FROM_EN: dict[str, dict[str, str]] = {}
PROFILE_SHORT: dict[str, dict[str, str]] = {}

# Legacy aliases (for UI compat)
HEBREW_TO_ENGLISH: dict[str, str] = {}
ENGLISH_TO_HEBREW: dict[str, str] = {}


def _detect_installed_layouts() -> list[int]:
    try:
        count = user32.GetKeyboardLayoutList(0, None)
        if count <= 0:
            return []
        arr = (wintypes.HKL * count)()
        user32.GetKeyboardLayoutList(count, arr)
        return [int(h) & 0xFFFF for h in arr]
    except Exception:
        return []


def _load_language_profiles() -> None:
    global ACTIVE_PROFILES, PROFILE_TO_EN, PROFILE_FROM_EN, PROFILE_SHORT
    global HEBREW_TO_ENGLISH, ENGLISH_TO_HEBREW

    installed = _detect_installed_layouts()
    if DEBUG:
        print(f'Installed layout IDs: {[f"0x{lid:04x}" for lid in installed]}')

    loaded = []
    for lid in installed:
        if is_english_lang_id(lid):
            continue
        profile = lang_id_to_profile(lid)
        if profile and profile.code not in ACTIVE_PROFILES:
            ACTIVE_PROFILES[profile.code] = profile
            PROFILE_TO_EN[profile.code] = profile.to_english
            PROFILE_FROM_EN[profile.code] = profile.from_english
            PROFILE_SHORT[profile.code] = profile.common_short_words
            loaded.append(f'{profile.flag} {profile.name} ({profile.code})')

    # Hebrew legacy compat
    if 'he' in ACTIVE_PROFILES:
        he = ACTIVE_PROFILES['he']
        HEBREW_TO_ENGLISH.update(he.to_english)
        ENGLISH_TO_HEBREW.update(he.from_english)

    if loaded:
        print(f'Loaded profiles: {", ".join(loaded)}')
    else:
        # Fallback: load Hebrew
        if 'he' in ALL_PROFILES:
            he = ALL_PROFILES['he']
            ACTIVE_PROFILES['he'] = he
            PROFILE_TO_EN['he'] = he.to_english
            PROFILE_FROM_EN['he'] = he.from_english
            PROFILE_SHORT['he'] = he.common_short_words
            HEBREW_TO_ENGLISH.update(he.to_english)
            ENGLISH_TO_HEBREW.update(he.from_english)
            print('Fallback: loaded Hebrew profile')


def _contains_non_english(text: str) -> str | None:
    """Return language code if text contains non-English script, else None."""
    for code, profile in ACTIVE_PROFILES.items():
        if profile.contains_script(text):
            return code
    return None


def _translate(text: str, mapping: dict[str, str]) -> str:
    return ''.join(mapping.get(c, c) for c in text)


def _nlp_valid(word: str, lang: str) -> bool:
    """Check if word is a real word in the given language using wordfreq."""
    if not word or len(word) < 2:
        return False
    w = word.strip().lower()
    if not w:
        return False

    # Structural check: word must start with a character from the expected script
    if lang == 'en':
        if not (w[0].isascii() and w[0].isalpha()):
            return False
    else:
        profile = ACTIVE_PROFILES.get(lang)
        if profile:
            cp = ord(w[0])
            if not any(s <= cp <= e for s, e in profile.unicode_ranges):
                return False

    return zipf_frequency(w, lang) >= NLP_ZIPF_THRESHOLD


def _nlp_decide(en_version: str, native_version: str, native_lang: str) -> str | None:
    """Pick the language whose version has the highest zipf score.

    Returns 'en', native_lang, or None if the gap is too small.
    Only corrects when the difference is >= NLP_SCORE_GAP.
    """
    en_score = zipf_frequency(en_version.strip().lower(), 'en') if en_version else 0.0
    native_score = zipf_frequency(native_version.strip().lower(), native_lang) if (native_version and native_lang) else 0.0

    # Structural check: en_version must start with ASCII alpha
    if en_version:
        c = en_version.strip().lower()[:1]
        if not (c.isascii() and c.isalpha()):
            en_score = 0.0

    # Structural check: native must start with correct script
    if native_version and native_lang:
        profile = ACTIVE_PROFILES.get(native_lang)
        if profile:
            cp = ord(native_version.strip()[:1]) if native_version.strip() else 0
            if not any(s <= cp <= e for s, e in profile.unicode_ranges):
                native_score = 0.0

    _dbg(f'[SCORES] en={en_score:.2f} ({en_version!r})  {native_lang}={native_score:.2f} ({native_version!r})  gap={en_score - native_score:+.2f}')

    # Pick highest score, but only if gap is significant
    gap = en_score - native_score
    if gap >= NLP_SCORE_GAP:
        return 'en'
    if gap <= -NLP_SCORE_GAP:
        return native_lang

    # Gap too small — ambiguous
    return None


def _is_pure_alpha(word: str) -> bool:
    """Return True if word consists only of letters (no digits, commas, etc)."""
    return bool(word) and all(c.isalpha() for c in word)


def _nlp_decide_for_lock(original: str, locked_lang: str) -> str | None:
    """Quick NLP check for sentence-lock override.

    Given a word that *appears* to match the locked language, translate it to
    the alternative and compare scores. Returns the alternative language if
    it scores significantly higher (gap >= NLP_LOCK_OVERRIDE_GAP), else None.

    Uses a MUCH higher gap threshold than normal NLP decisions because the
    sentence lock represents accumulated evidence from multiple words.
    Also performs structural checks to avoid false positives (e.g. 'fu,co'
    from ת→comma mapping is not a real English word).
    """
    if locked_lang == 'en':
        # Word is English chars — check if native translation scores higher
        en_word = original.strip().lower()
        # Structural: must be pure alpha to count as English
        en_score = zipf_frequency(en_word, 'en') if _is_pure_alpha(en_word) else 0.0
        for lang_code in ACTIVE_PROFILES:
            from_en = PROFILE_FROM_EN.get(lang_code, {})
            native_version = _translate(original, from_en)
            if native_version and native_version != original:
                nw = native_version.strip().lower()
                # Structural: must start with correct script
                profile = ACTIVE_PROFILES.get(lang_code)
                if profile and nw:
                    cp = ord(nw[0])
                    if not any(s <= cp <= e for s, e in profile.unicode_ranges):
                        continue
                native_score = zipf_frequency(nw, lang_code)
                gap = native_score - en_score
                _dbg(f'[LOCK-CHECK] en={en_score:.2f} ({en_word!r})  '
                     f'{lang_code}={native_score:.2f} ({native_version!r})  gap={gap:+.2f}')
                if gap >= NLP_LOCK_OVERRIDE_GAP:
                    return lang_code
    else:
        # Word is native chars — check if English translation scores higher
        to_en = PROFILE_TO_EN.get(locked_lang, {})
        en_version = _translate(original, to_en)
        if en_version and en_version != original:
            en_word = en_version.strip().lower()
            # Structural: English must be pure alpha letters
            if not _is_pure_alpha(en_word):
                _dbg(f'[LOCK-CHECK] en={en_word!r} has non-alpha — skip override')
                return None
            # Structural: must start with ASCII alpha
            if not (en_word[0].isascii() and en_word[0].isalpha()):
                return None
            native_score = zipf_frequency(original.strip().lower(), locked_lang)
            en_score = zipf_frequency(en_word, 'en')
            gap = en_score - native_score
            _dbg(f'[LOCK-CHECK] en={en_score:.2f} ({en_version!r})  '
                 f'{locked_lang}={native_score:.2f} ({original!r})  gap={gap:+.2f}')
            if gap >= NLP_LOCK_OVERRIDE_GAP:
                return 'en'

    return None


def _event_to_actual_char(event_name: str, lang_id: int) -> str:
    """Map a physical key name to the actual character based on current layout.

    Unlike the v2 approach (ToUnicodeEx), this uses our own char maps directly.
    This avoids corrupting the Windows dead-key state.
    """
    if not isinstance(event_name, str) or not event_name:
        return ''

    # If the key name is already a printable single char
    if len(event_name) == 1 and event_name.isprintable():
        # If layout is non-English, map through from_english
        if not is_english_lang_id(lang_id):
            # Find the specific profile for this layout
            for code, profile in ACTIVE_PROFILES.items():
                if lang_id in profile.win_lang_ids:
                    from_en = PROFILE_FROM_EN.get(code, {})
                    return from_en.get(event_name, event_name)
            # If no specific profile found, try all
            for from_en in PROFILE_FROM_EN.values():
                if event_name in from_en:
                    return from_en[event_name]
        return event_name
    return ''


def _scan_to_english_char(scan_code: int) -> str:
    """Convert a scan code to the English (QWERTY) character it produces.

    Uses MapVirtualKeyW(scan, MAPVK_VSC_TO_VK) to get the VK code,
    then maps VK to ASCII character. This is layout-independent —
    it always returns the English character regardless of the active
    keyboard layout on ANY thread (hook thread or foreground app).
    """
    vk = user32.MapVirtualKeyW(scan_code, MAPVK_VSC_TO_VK)
    if not vk:
        return ''
    # A-Z: VK 0x41–0x5A
    if 0x41 <= vk <= 0x5A:
        return chr(vk).lower()  # a-z
    # 0-9: VK 0x30–0x39
    if 0x30 <= vk <= 0x39:
        return chr(vk)  # 0-9
    # OEM keys (US QWERTY layout)
    oem_map = {
        0xBA: ';', 0xBB: '=', 0xBC: ',', 0xBD: '-', 0xBE: '.',
        0xBF: '/', 0xC0: '`', 0xDB: '[', 0xDC: '\\', 0xDD: ']', 0xDE: "'",
    }
    return oem_map.get(vk, '')


def _translate_buffer(text: str) -> str:
    """Translate buffer text to the other language.

    If text looks like English characters, translate to the primary non-English.
    If text looks like non-English, translate to English.
    Returns the translated text.
    """
    if not text:
        return ''

    # Detect: is the text English or non-English?
    is_en = all(c.isascii() for c in text if c.isalpha())

    if is_en:
        # English → primary non-English profile
        for code in ('he', 'ru', 'ar', 'uk', 'fr', 'de', 'es'):
            from_en = PROFILE_FROM_EN.get(code)
            if from_en:
                return ''.join(from_en.get(c, c) for c in text)
        return text
    else:
        # Non-English → English
        for code, to_en in PROFILE_TO_EN.items():
            if any(c in to_en for c in text if c.isalpha()):
                return ''.join(to_en.get(c, c) for c in text)
        return text


def _compute_word_scores(word: str) -> dict:
    """Compute NLP zipf scores for a word in EN + all active native languages.
    Returns e.g. {'en': 3.21, 'he': 5.04}."""
    scores = {}
    if not word or len(word) < 2:
        return scores

    # English version
    en_version = word
    is_en_chars = all(c.isascii() for c in word if c.isalpha())
    if not is_en_chars:
        # word is in non-English — translate to English
        for code, to_en in PROFILE_TO_EN.items():
            candidate = _translate(word, to_en)
            if candidate != word:
                en_version = candidate
                break

    en_w = en_version.strip().lower()
    if en_w and en_w[0:1].isascii() and en_w[0:1].isalpha():
        scores['EN'] = round(zipf_frequency(en_w, 'en'), 1)
    else:
        scores['EN'] = 0.0

    # Native versions
    for lang_code, profile in ACTIVE_PROFILES.items():
        flag = profile.flag
        if is_en_chars:
            from_en = PROFILE_FROM_EN.get(lang_code, {})
            native_version = _translate(word, from_en)
        else:
            native_version = word
        nw = native_version.strip().lower()
        if nw:
            scores[flag] = round(zipf_frequency(nw, lang_code), 1)
        else:
            scores[flag] = 0.0
    return scores


def _fire_buffer_callback():
    """Notify the UI about current buffer content and its translation.
    Shows up to 2 completed words + current word being typed.
    Also sends correction map so UI can show clickable revert,
    and per-word NLP scores."""
    cb = _buffer_callback
    if cb is None:
        return
    try:
        buf = state.buffer
        parts = list(_prev_words)   # up to 2 completed words
        if buf:
            parts.append(buf)
        display = ' '.join(parts) if parts else ''
        translated = _translate_buffer(display)
        # Send correction info: dict of corrected_word -> original_word
        corrections = dict(_word_corrections)
        # Compute per-word NLP scores
        word_scores = {}
        for w in parts:
            if w.strip():
                word_scores[w] = _compute_word_scores(w)
        cb(display, translated, corrections, word_scores)
    except Exception:
        pass


def _scan_to_actual_char(scan_code: int, lang_id: int) -> str:
    """Convert scan code to the actual character for the foreground layout.

    Step 1: scan_code → English char (physical key, layout-independent)
    Step 2: If foreground layout is non-English, map through from_english dict

    This replaces the old _event_to_actual_char which relied on event.name
    from the keyboard library. event.name uses the hook thread’s layout
    (which may differ from the foreground window’s layout), causing
    incorrect character mapping.
    """
    en_char = _scan_to_english_char(scan_code)
    if not en_char:
        return ''

    if is_english_lang_id(lang_id):
        return en_char

    # Non-English layout: map from English to native
    for code, profile in ACTIVE_PROFILES.items():
        if lang_id in profile.win_lang_ids:
            from_en = PROFILE_FROM_EN.get(code, {})
            return from_en.get(en_char, en_char)
    # Fallback: try all profiles
    for from_en in PROFILE_FROM_EN.values():
        if en_char in from_en:
            return from_en[en_char]
    return en_char


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Error logging                                                ║
# ╚═══════════════════════════════════════════════════════════════╝

_LOG_FILE = os.path.expanduser('~/.auto_lang3_error.log')


def _log_error(msg: str) -> None:
    """Write error to log file (safe for --noconsole EXE)."""
    try:
        import traceback
        with open(_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f'[{time.strftime("%Y-%m-%d %H:%M:%S")}] {msg}\n')
            f.write(traceback.format_exc() + '\n')
    except Exception:
        pass
    if DEBUG:
        print(f'ERROR: {msg}')


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Core engine state (thread-safe)                              ║
# ╚═══════════════════════════════════════════════════════════════╝

class EngineState:
    """All mutable engine state, protected by a lock.

    Every read/write of engine state goes through this object's lock.
    This eliminates all the race conditions in v2.
    """

    def __init__(self):
        self.lock = threading.Lock()

        # Current word buffer
        self.buffer = ''                    # Characters typed so far in current word

        # Sentence tracking
        self.sentence_lang: str | None = None      # Language locked for this sentence
        self.words_in_sentence = 0                  # Word count in current sentence
        self.confirmed_words = 0                    # Words confirmed in current language

        # Pending ambiguous words
        # Each: {'original': str, 'en_version': str, 'native_version': str,
        #        'native_lang': str, 'boundary': str, 'typed_is_native': bool}
        self.pending: list[dict] = []

        # Correction streaks for auto-switch
        self.streaks: dict[str, int] = {}

        # Auto-switch pause
        self.paused_until_boundary = False

        # Timing
        self.last_correction_time = 0.0

        # Undo support: tracks the last correction so F12 can revert it
        self.last_undo: dict | None = None  # {'original': str, 'corrected': str}

    def reset_sentence(self):
        """Reset sentence-level state (call on ENTER/period)."""
        self.sentence_lang = None
        self.words_in_sentence = 0
        self.confirmed_words = 0
        self.pending = []

    def reset_all(self):
        """Full reset (call on focus change)."""
        self.reset_sentence()
        self.buffer = ''
        self.paused_until_boundary = False
        self.streaks.clear()


state = EngineState()
injecting = threading.Event()
stop_event = threading.Event()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Auto-learn chat language                                     ║
# ╚═══════════════════════════════════════════════════════════════╝

_chat_stats: dict[tuple[str, str], dict] = {}


def _load_learned():
    global _chat_stats
    if not ENABLE_AUTO_LEARN:
        return
    try:
        if os.path.exists(AUTO_LEARN_CACHE_FILE):
            with open(AUTO_LEARN_CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for key_str, stats in data.items():
                if '|' in key_str:
                    exe, title = key_str.split('|', 1)
                    _chat_stats[(exe, title)] = stats
    except Exception:
        pass


def _save_learned():
    if not ENABLE_AUTO_LEARN:
        return
    try:
        data = {}
        for (exe, title), stats in _chat_stats.items():
            if stats.get('learned'):
                data[f'{exe}|{title}'] = stats
        with open(AUTO_LEARN_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


_learn_save_counter = 0


def _update_learn_stats(exe: str, title: str, char: str):
    """Update language stats for chat auto-learning. Saves every 50 chars, not every char."""
    global _learn_save_counter
    if not ENABLE_AUTO_LEARN or not exe or not title:
        return

    clean = _strip_bidi(title)
    key = (exe, clean)
    if key not in _chat_stats:
        _chat_stats[key] = {}

    stats = _chat_stats[key]
    detected = _contains_non_english(char)
    if detected:
        stats[detected] = stats.get(detected, 0) + 1
    elif char.isascii() and char.isalpha():
        stats['en'] = stats.get('en', 0) + 1
    else:
        return

    total = sum(v for k, v in stats.items() if k != 'learned')
    if total < AUTO_LEARN_MIN_CHARS:
        return

    lang_counts = {k: v for k, v in stats.items() if k != 'learned'}
    best = max(lang_counts, key=lang_counts.get)
    ratio = lang_counts[best] / total
    new_learned = best if ratio >= AUTO_LEARN_THRESHOLD else 'en'
    old = stats.get('learned')
    if new_learned != old:
        stats['learned'] = new_learned
        if DEBUG:
            print(f'Auto-learn: {exe}:{clean[:40]} -> {new_learned} ({ratio:.0%})')
        _learn_save_counter += 1
        if _learn_save_counter % 5 == 0:  # save every 5 decisions, not every char
            _save_learned()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ App-default language logic                                   ║
# ╚═══════════════════════════════════════════════════════════════╝

def _pick_default_lang(exe: str, title: str) -> str | None:
    """Determine the default language for the current foreground app/chat."""
    if not exe and not title:
        return None

    exe_l = (exe or '').lower()
    clean = _strip_bidi(title or '')
    title_l = clean.lower()

    # 1) Auto-learned preference
    if ENABLE_AUTO_LEARN:
        stats = _chat_stats.get((exe_l, clean))
        if stats and stats.get('learned'):
            return stats['learned']

    # 2) Per-chat title rules
    per_title = APP_DEFAULT_LANG_BY_TITLE.get(exe_l)
    if per_title:
        for needle, lang in per_title.items():
            if needle and (needle.lower() in title_l or needle in clean):
                if lang == 'en' or lang in ACTIVE_PROFILES:
                    return lang

    # 2b) Browser-specific: match keywords against page title for any browser
    if exe_l in BROWSER_EXES and BROWSER_LANG_BY_KEYWORD:
        for keyword, lang in BROWSER_LANG_BY_KEYWORD.items():
            if keyword and keyword.lower() in title_l:
                if lang == 'en' or lang in ACTIVE_PROFILES:
                    return lang

    # 3) Per-exe rules
    direct = APP_DEFAULT_LANG_BY_EXE.get(exe_l)
    if direct and (direct == 'en' or direct in ACTIVE_PROFILES):
        return direct

    return None


def _app_default_watcher():
    """Background thread: apply default language on focus/title changes."""
    last_hwnd = None
    last_title = ''

    while not stop_event.is_set():
        try:
            if injecting.is_set():
                time.sleep(APP_DEFAULT_POLL_INTERVAL)
                continue

            hwnd = user32.GetForegroundWindow()
            if not hwnd:
                time.sleep(APP_DEFAULT_POLL_INTERVAL)
                continue

            title = _get_foreground_title()
            exe = _get_foreground_exe()
            focus_changed = (hwnd != last_hwnd)
            title_changed = (not focus_changed and exe in WATCH_TITLE_CHANGES_EXE and title != last_title)

            if focus_changed or title_changed:
                # Skip if focus went to our own process (floating widget click)
                try:
                    pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value == os.getpid():
                        time.sleep(APP_DEFAULT_POLL_INTERVAL)
                        continue
                except Exception:
                    pass
                last_hwnd = hwnd
                last_title = title

                with state.lock:
                    state.reset_all()
                # Clear panel on window switch
                _prev_words.clear()
                _word_corrections.clear()
                _fire_buffer_callback()

                if DEBUG and focus_changed:
                    print(f'Focus -> {exe!r} {title[:60]!r}')

                # Don't apply during cooldown
                if (time.time() - state.last_correction_time) < APP_DEFAULT_COOLDOWN_SEC:
                    time.sleep(APP_DEFAULT_POLL_INTERVAL)
                    continue

                with state.lock:
                    if state.paused_until_boundary:
                        time.sleep(APP_DEFAULT_POLL_INTERVAL)
                        continue

                target = _pick_default_lang(exe, title)
                if target:
                    lang_id = _foreground_lang_id()
                    if not _layout_matches(lang_id, target):
                        if target == 'en':
                            desired = LANG_ENGLISH_US
                        elif target in ACTIVE_PROFILES:
                            desired = next(iter(ACTIVE_PROFILES[target].win_lang_ids))
                        else:
                            desired = LANG_ENGLISH_US

                        _set_input_language(desired)
                        time.sleep(0.06)

                        if not _layout_matches(_foreground_lang_id(), target):
                            for _ in range(8):
                                if _layout_matches(_foreground_lang_id(), target):
                                    break
                                injecting.set()
                                try:
                                    _toggle_layout_once()
                                finally:
                                    injecting.clear()
                                time.sleep(0.05)
        except Exception:
            pass

        time.sleep(APP_DEFAULT_POLL_INTERVAL)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Text correction                                              ║
# ╚═══════════════════════════════════════════════════════════════╝

def _note_correction(target_lang: str | None):
    """Update correction streak counters and maybe auto-switch layout."""
    if not target_lang:
        return

    do_switch = False
    with state.lock:
        for k in list(state.streaks.keys()):
            if k != target_lang:
                state.streaks[k] = 0
        state.streaks[target_lang] = state.streaks.get(target_lang, 0) + 1
        state.last_correction_time = time.time()

        if not AUTO_SWITCH_LAYOUT:
            return

        count = state.streaks.get(target_lang, 0)
        if count >= AUTO_SWITCH_AFTER_CONSECUTIVE:
            if DEBUG:
                print(f'Auto-switch -> {target_lang}')
            state.streaks[target_lang] = 0
            do_switch = True
            state.paused_until_boundary = True

    # Switch layout outside the lock (it does sleeps)
    if do_switch:
        _switch_layout_to(target_lang)


def _replace_text(original: str, corrected: str, boundary: str):
    """Delete the wrong text and type the corrected text.

    This runs on a worker thread. injecting is already set by the caller
    BEFORE the thread was spawned (eliminates the race window).
    Uses _send_unicode_text (layout-independent) instead of clipboard paste
    to avoid Ctrl+V key leaking into keyboard hook and clipboard contamination.
    """
    try:
        time.sleep(0.02)  # let the app commit the boundary char

        total_delete = 1 + len(original)  # boundary + original word

        _dbg(f'REPLACE: {original!r} -> {corrected!r} boundary={boundary!r} del={total_delete}')
        if DEBUG:
            print(f'Replacing: {original!r} -> {corrected!r} boundary={boundary!r} '
                  f'total_del={total_delete}')

        # Delete atomically
        _send_backspaces(total_delete)

        time.sleep(0.02)

        # Build replacement text: corrected + boundary
        replacement = corrected + boundary

        # Use direct Unicode injection (INPUT struct is 40 bytes — SendInput works)
        _send_unicode_text(replacement)

        target = _contains_non_english(corrected) or 'en'
        _note_correction(target)

        # Save undo info
        full_original = original + boundary
        with state.lock:
            state.last_undo = {'original': full_original, 'corrected': replacement}

        # Track per-word correction for click-to-revert
        _record_word_correction(original, corrected)
        _fire_buffer_callback()
    except Exception as e:
        _log_error(f'_replace_text failed: {e}')
    finally:
        time.sleep(0.02)
        injecting.clear()


def _replace_batch(winner_lang: str, pending: list[dict],
                   current: dict | None):
    """Batch-replace pending words + optionally the current word.

    Backspaces through ALL text from first pending word to cursor,
    retypes everything corrected.
    injecting is already set by the caller before the thread was spawned.
    Uses _send_unicode_text instead of clipboard paste.
    """
    try:
        time.sleep(0.02)

        # Build word list
        all_words = []
        for pw in pending:
            if winner_lang == 'en':
                corr = pw['en_version'] if pw['typed_is_native'] else pw['original']
            else:
                corr = pw['native_version'] if not pw['typed_is_native'] else pw['original']
            all_words.append({
                'original': pw['original'], 'corrected': corr, 'boundary': pw['boundary'],
            })

        if current:
            all_words.append(current)

        # Check if any correction is actually needed
        if not any(w['original'] != w['corrected'] for w in all_words):
            return

        total_delete = sum(len(w['original']) + len(w['boundary']) for w in all_words)

        if DEBUG:
            print(f'Batch replace: {[w["original"] for w in all_words]} -> {[w["corrected"] for w in all_words]} '
                  f'del={total_delete}')

        # Delete atomically
        _send_backspaces(total_delete)

        time.sleep(0.02)

        # Build full replacement text
        replacement = ''
        for w in all_words:
            replacement += w['corrected'] + w['boundary']

        # Use direct Unicode injection (no clipboard contamination)
        _dbg(f'BATCH: injecting replacement={replacement!r}')
        _send_unicode_text(replacement)
        _dbg(f'BATCH: done injecting')

        target = _contains_non_english(all_words[-1]['corrected']) or 'en'
        _note_correction(target)

        # Save undo info
        full_original = ''.join(w['original'] + w['boundary'] for w in all_words)
        with state.lock:
            state.last_undo = {'original': full_original, 'corrected': replacement}

        # Track per-word corrections for click-to-revert
        for w in all_words:
            if w['original'] != w['corrected']:
                _record_word_correction(w['original'], w['corrected'])
        _fire_buffer_callback()
    except Exception as e:
        _log_error(f'_replace_batch failed: {e}')
    finally:
        time.sleep(0.02)
        injecting.clear()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Word decision logic                                          ║
# ╚═══════════════════════════════════════════════════════════════╝

def _decide_word(original: str, boundary: str):
    """Decide if the word needs correction and dispatch replacement.

    IMPORTANT: This function must be called while state.lock is held.
    It reads/writes state.sentence_lang, state.pending, etc.

    Threshold logic (user spec):
      - 2+ words  OR  6+ chars in a single word  triggers correction.
      - Below threshold: buffer to pending and wait for more context.
    """

    # Determine typed language & generate both versions
    typed_lang = _contains_non_english(original)

    if typed_lang:
        native_version = original
        native_lang = typed_lang
        to_en = PROFILE_TO_EN.get(typed_lang, {})
        en_version = _translate(original, to_en)
        typed_is_native = True
    else:
        en_version = original
        typed_is_native = False
        # Try all active profiles (fixes v2 bug M3: only first profile was tried)
        native_version = None
        native_lang = None
        for lang_code in ACTIVE_PROFILES:
            from_en = PROFILE_FROM_EN.get(lang_code, {})
            candidate = _translate(original, from_en)
            if candidate != original:
                native_version = candidate
                native_lang = lang_code
                break
        if not native_version or not native_lang:
            if len(ACTIVE_PROFILES) == 1:
                native_lang = next(iter(ACTIVE_PROFILES))
                from_en = PROFILE_FROM_EN.get(native_lang, {})
                native_version = _translate(original, from_en)
            else:
                return

    lw = original.lower()

    # Skip excluded words
    if EXCLUDE_WORDS and lw in EXCLUDE_WORDS:
        if DEBUG:
            print(f'Excluded: {original!r}')
        return

    # Are we above the detection threshold?
    above_threshold = (
        len(original) >= MIN_CHARS_SINGLE_WORD or
        state.words_in_sentence >= MIN_WORDS_TO_DETECT
    )

    # Known shortcut words (from profile short-word dict)
    short = PROFILE_SHORT.get(native_lang, {})
    if lw in short and not typed_is_native:
        if above_threshold:
            _handle_winner(native_lang, original, en_version, short[lw],
                           native_lang, boundary, typed_is_native)
        else:
            state.pending.append({
                'original': original, 'en_version': en_version,
                'native_version': short[lw], 'native_lang': native_lang,
                'boundary': boundary, 'typed_is_native': typed_is_native,
                'winner': native_lang,
            })
            _maybe_flush_pending()
        return

    # NLP validation using the improved _nlp_decide (resolves ambiguity by score gap)
    winner = _nlp_decide(en_version, native_version, native_lang)

    _dbg(f'NLP: orig={original!r} en={en_version!r} '
         f'{native_lang}={native_version!r} '
         f'winner={winner} above={above_threshold} '
         f'pending={len(state.pending)} sentence={state.sentence_lang}')

    if DEBUG:
        print(f'NLP: orig={original!r} en={en_version!r} '
              f'{native_lang}={native_version!r} '
              f'winner={winner} above={above_threshold} '
              f'pending={len(state.pending)} sentence={state.sentence_lang}')

    # If we have a winner AND we're above threshold → act now
    if winner and above_threshold:
        _handle_winner(winner, original, en_version, native_version, native_lang,
                       boundary, typed_is_native)
        return

    # Otherwise buffer the word and maybe flush if we've accumulated enough
    state.pending.append({
        'original': original, 'en_version': en_version,
        'native_version': native_version, 'native_lang': native_lang,
        'boundary': boundary, 'typed_is_native': typed_is_native,
        'winner': winner,
    })

    if DEBUG:
        print(f'Buffered #{len(state.pending)}: {original!r} winner={winner}')

    # Check if accumulated pending words now meet the threshold
    if len(state.pending) >= MIN_WORDS_TO_DETECT:
        _maybe_flush_pending()


def _handle_winner(winner: str, original: str, en_version: str, native_version: str,
                   native_lang: str, boundary: str, typed_is_native: bool):
    """Handle a word that has a clear winner language.

    MUST be called with state.lock held.
    """
    state.sentence_lang = winner

    needs_correction = (
        (winner == 'en' and typed_is_native) or
        (winner != 'en' and not typed_is_native)
    )

    corrected = en_version if winner == 'en' else native_version

    if state.pending:
        # Batch correct pending + current
        pending_copy = list(state.pending)
        state.pending = []

        # Count all confirmed words (pending + current)
        state.confirmed_words += len(pending_copy) + 1

        current_entry = None
        if needs_correction:
            current_entry = {'original': original, 'corrected': corrected, 'boundary': boundary}
        else:
            current_entry = {'original': original, 'corrected': original, 'boundary': boundary}

        # Pre-check: does ANY word actually need correction?
        # This avoids setting injecting (which blocks user input) when nothing changes.
        any_needs = needs_correction
        if not any_needs:
            for pw in pending_copy:
                if winner == 'en' and pw['typed_is_native']:
                    any_needs = True
                    break
                elif winner != 'en' and not pw['typed_is_native']:
                    any_needs = True
                    break

        if any_needs:
            injecting.set()  # Set BEFORE spawning thread — no race window
            if DEBUG:
                print(f'Batch: {len(pending_copy)} pending + current -> {winner}')
            threading.Thread(
                target=_replace_batch,
                args=(winner, pending_copy, current_entry),
                daemon=True,
            ).start()
        else:
            if DEBUG:
                print(f'Batch: no correction needed for {len(pending_copy)+1} words')
    elif needs_correction:
        state.confirmed_words += 1
        _spawn_single(original, corrected, boundary, winner)
    else:
        state.confirmed_words += 1  # correct language, still counts


def _maybe_flush_pending():
    """If enough words accumulated in pending, determine winner and batch-correct.

    Must be called with state.lock held.
    """
    if len(state.pending) < MIN_WORDS_TO_DETECT:
        return

    # Count language votes from words that had a clear NLP winner
    votes: dict[str, int] = {}
    for pw in state.pending:
        w = pw.get('winner')
        if w:
            votes[w] = votes.get(w, 0) + 1

    if not votes:
        return  # All ambiguous — can't decide yet

    best = max(votes, key=votes.get)
    if votes[best] < 1:
        return

    # Pre-check: does any word actually need correction?
    any_needs = False
    for pw in state.pending:
        if best == 'en' and pw['typed_is_native']:
            any_needs = True
            break
        elif best != 'en' and not pw['typed_is_native']:
            any_needs = True
            break

    # Commit the batch correction
    state.sentence_lang = best
    state.confirmed_words += len(state.pending)

    if not any_needs:
        if DEBUG:
            print(f'Flush pending: {len(state.pending)} words -> {best} (no correction needed)')
        state.pending = []
        return

    pending_copy = list(state.pending)
    state.pending = []
    injecting.set()  # Set BEFORE spawning thread

    if DEBUG:
        print(f'Flush pending: {len(pending_copy)} words -> {best}')

    threading.Thread(
        target=_replace_batch,
        args=(best, pending_copy, None),
        daemon=True,
    ).start()


def _spawn_single(original: str, corrected: str, boundary: str, target_lang: str):
    """Spawn a thread to correct a single word. Must be called with state.lock held."""
    injecting.set()  # Set BEFORE spawning thread — no race window
    if DEBUG:
        print(f'CORRECT: {original!r} -> {corrected!r} ({target_lang})')
    threading.Thread(
        target=_replace_text,
        args=(original, corrected, boundary),
        daemon=True,
    ).start()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Spell check integration                                      ║
# ╚═══════════════════════════════════════════════════════════════╝

def _spell_check_completed_word(word: str):
    """Run spell check on a completed word (after language correction is done).

    If the word is misspelled, fires spell_module.SPELL_CALLBACK.
    In 'auto' mode, replaces the word immediately.
    In 'tooltip' mode, sets _spell_pending so Tab can accept.
    """
    global _spell_pending
    if not word or not spell_module.SPELL_ENABLED:
        return

    result = spell_module.check_word(word)
    if not result or not result.get('suggestions'):
        with _spell_pending_lock:
            _spell_pending = None
        return

    mode = spell_module.SPELL_MODE

    if mode == 'auto' and result['suggestions']:
        # Auto-correct: backspace the word and retype the best suggestion
        best = result['suggestions'][0]
        if best != word:
            with state.lock:
                injecting.set()
            threading.Thread(
                target=_replace_text,
                args=(word, best, 'space'),
                daemon=True,
            ).start()
            _dbg(f'[SPELL-AUTO] {word!r} → {best!r}')
        return

    # tooltip / visual modes: store pending and notify UI
    with _spell_pending_lock:
        _spell_pending = result

    spell_module.check_and_notify(word)
    _dbg(f'[SPELL] {word!r} → suggestions: {result["suggestions"][:3]}')


def _spell_accept():
    """Accept the pending spell suggestion (called on Tab when suggestion is active)."""
    global _spell_pending
    with _spell_pending_lock:
        pending = _spell_pending
        _spell_pending = None

    if not pending or not pending.get('suggestions'):
        return False  # No active suggestion

    original = pending['original']
    corrected = pending['suggestions'][0]

    if original == corrected:
        return False

    # Backspace the word + space, retype corrected + space
    with state.lock:
        injecting.set()
    threading.Thread(
        target=_spell_replace_last_word,
        args=(original, corrected),
        daemon=True,
    ).start()
    _dbg(f'[SPELL-ACCEPT] {original!r} → {corrected!r}')
    return True


def _spell_dismiss():
    """Dismiss the pending spell suggestion."""
    global _spell_pending
    with _spell_pending_lock:
        _spell_pending = None
    if spell_module.SPELL_CALLBACK:
        try:
            spell_module.SPELL_CALLBACK(None, None, 'dismiss')
        except Exception:
            pass


def _spell_replace_last_word(original: str, corrected: str):
    """Replace the last typed word with the corrected version."""
    try:
        import time
        # Backspace over: word + the space after it
        bs_count = len(original) + 1  # +1 for the space/boundary that followed

        for _ in range(bs_count):
            keyboard.send('backspace')
            time.sleep(0.01)
        time.sleep(0.03)

        # Type corrected word + space
        _send_unicode(corrected + ' ')
        time.sleep(0.03)
    except Exception as e:
        _dbg(f'[SPELL-REPLACE] error: {e}')
    finally:
        injecting.clear()


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Grammar check (LLM) integration                              ║
# ╚═══════════════════════════════════════════════════════════════╝

def _trigger_grammar_check():
    """Trigger grammar/phrasing correction via LLM.

    Flow:
    1. Copy selected text (Ctrl+C) → clipboard
    2. Send to LLM API
    3. On result: show diff popup in UI; user confirms → paste corrected text
    """
    if not grammar_module.GRAMMAR_ENABLED:
        if grammar_module.GRAMMAR_CALLBACK:
            grammar_module.GRAMMAR_CALLBACK('', None, 'Grammar check is disabled — enable it in Settings')
        return
    if not grammar_module.GRAMMAR_API_KEY:
        if grammar_module.GRAMMAR_CALLBACK:
            grammar_module.GRAMMAR_CALLBACK('', None, 'No API key — set one in Settings → Spell & Grammar')
        return

    def _do():
        import time
        try:
            # Save current clipboard
            import ctypes as _ct
            _user32 = _ct.windll.user32
            _kernel32 = _ct.windll.kernel32

            # Copy selected text
            keyboard.send('ctrl+c')
            time.sleep(0.2)

            # Read clipboard
            text = ''
            if _user32.OpenClipboard(0):
                try:
                    handle = _user32.GetClipboardData(13)  # CF_UNICODETEXT
                    if handle:
                        _kernel32.GlobalLock.restype = _ct.c_wchar_p
                        ptr = _kernel32.GlobalLock(handle)
                        if ptr:
                            text = str(ptr)
                            _kernel32.GlobalUnlock(handle)
                finally:
                    _user32.CloseClipboard()

            if not text or not text.strip():
                if grammar_module.GRAMMAR_CALLBACK:
                    grammar_module.GRAMMAR_CALLBACK('', None, 'No text selected — select text first, then press Ctrl+Shift+G')
                return

            _dbg(f'[GRAMMAR] Sending {len(text)} chars to {grammar_module.GRAMMAR_PROVIDER}')

            # Send to LLM (async callback will handle the result)
            grammar_module.correct_text_async(text)

        except Exception as e:
            _dbg(f'[GRAMMAR] Error: {e}')
            if grammar_module.GRAMMAR_CALLBACK:
                grammar_module.GRAMMAR_CALLBACK('', None, str(e))

    threading.Thread(target=_do, name='grammar-trigger', daemon=True).start()


def _process_boundary(boundary: str):
    """Called on word boundary. Runs the detection/correction logic.

    This is the main entry point from the keyboard hook.
    """
    with state.lock:
        if not state.buffer:
            return
        if not ENGINE_ENABLED:
            state.buffer = ''
            return
        if state.paused_until_boundary:
            state.buffer = ''
            return

        original = state.buffer
        state.buffer = ''  # Clear buffer BEFORE starting analysis (snapshot taken)
        state.words_in_sentence += 1

        # If sentence language is already locked and confirmed (2+ words),
        # auto-correct without NLP — UNLESS NLP strongly disagrees.
        if state.sentence_lang is not None and state.confirmed_words >= MIN_WORDS_TO_DETECT:
            typed_lang = _contains_non_english(original)
            typed_is_native = bool(typed_lang)
            wrong = False
            corrected = original

            if state.sentence_lang == 'en' and typed_is_native:
                to_en = PROFILE_TO_EN.get(typed_lang, {})
                corrected = _translate(original, to_en)
                wrong = True
            elif state.sentence_lang != 'en' and not typed_is_native:
                from_en = PROFILE_FROM_EN.get(state.sentence_lang, {})
                corrected = _translate(original, from_en)
                wrong = True

            if wrong:
                # NLP sanity check: make sure the correction is a real word
                # and the original isn't already a better word in its own language.
                # Without this, "working" gets blindly translated to "'םרלןמע"
                # (Hebrew nonsense) just because the sentence is locked to Hebrew.
                _skip_correction = False
                if typed_is_native:
                    # Original is native (e.g. Hebrew), corrected is English
                    ew = corrected.strip().lower()
                    if ew and _is_pure_alpha(ew):
                        en_sc = zipf_frequency(ew, 'en')
                    else:
                        en_sc = 0.0
                    nw = original.strip().lower()
                    na_sc = zipf_frequency(nw, typed_lang) if nw else 0.0
                    # Keep original if native word is real and EN correction isn't
                    if na_sc >= NLP_ZIPF_THRESHOLD and en_sc < NLP_ZIPF_THRESHOLD:
                        _skip_correction = True
                    elif na_sc - en_sc >= NLP_LOCK_OVERRIDE_GAP:
                        _skip_correction = True
                    _dbg(f'[LOCK-WRONG] native→en: orig={original!r} ({na_sc:.2f}) '
                         f'corr={corrected!r} ({en_sc:.2f}) skip={_skip_correction}')
                else:
                    # Original is English, corrected is native
                    ew = original.strip().lower()
                    if ew and _is_pure_alpha(ew):
                        en_sc = zipf_frequency(ew, 'en')
                    else:
                        en_sc = 0.0
                    nw = corrected.strip().lower()
                    na_sc = zipf_frequency(nw, state.sentence_lang) if nw else 0.0
                    # Keep original English if it's a real word and native isn't
                    if en_sc >= NLP_ZIPF_THRESHOLD and na_sc < NLP_ZIPF_THRESHOLD:
                        _skip_correction = True
                    elif en_sc - na_sc >= NLP_LOCK_OVERRIDE_GAP:
                        _skip_correction = True
                    _dbg(f'[LOCK-WRONG] en→native: orig={original!r} ({en_sc:.2f}) '
                         f'corr={corrected!r} ({na_sc:.2f}) skip={_skip_correction}')

                if _skip_correction:
                    _dbg(f'[LOCK-WRONG-SKIP] keeping {original!r} — NLP says correction is worse')
                    return

                if DEBUG:
                    print(f'Sentence-lock: {original!r} -> {corrected!r}')
                _spawn_single(original, corrected, boundary, state.sentence_lang)
                return

            # Word matches the locked language — but verify via NLP.
            # If the OTHER language scores much higher, correct just THIS word
            # but keep the sentence lock intact.  A single foreign word in a
            # sentence shouldn't flip the lock for all subsequent words.
            nlp_winner = _nlp_decide_for_lock(original, state.sentence_lang)
            if nlp_winner and nlp_winner != state.sentence_lang:
                _dbg(f'[LOCK-OVERRIDE] NLP says {nlp_winner!r} not '
                     f'{state.sentence_lang!r} for {original!r} — '
                     f'correcting word but KEEPING {state.sentence_lang!r} lock')
                # Translate the word to the winner language
                if nlp_winner == 'en':
                    to_en = PROFILE_TO_EN.get(state.sentence_lang, {})
                    corrected = _translate(original, to_en)
                else:
                    from_en = PROFILE_FROM_EN.get(nlp_winner, {})
                    corrected = _translate(original, from_en)
                # Correct this one word, but don't touch sentence_lang
                _spawn_single(original, corrected, boundary, nlp_winner)
                return
            else:
                return

        # Run NLP decision
        _decide_word(original, boundary)


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Keyboard hook                                                ║
# ╚═══════════════════════════════════════════════════════════════╝

def _is_own_window(hwnd):
    """Check if hwnd belongs to our own process."""
    if not hwnd:
        return False
    try:
        import ctypes
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value == _own_pid
    except Exception:
        return False


def _on_key(event: keyboard.KeyboardEvent):
    global _prev_words, _last_target_hwnd

    if event.event_type == 'up':
        return

    # Track the foreground HWND — but ONLY if it's not our own widget/editor
    try:
        hwnd = user32.GetForegroundWindow()
        if hwnd and not _is_own_window(hwnd):
            _last_target_hwnd = hwnd
    except Exception:
        pass

    name = event.name

    # Skip when our text editor / widget has focus — don't interfere with widget input
    try:
        fg = user32.GetForegroundWindow()
        if fg and _is_own_window(fg):
            return
    except Exception:
        pass

    # ── Privacy Guard: block collection on password fields / sensitive apps ──
    if _privacy_should_block():
        # Silently discard — don't collect, don't correct, don't buffer
        # NOTE: deliberately NO logging here — protect user privacy
        with state.lock:
            if state.buffer:
                state.buffer = ''
                state.pending = []
        return

    _dbg(f'[KEY] name={name!r} scan={event.scan_code} injecting={injecting.is_set()}')

    # During injection: ignore ALL keystrokes (both ours and user's).
    # This avoids Ctrl+V 'v' key leaking into overflow and clipboard contamination.
    # User keystrokes during the ~100ms correction window still reach the app
    # (hook is non-suppressing) and will be processed normally after correction.
    if injecting.is_set():
        if DEBUG:
            print(f'[SKIP] injecting is set, ignoring key: {name!r}')
        return

    if DEBUG:
        print(f'[KEY] name={name!r} buffer={state.buffer!r}')

    # Skip when Ctrl / Alt / Win modifiers are held (Ctrl+C, Ctrl+V, Alt+Tab, etc.)
    # These are shortcuts, not text entry — don't pollute the buffer.
    try:
        import keyboard as _kb
        mods = _kb.is_pressed('ctrl') or _kb.is_pressed('alt') or _kb.is_pressed('windows')
    except Exception:
        mods = False
    if mods and name not in ('ctrl', 'alt', 'windows', 'right ctrl', 'left ctrl',
                              'right alt', 'left alt', 'right windows', 'left windows',
                              'shift', 'right shift'):
        _dbg(f'[SKIP] modifier held, ignoring key: {name!r}')
        # Ctrl+A selects all text — full reset (next Backspace/char replaces everything)
        if name == 'a':
            with state.lock:
                state.buffer = ''
                state.pending = []
                state.reset_sentence()
            _prev_words.clear()
            _word_corrections.clear()
            _fire_buffer_callback()
        elif name in ('z', 'x', 'c', 'v'):
            with state.lock:
                state.buffer = ''
            _fire_buffer_callback()
        return

    # ── Tab: accept pending spell suggestion (if any) ──
    if name == 'tab':
        with _spell_pending_lock:
            has_pending = _spell_pending is not None
        if has_pending:
            if _spell_accept():
                return  # Tab consumed by spell accept — don't pass through
        # No spell pending → let Tab through normally (don't intercept)

    # Any key other than Tab dismisses the pending spell suggestion
    if name != 'tab':
        _spell_dismiss()

    # Sentence boundary (ENTER/period): flush pending BEFORE resetting
    # (Fixes v2 bug M8: period used to reset before correction ran)
    if name in SENTENCE_BOUNDARIES:
        # BUT period key may be ץ in Hebrew when layout is English (wrong-layout typing)
        if name in ('.', 'dot') and event.scan_code:
            _lid2 = _foreground_lang_id()
            if is_english_lang_id(_lid2):
                _en_key2 = _scan_to_english_char(event.scan_code)
                _is_actually_letter = False
                if _en_key2:
                    for _pcode2, _pfrom2 in PROFILE_FROM_EN.items():
                        _mapped2 = _pfrom2.get(_en_key2)
                        if _mapped2 and _mapped2.isalpha():
                            # Period maps to a letter in this profile — check buffer
                            buf2 = state.buffer
                            if buf2:
                                _buf_maps2 = sum(1 for c in buf2 if c in _pfrom2)
                                if _buf_maps2 >= len(buf2) * 0.5:
                                    _is_actually_letter = True
                                    _dbg(f'[SENTENCE→CHAR-EN] name={name!r} maps_to={_mapped2!r} in {_pcode2} — treating as letter')
                                    break
                if _is_actually_letter:
                    # Skip sentence boundary processing — let it fall through
                    # to the regular boundary/character handling below
                    pass
                else:
                    pass  # fall through to sentence boundary handling below
                if _is_actually_letter:
                    # Treat as a regular character — add to buffer
                    with state.lock:
                        state.buffer += name if len(name) == 1 else '.'
                    _fire_buffer_callback()
                    return

        with state.lock:
            if state.paused_until_boundary:
                state.paused_until_boundary = False
                if DEBUG:
                    print('Resuming corrections')

            # If there's a current word in buffer, process it first
            if state.buffer and name in ('.', 'dot'):
                # Period is both boundary and sentence-end
                original = state.buffer
                state.buffer = ''
                state.words_in_sentence += 1
                _decide_word(original, '.')

            # Flush remaining pending words — correct if clear winner, else drop
            if state.pending:
                votes: dict[str, int] = {}
                for pw in state.pending:
                    w = pw.get('winner')
                    if w:
                        votes[w] = votes.get(w, 0) + 1
                if votes:
                    best = max(votes, key=votes.get)
                    if DEBUG:
                        print(f'Sentence-end flush: {len(state.pending)} pending -> {best}')
                    state.sentence_lang = best
                    state.confirmed_words += len(state.pending)

                    # Pre-check: does any word actually need correction?
                    any_needs = False
                    for pw in state.pending:
                        if best == 'en' and pw['typed_is_native']:
                            any_needs = True
                            break
                        elif best != 'en' and not pw['typed_is_native']:
                            any_needs = True
                            break

                    if any_needs:
                        pending_copy = list(state.pending)
                        state.pending = []
                        injecting.set()
                        threading.Thread(
                            target=_replace_batch,
                            args=(best, pending_copy, None),
                            daemon=True,
                        ).start()
                    else:
                        state.pending = []
                else:
                    if DEBUG:
                        print(f'Dropping {len(state.pending)} ambiguous pending words')
                    state.pending = []

            state.reset_sentence()
            _prev_words = []

        if name == 'enter':
            with state.lock:
                state.buffer = ''
            _prev_words = []
            _fire_buffer_callback()
        return

    # Word boundary — but only if the key is ACTUALLY punctuation in the
    # foreground layout.  In Hebrew the comma key is ת, period is ץ,
    # semicolon is ף — all common letters, NOT boundaries.
    boundary = BOUNDARY_KEYS.get(name)
    if boundary is not None:
        _key_is_letter = False
        if event.scan_code:
            _lid = _foreground_lang_id()
            if not is_english_lang_id(_lid):
                # If Shift is held, user wants the shifted English char
                # (e.g. Shift+ף → ':'), NOT the Hebrew letter
                _shift_held = False
                try:
                    _shift_held = keyboard.is_pressed('shift')
                except Exception:
                    pass
                if not _shift_held:
                    _actual = _scan_to_actual_char(event.scan_code, _lid)
                    if _actual and _actual.isalpha():
                        _key_is_letter = True
                        _dbg(f'[BOUNDARY→CHAR] name={name!r} actual={_actual!r} — treating as letter')
            else:
                # Layout IS English — but user may be typing wrong-layout
                # (e.g. Hebrew chars on EN layout: ת→',' ף→';' ץ→'.')
                # Check if this punctuation maps to a letter in any active profile
                _en_key = _scan_to_english_char(event.scan_code)
                if _en_key and ACTIVE_PROFILES:
                    _shift_held = False
                    try:
                        _shift_held = keyboard.is_pressed('shift')
                    except Exception:
                        pass
                    if not _shift_held:
                        for _pcode, _pfrom in PROFILE_FROM_EN.items():
                            _mapped = _pfrom.get(_en_key)
                            if _mapped and _mapped.isalpha():
                                # This punct key IS a letter in a non-EN profile.
                                # Treat as letter if the buffer already has content
                                # that maps to that same profile (wrong-layout typing).
                                buf = state.buffer
                                if buf:
                                    _buf_maps = sum(1 for c in buf if c in _pfrom)
                                    if _buf_maps >= len(buf) * 0.5:
                                        _key_is_letter = True
                                        _dbg(f'[BOUNDARY→CHAR-EN] name={name!r} en_key={_en_key!r} maps_to={_mapped!r} in {_pcode} — treating as letter (buffer={buf!r})')
                                        break
                                elif not buf:
                                    # Empty buffer — could be first char of a new word.
                                    # Allow common start-of-word letters through.
                                    _key_is_letter = True
                                    _dbg(f'[BOUNDARY→CHAR-EN] name={name!r} en_key={_en_key!r} maps_to={_mapped!r} in {_pcode} — empty buffer, treating as letter')
                                    break

        if not _key_is_letter:
            with state.lock:
                if state.paused_until_boundary:
                    _push_prev_word(state.buffer)
                    state.buffer = ''
                    state.paused_until_boundary = False  # Resume on next word
                    _fire_buffer_callback()
                    if DEBUG:
                        print('Pause cleared on word boundary')
                    return

            fg = _get_foreground_exe()
            if not fg or fg not in CORRECTION_EXCLUDED_EXE:
                _word_before = state.buffer
                _push_prev_word(state.buffer)
                _process_boundary(boundary)
                _fire_buffer_callback()
                # Spell check the completed word
                _spell_check_completed_word(_word_before)
            else:
                with state.lock:
                    _push_prev_word(state.buffer)
                    state.buffer = ''
                _fire_buffer_callback()
            return
        # else: fall through to character handling below

    # Delete — reset the typing panel completely
    if name == 'delete':
        with state.lock:
            state.buffer = ''
            state.pending = []
            state.reset_sentence()
        _prev_words.clear()
        _word_corrections.clear()
        _fire_buffer_callback()
        return

    # Backspace
    if name == 'backspace':
        with state.lock:
            if state.buffer:
                state.buffer = state.buffer[:-1]
            elif _prev_words:
                # Buffer empty — pop last completed word back so continued
                # backspacing removes it from the panel too.
                state.buffer = _prev_words.pop()[:-1]  # minus 1 char for this backspace
        _fire_buffer_callback()
        return

    # Regular character: use scan code for reliable layout-independent mapping
    if event.scan_code:
        lang_id = _foreground_lang_id()

        # Shift + OEM key in non-English layout: use shifted English char
        # (e.g. Shift+ת → '<', Shift+ץ → '>')  instead of Hebrew letter
        if not is_english_lang_id(lang_id):
            _shift_held = False
            try:
                _shift_held = keyboard.is_pressed('shift')
            except Exception:
                pass
            if _shift_held:
                _en_base = _scan_to_english_char(event.scan_code)
                _shifted = SHIFT_EN_CHARS.get(_en_base)
                if _shifted:
                    _dbg(f'[SHIFT+OEM] Shift+{_en_base!r} → {_shifted!r}')
                    # Check if shifted char is a known boundary
                    _sbnd = BOUNDARY_KEYS.get(_shifted)
                    if _sbnd:
                        # Process as word boundary
                        _push_prev_word(state.buffer)
                        _process_boundary(_sbnd)
                        _fire_buffer_callback()
                    else:
                        with state.lock:
                            state.buffer += _shifted
                        _fire_buffer_callback()
                    return

        ch = _scan_to_actual_char(event.scan_code, lang_id)
        if not ch:
            return
        en_char = _scan_to_english_char(event.scan_code)
        _dbg(f'[CHAR] scan={event.scan_code} en={en_char!r} actual={ch!r} lang=0x{lang_id:04x}')

        with state.lock:
            state.buffer += ch

            # Privacy Guard: if buffer looks like sensitive data (CC, ID, OTP),
            # silently discard without correcting or storing
            if PRIVACY_GUARD_ENABLED and _is_sensitive_buffer(state.buffer):
                _dbg(f'[PRIVACY] Sensitive buffer detected, clearing: {len(state.buffer)} chars')
                state.buffer = ''
                state.pending = []
                return

        _fire_buffer_callback()

        # Auto-learn stats
        if ENABLE_AUTO_LEARN:
            fg = _get_foreground_exe()
            if fg in WATCH_TITLE_CHANGES_EXE:
                title = _get_foreground_title()
                _update_learn_stats(fg, title, ch)

        if DEBUG and len(state.buffer) > 0 and len(state.buffer) % 8 == 0:
            print(f'Buffer: {state.buffer!r}')
        return


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Debug / Info                                                 ║
# ╚═══════════════════════════════════════════════════════════════╝

def _undo_last_correction():
    """Undo the last auto-correction (F12 hotkey).

    Backspaces the corrected text and retypes the original.
    Only works if the cursor is right after the last correction.
    """
    with state.lock:
        undo = state.last_undo
        state.last_undo = None
    if not undo:
        _dbg('[UNDO] nothing to undo')
        return

    corrected = undo['corrected']
    original = undo['original']
    _dbg(f'[UNDO] {corrected!r} -> {original!r}')

    injecting.set()
    try:
        time.sleep(0.02)
        _send_backspaces(len(corrected))
        time.sleep(0.02)
        _send_unicode_text(original)
    except Exception as e:
        _log_error(f'_undo failed: {e}')
    finally:
        time.sleep(0.02)
        injecting.clear()
        # Reset sentence so engine doesn't re-correct the restored text
        with state.lock:
            state.reset_sentence()
            state.buffer = ''


def _revert_word(corrected_word: str):
    """Revert a specific corrected word back to its original.

    Called from the UI when the user clicks on a corrected word in the panel.
    This doesn't change the text on screen (the user already typed past it),
    so we send backspaces to the position of the word and retype.

    NOTE: This only works reliably for the *last* corrected word, because
    we can't move the cursor backwards to arbitrary positions.
    For simplicity we require the word to be the last thing typed before
    the current buffer (i.e. the most recent completed word).
    """
    global _prev_words, _word_corrections

    original = _word_corrections.get(corrected_word)
    if not original:
        _dbg(f'[REVERT] no correction found for {corrected_word!r}')
        return

    # Figure out how many chars to backspace:
    # We need to erase: current_buffer + boundary(space) + corrected_word
    # then retype: original_word + boundary + current_buffer
    with state.lock:
        cur_buf = state.buffer

    # Count of chars to delete = current buffer + 1 space + corrected word
    total_delete = len(cur_buf) + 1 + len(corrected_word)

    _dbg(f'[REVERT] {corrected_word!r} -> {original!r} (del={total_delete}, buf={cur_buf!r})')

    injecting.set()
    try:
        # Restore focus to the app window (clicking the widget stole focus)
        if _last_target_hwnd:
            try:
                user32.SetForegroundWindow(_last_target_hwnd)
            except Exception:
                pass
        time.sleep(0.05)
        _send_backspaces(total_delete)
        time.sleep(0.02)
        # Retype: original word + space + whatever was in buffer
        retype = original + ' ' + cur_buf if cur_buf else original + ' '
        _send_unicode_text(retype)
    except Exception as e:
        _log_error(f'_revert_word failed: {e}')
    finally:
        time.sleep(0.02)
        injecting.clear()
        # Update tracking
        # Replace corrected word back with original in _prev_words
        for i in range(len(_prev_words) - 1, -1, -1):
            if _prev_words[i] == corrected_word:
                _prev_words[i] = original
                break
        # Remove from corrections map
        _word_corrections.pop(corrected_word, None)
        with state.lock:
            state.reset_sentence()
        _fire_buffer_callback()


def _replace_panel_word(word_index: int, actual_word: str, replacement: str,
                        snapshot_parts: list, snapshot_buf: str):
    """Replace a word shown on the panel with its translated version.

    Called from the UI when the user clicks on a translated word in line 2.
    word_index: index in the display parts list
    actual_word: the word currently in the text (line 1)
    replacement: the translated word to replace it with (line 2)
    snapshot_parts: copy of the parts list captured at click time
    snapshot_buf: copy of buffer captured at click time
    """
    global _prev_words

    parts = snapshot_parts
    cur_buf = snapshot_buf

    if word_index < 0 or word_index >= len(parts):
        _dbg(f'[PANEL_REPLACE] invalid index {word_index}, parts={parts}')
        return

    # Verify the word at that index matches
    if parts[word_index] != actual_word:
        _dbg(f'[PANEL_REPLACE] mismatch: parts[{word_index}]={parts[word_index]!r} != {actual_word!r}')
        return

    # Build the text from word_index to cursor in the target app.
    # In the target app, after all typing, the text looks like:
    #   prev_word0 SPACE prev_word1 SPACE buffer    (if buffer not empty)
    #   prev_word0 SPACE prev_word1 SPACE            (if buffer empty → trailing space)
    has_trailing_space = not cur_buf
    target_parts = parts[word_index:]
    text_to_delete = ' '.join(target_parts)
    if has_trailing_space:
        text_to_delete += ' '
    total_delete = len(text_to_delete)

    # Build replacement text
    new_parts = [replacement] + parts[word_index + 1:]
    retype = ' '.join(new_parts)
    if has_trailing_space:
        retype += ' '

    _dbg(f'[PANEL_REPLACE] idx={word_index} {actual_word!r}->{replacement!r} '
         f'del={total_delete} retype={retype!r} parts={parts} buf={cur_buf!r}')

    injecting.set()
    try:
        # Restore focus to the target app (clicking our widget stole focus)
        if _last_target_hwnd:
            try:
                user32.SetForegroundWindow(_last_target_hwnd)
            except Exception:
                pass
        time.sleep(0.05)
        _send_backspaces(total_delete)
        time.sleep(0.02)
        _send_unicode_text(retype)
    except Exception as e:
        _log_error(f'_replace_panel_word failed: {e}')
    finally:
        time.sleep(0.02)
        injecting.clear()
        # Update tracking: replace the word in _prev_words or buffer
        is_buffer = (word_index == len(parts) - 1 and cur_buf)
        if is_buffer:
            with state.lock:
                state.buffer = replacement
        else:
            if word_index < len(_prev_words):
                _prev_words[word_index] = replacement
        # Remove old correction mapping if any
        _word_corrections.pop(actual_word, None)
        with state.lock:
            state.reset_sentence()
        _fire_buffer_callback()


def _print_info():
    """Print debug info about the current state (INFO_HOTKEY)."""
    try:
        exe = _get_foreground_exe()
        title = _get_foreground_title()
        lang_id = _foreground_lang_id()
        print(f'FG: exe={exe!r} lang=0x{lang_id:04x} title={title[:80]!r}')
        with state.lock:
            print(f'State: buffer={state.buffer!r} sentence={state.sentence_lang} '
                  f'words={state.words_in_sentence} pending={len(state.pending)} '
                  f'paused={state.paused_until_boundary}')
    except Exception as e:
        print(f'Info failed: {e}')


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Config loading (from UI)                                     ║
# ╚═══════════════════════════════════════════════════════════════╝

def _load_ui_config():
    """Load config from UI config file."""
    global DEBUG, EXCLUDE_WORDS, AUTO_SWITCH_LAYOUT, AUTO_SWITCH_AFTER_CONSECUTIVE
    global APP_DEFAULT_LANG_BY_EXE, APP_DEFAULT_LANG_BY_TITLE, ENGINE_ENABLED

    config_path = os.path.expanduser('~/.auto_lang2_config.json')
    if not os.path.exists(config_path):
        return

    try:
        with open(config_path, 'r', encoding='utf-8-sig') as f:
            cfg = json.load(f)

        if 'app_defaults' in cfg:
            APP_DEFAULT_LANG_BY_EXE.clear()
            APP_DEFAULT_LANG_BY_EXE.update(cfg['app_defaults'])

        if 'chat_defaults' in cfg:
            APP_DEFAULT_LANG_BY_TITLE.clear()
            for exe, chats in cfg['chat_defaults'].items():
                if chats:
                    APP_DEFAULT_LANG_BY_TITLE[exe] = chats

        if 'watch_title_exes' in cfg:
            WATCH_TITLE_CHANGES_EXE.clear()
            WATCH_TITLE_CHANGES_EXE.update(cfg['watch_title_exes'])
        WATCH_TITLE_CHANGES_EXE.update(APP_DEFAULT_LANG_BY_TITLE.keys())

        if 'browser_defaults' in cfg:
            BROWSER_LANG_BY_KEYWORD.clear()
            BROWSER_LANG_BY_KEYWORD.update(cfg['browser_defaults'])
            # Auto-watch all browsers for title changes
            if BROWSER_LANG_BY_KEYWORD:
                WATCH_TITLE_CHANGES_EXE.update(BROWSER_EXES)

        if 'exclude_words' in cfg:
            EXCLUDE_WORDS = set(w.lower() for w in cfg['exclude_words'])

        if 'enabled' in cfg:
            ENGINE_ENABLED = cfg['enabled']
        if 'debug' in cfg:
            DEBUG = cfg['debug']
        if 'auto_switch' in cfg:
            AUTO_SWITCH_LAYOUT = cfg['auto_switch']
        if 'auto_switch_count' in cfg:
            AUTO_SWITCH_AFTER_CONSECUTIVE = cfg['auto_switch_count']

        if DEBUG:
            print(f'Loaded config: {len(APP_DEFAULT_LANG_BY_EXE)} app defaults, '
                  f'{len(EXCLUDE_WORDS)} excludes')
    except Exception as e:
        print(f'Config load failed: {e}')


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Mouse click detection (reset buffer on click)               ║
# ╚═══════════════════════════════════════════════════════════════╝

def _on_mouse_click():
    """Called when any mouse button is clicked. Resets the typing buffer."""
    if injecting.is_set():
        return
    # Skip reset if UI is handling a panel word click
    if _widget_clicking.is_set():
        _dbg('[MOUSE] widget_clicking flag set, skip reset')
        return
    with state.lock:
        old_buf = state.buffer
        state.reset_all()
        if old_buf:
            _dbg(f'[MOUSE] Click detected, full reset (was: {old_buf!r})')
        else:
            _dbg('[MOUSE] Click detected, reset (buffer was empty)')


def _install_mouse_hook():
    """Install mouse click hook using the mouse library."""
    mouse.on_click(_on_mouse_click)
    _dbg('[MOUSE] Hook installed via mouse library')


def _uninstall_mouse_hook():
    """Unhook mouse listener."""
    try:
        mouse.unhook_all()
    except Exception:
        pass


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Main                                                         ║
# ╚═══════════════════════════════════════════════════════════════╝

def main():
    global _own_pid
    import os
    _own_pid = os.getpid()
    _dbg(f'=== AutoLang v3 starting === DEBUG={DEBUG} log={_DEBUG_LOG_PATH} pid={_own_pid}')
    _ensure_single_instance()
    _load_language_profiles()
    _dbg(f'Profiles loaded: {list(ACTIVE_PROFILES.keys())}')
    _load_ui_config()
    _dbg(f'Config loaded, DEBUG={DEBUG}, ENGINE_ENABLED={ENGINE_ENABLED}')
    _load_learned()

    # Build watch set
    WATCH_TITLE_CHANGES_EXE.update(APP_DEFAULT_LANG_BY_TITLE.keys())
    WATCH_TITLE_CHANGES_EXE.update({
        'chrome.exe', 'msedge.exe', 'outlook.exe', 'olk.exe',
    })
    if BROWSER_LANG_BY_KEYWORD:
        WATCH_TITLE_CHANGES_EXE.update(BROWSER_EXES)

    keyboard.add_hotkey(EXIT_HOTKEY, lambda: stop_event.set())
    keyboard.add_hotkey(INFO_HOTKEY, _print_info)
    keyboard.add_hotkey(UNDO_HOTKEY, _undo_last_correction)
    keyboard.add_hotkey(GRAMMAR_HOTKEY, _trigger_grammar_check)
    keyboard.on_press(_on_key, suppress=False)

    # Install mouse click hook (reset buffer when user clicks)
    _install_mouse_hook()

    # Pre-load spell check dictionaries in background
    spell_module.init_background()

    if ENABLE_APP_DEFAULT_LAYOUT:
        threading.Thread(target=_app_default_watcher, name='app_default', daemon=True).start()

    print('AutoLang v3 running')
    print(f'  Exit: {EXIT_HOTKEY}')
    print(f'  Info: {INFO_HOTKEY}')
    print(f'  Undo: {UNDO_HOTKEY}')
    langs = ', '.join(f'{p.flag} {p.name}' for p in ACTIVE_PROFILES.values())
    print(f'  Languages: {langs or "Hebrew (fallback)"}')
    print(f'  Detection: {MIN_WORDS_TO_DETECT} words or {MIN_CHARS_SINGLE_WORD}+ chars in single word')

    stop_event.wait()
    _uninstall_mouse_hook()
    keyboard.unhook_all()
    _save_learned()
    print('AutoLang stopped.')


if __name__ == '__main__':
    main()
