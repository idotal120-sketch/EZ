"""
AutoLang UI - PySide6 Version
============================================
ממשק משתמש לסקריפט auto_lang.py:
- אייקון ב-System Tray עם חיווי שהתוכנה רצה
- וידג'ט צף עם 4 כפתורים + עורך תרגום
- חלון הגדרות עם טאבים
"""

import json
import math
import os
import sys
import threading
import time

# When running as --noconsole EXE (PyInstaller), sys.stdout/stderr are None.
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

_UI_LOG_FILE = os.path.expanduser('~/.auto_lang_ui.log')


def _ui_log(msg: str):
    try:
        with open(_UI_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        pass

from PySide6.QtWidgets import (
    QApplication, QWidget, QSystemTrayIcon, QMenu,
    QTabWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QTextEdit, QCheckBox, QSpinBox, QComboBox,
    QListWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QFrame, QProgressBar, QScrollArea,
)
from PySide6.QtCore import (
    Qt, QTimer, QPoint, QRect, QRectF, Signal,
)
from PySide6.QtGui import (
    QIcon, QPixmap, QPainter, QPen, QColor, QFont,
    QFontMetrics, QPainterPath, QGuiApplication, QCursor,
)

# ----------------------------
# Dark title bar (Windows 10/11)
# ----------------------------

def _apply_dark_title_bar(widget: QWidget):
    """Force the Windows title bar to dark mode using DWM API."""
    try:
        import ctypes
        hwnd = int(widget.winId())
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20          # Win10 build >=19041
        DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19      # older Win10 builds
        value = ctypes.c_int(1)
        hr = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value), ctypes.sizeof(value))
        if hr != 0:  # fallback for older builds
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD,
                ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass


# ----------------------------
# Configuration file path
# ----------------------------
CONFIG_FILE = os.path.expanduser('~/.auto_lang2_config.json')

DEFAULT_CONFIG = {
    'app_defaults': {
        'whatsapp.exe': 'he',
        'whatsapp.root.exe': 'he',
        'telegram.exe': 'he',
        'winword.exe': 'he',
        'teams.exe': 'he',
        'ms-teams.exe': 'he',
        'code.exe': 'en',
        'pycharm64.exe': 'en',
        'windowsterminal.exe': 'en',
        'powershell.exe': 'en',
        'cmd.exe': 'en',
        'slack.exe': 'en',
        'excel.exe': 'en',
        'outlook.exe': 'en',
    },
    'chat_defaults': {
        'whatsapp.exe': {},
        'ms-teams.exe': {},
        'teams.exe': {},
    },
    'watch_title_exes': [
        'chrome.exe',
        'msedge.exe',
        'outlook.exe',
        'olk.exe',
    ],
    'browser_defaults': {},
    'exclude_words': [],
    'enabled': True,
    'debug': False,
    'auto_switch': True,
    'auto_switch_count': 2,
    'hide_scores': False,
    'show_typing_panel': False,
    'privacy_guard': True,
    'privacy_blocked_exes': [
        '1password.exe', 'keeper.exe', 'keepass.exe', 'keepassxc.exe',
        'lastpass.exe', 'bitwarden.exe', 'dashlane.exe', 'roboform.exe',
        'enpass.exe', 'mstsc.exe', 'vmconnect.exe', 'vmware-vmx.exe',
        'virtualbox.exe', 'anydesk.exe', 'teamviewer.exe',
    ],
    # Spell check
    'spell_enabled': True,
    'spell_mode': 'tooltip',       # 'tooltip' | 'auto' | 'visual'
    # Grammar / LLM
    'grammar_enabled': False,
    'grammar_provider': 'openai',  # 'openai' | 'anthropic' | 'gemini'
    'grammar_api_key': '',
    'grammar_model': '',           # empty → use provider default
}


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            config = DEFAULT_CONFIG.copy()
            for key in DEFAULT_CONFIG:
                if key in saved:
                    config[key] = saved[key]
            return config
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(config: dict) -> None:
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f'Failed to save config: {e}')


# ----------------------------
# Multi-language support
# ----------------------------
LANG_DISPLAY = {
    'en': ('English', '\U0001f1fa\U0001f1f8'),
    'he': ('\u05e2\u05d1\u05e8\u05d9\u05ea', '\U0001f1ee\U0001f1f1'),
    'ru': ('\u0420\u0443\u0441\u0441\u043a\u0438\u0439', '\U0001f1f7\U0001f1fa'),
    'ar': ('\u0627\u0644\u0639\u0631\u0628\u064a\u0629', '\U0001f1f8\U0001f1e6'),
    'uk': ('\u0423\u043a\u0440\u0430\u0457\u043d\u0441\u044c\u043a\u0430', '\U0001f1fa\U0001f1e6'),
    'fr': ('Fran\u00e7ais', '\U0001f1eb\U0001f1f7'),
    'de': ('Deutsch', '\U0001f1e9\U0001f1ea'),
    'es': ('Espa\u00f1ol', '\U0001f1ea\U0001f1f8'),
    'el': ('\u0395\u03bb\u03bb\u03b7\u03bd\u03b9\u03ba\u03ac', '\U0001f1ec\U0001f1f7'),
    'fa': ('\u0641\u0627\u0631\u0633\u06cc', '\U0001f1ee\U0001f1f7'),
    'tr': ('T\u00fcrk\u00e7e', '\U0001f1f9\U0001f1f7'),
    'th': ('\u0e20\u0e32\u0e29\u0e32\u0e44\u0e17\u0e22', '\U0001f1f9\U0001f1ed'),
    'hi': ('\u0939\u093f\u0928\u094d\u0926\u0940', '\U0001f1ee\U0001f1f3'),
    'ko': ('\ud55c\uad6d\uc5b4', '\U0001f1f0\U0001f1f7'),
    'pl': ('Polski', '\U0001f1f5\U0001f1f1'),
}


def _lang_display(code: str) -> str:
    info = LANG_DISPLAY.get(code)
    if info:
        return f'{info[0]} {info[1]}'
    return code


def _lang_code_from_display(display: str) -> str:
    for code, (name, flag) in LANG_DISPLAY.items():
        if f'{name} {flag}' == display or code == display:
            return code
    return display


def _get_available_lang_codes() -> list:
    codes = ['en', 'he']
    try:
        from keyboard_maps import ALL_PROFILES
        for code in ALL_PROFILES:
            if code not in codes:
                codes.append(code)
    except ImportError:
        pass
    return codes


# ──────────────────────────────────────────────────────────
# Dark Theme Color Palette
# ──────────────────────────────────────────────────────────
COLORS = {
    'bg_dark':       '#0a1628',
    'bg_medium':     '#0f2137',
    'bg_light':      '#162d4a',
    'bg_highlight':  '#1e3d5f',
    'accent':        '#4ec9f0',
    'accent_hover':  '#7dd8f5',
    'accent_dim':    '#1a6a8f',
    'translate':     '#ffe08a',
    'text_primary':  '#e0e8f0',
    'text_secondary': '#7a8fa5',
    'text_muted':    '#4a6580',
    'border':        '#1a3a5c',
    'success':       '#4ade80',
    'error':         '#f87171',
    'warning':       '#fbbf24',
    'tab_active':    '#4ec9f0',
    'tab_inactive':  '#162d4a',
    'btn_primary':   '#4ec9f0',
    'btn_danger':    '#f87171',
    'tree_stripe':   '#0c1f35',
    'scrollbar':     '#2a4a6a',
}

# ──────────────────────────────────────────────────────────
# QSS Dark Stylesheet — replaces ALL the ttk/clam hacks
# ──────────────────────────────────────────────────────────
DARK_QSS = f"""
QWidget {{
    background-color: {COLORS['bg_dark']};
    color: {COLORS['text_primary']};
    font-family: 'Segoe UI';
    font-size: 10pt;
}}
QTabWidget::pane {{
    border: none;
    background-color: {COLORS['bg_dark']};
}}
QTabBar::tab {{
    background-color: {COLORS['tab_inactive']};
    color: {COLORS['text_secondary']};
    padding: 8px 16px;
    border: none;
    font-size: 11pt;
}}
QTabBar::tab:selected {{
    background-color: {COLORS['bg_dark']};
    color: {COLORS['accent']};
}}
QTabBar::tab:hover:!selected {{
    background-color: {COLORS['bg_highlight']};
}}
QTableWidget {{
    background-color: {COLORS['bg_medium']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    gridline-color: {COLORS['border']};
    selection-background-color: {COLORS['accent_dim']};
    selection-color: #ffffff;
    alternate-background-color: {COLORS['tree_stripe']};
    font-size: 10pt;
}}
QHeaderView::section {{
    background-color: {COLORS['bg_light']};
    color: {COLORS['accent']};
    font-weight: bold;
    border: none;
    border-bottom: 1px solid {COLORS['border']};
    padding: 6px 8px;
    font-size: 10pt;
}}
QPushButton {{
    background-color: {COLORS['bg_light']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 10pt;
}}
QPushButton:hover {{
    background-color: {COLORS['bg_highlight']};
    color: {COLORS['accent_hover']};
}}
QPushButton:pressed {{
    background-color: {COLORS['accent_dim']};
}}
QPushButton[cssClass="accent"] {{
    background-color: {COLORS['accent']};
    color: {COLORS['bg_dark']};
    font-weight: bold;
}}
QPushButton[cssClass="accent"]:hover {{
    background-color: {COLORS['accent_hover']};
}}
QPushButton[cssClass="danger"] {{
    background-color: {COLORS['error']};
    color: #ffffff;
}}
QPushButton[cssClass="danger"]:hover {{
    background-color: #ff7070;
}}
QLineEdit, QSpinBox {{
    background-color: {COLORS['bg_light']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 3px;
    padding: 4px 6px;
    font-size: 10pt;
}}
QLineEdit:focus, QSpinBox:focus {{
    border-color: {COLORS['accent']};
}}
QComboBox {{
    background-color: {COLORS['bg_light']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 3px;
    padding: 4px 6px;
    font-size: 10pt;
}}
QComboBox:focus {{
    border-color: {COLORS['accent']};
}}
QComboBox::drop-down {{
    border: none;
}}
QComboBox QAbstractItemView {{
    background-color: {COLORS['bg_light']};
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['accent_dim']};
    selection-color: #ffffff;
    border: 1px solid {COLORS['border']};
}}
QCheckBox {{
    color: {COLORS['text_primary']};
    spacing: 8px;
    font-size: 10pt;
}}
QCheckBox::indicator {{
    width: 18px;
    height: 18px;
    border: 2px solid {COLORS['border']};
    border-radius: 3px;
    background-color: {COLORS['bg_light']};
}}
QCheckBox::indicator:checked {{
    background-color: {COLORS['accent']};
    border-color: {COLORS['accent']};
}}
QCheckBox::indicator:hover {{
    border-color: {COLORS['accent_hover']};
}}
QListWidget {{
    background-color: {COLORS['bg_medium']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    selection-background-color: {COLORS['accent_dim']};
    selection-color: #ffffff;
    font-size: 11pt;
}}
QScrollBar:vertical {{
    background-color: {COLORS['bg_dark']};
    width: 10px;
    border: none;
}}
QScrollBar::handle:vertical {{
    background-color: {COLORS['scrollbar']};
    border-radius: 5px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
    background-color: {COLORS['accent_dim']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background-color: {COLORS['bg_dark']};
    height: 10px;
    border: none;
}}
QScrollBar::handle:horizontal {{
    background-color: {COLORS['scrollbar']};
    border-radius: 5px;
    min-width: 20px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}
QProgressBar {{
    background-color: {COLORS['bg_medium']};
    border: none;
    border-radius: 4px;
    text-align: center;
    color: {COLORS['text_primary']};
}}
QProgressBar::chunk {{
    background-color: {COLORS['accent']};
    border-radius: 4px;
}}
QTextEdit {{
    background-color: {COLORS['bg_medium']};
    color: {COLORS['text_primary']};
    border: none;
    font-size: 10pt;
}}
QMenu {{
    background-color: {COLORS['bg_medium']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    padding: 4px;
    font-size: 10pt;
}}
QMenu::item {{
    padding: 6px 20px;
}}
QMenu::item:selected {{
    background-color: {COLORS['accent_dim']};
    color: #ffffff;
}}
QMenu::separator {{
    height: 1px;
    background-color: {COLORS['border']};
    margin: 4px 8px;
}}
"""


# ──────────────────────────────────────────────────────────
# System Tray Icon (QPainter replaces PIL)
# ──────────────────────────────────────────────────────────

def create_tray_icon(enabled: bool = True) -> QIcon:
    """Create the system tray icon using QPainter."""
    size = 64
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    p = QPainter(pixmap)
    p.setRenderHint(QPainter.Antialiasing)

    if enabled:
        p.setBrush(QColor(15, 60, 120, 80))
        p.setPen(Qt.NoPen)
        p.drawEllipse(2, 2, 60, 60)
        p.setBrush(QColor(20, 75, 140))
        p.drawEllipse(5, 5, 54, 54)
        p.setBrush(QColor(78, 201, 240))
        p.drawEllipse(8, 8, 48, 48)
    else:
        p.setBrush(QColor(120, 50, 50, 80))
        p.setPen(Qt.NoPen)
        p.drawEllipse(2, 2, 60, 60)
        p.setBrush(QColor(160, 60, 60))
        p.drawEllipse(5, 5, 54, 54)
        p.setBrush(QColor(248, 113, 113))
        p.drawEllipse(8, 8, 48, 48)

    font = QFont('Arial', 16)
    font.setBold(True)
    p.setFont(font)
    p.setPen(QColor(255, 255, 255))
    p.drawText(QRect(0, 0, size, size), Qt.AlignCenter, '\u05d0a')

    p.end()
    return QIcon(pixmap)


# ──────────────────────────────────────────────────────────
# Custom QTextEdit for translator input (Enter = translate)
# ──────────────────────────────────────────────────────────

class _TranslatorInput(QTextEdit):
    """Text input that emits enter_pressed on Enter (Shift+Enter = newline)."""
    enter_pressed = Signal()
    boundary_typed = Signal()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Ensure pasted content does not keep source formatting/colors
        self.setAcceptRichText(False)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if not (event.modifiers() & Qt.ShiftModifier):
                self.enter_pressed.emit()
                return
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        if event.key() in (Qt.Key_Space, Qt.Key_Return, Qt.Key_Period,
                           Qt.Key_Comma, Qt.Key_Semicolon, Qt.Key_Colon,
                           Qt.Key_Exclam, Qt.Key_Question):
            self.boundary_typed.emit()
        super().keyReleaseEvent(event)

    def insertFromMimeData(self, source):
        # Force plain-text paste to keep consistent color/formatting
        text = source.text() if source is not None else ''
        if text:
            try:
                self.setTextColor(QColor(COLORS['accent']))
            except Exception:
                pass
            self.insertPlainText(text)
        else:
            super().insertFromMimeData(source)


# ──────────────────────────────────────────────────────────
# Translate Popup (right-click quick translate)
# ──────────────────────────────────────────────────────────

class _TranslatePopup(QWidget):
    """Small frameless popup near cursor with a 'Translate' button."""
    translate_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        try:
            self.setWindowFlag(Qt.WindowDoesNotAcceptFocus, True)
        except Exception:
            pass
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFocusPolicy(Qt.NoFocus)
        self.setFixedSize(110, 36)

        self._frame = QFrame(self)
        self._frame.setGeometry(0, 0, 110, 36)
        self._frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_medium']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
            }}
        """)

        btn = QPushButton('\U0001f50d  \u05ea\u05e8\u05d2\u05dd', self._frame)
        btn.setGeometry(2, 2, 106, 32)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFocusPolicy(Qt.NoFocus)
        btn.setAutoDefault(False)
        btn.setDefault(False)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['accent']};
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 10pt;
            }}
            QPushButton:hover {{
                background-color: {COLORS['bg_light']};
            }}
        """)
        btn.clicked.connect(self.translate_clicked.emit)


# ──────────────────────────────────────────────────────────
# Floating Widget (always-on-top draggable icon)
# ──────────────────────────────────────────────────────────

class FloatingWidget(QWidget):
    """
    חלון צף קטן עם האייקון של התוכנה.
    תמיד מעל כל החלונות, ניתן לגרירה, לחיצה פותחת תפריט פעולות.
    """

    WIDGET_SIZE = 68
    OPACITY = 0.92
    EDGE_SNAP_MARGIN = 8
    PANEL_WIDTH = 220

    # Thread-safe signals
    _buffer_signal = Signal(str, str, object, object)
    _state_signal = Signal(bool)
    _models_ready_signal = Signal()
    _speech_text_signal = Signal(str, str)
    _speech_error_signal = Signal(str)
    _spell_signal = Signal(object, object, str)       # (word, suggestions, mode)
    _grammar_signal = Signal(str, object, object)     # (original, corrected, error)
    _translate_result_signal = Signal(str, bool, str) # (text, error, source)
    _rc_signal = Signal()
    _translate_fill_signal = Signal(str)

    def __init__(self, tray_app: 'AutoLangTray'):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(self.OPACITY)
        self.setMouseTracking(True)

        self.tray = tray_app
        _cfg = tray_app.config if tray_app else {}
        self._panel_visible = _cfg.get('show_typing_panel', False)
        self._hide_scores = _cfg.get('hide_scores', False)

        # State
        self._hovered_quadrant = None
        self._drag_start = None
        self._drag_moved = False
        self._cur_line1 = ' '
        self._cur_line2 = ' '
        self._corrections = {}
        self._word_scores = {}
        self._line2_word_rects = []   # [(word_index, x_left, x_right, word_str), ...]
        self._scroll_x = 0   # manual horizontal scroll offset (pixels)

        # Editor
        self._edit_mode = False
        self._editor_height = 0
        self._editor_min_h = 148
        self._editor_max_h = 500
        self._text_widget = None
        self._trans_widget = None
        self._trans_container = None
        self._trans_sep = None
        self._toggle_trans_btn = None
        self._trans_visible = True
        self._status_label = None
        self._send_btn = None
        self._editor_container = None
        self._resize_drag_y = None
        self._send_target_hwnd = None

        # Timers
        self._translate_debounce = QTimer(self)
        self._translate_debounce.setSingleShot(True)
        self._translate_debounce.setInterval(150)
        self._translate_debounce.timeout.connect(self._do_translate)

        self._pulse_timer = QTimer(self)
        self._pulse_timer.setInterval(500)
        self._pulse_timer.timeout.connect(self._pulse_animation)
        self._pulse_step = 0

        # Speech
        self._speech_recording = False
        self._last_speech_text = ''
        self._last_speech_time = 0

        # Translate popup
        self._translate_popup = None
        self._popup_hide_timer = QTimer(self)
        self._popup_hide_timer.setSingleShot(True)
        self._popup_hide_timer.timeout.connect(self._hide_translate_popup)

        # Spell tooltip
        self._spell_tooltip = None
        self._spell_hide_timer = QTimer(self)
        self._spell_hide_timer.setSingleShot(True)
        self._spell_hide_timer.setInterval(8000)
        self._spell_hide_timer.timeout.connect(self._hide_spell_tooltip)

        # Grammar result popup
        self._grammar_popup = None
        self._last_rc_time = 0.0
        self._rc_hook_thread = None
        self._rc_hook_id = None
        self._rc_hook_proc = None
        self._copy_in_flight = False
        self._rc_hwnd_point = None
        self._rc_hwnd_root = None
        self._rc_cache_in_flight = False
        self._last_rc_pos = None
        self._last_selection_text = ''
        self._last_selection_time = 0.0

        # Tooltip
        self._tooltip_win = None
        self._tooltip_timer = QTimer(self)
        self._tooltip_timer.setSingleShot(True)
        self._tooltip_timer.timeout.connect(self._do_show_tooltip)
        self._tooltip_tag = ''
        self._tooltip_texts = {
            'q_top':    '\u05d4\u05e4\u05e2\u05dc / \u05d4\u05e9\u05d1\u05ea \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4',
            'q_left':   '\u05d1\u05d8\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df',
            'q_bottom': '\u05e4\u05ea\u05d7 \u05e2\u05d5\u05e8\u05da \u05ea\u05e8\u05d2\u05d5\u05dd',
            'q_right':  '\u05d4\u05d2\u05d3\u05e8\u05d5\u05ea',
            'q_center': '\u05d4\u05e7\u05dc\u05d8\u05d4 \u05e7\u05d5\u05dc\u05d9\u05ea',
        }

        # Connect signals
        self._buffer_signal.connect(self._set_panel_text)
        self._state_signal.connect(self._apply_state)
        self._models_ready_signal.connect(self._update_ready_status)
        self._speech_text_signal.connect(self._handle_speech_result)
        self._speech_error_signal.connect(lambda msg: self._stop_speech())
        self._spell_signal.connect(self._show_spell_tooltip)
        self._grammar_signal.connect(self._show_grammar_result)
        self._translate_result_signal.connect(self._show_translation_result)
        self._rc_signal.connect(self._on_right_click_global)
        self._translate_fill_signal.connect(self._fill_and_translate_text)

        # Context menu
        self._menu = QMenu(self)
        self._build_menu()

        # Position: bottom-right, above taskbar
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = pw + sz
        screen = QGuiApplication.primaryScreen().availableGeometry()
        x = screen.width() - total_w - self.EDGE_SNAP_MARGIN - 12
        y = screen.height() - sz - 20
        self.setGeometry(x, y, total_w, sz)

    def start(self):
        """Initialize engine connection and show the widget."""
        # Register engine callback + tell engine our HWND
        try:
            import auto_lang
            auto_lang._buffer_callback = self._on_buffer_update
            auto_lang._widget_hwnd = int(self.winId())
        except Exception:
            pass

        # Register spell & grammar callbacks
        try:
            import spell_module
            spell_module.SPELL_CALLBACK = self._on_spell_result
        except Exception:
            pass
        try:
            import grammar_module
            grammar_module.GRAMMAR_CALLBACK = self._on_grammar_result
        except Exception:
            pass

        # Preload Whisper model in background
        try:
            import speech_module
            speech_module.preload_model()
        except Exception:
            pass

        # Right-click translate hook
        self._start_right_click_hook()
        try:
            import mouse as _mouse
            _ui_log('[RC] mouse module loaded')

            def _rc_fire():
                now = time.time()
                if now - self._last_rc_time > 0.4:
                    self._last_rc_time = now
                    self._on_right_click_global()

            def _rc_cb(*_a, **_k):
                _rc_fire()

            if hasattr(_mouse, 'on_right_click'):
                try:
                    _mouse.on_right_click(_rc_cb)
                    _ui_log('[RC] mouse.on_right_click registered')
                except Exception:
                    pass
            if hasattr(_mouse, 'on_click'):
                try:
                    _mouse.on_click(_rc_cb, buttons=('right',))
                    _ui_log('[RC] mouse.on_click right registered')
                except Exception:
                    try:
                        _mouse.on_click(_rc_cb)
                        _ui_log('[RC] mouse.on_click registered')
                    except Exception:
                        pass
            if hasattr(_mouse, 'hook'):
                def _rc_hook(event):
                    try:
                        btn = getattr(event, 'button', None)
                        et = getattr(event, 'event_type', None)
                        if btn in ('right', getattr(_mouse, 'RIGHT', None), 2) and et in ('up', 'down', 'click'):
                            _rc_fire()
                    except Exception:
                        pass
                try:
                    _mouse.hook(_rc_hook)
                    _ui_log('[RC] mouse.hook registered')
                except Exception:
                    pass
        except Exception:
            _ui_log('[RC] mouse module failed to load')
            pass

        self._apply_state(self.tray.enabled)

    # ── public API (called from any thread) ──────────────

    def update_state(self, enabled: bool):
        try:
            self._state_signal.emit(enabled)
        except Exception:
            pass

    def _on_buffer_update(self, actual: str, translated: str,
                          corrections=None, word_scores=None):
        try:
            self._buffer_signal.emit(actual, translated,
                                     corrections or {}, word_scores or {})
        except Exception:
            pass

    def _dbg_click(self, msg):
        """Write click debug message (no-op in production)."""
        pass

    # ── painting ─────────────────────────────────────────

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        w = self.width()
        sz = self.WIDGET_SIZE
        total_h = sz + self._editor_height
        enabled = self.tray.enabled

        # ── Draw background shape ──
        if self._editor_height > 0:
            path = self._build_unified_path(w, total_h)
        else:
            path = QPainterPath()
            path.addRoundedRect(QRectF(1, 1, w - 2, sz - 2), (sz - 2) / 2, (sz - 2) / 2)

        p.fillPath(path, QColor(COLORS['bg_medium']))
        p.setPen(QPen(QColor(COLORS['border']), 1))
        p.drawPath(path)

        # ── Separator between text panel and icon ──
        if self._panel_visible:
            sep_x = self.PANEL_WIDTH
            p.setPen(QPen(QColor(COLORS['border']), 1))
            p.drawLine(sep_x, 10, sep_x, sz - 10)

            # Draw text words
            self._paint_text_words(p)

        # ── 4-quadrant icon ──
        icon_x = self.PANEL_WIDTH if self._panel_visible else 0
        cx = icon_x + sz // 2
        cy = sz // 2
        r = sz // 2 - 4
        r_center = 13
        r_text = (r_center + r) // 2 + 1

        gap = 4
        extent = 90 - gap

        q_toggle = COLORS['accent'] if enabled else COLORS['error']
        q_base = '#1a5276'

        quadrants = [
            (90,  'q_top',    q_toggle, '\u23fb', 0, -1),
            (180, 'q_left',   q_base,   '\u21a9', -1, 0),
            (270, 'q_bottom', q_base,   '\u270e', 0,  1),
            (0,   'q_right',  q_base,   '\u2699', 1,  0),
        ]

        icon_rect = QRect(cx - r, cy - r, 2 * r, 2 * r)

        for angle, tag, color, label, dx, dy in quadrants:
            # Hover effect
            if tag == self._hovered_quadrant:
                if tag == 'q_top':
                    fill = '#6dd5f5' if enabled else '#f9a3a3'
                elif tag == 'q_center':
                    fill = '#e74c3c' if self._speech_recording else '#162d4a'
                else:
                    fill = '#2980b9'
            else:
                fill = color

            start_angle = int((angle - extent / 2) * 16)
            span_angle = int(extent * 16)

            p.setBrush(QColor(fill))
            p.setPen(QPen(QColor(COLORS['bg_dark']), 2))
            p.drawPie(icon_rect, start_angle, span_angle)

            # Label text
            label_font = QFont('Segoe UI Symbol', 10)
            p.setFont(label_font)
            p.setPen(QColor('#ffffff'))
            label_rect = QRectF(cx + dx * r_text - 12, cy + dy * r_text - 10, 24, 20)
            p.drawText(label_rect, Qt.AlignCenter, label)

        # ── Center circle ──
        if self._hovered_quadrant == 'q_center':
            center_fill = '#e74c3c' if self._speech_recording else '#162d4a'
            center_outline = '#e74c3c' if self._speech_recording else COLORS['accent']
        elif self._speech_recording:
            center_fill = '#c0392b'
            center_outline = '#e74c3c'
        else:
            center_fill = '#0d1f38'
            center_outline = COLORS['accent']

        p.setBrush(QColor(center_fill))
        p.setPen(QPen(QColor(center_outline), 2))
        p.drawEllipse(cx - r_center, cy - r_center, 2 * r_center, 2 * r_center)

        # Center indicator (dot or square)
        if self._speech_recording:
            # Blinking square
            sq = 5
            show = (self._pulse_step % 2 == 0) if self._pulse_timer.isActive() else True
            if show:
                p.setBrush(QColor('#e74c3c'))
                p.setPen(Qt.NoPen)
                p.drawRect(cx - sq, cy - sq, 2 * sq, 2 * sq)
        else:
            # Red dot
            dot_r = 5
            p.setBrush(QColor('#e74c3c'))
            p.setPen(Qt.NoPen)
            p.drawEllipse(cx - dot_r, cy - dot_r, 2 * dot_r, 2 * dot_r)

        # ── Editor resize grip ──
        if self._editor_height > 0:
            grip_y = total_h - 10
            grip_font = QFont('Segoe UI', 9)
            p.setFont(grip_font)
            p.setPen(QColor(COLORS['text_muted']))
            p.drawText(QRectF(0, grip_y - 4, w, 14), Qt.AlignCenter, '\u22ef')

        p.end()

    def _build_unified_path(self, w, total_h):
        """Build the unified shape path: pill top + editor bottom with rounded corners."""
        sz = self.WIDGET_SIZE
        pill_r = (sz - 2) / 2
        br = 20

        path = QPainterPath()
        # Start at top center of left semicircle
        path.moveTo(1 + pill_r, 1)
        # Top line
        path.lineTo(w - 1 - pill_r, 1)
        # Top-right semicircle (90° → 0°, clockwise)
        path.arcTo(QRectF(w - 1 - 2 * pill_r, 1, 2 * pill_r, 2 * pill_r), 90, -90)
        # Right side down
        path.lineTo(w - 1, total_h - 1 - br)
        # Bottom-right corner
        path.arcTo(QRectF(w - 1 - 2 * br, total_h - 1 - 2 * br, 2 * br, 2 * br), 0, -90)
        # Bottom line
        path.lineTo(1 + br, total_h - 1)
        # Bottom-left corner
        path.arcTo(QRectF(1, total_h - 1 - 2 * br, 2 * br, 2 * br), 270, -90)
        # Left side up
        path.lineTo(1, 1 + pill_r)
        # Top-left semicircle (180° → 90°, clockwise)
        path.arcTo(QRectF(1, 1, 2 * pill_r, 2 * pill_r), 180, -90)
        path.closeSubpath()
        return path

    def _paint_text_words(self, p):
        """Paint the per-word text on the panel area with scroll & RTL support."""
        pw = self.PANEL_WIDTH
        margin = 10
        available = pw - 2 * margin
        y1, y2, y3 = 14, 34, 54
        gap = 6

        words1 = self._cur_line1.split() if self._cur_line1.strip() else []
        words2 = self._cur_line2.split() if self._cur_line2.strip() else []
        is_rtl = self._is_rtl_text(self._cur_line1)

        # ── helper: compute total width ──
        def _total(words, fm):
            if not words:
                return 0
            return sum(fm.horizontalAdvance(w) for w in words) + gap * (len(words) - 1)

        # ── helper: auto-scroll shift for a line ──
        def _auto_shift(total_w):
            if total_w <= available:
                return 0
            return total_w - available  # pixels to skip so latest text is visible

        # Clip text area
        p.save()
        p.setClipRect(QRectF(margin, 0, available, self.WIDGET_SIZE))

        # ── Line 1: actual text ──
        font_normal = QFont('Segoe UI', 10)
        font_corrected = QFont('Segoe UI', 10)
        font_corrected.setUnderline(True)
        fm1 = QFontMetrics(font_normal)
        total1 = _total(words1, fm1)
        shift1 = _auto_shift(total1) + self._scroll_x
        shift1 = max(0, min(shift1, max(0, total1 - available)))

        if is_rtl:
            x = pw - margin
            for word in words1:
                is_corr = word in self._corrections
                p.setFont(font_corrected if is_corr else font_normal)
                p.setPen(QColor(COLORS['accent']))
                ww = QFontMetrics(p.font()).horizontalAdvance(word)
                draw_x = x - ww + shift1
                p.drawText(int(draw_x), y1, word)
                x -= ww + gap
        else:
            x = margin
            for word in words1:
                is_corr = word in self._corrections
                p.setFont(font_corrected if is_corr else font_normal)
                p.setPen(QColor(COLORS['accent']))
                ww = QFontMetrics(p.font()).horizontalAdvance(word)
                draw_x = x - shift1
                p.drawText(int(draw_x), y1, word)
                x += ww + gap

        # ── Line 2: translated text (store rects for click-to-replace) ──
        font2 = QFont('Segoe UI', 9)
        p.setFont(font2)
        p.setPen(QColor(COLORS['translate']))
        fm2 = QFontMetrics(font2)
        total2 = _total(words2, fm2)
        shift2 = _auto_shift(total2) + self._scroll_x
        shift2 = max(0, min(shift2, max(0, total2 - available)))
        rects = []

        if is_rtl:
            x = pw - margin
            for i, word in enumerate(words2):
                ww = fm2.horizontalAdvance(word)
                draw_x = x - ww + shift2
                p.drawText(int(draw_x), y2, word)
                rects.append((i, int(draw_x), int(draw_x + ww), word))
                x -= ww + gap
        else:
            x = margin
            for i, word in enumerate(words2):
                ww = fm2.horizontalAdvance(word)
                draw_x = x - shift2
                p.drawText(int(draw_x), y2, word)
                rects.append((i, int(draw_x), int(draw_x + ww), word))
                x += ww + gap
        self._line2_word_rects = rects

        p.restore()  # remove clip

        # ── Line 3: NLP scores (unchanged, always right-aligned, truncated) ──
        if not self._hide_scores:
            font_score = QFont('Segoe UI', 8)
            p.setFont(font_score)
            p.setPen(QColor('#8899aa'))
            score_parts = []
            orig_words = self._cur_line1.split() if self._cur_line1.strip() else []
            for w in orig_words:
                ws = self._word_scores.get(w)
                if ws:
                    pairs = ' '.join(f'{k}:{v}' for k, v in ws.items())
                    score_parts.append(pairs)
            if score_parts:
                score_text = ' | '.join(score_parts)
                if len(score_text) > 40:
                    score_text = score_text[-40:]
                sw = QFontMetrics(font_score).horizontalAdvance(score_text)
                p.drawText(int(pw - margin - sw), y3, score_text)

    # ── hit testing ──────────────────────────────────────

    def _hit_test_line2(self, click_x, click_y=None):
        """Check if click hits a word on line 2 (translated text).
        Uses stored rects from the last paint pass.
        Returns (word_index, word_str) or None."""
        if click_y is not None and not (18 <= click_y <= 52):
            return None
        margin = 10
        pw = self.PANEL_WIDTH
        if click_x < margin or click_x > pw - margin:
            return None
        for (idx, x_left, x_right, word) in self._line2_word_rects:
            if x_left <= click_x <= x_right:
                return (idx, word)
        return None

    def _get_quadrant_at(self, x, y):
        """Return the quadrant tag at widget coordinates, or None."""
        icon_x = self.PANEL_WIDTH if self._panel_visible else 0
        sz = self.WIDGET_SIZE
        cx = icon_x + sz // 2
        cy = sz // 2
        r = sz // 2 - 4
        r_center = 13

        dx = x - cx
        dy = y - cy
        dist = math.sqrt(dx * dx + dy * dy)

        if dist <= r_center:
            return 'q_center'
        if dist > r:
            return None

        angle = math.degrees(math.atan2(-dy, dx)) % 360
        if 45 <= angle < 135:
            return 'q_top'
        elif 135 <= angle < 225:
            return 'q_left'
        elif 225 <= angle < 315:
            return 'q_bottom'
        else:
            return 'q_right'

    # ── mouse events ─────────────────────────────────────

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start = event.globalPosition().toPoint()
            self._drag_moved = False
            # If clicking on panel area, set flag so engine mouse hook won't reset
            pos = event.position().toPoint()
            if self._panel_visible and pos.x() < self.PANEL_WIDTH and pos.y() < self.WIDGET_SIZE:
                try:
                    import auto_lang
                    auto_lang._widget_clicking.set()
                except Exception:
                    pass

    def mouseMoveEvent(self, event):
        pos = event.position().toPoint()
        # Handle dragging
        if self._drag_start is not None and event.buttons() & Qt.LeftButton:
            delta = event.globalPosition().toPoint() - self._drag_start
            if delta.manhattanLength() > 3:
                self._drag_moved = True
                self.move(self.pos() + delta)
                self._drag_start = event.globalPosition().toPoint()
            return

        # Handle resize grip dragging
        if self._resize_drag_y is not None and event.buttons() & Qt.LeftButton:
            dy = event.globalPosition().toPoint().y() - self._resize_drag_y
            self._resize_drag_y = event.globalPosition().toPoint().y()
            new_eh = max(self._editor_min_h, min(self._editor_max_h, self._editor_height + dy))
            if new_eh != self._editor_height:
                self._editor_height = new_eh
                self._relayout()
            return

        # Hover tracking
        px, py = pos.x(), pos.y()

        # Check for line 2 word hover on panel (show hand cursor)
        if self._panel_visible and px < self.PANEL_WIDTH and py < self.WIDGET_SIZE:
            if 18 <= py <= 52 and self._hit_test_line2(px, py) is not None:
                self.setCursor(Qt.PointingHandCursor)
                self._hovered_quadrant = None
                self._tooltip_timer.stop()
                self._hide_tooltip()
                return

        q = self._get_quadrant_at(px, py)
        if q != self._hovered_quadrant:
            self._hovered_quadrant = q
            self.setCursor(Qt.PointingHandCursor if q else Qt.ArrowCursor)
            # Tooltip
            if q:
                self._tooltip_tag = q
                self._tooltip_timer.start(800)
            else:
                self._tooltip_timer.stop()
                self._hide_tooltip()
            self.update()

        # Resize cursor at bottom edge
        if self._editor_height > 0:
            total_h = self.WIDGET_SIZE + self._editor_height
            if pos.y() > total_h - 14:
                self.setCursor(Qt.SizeVerCursor)
            elif not q:
                self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        # Always clear the widget_clicking flag
        try:
            import auto_lang
            auto_lang._widget_clicking.clear()
        except Exception:
            pass

        if event.button() == Qt.LeftButton:
            was_dragging = self._drag_moved
            self._drag_start = None
            self._resize_drag_y = None

            if was_dragging:
                return

            pos = event.position().toPoint()
            x, y = pos.x(), pos.y()

            # Check resize grip
            if self._editor_height > 0:
                total_h = self.WIDGET_SIZE + self._editor_height
                if y > total_h - 14:
                    self._resize_drag_y = event.globalPosition().toPoint().y()
                    return

            # Click on panel area — check for word click first, then toggle editor
            if self._panel_visible and x < self.PANEL_WIDTH and y < self.WIDGET_SIZE:
                # Hit-test line 2 words on-the-fly (any click on panel)
                hit = self._hit_test_line2(x, y)
                if hit is not None:
                    idx, word = hit
                    self._on_translated_word_click(idx, word)
                    return
                self._toggle_edit_mode()
                return

            # Click on quadrant
            q = self._get_quadrant_at(x, y)
            if q == 'q_center':
                self._toggle_speech()
            elif q == 'q_top':
                self._toggle()
            elif q == 'q_left':
                self._undo()
            elif q == 'q_right':
                self._open_settings()
            elif q == 'q_bottom':
                self._toggle_edit_mode()
            else:
                self._show_context_menu(event.globalPosition().toPoint())

    def wheelEvent(self, event):
        """Horizontal scroll through panel text when content overflows."""
        if not self._panel_visible:
            return super().wheelEvent(event)
        pos = event.position().toPoint()
        if pos.x() >= self.PANEL_WIDTH or pos.y() >= self.WIDGET_SIZE:
            return super().wheelEvent(event)
        delta = event.angleDelta().y()  # positive = scroll up / right, negative = down / left
        step = 30
        if delta > 0:
            self._scroll_x = max(self._scroll_x - step, -9999)
        else:
            self._scroll_x = min(self._scroll_x + step, 9999)
        self.update()

    def contextMenuEvent(self, event):
        self._show_context_menu(event.globalPos())

    # ── context menu ─────────────────────────────────────

    def _build_menu(self):
        self._menu = QMenu(self)
        self._menu.setStyleSheet(f"""
            QMenu {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']};
                padding: 4px;
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
            QMenu::item {{
                padding: 6px 16px;
            }}
            QMenu::item:selected {{
                background-color: {COLORS['accent_dim']};
                color: #ffffff;
            }}
            QMenu::separator {{
                height: 1px;
                background-color: {COLORS['border']};
                margin: 4px 8px;
            }}
        """)
        self._toggle_action = self._menu.addAction('\u2705  \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4 \u05e4\u05e2\u05d9\u05dc')
        self._toggle_action.triggered.connect(self._toggle)
        self._menu.addSeparator()
        self._panel_action = self._menu.addAction('\U0001f50d  \u05d4\u05e6\u05d2/\u05d4\u05e1\u05ea\u05e8 \u05ea\u05d9\u05d1\u05ea \u05d4\u05e7\u05dc\u05d3\u05d4')
        self._panel_action.triggered.connect(self._toggle_panel)
        self._menu.addSeparator()
        self._scores_action = self._menu.addAction('\U0001f4ca  \u05d4\u05e6\u05d2/\u05d4\u05e1\u05ea\u05e8 \u05e6\u05d9\u05d5\u05e0\u05d9\u05dd')
        self._scores_action.triggered.connect(self._toggle_scores)
        self._menu.addSeparator()
        self._menu.addAction('\u23ea  \u05d1\u05d8\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df  (F12)', self._undo)
        self._menu.addSeparator()
        self._menu.addAction('\u2699\ufe0f  \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea', self._open_settings)
        self._menu.addSeparator()
        self._menu.addAction('\u2753  \u05de\u05d3\u05e8\u05d9\u05da \u05dc\u05de\u05e9\u05ea\u05de\u05e9', self._show_help)
        self._menu.addSeparator()
        self._menu.addAction('\U0001f6aa  \u05d9\u05e6\u05d9\u05d0\u05d4', self._quit)

    def _refresh_menu_labels(self):
        if self.tray.enabled:
            self._toggle_action.setText('\u2705  \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4 \u05e4\u05e2\u05d9\u05dc')
        else:
            self._toggle_action.setText('\u274c  \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4 \u05de\u05d5\u05e9\u05d1\u05ea')
        if self._panel_visible:
            self._panel_action.setText('\U0001f50d  \u05d4\u05e1\u05ea\u05e8 \u05ea\u05d9\u05d1\u05ea \u05d4\u05e7\u05dc\u05d3\u05d4')
        else:
            self._panel_action.setText('\U0001f50d  \u05d4\u05e6\u05d2 \u05ea\u05d9\u05d1\u05ea \u05d4\u05e7\u05dc\u05d3\u05d4')
        if self._hide_scores:
            self._scores_action.setText('\U0001f4ca  \u05d4\u05e6\u05d2 \u05e6\u05d9\u05d5\u05e0\u05d9\u05dd')
        else:
            self._scores_action.setText('\U0001f4ca  \u05d4\u05e1\u05ea\u05e8 \u05e6\u05d9\u05d5\u05e0\u05d9\u05dd')

    def _show_context_menu(self, global_pos):
        self._refresh_menu_labels()
        self._menu.popup(global_pos)

    # ── tooltip ──────────────────────────────────────────

    def _do_show_tooltip(self):
        text = self._tooltip_texts.get(self._tooltip_tag, '')
        if not text:
            return
        self._hide_tooltip()
        tw = QWidget(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.ToolTip)
        tw.setAttribute(Qt.WA_TranslucentBackground)
        layout = QVBoxLayout(tw)
        layout.setContentsMargins(1, 1, 1, 1)
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_dark']};
                border: 1px solid {COLORS['border']};
                border-radius: 4px;
            }}
        """)
        fl = QHBoxLayout(frame)
        fl.setContentsMargins(8, 4, 8, 4)
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 9pt; background: transparent;")
        fl.addWidget(lbl)
        layout.addWidget(frame)
        tw.adjustSize()

        x = self.mapToGlobal(QPoint(self.width() // 2 - 60, self.WIDGET_SIZE + 4))
        tw.move(x)
        tw.show()
        self._tooltip_win = tw
        QTimer.singleShot(3000, self._hide_tooltip)

    def _hide_tooltip(self):
        if self._tooltip_win:
            self._tooltip_win.close()
            self._tooltip_win.deleteLater()
            self._tooltip_win = None

    # ── spell tooltip ────────────────────────────────────

    def _on_spell_result(self, word, suggestions, mode):
        """Thread-safe callback from spell_module (called from hook thread)."""
        if word is None:  # dismiss
            self._spell_signal.emit(None, None, 'dismiss')
            return
        if suggestions:
            self._spell_signal.emit(word, suggestions, mode)

    def _show_spell_tooltip(self, word, suggestions, mode):
        """Show a spell suggestion tooltip near the widget (runs on Qt thread)."""
        self._hide_spell_tooltip()

        if not word or not suggestions or mode == 'dismiss':
            return
        if mode == 'visual':
            # Just update panel with underline info (future enhancement)
            self.update()
            return

        # Build tooltip (tooltip mode)
        tw = QWidget(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.ToolTip)
        tw.setAttribute(Qt.WA_TranslucentBackground)
        outer = QVBoxLayout(tw)
        outer.setContentsMargins(1, 1, 1, 1)

        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_dark']};
                border: 1px solid {COLORS['accent']};
                border-radius: 6px;
            }}
        """)
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(10, 6, 10, 6)

        # Header
        hdr = QLabel(f'\U0001f4dd \u05ea\u05d9\u05e7\u05d5\u05df \u05db\u05ea\u05d9\u05d1:  {word}')
        hdr.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
        hdr.setAlignment(Qt.AlignRight)
        fl.addWidget(hdr)

        # Show top suggestion prominently
        top = suggestions[0] if suggestions else ''
        sugg_lbl = QLabel(f'\u27a1 {top}')
        sugg_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 12pt; font-weight: bold; background: transparent;")
        sugg_lbl.setAlignment(Qt.AlignRight)
        fl.addWidget(sugg_lbl)

        # More suggestions (if any)
        if len(suggestions) > 1:
            others = ', '.join(suggestions[1:4])
            more_lbl = QLabel(others)
            more_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 8pt; background: transparent;")
            more_lbl.setAlignment(Qt.AlignRight)
            fl.addWidget(more_lbl)

        # Hint
        hint = QLabel('Tab \u2190 \u05dc\u05e7\u05d1\u05dc')
        hint.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 8pt; background: transparent;")
        hint.setAlignment(Qt.AlignRight)
        fl.addWidget(hint)

        outer.addWidget(frame)
        tw.adjustSize()

        # Position above the widget
        pos = self.mapToGlobal(QPoint(self.width() // 2 - tw.width() // 2,
                                      -tw.height() - 4))
        tw.move(pos)
        tw.show()
        self._spell_tooltip = tw
        self._spell_hide_timer.start()

    def _hide_spell_tooltip(self):
        self._spell_hide_timer.stop()
        if self._spell_tooltip:
            self._spell_tooltip.close()
            self._spell_tooltip.deleteLater()
            self._spell_tooltip = None

    # ── grammar result ───────────────────────────────────

    def _on_grammar_result(self, original, corrected, error):
        """Thread-safe callback from grammar_module (called from bg thread)."""
        self._grammar_signal.emit(
            original or '',
            corrected if corrected else '',
            error if error else '',
        )

    def _show_grammar_result(self, original, corrected, error):
        """Show grammar correction result popup (runs on Qt thread)."""
        self._hide_grammar_popup()

        popup = QWidget(None, Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        popup.setAttribute(Qt.WA_TranslucentBackground)
        popup.setMinimumWidth(400)

        outer = QVBoxLayout(popup)
        outer.setContentsMargins(2, 2, 2, 2)

        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_dark']};
                border: 2px solid {COLORS['accent']};
                border-radius: 8px;
            }}
        """)
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(14, 10, 14, 10)

        # Title
        title = QLabel('\U0001f4ac \u05ea\u05d9\u05e7\u05d5\u05df \u05e0\u05d9\u05e1\u05d5\u05d7 / \u05d3\u05e7\u05d3\u05d5\u05e7')
        title.setStyleSheet(f"color: {COLORS['accent']}; font-size: 12pt; font-weight: bold; background: transparent;")
        title.setAlignment(Qt.AlignRight)
        fl.addWidget(title)

        if error:
            err_lbl = QLabel(f'\u274c {error}')
            err_lbl.setStyleSheet(f"color: #ff6b6b; font-size: 10pt; background: transparent;")
            err_lbl.setWordWrap(True)
            err_lbl.setAlignment(Qt.AlignRight)
            fl.addWidget(err_lbl)
        elif corrected:
            if corrected.strip() == original.strip():
                ok_lbl = QLabel('\u2705 \u05d4\u05d8\u05e7\u05e1\u05d8 \u05ea\u05e7\u05d9\u05df \u2014 \u05dc\u05dc\u05d0 \u05e9\u05d2\u05d9\u05d0\u05d5\u05ea')
                ok_lbl.setStyleSheet(f"color: {COLORS['success']}; font-size: 10pt; background: transparent;")
                ok_lbl.setAlignment(Qt.AlignRight)
                fl.addWidget(ok_lbl)
            else:
                # Show original
                orig_lbl = QLabel('\u05de\u05e7\u05d5\u05e8:')
                orig_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
                orig_lbl.setAlignment(Qt.AlignRight)
                fl.addWidget(orig_lbl)
                orig_text = QLabel(original[:300])
                orig_text.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 10pt; background: transparent; "
                                         f"text-decoration: line-through;")
                orig_text.setWordWrap(True)
                orig_text.setAlignment(Qt.AlignRight)
                fl.addWidget(orig_text)

                # Show corrected
                corr_lbl = QLabel('\u05de\u05ea\u05d5\u05e7\u05df:')
                corr_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
                corr_lbl.setAlignment(Qt.AlignRight)
                fl.addWidget(corr_lbl)
                corr_text = QLabel(corrected[:300])
                corr_text.setStyleSheet(f"color: {COLORS['accent']}; font-size: 10pt; font-weight: bold; background: transparent;")
                corr_text.setWordWrap(True)
                corr_text.setAlignment(Qt.AlignRight)
                fl.addWidget(corr_text)

                # Buttons
                btn_row = QHBoxLayout()
                btn_row.addStretch()

                dismiss_btn = QPushButton('\u274c \u05d1\u05d9\u05d8\u05d5\u05dc')
                dismiss_btn.setStyleSheet(f"padding: 4px 12px; font-size: 9pt;")
                dismiss_btn.clicked.connect(self._hide_grammar_popup)
                btn_row.addWidget(dismiss_btn)

                accept_btn = QPushButton('\u2705 \u05d4\u05d7\u05dc\u05e3')
                accept_btn.setStyleSheet(f"padding: 4px 12px; font-size: 9pt; "
                                          f"background-color: {COLORS['accent']}; color: white; font-weight: bold;")
                accept_btn.clicked.connect(lambda: self._accept_grammar(corrected))
                btn_row.addWidget(accept_btn)

                fl.addLayout(btn_row)

        outer.addWidget(frame)
        popup.adjustSize()

        # Position near widget
        screen = QGuiApplication.primaryScreen().availableGeometry()
        pos = self.mapToGlobal(QPoint(
            self.width() // 2 - popup.width() // 2,
            -popup.height() - 8
        ))
        # Keep on screen
        if pos.x() < 0:
            pos.setX(4)
        if pos.y() < 0:
            pos.setY(screen.height() - popup.height() - 60)
        popup.move(pos)
        popup.show()
        self._grammar_popup = popup

        # Auto-hide after 30s if no action
        QTimer.singleShot(30000, self._hide_grammar_popup)

    def _accept_grammar(self, corrected_text):
        """Paste the corrected text (replaces selected text)."""
        self._hide_grammar_popup()
        import threading
        def _paste():
            import time
            try:
                # Put corrected text on clipboard
                import ctypes as _ct
                _user32 = _ct.windll.user32
                _kernel32 = _ct.windll.kernel32
                if _user32.OpenClipboard(0):
                    try:
                        _user32.EmptyClipboard()
                        data = corrected_text.encode('utf-16-le') + b'\\x00\\x00'
                        h = _kernel32.GlobalAlloc(0x0042, len(data))
                        _kernel32.GlobalLock.restype = _ct.c_void_p
                        ptr = _kernel32.GlobalLock(h)
                        _ct.memmove(ptr, data, len(data))
                        _kernel32.GlobalUnlock(h)
                        _user32.SetClipboardData(13, h)  # CF_UNICODETEXT
                    finally:
                        _user32.CloseClipboard()
                time.sleep(0.1)
                import keyboard as _kb
                _kb.send('ctrl+v')
            except Exception:
                pass
        threading.Thread(target=_paste, daemon=True).start()

    def _hide_grammar_popup(self):
        if self._grammar_popup:
            self._grammar_popup.close()
            self._grammar_popup.deleteLater()
            self._grammar_popup = None

    # ── panel toggle ─────────────────────────────────────

    @staticmethod
    def _is_rtl_text(text: str) -> bool:
        """Return True if text starts with Hebrew/Arabic characters (RTL)."""
        for ch in text:
            if '\u0590' <= ch <= '\u05FF' or '\u0600' <= ch <= '\u06FF':
                return True
            if 'A' <= ch <= 'Z' or 'a' <= ch <= 'z':
                return False
        return False

    def _set_panel_text(self, actual, translated, corrections, word_scores):
        self._cur_line1 = actual if actual else ' '
        self._cur_line2 = translated if translated else ' '
        self._corrections = corrections or {}
        self._word_scores = word_scores or {}
        self._scroll_x = 0   # auto-scroll to latest text
        self.update()

    def _toggle_panel(self):
        self._panel_visible = not self._panel_visible
        self.tray.config['show_typing_panel'] = self._panel_visible
        save_config(self.tray.config)
        self._relayout()

    def _toggle_scores(self):
        self._hide_scores = not self._hide_scores
        self.tray.config['hide_scores'] = self._hide_scores
        save_config(self.tray.config)
        self.update()

    def _relayout(self):
        """Recalculate widget size and reposition editor widgets."""
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = pw + sz
        total_h = sz + self._editor_height

        # Keep right edge stable when toggling panel
        old_right = self.x() + self.width()
        self.setFixedSize(total_w, total_h)
        self.move(old_right - total_w, self.y())

        # Reposition editor widgets
        if self._editor_height > 0 and self._editor_container:
            self._editor_container.setGeometry(0, sz, total_w, self._editor_height)
        self.update()

    # ── state ────────────────────────────────────────────

    def _apply_state(self, enabled: bool):
        if enabled:
            self.show()
        else:
            if self._edit_mode:
                self._edit_mode = False
                self._hide_edit_window()
            self.hide()
        self.update()

    # ── actions ──────────────────────────────────────────

    def _toggle(self):
        self.tray._toggle_enabled()
        self._apply_state(self.tray.enabled)

    def _undo(self):
        try:
            import auto_lang
            auto_lang._undo_last_correction()
        except Exception:
            pass

    def _on_translated_word_click(self, word_index: int, translated_word: str):
        """User clicked a word on line 2 (translated).  Replace the
        corresponding actual word with this translated version."""
        words1 = self._cur_line1.split() if self._cur_line1.strip() else []
        if word_index < 0 or word_index >= len(words1):
            return
        actual_word = words1[word_index]
        if actual_word == translated_word:
            return  # nothing to do
        try:
            import auto_lang
            import threading
            # Snapshot state NOW (before engine mouse hook fires and resets)
            with auto_lang.state.lock:
                snap_buf = auto_lang.state.buffer
            snap_parts = list(auto_lang._prev_words)
            if snap_buf:
                snap_parts.append(snap_buf)
            threading.Thread(
                target=auto_lang._replace_panel_word,
                args=(word_index, actual_word, translated_word,
                      snap_parts, snap_buf),
                daemon=True,
            ).start()
        except Exception:
            pass

    def _open_settings(self):
        self.tray._open_settings()

    def _quit(self):
        self.tray._quit()

    # ── text editor ──────────────────────────────────────

    def _toggle_edit_mode(self):
        self._edit_mode = not self._edit_mode
        if self._edit_mode:
            self._trans_visible = True
            self._show_edit_window()
        else:
            self._hide_edit_window()

    def _show_edit_window(self):
        if self._editor_height > 0 and self._editor_container and self._text_widget and self._trans_widget:
            try:
                self._editor_container.show()
                self.show()
            except Exception:
                pass
            return

        try:
            import auto_lang
            self._send_target_hwnd = auto_lang._last_target_hwnd
        except Exception:
            self._send_target_hwnd = None

        try:
            import translator
            translator.ensure_models_loaded(callback=lambda: self._models_ready_signal.emit())
        except Exception as e:
            print(f'[UI] translator import failed: {e}')

        if self._editor_height == 0:
            self._editor_height = 250
        eh = self._editor_height
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = pw + sz
        total_h = sz + eh

        # Move up if would go below screen
        screen = QGuiApplication.primaryScreen().availableGeometry()
        cur_y = self.y()
        if cur_y + total_h > screen.height():
            cur_y = max(0, screen.height() - total_h)
            self.move(self.x(), cur_y)

        self.setFixedSize(total_w, total_h)
        self._create_editor_widgets()

        # Focus
        self.activateWindow()
        self.raise_()
        if self._text_widget:
            self._text_widget.setFocus()

        try:
            import auto_lang
            auto_lang._widget_edit_hwnd = int(self.winId())
        except Exception:
            pass

    def _create_editor_widgets(self):
        """Create the editor child widgets inside the widget."""
        if self._editor_container:
            self._editor_container.deleteLater()

        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = pw + sz
        eh = self._editor_height

        container = QWidget(self)
        container.setGeometry(0, sz, total_w, eh)
        container.setStyleSheet("background: transparent;")

        layout = QVBoxLayout(container)
        layout.setContentsMargins(14, 6, 14, 14)
        layout.setSpacing(4)

        # Input
        self._text_widget = _TranslatorInput()
        self._text_widget.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_dark']};
                color: {COLORS['translate']};
                border: none;
                border-radius: 4px;
                padding: 4px 6px;
                font-family: 'Segoe UI';
                font-size: 11pt;
            }}
        """)
        self._text_widget.setPlaceholderText('\u05d4\u05e7\u05dc\u05d3 \u05d8\u05e7\u05e1\u05d8 \u05dc\u05ea\u05e8\u05d2\u05d5\u05dd...')
        self._text_widget.enter_pressed.connect(self._do_translate)
        self._text_widget.boundary_typed.connect(self._on_boundary_typed)
        layout.addWidget(self._text_widget, 1)

        # Translation container (can be shown/hidden independently)
        self._trans_container = QWidget()
        trans_layout = QVBoxLayout(self._trans_container)
        trans_layout.setContentsMargins(0, 0, 0, 0)
        trans_layout.setSpacing(4)

        self._trans_sep = QFrame()
        self._trans_sep.setFixedHeight(1)
        self._trans_sep.setStyleSheet(f"background-color: {COLORS['accent']}; border: none;")
        trans_layout.addWidget(self._trans_sep)

        # Translation output
        self._trans_widget = QTextEdit()
        self._trans_widget.setReadOnly(True)
        self._trans_widget.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_dark']};
                color: {COLORS['accent']};
                border: none;
                border-radius: 4px;
                padding: 4px 6px;
                font-family: 'Segoe UI';
                font-size: 11pt;
            }}
        """)
        trans_layout.addWidget(self._trans_widget, 1)

        layout.addWidget(self._trans_container, 1)

        # Bottom bar
        bar = QWidget()
        bar.setFixedHeight(24)
        bar.setStyleSheet("background: transparent;")
        bar_layout = QHBoxLayout(bar)
        bar_layout.setContentsMargins(0, 0, 0, 0)

        self._send_btn = QPushButton('\U0001f4e4 \u05e9\u05dc\u05d7')
        self._send_btn.setFixedSize(60, 22)
        self._send_btn.setCursor(Qt.PointingHandCursor)
        self._send_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['accent_dim']};
                color: #ffffff;
                border: none;
                border-radius: 3px;
                font-weight: bold;
                font-size: 9pt;
            }}
            QPushButton:hover {{
                background-color: {COLORS['accent']};
            }}
        """)
        self._send_btn.clicked.connect(self._send_translation)

        self._status_label = QLabel('\u2705 \u05de\u05d5\u05db\u05df \u05dc\u05ea\u05e8\u05d2\u05d5\u05dd')
        self._status_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 8pt; background: transparent;")

        self._toggle_trans_btn = QPushButton('\u05d4\u05e1\u05ea\u05e8')
        self._toggle_trans_btn.setFixedSize(44, 22)
        self._toggle_trans_btn.setCursor(Qt.PointingHandCursor)
        self._toggle_trans_btn.setFocusPolicy(Qt.NoFocus)
        self._toggle_trans_btn.setAutoDefault(False)
        self._toggle_trans_btn.setDefault(False)
        self._toggle_trans_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {COLORS['text_muted']};
                border: 1px solid {COLORS['border']};
                border-radius: 3px;
                font-size: 8pt;
            }}
            QPushButton:hover {{
                color: {COLORS['accent']};
                border-color: {COLORS['accent']};
            }}
        """)
        self._toggle_trans_btn.clicked.connect(self._toggle_translation_box)

        bar_layout.addWidget(self._status_label)
        bar_layout.addStretch()
        bar_layout.addWidget(self._toggle_trans_btn)
        bar_layout.addWidget(self._send_btn)
        layout.addWidget(bar)

        container.show()
        self._editor_container = container

        self._set_translation_visible(self._trans_visible)

        # Escape key
        self._text_widget.installEventFilter(self)
        if self._trans_widget:
            self._trans_widget.installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj == self._text_widget and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key_Escape:
                self._toggle_edit_mode()
                return True
        if obj == self._trans_widget and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key_Escape:
                self._set_translation_visible(False)
                return True
        return super().eventFilter(obj, event)

    def _toggle_translation_box(self):
        self._set_translation_visible(not self._trans_visible)

    def _set_translation_visible(self, visible: bool):
        self._trans_visible = bool(visible)
        if self._trans_container:
            self._trans_container.setVisible(self._trans_visible)
        if self._toggle_trans_btn:
            self._toggle_trans_btn.setText('\u05d4\u05e1\u05ea\u05e8' if self._trans_visible else '\u05d4\u05e6\u05d2')
        if self._send_btn:
            self._send_btn.setEnabled(self._trans_visible)

    def _hide_edit_window(self):
        self._edit_mode = False
        self._translate_debounce.stop()

        try:
            import auto_lang
            auto_lang._widget_edit_hwnd = 0
        except Exception:
            pass

        if self._editor_container:
            self._editor_container.deleteLater()
            self._editor_container = None
        self._text_widget = None
        self._trans_widget = None
        self._trans_container = None
        self._trans_sep = None
        self._status_label = None
        self._send_btn = None
        self._toggle_trans_btn = None

        self._editor_height = 0
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        self.setFixedSize(pw + sz, sz)
        self.update()

    def _on_boundary_typed(self):
        self._translate_debounce.start()

    def _do_translate(self):
        if self._text_widget is None:
            return
        text = self._text_widget.toPlainText().strip()
        if not text:
            return
        if self._status_label:
            self._status_label.setText('\u23f3 \u05de\u05ea\u05e8\u05d2\u05dd...')
        threading.Thread(target=self._translate_worker, args=(text,), daemon=True).start()

    def _translate_worker(self, text: str):
        try:
            import translator
            result = translator.translate(text)
            if result is None:
                self._translate_result_signal.emit(
                    '\u274c \u05ea\u05e8\u05d2\u05d5\u05dd \u05e0\u05db\u05e9\u05dc', True, '')
            else:
                source = 'Google' if translator._online else 'Argos'
                self._translate_result_signal.emit(result, False, source)
        except Exception as e:
            self._translate_result_signal.emit(
                f'\u05e9\u05d2\u05d9\u05d0\u05d4: {e}', True, '')

    def _show_translation_result(self, text: str, error: bool = False, source: str = ''):
        if self._editor_height == 0 or self._text_widget is None or self._trans_widget is None:
            self._show_edit_window()
        self._set_translation_visible(True)
        if self._trans_widget is None:
            return
        self._trans_widget.setPlainText(text)
        color = COLORS['error'] if error else COLORS['translate']
        self._trans_widget.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_dark']};
                color: {color};
                border: none;
                border-radius: 4px;
                padding: 4px 6px;
                font-family: 'Segoe UI';
                font-size: 11pt;
            }}
        """)
        if self._status_label:
            if error:
                self._status_label.setText('\u274c \u05e9\u05d2\u05d9\u05d0\u05d4')
            elif self._text_widget:
                try:
                    import translator
                    pair = translator.detect_direction(self._text_widget.toPlainText().strip())
                    direction = 'EN \u2192 \u05e2\u05d1' if pair == 'en-he' else '\u05e2\u05d1 \u2192 EN'
                    src_label = f' ({source})' if source else ''
                    self._status_label.setText(f'\u2705 {direction}{src_label}')
                except Exception:
                    self._status_label.setText('\u2705 \u05ea\u05d5\u05e8\u05d2\u05dd')

    def _update_ready_status(self):
        if self._status_label:
            self._status_label.setText('\u2705 \u05de\u05d5\u05db\u05df \u05dc\u05ea\u05e8\u05d2\u05d5\u05dd')

    def _send_translation(self):
        if self._trans_widget is None:
            return
        translated = self._trans_widget.toPlainText().strip()
        if not translated:
            self._do_translate()
            return
        if translated.startswith('\u274c') or translated.startswith('\u05e9\u05d2\u05d9\u05d0\u05d4'):
            return

        self._edit_mode = False
        self._hide_edit_window()
        QTimer.singleShot(100, lambda t=translated: threading.Thread(
            target=self._paste_to_target, args=(t,), daemon=True).start())

    def _paste_to_target(self, text: str):
        try:
            import auto_lang
            import ctypes
            send_t = self._send_target_hwnd
            last_t = auto_lang._last_target_hwnd
            hwnd = send_t or last_t
            if not hwnd:
                hwnd = self._find_target_hwnd()
            if not hwnd:
                return
            ctypes.windll.user32.AllowSetForegroundWindow(-1)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            auto_lang.injecting.set()
            try:
                auto_lang._paste_text(text)
            finally:
                time.sleep(0.05)
                auto_lang.injecting.clear()
        except Exception as e:
            print(f'[Translator] Paste failed: {e}')

    def _start_right_click_hook(self):
        """WinAPI low-level hook to catch right-clicks even if mouse lib fails."""
        if self._rc_hook_thread is not None:
            return

        def _thread():
            try:
                import ctypes
                from ctypes import wintypes

                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32
                WH_MOUSE_LL = 14
                WM_RBUTTONUP = 0x0205

                LowLevelMouseProc = ctypes.WINFUNCTYPE(
                    ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
                )

                def _proc(nCode, wParam, lParam):
                    try:
                        if nCode == 0 and wParam == WM_RBUTTONUP:
                            now = time.time()
                            if now - self._last_rc_time > 0.4:
                                self._last_rc_time = now
                                _ui_log('[RC] winapi hook fired')
                                self._rc_signal.emit()
                    except Exception:
                        pass
                    return user32.CallNextHookEx(self._rc_hook_id, nCode, wParam, lParam)

                self._rc_hook_proc = LowLevelMouseProc(_proc)
                hmod = kernel32.GetModuleHandleW(None)
                self._rc_hook_id = user32.SetWindowsHookExW(
                    WH_MOUSE_LL, self._rc_hook_proc, hmod, 0
                )
                if not self._rc_hook_id:
                    _ui_log(f'[RC] SetWindowsHookExW failed hmod={hmod} err={ctypes.get_last_error()}')
                    self._rc_hook_id = user32.SetWindowsHookExW(
                        WH_MOUSE_LL, self._rc_hook_proc, 0, 0
                    )
                    if not self._rc_hook_id:
                        _ui_log(f'[RC] SetWindowsHookExW failed hmod=0 err={ctypes.get_last_error()}')
                        return
                _ui_log(f'[RC] WinAPI hook installed id={self._rc_hook_id}')

                msg = wintypes.MSG()
                while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
            except Exception:
                _ui_log('[RC] WinAPI hook thread crashed')
                pass

        self._rc_hook_thread = threading.Thread(target=_thread, daemon=True)
        self._rc_hook_thread.start()

    # ── Right-click translate popup ──────────────────────

    def _on_right_click_global(self):
        x = y = None
        try:
            import mouse as _mouse
            x, y = _mouse.get_position()
        except Exception:
            try:
                pos = QCursor.pos()
                x, y = pos.x(), pos.y()
            except Exception:
                try:
                    import ctypes
                    from ctypes import wintypes
                    pt = wintypes.POINT()
                    if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
                        x, y = pt.x, pt.y
                except Exception:
                    pass
        if x is None or y is None:
            return
        self._last_rc_pos = (int(x), int(y))
        self._rc_hwnd_point = None
        self._rc_hwnd_root = None

        # Prefer window under cursor as target (more accurate than last focus)
        try:
            import ctypes
            from ctypes import wintypes
            pt = wintypes.POINT(int(x), int(y))
            hwnd_point = ctypes.windll.user32.WindowFromPoint(pt)
            if hwnd_point:
                GA_ROOT = 2
                hwnd_root = ctypes.windll.user32.GetAncestor(hwnd_point, GA_ROOT)
                self._rc_hwnd_point = hwnd_point
                self._rc_hwnd_root = hwnd_root or hwnd_point
                self._send_target_hwnd = self._rc_hwnd_root
                _ui_log(f'[RC] hwnd under cursor={int(hwnd_point)} root={int(self._rc_hwnd_root)}')
        except Exception:
            pass

        try:
            import auto_lang
            if not self._send_target_hwnd:
                self._send_target_hwnd = auto_lang._last_target_hwnd
        except Exception:
            self._send_target_hwnd = None
        if not self._send_target_hwnd:
            try:
                import ctypes
                self._send_target_hwnd = ctypes.windll.user32.GetForegroundWindow()
            except Exception:
                self._send_target_hwnd = None
        if not self._rc_hwnd_root:
            self._rc_hwnd_root = self._send_target_hwnd

        # Try to capture selection via UIA immediately
        text = ''
        try:
            _ui_log('[RC] UIA try (rc)')
            text = self._capture_selection_uia()
            if text:
                self._last_selection_text = text
                self._last_selection_time = time.time()
                _ui_log(f'[RC] cached selection len={len(text)}')
        except Exception:
            pass
        if not text:
            self._cache_selection_clipboard_async()

        _ui_log(f'[RC] popup show at {x},{y}')
        QTimer.singleShot(80, lambda: self._show_translate_popup(x, y))

    def _show_translate_popup(self, x, y):
        self._hide_translate_popup()
        popup = _TranslatePopup(None)
        popup.translate_clicked.connect(self._on_popup_translate_click)
        try:
            popup.setAttribute(Qt.WA_ShowWithoutActivating, True)
        except Exception:
            pass
        # Prefer above cursor; clamp to screen bounds (never below)
        screen_obj = QGuiApplication.screenAt(QPoint(x, y))
        screen = (screen_obj.availableGeometry()
                  if screen_obj else QGuiApplication.primaryScreen().availableGeometry())
        px = x + 12
        py = max(screen.top() + 8, y - 140)
        if px + popup.width() > screen.right():
            px = screen.right() - popup.width() - 8
        if px < screen.left():
            px = screen.left() + 8
        if py + popup.height() > screen.bottom():
            py = screen.bottom() - popup.height() - 8
        popup.move(px, py)
        popup.show()
        try:
            popup.raise_()
            import ctypes
            HWND_TOPMOST = -1
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_SHOWWINDOW = 0x0040
            SWP_NOACTIVATE = 0x0010
            ctypes.windll.user32.SetWindowPos(
                int(popup.winId()),
                HWND_TOPMOST,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE,
            )
        except Exception:
            pass
        self._translate_popup = popup
        self._popup_hide_timer.start(4000)

    def _hide_translate_popup(self):
        self._popup_hide_timer.stop()
        if self._translate_popup:
            self._translate_popup.close()
            self._translate_popup.deleteLater()
            self._translate_popup = None

    def _on_popup_translate_click(self):
        _ui_log('[RC] popup clicked')
        self._hide_translate_popup()
        self._translate_selection()

    def _translate_selection(self):
        if self._copy_in_flight:
            return
        self._copy_in_flight = True
        target_hwnd = self._send_target_hwnd
        rc_point = self._rc_hwnd_point
        rc_root = self._rc_hwnd_root or target_hwnd

        def _do_copy():
            _ui_log(f'[RC] click copy root={int(rc_root) if rc_root else 0} point={int(rc_point) if rc_point else 0}')

            # Use cached UIA selection if fresh
            if self._last_selection_text and (time.time() - self._last_selection_time) < 3.0:
                _ui_log(f'[RC] use cached selection len={len(self._last_selection_text)}')
                self._translate_fill_signal.emit(self._last_selection_text)
                self._last_selection_text = ''
                return
            # Try UIAutomation selection (no clipboard dependency)
            try:
                _ui_log('[RC] UIA try')
                text_uia = self._capture_selection_uia()
                if text_uia:
                    _ui_log(f'[RC] UIA selection len={len(text_uia)}')
                    self._translate_fill_signal.emit(text_uia)
                    return
                _ui_log('[RC] UIA selection empty')
            except Exception as e:
                _ui_log(f'[RC] UIA failed: {e!r}')

            try:
                text = self._capture_selection_clipboard(rc_point, rc_root, log_prefix='[RC]')
                if text:
                    self._translate_fill_signal.emit(text)
                else:
                    self._translate_fill_signal.emit('')
            except Exception:
                self._translate_fill_signal.emit('')
            finally:
                self._copy_in_flight = False

        threading.Thread(target=_do_copy, daemon=True).start()

    def _capture_selection_clipboard(self, hwnd_point=None, hwnd_root=None, log_prefix='[RC]'):
        import ctypes
        import time
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        CF_UNICODETEXT = 13
        WM_COPY = 0x0301
        GMEM_MOVEABLE = 0x0002

        class GUITHREADINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_uint),
                ("flags", ctypes.c_uint),
                ("hwndActive", ctypes.c_void_p),
                ("hwndFocus", ctypes.c_void_p),
                ("hwndCapture", ctypes.c_void_p),
                ("hwndMenuOwner", ctypes.c_void_p),
                ("hwndMoveSize", ctypes.c_void_p),
                ("hwndCaret", ctypes.c_void_p),
                ("rcCaret", ctypes.c_long * 4),
            ]

        def _log(msg: str):
            try:
                _ui_log(f'{log_prefix} {msg}')
            except Exception:
                pass

        def _get_focus_hwnd(hwnd):
            try:
                tid = user32.GetWindowThreadProcessId(hwnd, 0)
                if not tid:
                    return None
                info = GUITHREADINFO()
                info.cbSize = ctypes.sizeof(GUITHREADINFO)
                if user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
                    return info.hwndFocus or info.hwndActive
            except Exception:
                pass
            return None

        def _read_clipboard_text():
            text_local = ''
            try:
                if user32.OpenClipboard(0):
                    try:
                        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                            return ''
                        handle = user32.GetClipboardData(CF_UNICODETEXT)
                        if handle:
                            kernel32.GlobalLock.restype = ctypes.c_void_p
                            ptr = kernel32.GlobalLock(handle)
                            if ptr:
                                text_local = ctypes.wstring_at(ptr)
                                kernel32.GlobalUnlock(handle)
                    finally:
                        user32.CloseClipboard()
            except Exception:
                pass
            return text_local

        def _set_clipboard_text(text: str) -> bool:
            if text is None:
                text = ''
            data = text + '\x00'
            size = len(data) * ctypes.sizeof(ctypes.c_wchar)
            if not user32.OpenClipboard(0):
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
                hglob = None
                return True
            finally:
                if hglob:
                    kernel32.GlobalFree(hglob)
                user32.CloseClipboard()

        def _send_ctrl_c():
            INPUT_KEYBOARD = 1
            KEYEVENTF_KEYUP = 0x0002
            VK_CONTROL = 0x11
            VK_C = 0x43

            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [
                    ("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
                ]

            class INPUT(ctypes.Structure):
                _fields_ = [("type", ctypes.c_ulong), ("ki", KEYBDINPUT)]

            extra = ctypes.c_ulong(0)
            inputs = (INPUT * 4)(
                INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_CONTROL, 0, 0, 0, ctypes.pointer(extra))),
                INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_C, 0, 0, 0, ctypes.pointer(extra))),
                INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_C, 0, KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))),
                INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))),
            )
            user32.SendInput(4, ctypes.byref(inputs), ctypes.sizeof(INPUT))

        if not hwnd_root:
            hwnd_root = self._send_target_hwnd
        if not hwnd_root and hwnd_point:
            hwnd_root = hwnd_point
        if not hwnd_root:
            try:
                hwnd_root = user32.GetForegroundWindow()
            except Exception:
                hwnd_root = None
        hwnd_copy = hwnd_point or hwnd_root

        _log(f'copy start root={int(hwnd_root) if hwnd_root else 0} point={int(hwnd_point) if hwnd_point else 0}')

        hwnd_focus = None
        try:
            user32.AllowSetForegroundWindow(-1)
        except Exception:
            pass
        try:
            if hwnd_root:
                pid = ctypes.c_uint(0)
                tid_target = user32.GetWindowThreadProcessId(hwnd_root, ctypes.byref(pid))
                tid_self = kernel32.GetCurrentThreadId()
                if tid_target:
                    user32.AttachThreadInput(tid_self, tid_target, True)
                try:
                    user32.SetForegroundWindow(hwnd_root)
                    hwnd_focus = _get_focus_hwnd(hwnd_root)
                    if hwnd_focus:
                        user32.SetFocus(hwnd_focus)
                finally:
                    if tid_target:
                        user32.AttachThreadInput(tid_self, tid_target, False)
        except Exception:
            pass
        if not hwnd_focus and hwnd_root:
            hwnd_focus = _get_focus_hwnd(hwnd_root)

        has_text = False
        try:
            has_text = bool(user32.IsClipboardFormatAvailable(CF_UNICODETEXT))
        except Exception:
            has_text = False
        prev_text = _read_clipboard_text() if has_text else ''
        try:
            prev_seq = user32.GetClipboardSequenceNumber()
        except Exception:
            prev_seq = None

        sentinel = None
        if has_text:
            sentinel = f'__AUTO_LANG_SENTINEL__{time.time()}__'
            if not _set_clipboard_text(sentinel):
                sentinel = None
                has_text = False

        try:
            if hwnd_focus:
                user32.SendMessageW(hwnd_focus, WM_COPY, 0, 0)
            if hwnd_copy and hwnd_copy != hwnd_focus:
                user32.SendMessageW(hwnd_copy, WM_COPY, 0, 0)
        except Exception:
            pass
        _send_ctrl_c()

        text = ''
        for _ in range(12):
            if has_text and sentinel is not None:
                text = _read_clipboard_text()
                if text and text != sentinel:
                    break
            else:
                try:
                    if prev_seq is not None and user32.GetClipboardSequenceNumber() != prev_seq:
                        text = _read_clipboard_text()
                        if text:
                            break
                except Exception:
                    pass
            time.sleep(0.05)

        if has_text and sentinel is not None:
            try:
                _set_clipboard_text(prev_text)
            except Exception:
                pass

        if text and text.strip() and text != sentinel:
            text = text.strip()
            _log(f'copy got len={len(text)}')
            return text

        _log('copy empty')
        return ''

    def _cache_selection_clipboard_async(self):
        if self._rc_cache_in_flight:
            return
        self._rc_cache_in_flight = True
        hwnd_point = self._rc_hwnd_point
        hwnd_root = self._rc_hwnd_root or self._send_target_hwnd

        def _do():
            try:
                text = self._capture_selection_clipboard(hwnd_point, hwnd_root, log_prefix='[RC] cache')
                if text:
                    self._last_selection_text = text
                    self._last_selection_time = time.time()
                    _ui_log(f'[RC] cached selection len={len(text)}')
                else:
                    _ui_log('[RC] cache empty')
            finally:
                self._rc_cache_in_flight = False

        threading.Thread(target=_do, daemon=True).start()

    def _capture_selection_uia(self) -> str:
        try:
            import uiautomation as auto
        except Exception as e:
            _ui_log(f'[RC] UIA import failed: {e!r}')
            return ''
        text_uia = ''

        def _get_sel(control):
            if not control:
                return ''
            try:
                tp = control.GetPattern(auto.PatternId.TextPattern)
                if not tp:
                    tp = control.GetPattern(auto.PatternId.TextPattern2)
                if not tp:
                    return ''
                ranges = tp.GetSelection()
                if ranges:
                    parts = []
                    for r in ranges:
                        t = r.GetText(-1)
                        if t:
                            parts.append(t)
                    return ''.join(parts).strip()
            except Exception:
                return ''
            return ''

        if self._last_rc_pos:
            cx, cy = self._last_rc_pos
        else:
            cx = cy = None
        try:
            with auto.UIAutomationInitializerInThread():
                if cx is not None and cy is not None:
                    ctrl = auto.ControlFromPoint(cx, cy)
                    try:
                        if ctrl:
                            _ui_log(f'[RC] UIA ctrl={ctrl.ControlTypeName} name={ctrl.Name} class={ctrl.ClassName}')
                    except Exception:
                        pass
                    text_uia = _get_sel(ctrl)
                if not text_uia:
                    focused = auto.GetFocusedControl()
                    try:
                        if focused:
                            _ui_log(f'[RC] UIA focused={focused.ControlTypeName} name={focused.Name} class={focused.ClassName}')
                    except Exception:
                        pass
                    text_uia = _get_sel(focused)
        except Exception as e:
            _ui_log(f'[RC] UIA exception: {e!r}')
            return ''
        return text_uia
    def _fill_and_translate_text(self, text: str):
        if not text:
            if self._text_widget:
                self._text_widget.setPlainText('')
            if self._trans_widget:
                self._trans_widget.setPlainText('')
            if self._status_label:
                self._status_label.setText('\u26a0\ufe0f \u05d1\u05d7\u05e8 \u05d8\u05e7\u05e1\u05d8 \u05dc\u05ea\u05e8\u05d2\u05d5\u05dd')
            return

        if self._editor_height == 0 or self._text_widget is None or self._trans_widget is None:
            self._show_edit_window()
        self._set_translation_visible(True)

        if self._text_widget:
            self._text_widget.setPlainText(text)

        QTimer.singleShot(100, self._do_translate)

    # ── speech-to-text ───────────────────────────────────

    def _toggle_speech(self):
        if self._speech_recording:
            self._stop_speech()
        else:
            self._start_speech()

    def _start_speech(self):
        if self._speech_recording:
            return
        self._speech_recording = True

        try:
            import auto_lang
            hwnd = auto_lang._last_target_hwnd
            if not hwnd:
                hwnd = self._find_target_hwnd()
            self._send_target_hwnd = hwnd
        except Exception:
            pass

        self._pulse_step = 0
        self._pulse_timer.start()
        self.update()

        try:
            import speech_module
            speech_module.start_recording(
                callback=lambda text, lang: self._speech_text_signal.emit(text, lang),
                on_error=lambda msg: self._speech_error_signal.emit(msg),
                on_state=lambda s: None,
            )
        except Exception as e:
            print(f'[Speech] Failed: {e}')
            self._speech_recording = False
            self._pulse_timer.stop()
            self.update()

    def _stop_speech(self):
        self._speech_recording = False
        self._pulse_timer.stop()
        self.update()
        try:
            import speech_module
            speech_module.stop_recording()
        except Exception:
            pass

    def _pulse_animation(self):
        self._pulse_step += 1
        self.update()

    def _handle_speech_result(self, text: str, lang_code: str):
        if not text:
            return

        now = time.time()
        if text == self._last_speech_text and now - self._last_speech_time < 3.0:
            return
        self._last_speech_text = text
        self._last_speech_time = now

        if self._edit_mode and self._text_widget:
            current = self._text_widget.toPlainText()
            if current and not current.endswith(' ') and not current.endswith('\n'):
                self._text_widget.insertPlainText(' ')
            self._text_widget.insertPlainText(text)
            self._translate_debounce.start(300)
        else:
            threading.Thread(
                target=self._paste_speech_to_target,
                args=(text, lang_code),
                daemon=True,
            ).start()

    def _paste_speech_to_target(self, text: str, lang_code: str):
        try:
            import auto_lang
            import ctypes
            hwnd = self._send_target_hwnd or auto_lang._last_target_hwnd
            if not hwnd:
                hwnd = self._find_target_hwnd()
            if not hwnd:
                return
            ctypes.windll.user32.AllowSetForegroundWindow(-1)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.2)
            auto_lang.injecting.set()
            try:
                auto_lang._paste_text(text)
            finally:
                time.sleep(0.05)
                auto_lang.injecting.clear()
        except Exception:
            pass

    # ── help window ──────────────────────────────────────

    def _show_help(self):
        win = QWidget(None, Qt.Window)
        win.setWindowTitle('AutoLang \u2014 \u05de\u05d3\u05e8\u05d9\u05da \u05dc\u05de\u05e9\u05ea\u05de\u05e9')
        win.setGeometry(200, 100, 520, 620)
        win.setStyleSheet(DARK_QSS)
        win.setWindowFlag(Qt.WindowStaysOnTopHint)
        _apply_dark_title_bar(win)

        layout = QVBoxLayout(win)
        layout.setContentsMargins(14, 14, 14, 14)

        title = QLabel('\U0001f4d6  \u05de\u05d3\u05e8\u05d9\u05da \u05dc\u05de\u05e9\u05ea\u05de\u05e9')
        title.setStyleSheet(f"color: {COLORS['accent']}; font-size: 16pt; font-weight: bold;")
        title.setAlignment(Qt.AlignRight)
        layout.addWidget(title)

        sep = QFrame()
        sep.setFixedHeight(2)
        sep.setStyleSheet(f"background-color: {COLORS['accent']};")
        layout.addWidget(sep)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['text_primary']};
                border: none;
                padding: 10px;
                font-size: 10pt;
            }}
        """)

        sections = [
            ('\u05de\u05d4 \u05d6\u05d4 AutoLang?',
             'AutoLang \u05d4\u05d9\u05d0 \u05ea\u05d5\u05db\u05e0\u05d4 \u05e9\u05de\u05ea\u05e7\u05e0\u05ea \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9\u05ea \u05d8\u05e7\u05e1\u05d8 \u05e9\u05d4\u05d5\u05e7\u05dc\u05d3 \u05d1\u05e9\u05e4\u05d4 \u05d4\u05dc\u05d0 \u05e0\u05db\u05d5\u05e0\u05d4.\n'
             '\u05dc\u05de\u05e9\u05dc \u2014 \u05d0\u05dd \u05d4\u05de\u05e7\u05dc\u05d3\u05ea \u05e9\u05dc\u05da \u05e2\u05dc \u05d0\u05e0\u05d2\u05dc\u05d9\u05ea \u05d5\u05d0\u05ea\u05d4 \u05de\u05e7\u05dc\u05d9\u05d3 \u05e2\u05d1\u05e8\u05d9\u05ea, \u05d4\u05ea\u05d5\u05db\u05e0\u05d4 \u05de\u05d6\u05d4\u05d4 \u05d0\u05ea \u05d4\u05d8\u05e2\u05d5\u05ea \u05d5\u05de\u05ea\u05e7\u05e0\u05ea \u05d0\u05d5\u05ea\u05d4 \u05d1\u05d6\u05de\u05df \u05d0\u05de\u05ea.'),
            ('\U0001f535 \u05db\u05e4\u05ea\u05d5\u05e8\u05d9 \u05d4\u05e2\u05d9\u05d2\u05d5\u05dc',
             '\u23fb  \u05dc\u05de\u05e2\u05dc\u05d4 \u2014 \u05d4\u05e4\u05e2\u05dc\u05d4 / \u05d4\u05e9\u05d1\u05ea\u05d4 \u05e9\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d4\u05e9\u05e4\u05d4\n'
             '\u21a9  \u05e9\u05de\u05d0\u05dc \u2014 \u05d1\u05d9\u05d8\u05d5\u05dc \u05d4\u05ea\u05d9\u05e7\u05d5\u05df \u05d4\u05d0\u05d7\u05e8\u05d5\u05df (\u05d2\u05dd F12)\n'
             '\u270e  \u05dc\u05de\u05d8\u05d4 \u2014 \u05e4\u05ea\u05d9\u05d7\u05ea \u05e2\u05d5\u05e8\u05da \u05ea\u05e8\u05d2\u05d5\u05dd\n'
             '\u2699  \u05d9\u05de\u05d9\u05df \u2014 \u05e4\u05ea\u05d9\u05d7\u05ea \u05d7\u05dc\u05d5\u05df \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea\n'
             '\U0001f534  \u05de\u05e8\u05db\u05d6 \u2014 \u05d4\u05e7\u05dc\u05d8\u05d4 \u05e7\u05d5\u05dc\u05d9\u05ea (Speech-to-Text)'),
            ('\u2328\ufe0f \u05e7\u05d9\u05e6\u05d5\u05e8\u05d9 \u05de\u05e7\u05e9\u05d9\u05dd',
             'F12 \u2014 \u05d1\u05d9\u05d8\u05d5\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df\n'
             'Ctrl+Alt+Q \u2014 \u05d9\u05e6\u05d9\u05d0\u05d4 \u05de\u05d4\u05ea\u05d5\u05db\u05e0\u05d4\n'
             'Ctrl+Alt+I \u2014 \u05d4\u05e6\u05d2\u05ea \u05de\u05d9\u05d3\u05e2 (Debug)'),
        ]

        html = ''
        for title_text, body in sections:
            html += f'<h3 style="color: {COLORS["accent"]}; margin-top: 12px;">{title_text}</h3>'
            body_html = body.replace('\n', '<br>')
            html += f'<p style="color: {COLORS["text_primary"]};">{body_html}</p>'

        help_text.setHtml(html)
        layout.addWidget(help_text)

        close_btn = QPushButton('\u05e1\u05d2\u05d5\u05e8')
        close_btn.setFixedWidth(80)
        close_btn.clicked.connect(win.close)
        layout.addWidget(close_btn, alignment=Qt.AlignCenter)

        win.show()
        self._help_win = win  # prevent garbage collection

    # ── utility ──────────────────────────────────────────

    def _find_target_hwnd(self):
        try:
            import ctypes
            from ctypes import wintypes
            _user32 = ctypes.windll.user32
            _kernel32 = ctypes.windll.kernel32
            own_pid = os.getpid()
            WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            result = [None]
            def _enum_cb(hwnd, _):
                if not _user32.IsWindowVisible(hwnd):
                    return True
                pid = wintypes.DWORD(0)
                _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value == own_pid:
                    return True
                length = _user32.GetWindowTextLengthW(hwnd)
                if length <= 0:
                    return True
                result[0] = hwnd
                return False
            cb = WNDENUMPROC(_enum_cb)
            _user32.EnumWindows(cb, 0)
            return result[0]
        except Exception:
            return None


# ──────────────────────────────────────────────────────────
# Settings Window
# ──────────────────────────────────────────────────────────

class SettingsWindow(QWidget):
    """\u05d7\u05dc\u05d5\u05df \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea \u05e2\u05dd \u05d8\u05d0\u05d1\u05d9\u05dd."""

    def __init__(self, config: dict, on_save_callback=None):
        super().__init__(None, Qt.Window)
        self.config = config
        self.on_save = on_save_callback

        self.setWindowTitle('AutoLang - \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea')
        self.setGeometry(200, 100, 740, 600)
        self.setStyleSheet(DARK_QSS)
        _apply_dark_title_bar(self)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 14, 20, 14)

        # ── Title ──
        title_row = QHBoxLayout()
        title_row.addStretch()
        title_text = QVBoxLayout()
        lbl_title = QLabel('AutoLang')
        lbl_title.setStyleSheet(f"color: #ffffff; font-size: 20pt; font-weight: bold; background: transparent;")
        lbl_title.setAlignment(Qt.AlignRight)
        title_text.addWidget(lbl_title)
        lbl_sub = QLabel('\u05d4\u05d2\u05d3\u05e8\u05d5\u05ea \u05de\u05ea\u05e7\u05d3\u05de\u05d5\u05ea')
        lbl_sub.setStyleSheet(f"color: {COLORS['accent']}; font-size: 9pt; background: transparent;")
        lbl_sub.setAlignment(Qt.AlignRight)
        title_text.addWidget(lbl_sub)
        title_row.addLayout(title_text)
        lbl_icon = QLabel('\u2328\ufe0f')
        lbl_icon.setStyleSheet(f"font-size: 20pt; background: transparent;")
        title_row.addWidget(lbl_icon)
        main_layout.addLayout(title_row)

        sep = QFrame()
        sep.setFixedHeight(2)
        sep.setStyleSheet(f"background-color: {COLORS['accent']}; border: none;")
        main_layout.addWidget(sep)

        # ── Tabs ──
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self._build_app_defaults_tab()
        self._build_chat_defaults_tab()
        self._build_browser_defaults_tab()
        self._build_exclude_words_tab()
        self._build_general_tab()
        self._build_spell_grammar_tab()
        self._build_help_tab()

        # ── Bottom buttons ──
        btn_row = QHBoxLayout()
        self.status_label = QLabel('')
        self.status_label.setStyleSheet(f"color: {COLORS['success']}; font-size: 9pt; background: transparent;")
        btn_row.addWidget(self.status_label)
        btn_row.addStretch()

        cancel_btn = QPushButton('\u274c \u05d1\u05d8\u05dc')
        cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(cancel_btn)

        save_btn = QPushButton('\U0001f4be \u05e9\u05de\u05d5\u05e8')
        save_btn.setProperty('cssClass', 'accent')
        save_btn.clicked.connect(self._save)
        btn_row.addWidget(save_btn)

        main_layout.addLayout(btn_row)

    def _make_card(self):
        """Create a dark card container frame."""
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['bg_medium']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
            }}
        """)
        return card

    def _make_table(self, columns, headers):
        table = QTableWidget(0, columns)
        table.setHorizontalHeaderLabels(headers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.horizontalHeader().setStretchLastSection(True)
        table.verticalHeader().setVisible(False)
        table.setLayoutDirection(Qt.RightToLeft)
        return table

    # ─── Tab 1: App Defaults ─────────────────────────────

    def _build_app_defaults_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f5a5\ufe0f \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d5\u05ea')
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        lbl = QLabel('\u05e9\u05e4\u05ea \u05d1\u05e8\u05d9\u05e8\u05ea \u05de\u05d7\u05d3\u05dc \u05dc\u05db\u05dc \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4')
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(lbl)

        desc = QLabel('\u05d4\u05d2\u05d3\u05e8 \u05e9\u05e4\u05ea \u05d1\u05e8\u05d9\u05e8\u05ea \u05de\u05d7\u05d3\u05dc \u05dc\u05e4\u05d9 \u05e9\u05dd \u05d4\u05ea\u05d4\u05dc\u05d9\u05da (exe). \u05d4\u05ea\u05d5\u05db\u05e0\u05d4 \u05ea\u05d7\u05dc\u05d9\u05e3 \u05d0\u05ea \u05d4\u05e9\u05e4\u05d4 \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9\u05ea \u05db\u05e9\u05ea\u05e2\u05d1\u05d5\u05e8 \u05dc\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4.')
        desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        desc.setAlignment(Qt.AlignRight)
        desc.setWordWrap(True)
        card_layout.addWidget(desc)

        self.app_table = self._make_table(2, ['\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4 (exe)', '\u05e9\u05e4\u05d4'])
        self.app_table.setColumnWidth(0, 350)
        for app, lang in sorted(self.config.get('app_defaults', {}).items()):
            row = self.app_table.rowCount()
            self.app_table.insertRow(row)
            self.app_table.setItem(row, 0, QTableWidgetItem(app))
            self.app_table.setItem(row, 1, QTableWidgetItem(_lang_display(lang)))
        card_layout.addWidget(self.app_table)

        # Add row
        add_row = QHBoxLayout()
        add_row.addStretch()

        self.app_exe_edit = QLineEdit()
        self.app_exe_edit.setPlaceholderText('\u05e9\u05dd \u05ea\u05d4\u05dc\u05d9\u05da')
        self.app_exe_edit.setFixedWidth(180)
        add_row.addWidget(self.app_exe_edit)

        self.app_lang_combo = QComboBox()
        self.app_lang_combo.addItems(_get_available_lang_codes())
        self.app_lang_combo.setCurrentText('he')
        self.app_lang_combo.setFixedWidth(70)
        add_row.addWidget(self.app_lang_combo)

        add_btn = QPushButton('\u2795 \u05d4\u05d5\u05e1\u05e3')
        add_btn.clicked.connect(self._add_app)
        add_row.addWidget(add_btn)

        del_btn = QPushButton('\U0001f5d1\ufe0f \u05d4\u05e1\u05e8')
        del_btn.setProperty('cssClass', 'danger')
        del_btn.clicked.connect(self._remove_app)
        add_row.addWidget(del_btn)

        card_layout.addLayout(add_row)
        layout.addWidget(card)

    def _add_app(self):
        exe = self.app_exe_edit.text().strip()
        lang = self.app_lang_combo.currentText()
        if not exe:
            return
        # Check if exists
        for row in range(self.app_table.rowCount()):
            if self.app_table.item(row, 0).text() == exe:
                self.app_table.setItem(row, 1, QTableWidgetItem(_lang_display(lang)))
                self.app_exe_edit.clear()
                return
        row = self.app_table.rowCount()
        self.app_table.insertRow(row)
        self.app_table.setItem(row, 0, QTableWidgetItem(exe))
        self.app_table.setItem(row, 1, QTableWidgetItem(_lang_display(lang)))
        self.app_exe_edit.clear()

    def _remove_app(self):
        rows = sorted(set(idx.row() for idx in self.app_table.selectionModel().selectedRows()), reverse=True)
        for row in rows:
            self.app_table.removeRow(row)

    # ─── Tab 2: Chat Defaults ────────────────────────────

    def _build_chat_defaults_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f4ac \u05e6\'\u05d0\u05d8\u05d9\u05dd')
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        lbl = QLabel('\u05e9\u05e4\u05ea \u05d1\u05e8\u05d9\u05e8\u05ea \u05de\u05d7\u05d3\u05dc \u05dc\u05db\u05dc \u05e6\'\u05d0\u05d8')
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(lbl)

        desc = QLabel('\u05d4\u05d2\u05d3\u05e8 \u05e9\u05e4\u05d4 \u05dc\u05e4\u05d9 \u05e9\u05dd \u05e6\'\u05d0\u05d8 (\u05d7\u05dc\u05e7 \u05de\u05db\u05d5\u05ea\u05e8\u05ea \u05d4\u05d7\u05dc\u05d5\u05df). \u05dc\u05de\u05e9\u05dc: "\u05de\u05e8\u05d9\u05d4" = English')
        desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        desc.setAlignment(Qt.AlignRight)
        card_layout.addWidget(desc)

        self.chat_table = self._make_table(3, ['\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4', '\u05e9\u05dd \u05e6\'\u05d0\u05d8 (\u05de\u05d9\u05dc\u05ea \u05de\u05e4\u05ea\u05d7)', '\u05e9\u05e4\u05d4'])
        self.chat_table.setColumnWidth(0, 180)
        self.chat_table.setColumnWidth(1, 250)
        for exe, chats in self.config.get('chat_defaults', {}).items():
            for title, lang in chats.items():
                row = self.chat_table.rowCount()
                self.chat_table.insertRow(row)
                self.chat_table.setItem(row, 0, QTableWidgetItem(exe))
                self.chat_table.setItem(row, 1, QTableWidgetItem(title))
                self.chat_table.setItem(row, 2, QTableWidgetItem(_lang_display(lang)))
        card_layout.addWidget(self.chat_table)

        # Add row
        add_row = QHBoxLayout()
        add_row.addStretch()

        known_exes = set(self.config.get('app_defaults', {}).keys())
        known_exes.update(self.config.get('chat_defaults', {}).keys())
        known_exes.update(self.config.get('watch_title_exes', []))
        known_exes.update(['chrome.exe', 'msedge.exe', 'outlook.exe', 'olk.exe', 'firefox.exe'])

        self.chat_exe_combo = QComboBox()
        self.chat_exe_combo.setEditable(True)
        self.chat_exe_combo.addItems(sorted(known_exes))
        self.chat_exe_combo.setCurrentText('ms-teams.exe')
        self.chat_exe_combo.setFixedWidth(160)
        add_row.addWidget(self.chat_exe_combo)

        self.chat_title_edit = QLineEdit()
        self.chat_title_edit.setPlaceholderText('\u05e9\u05dd \u05e6\'\u05d0\u05d8')
        self.chat_title_edit.setFixedWidth(150)
        add_row.addWidget(self.chat_title_edit)

        self.chat_lang_combo = QComboBox()
        self.chat_lang_combo.addItems(_get_available_lang_codes())
        self.chat_lang_combo.setCurrentText('he')
        self.chat_lang_combo.setFixedWidth(70)
        add_row.addWidget(self.chat_lang_combo)

        add_btn = QPushButton('\u2795 \u05d4\u05d5\u05e1\u05e3')
        add_btn.clicked.connect(self._add_chat)
        add_row.addWidget(add_btn)

        del_btn = QPushButton('\U0001f5d1\ufe0f \u05d4\u05e1\u05e8')
        del_btn.setProperty('cssClass', 'danger')
        del_btn.clicked.connect(self._remove_chat)
        add_row.addWidget(del_btn)

        card_layout.addLayout(add_row)

        # Test button
        test_btn = QPushButton('\U0001f50d \u05d1\u05d3\u05d5\u05e7 \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4 \u2014 \u05d4\u05d0\u05dd \u05ea\u05d5\u05de\u05db\u05ea \u05d1\u05d6\u05d9\u05d4\u05d5\u05d9 \u05e6\'\u05d0\u05d8\u05d9\u05dd?')
        test_btn.clicked.connect(self._test_app_title_support)
        card_layout.addWidget(test_btn)

        layout.addWidget(card)

    def _add_chat(self):
        exe = self.chat_exe_combo.currentText().strip()
        title = self.chat_title_edit.text().strip()
        lang = self.chat_lang_combo.currentText()
        if not exe or not title:
            return
        row = self.chat_table.rowCount()
        self.chat_table.insertRow(row)
        self.chat_table.setItem(row, 0, QTableWidgetItem(exe))
        self.chat_table.setItem(row, 1, QTableWidgetItem(title))
        self.chat_table.setItem(row, 2, QTableWidgetItem(_lang_display(lang)))
        self.chat_title_edit.clear()

    def _remove_chat(self):
        rows = sorted(set(idx.row() for idx in self.chat_table.selectionModel().selectedRows()), reverse=True)
        for row in rows:
            self.chat_table.removeRow(row)

    def _test_app_title_support(self):
        """Monitor foreground window for 12s to check title changes."""
        import ctypes
        import ctypes.wintypes as wintypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        _user32 = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32

        def _pid_to_exe(pid_val):
            h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid_val)
            if not h:
                return '?'
            try:
                bl = wintypes.DWORD(260)
                b = ctypes.create_unicode_buffer(260)
                if _kernel32.QueryFullProcessImageNameW(h, 0, b, ctypes.byref(bl)):
                    return b.value.rsplit('\\', 1)[-1].lower()
                return '?'
            finally:
                _kernel32.CloseHandle(h)

        def _get_fg():
            hwnd = _user32.GetForegroundWindow()
            if not hwnd:
                return '', ''
            pid = wintypes.DWORD(0)
            _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            exe = _pid_to_exe(pid.value)
            length = _user32.GetWindowTextLengthW(hwnd)
            title = ''
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                _user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
            return exe, title

        # Popup window
        popup = QWidget(None, Qt.Window | Qt.WindowStaysOnTopHint)
        popup.setWindowTitle('\u05d1\u05d3\u05d9\u05e7\u05ea \u05ea\u05de\u05d9\u05db\u05d4 \u05d1\u05d6\u05d9\u05d4\u05d5\u05d9 \u05e6\'\u05d0\u05d8\u05d9\u05dd')
        popup.setFixedSize(540, 400)
        popup.setStyleSheet(DARK_QSS)
        _apply_dark_title_bar(popup)

        p_layout = QVBoxLayout(popup)
        p_title = QLabel('\U0001f50d \u05d1\u05d3\u05d9\u05e7\u05ea \u05ea\u05de\u05d9\u05db\u05d4 \u05d1\u05d6\u05d9\u05d4\u05d5\u05d9 \u05e6\'\u05d0\u05d8\u05d9\u05dd')
        p_title.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold;")
        p_title.setAlignment(Qt.AlignCenter)
        p_layout.addWidget(p_title)

        info_label = QLabel('\u23f3 \u05e2\u05d1\u05d5\u05e8 \u05dc\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4 \u05d5\u05d4\u05d7\u05dc\u05e3 \u05d1\u05d9\u05df \u05e6\'\u05d0\u05d8\u05d9\u05dd/\u05de\u05d9\u05d9\u05dc\u05d9\u05dd/\u05d8\u05d0\u05d1\u05d9\u05dd...\n\n\u05d4\u05d1\u05d3\u05d9\u05e7\u05d4 \u05ea\u05d9\u05de\u05e9\u05da 12 \u05e9\u05e0\u05d9\u05d5\u05ea')
        info_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 10pt;")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setWordWrap(True)
        p_layout.addWidget(info_label)

        progress = QProgressBar()
        progress.setMaximum(12)
        progress.setValue(0)
        p_layout.addWidget(progress)

        result_text = QTextEdit()
        result_text.setReadOnly(True)
        result_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']};
                font-family: Consolas;
                font-size: 9pt;
            }}
        """)
        p_layout.addWidget(result_text)

        popup.show()
        self._test_popup = popup  # prevent GC

        monitor_data = {'seen': {}, 'step': 0, 'done': False, 'report': '', 'supported_exe': ''}

        def _monitor_thread():
            seen = {}
            for i in range(24):
                try:
                    exe, title = _get_fg()
                    if exe and exe not in ('autolang.exe', 'python.exe', 'pythonw.exe'):
                        if exe not in seen:
                            seen[exe] = set()
                        seen[exe].add(title)
                except Exception:
                    pass
                monitor_data['seen'] = seen
                monitor_data['step'] = i + 1
                time.sleep(0.5)

            lines = []
            supported_exe = ''
            for exe, titles in sorted(seen.items()):
                count = len(titles)
                if count >= 2:
                    if not supported_exe:
                        supported_exe = exe
                    lines.append(f'\u2705 {exe} \u2014 \u05e0\u05de\u05e6\u05d0\u05d5 {count} \u05db\u05d5\u05ea\u05e8\u05d5\u05ea \u05e9\u05d5\u05e0\u05d5\u05ea \u2014 \u05ea\u05d5\u05de\u05da!')
                    for t in list(titles)[:3]:
                        short = t[:55] + '...' if len(t) > 55 else t
                        lines.append(f'    \u2022 {short}')
                else:
                    t = list(titles)[0] if titles else '(\u05e8\u05d9\u05e7)'
                    short = t[:45] + '...' if len(t) > 45 else t
                    lines.append(f'\u274c {exe} \u2014 \u05db\u05d5\u05ea\u05e8\u05ea \u05e7\u05d1\u05d5\u05e2\u05d4: "{short}"')

            if not lines:
                lines.append('\u05dc\u05d0 \u05d6\u05d5\u05d4\u05d5 \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d5\u05ea. \u05d5\u05d3\u05d0 \u05e9\u05e2\u05d1\u05e8\u05ea \u05dc\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d4 \u05d0\u05d7\u05e8\u05ea.')

            monitor_data['report'] = '\n'.join(lines)
            monitor_data['supported_exe'] = supported_exe
            monitor_data['done'] = True

        def _poll():
            if not popup.isVisible():
                return
            progress.setValue(monitor_data['step'] // 2)
            if monitor_data['done']:
                has_supported = bool(monitor_data['supported_exe'])
                info_label.setText('\u2705 \u05d4\u05d1\u05d3\u05d9\u05e7\u05d4 \u05d4\u05e1\u05ea\u05d9\u05d9\u05de\u05d4!' if has_supported else '\u26a0\ufe0f \u05d4\u05d1\u05d3\u05d9\u05e7\u05d4 \u05d4\u05e1\u05ea\u05d9\u05d9\u05de\u05d4 \u2014 \u05dc\u05d0 \u05e0\u05de\u05e6\u05d0\u05d5 \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d5\u05ea \u05ea\u05d5\u05de\u05db\u05d5\u05ea')
                info_label.setStyleSheet(f"color: {COLORS['success'] if has_supported else COLORS['warning']}; font-size: 10pt;")
                progress.setValue(12)
                result_text.setPlainText(monitor_data['report'])
                if has_supported:
                    self.chat_exe_combo.setCurrentText(monitor_data['supported_exe'])
            else:
                QTimer.singleShot(300, _poll)

        threading.Thread(target=_monitor_thread, daemon=True).start()
        QTimer.singleShot(300, _poll)

    # ─── Tab 3: Browser Defaults ─────────────────────────

    def _build_browser_defaults_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f310 \u05d0\u05ea\u05e8\u05d9\u05dd')
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        lbl = QLabel('\u05e9\u05e4\u05ea \u05d1\u05e8\u05d9\u05e8\u05ea \u05de\u05d7\u05d3\u05dc \u05dc\u05e4\u05d9 \u05d0\u05ea\u05e8 / \u05d3\u05e3')
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(lbl)

        desc = QLabel('\u05d4\u05d2\u05d3\u05e8 \u05e9\u05e4\u05d4 \u05dc\u05e4\u05d9 \u05de\u05d9\u05dc\u05ea \u05de\u05e4\u05ea\u05d7 \u05d1\u05db\u05d5\u05ea\u05e8\u05ea \u05d4\u05d3\u05e3 \u05d1\u05d3\u05e4\u05d3\u05e4\u05df.\n\u05e2\u05d5\u05d1\u05d3 \u05e2\u05dd \u05db\u05dc \u05d3\u05e4\u05d3\u05e4\u05df (Chrome, Edge, Firefox \u05d5\u05e2\u05d5\u05d3).')
        desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        desc.setAlignment(Qt.AlignRight)
        desc.setWordWrap(True)
        card_layout.addWidget(desc)

        self.browser_table = self._make_table(2, ['\u05d0\u05ea\u05e8 / \u05de\u05d9\u05dc\u05ea \u05de\u05e4\u05ea\u05d7', '\u05e9\u05e4\u05d4'])
        self.browser_table.setColumnWidth(0, 350)
        for keyword, lang in self.config.get('browser_defaults', {}).items():
            row = self.browser_table.rowCount()
            self.browser_table.insertRow(row)
            self.browser_table.setItem(row, 0, QTableWidgetItem(keyword))
            self.browser_table.setItem(row, 1, QTableWidgetItem(_lang_display(lang)))
        card_layout.addWidget(self.browser_table)

        add_row = QHBoxLayout()
        add_row.addStretch()

        self.browser_keyword_edit = QLineEdit()
        self.browser_keyword_edit.setPlaceholderText('\u05d0\u05ea\u05e8 / \u05d3\u05d5\u05de\u05d9\u05d9\u05df / \u05e9\u05dd \u05d3\u05e3')
        self.browser_keyword_edit.setFixedWidth(200)
        add_row.addWidget(self.browser_keyword_edit)

        self.browser_lang_combo = QComboBox()
        self.browser_lang_combo.addItems(_get_available_lang_codes())
        self.browser_lang_combo.setCurrentText('he')
        self.browser_lang_combo.setFixedWidth(70)
        add_row.addWidget(self.browser_lang_combo)

        add_btn = QPushButton('\u2795 \u05d4\u05d5\u05e1\u05e3')
        add_btn.clicked.connect(self._add_browser_rule)
        add_row.addWidget(add_btn)

        del_btn = QPushButton('\U0001f5d1\ufe0f \u05d4\u05e1\u05e8')
        del_btn.setProperty('cssClass', 'danger')
        del_btn.clicked.connect(self._remove_browser_rule)
        add_row.addWidget(del_btn)

        card_layout.addLayout(add_row)

        hint = QLabel('\U0001f4a1 \u05d4\u05db\u05d5\u05ea\u05e8\u05ea \u05d1\u05db\u05e8\u05d5\u05dd \u05e0\u05e8\u05d0\u05d9\u05ea \u05dc\u05de\u05e9\u05dc: "YouTube - Google Chrome"\n     \u05de\u05e1\u05e4\u05d9\u05e7 \u05dc\u05db\u05ea\u05d5\u05d1 "YouTube" \u05d0\u05d5 "youtube.com" \u05db\u05d3\u05d9 \u05dc\u05d6\u05d4\u05d5\u05ea \u05d0\u05ea \u05d4\u05d0\u05ea\u05e8.')
        hint.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
        hint.setAlignment(Qt.AlignRight)
        card_layout.addWidget(hint)

        layout.addWidget(card)

    def _add_browser_rule(self):
        keyword = self.browser_keyword_edit.text().strip()
        lang = self.browser_lang_combo.currentText()
        if not keyword:
            return
        row = self.browser_table.rowCount()
        self.browser_table.insertRow(row)
        self.browser_table.setItem(row, 0, QTableWidgetItem(keyword))
        self.browser_table.setItem(row, 1, QTableWidgetItem(_lang_display(lang)))
        self.browser_keyword_edit.clear()

    def _remove_browser_rule(self):
        rows = sorted(set(idx.row() for idx in self.browser_table.selectionModel().selectedRows()), reverse=True)
        for row in rows:
            self.browser_table.removeRow(row)

    # ─── Tab 4: Exclude Words ────────────────────────────

    def _build_exclude_words_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f6ab \u05de\u05d9\u05dc\u05d9\u05dd \u05dc\u05d0 \u05dc\u05d4\u05de\u05e8\u05d4')
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        lbl = QLabel('\u05de\u05d9\u05dc\u05d9\u05dd/\u05e7\u05d9\u05e6\u05d5\u05e8\u05d9\u05dd \u05e9\u05dc\u05d0 \u05dc\u05d4\u05de\u05d9\u05e8')
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(lbl)

        desc = QLabel('\u05d4\u05d5\u05e1\u05e3 \u05de\u05d9\u05dc\u05d9\u05dd \u05d0\u05d5 \u05e7\u05d9\u05e6\u05d5\u05e8\u05d9\u05dd \u05e9\u05d4\u05ea\u05d5\u05db\u05e0\u05d4 \u05dc\u05d0 \u05ea\u05e0\u05e1\u05d4 \u05dc\u05d4\u05de\u05d9\u05e8 (\u05dc\u05de\u05e9\u05dc: lol, ok, brb)')
        desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        desc.setAlignment(Qt.AlignRight)
        card_layout.addWidget(desc)

        self.exclude_list = QListWidget()
        self.exclude_list.setLayoutDirection(Qt.RightToLeft)
        self.exclude_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        for word in self.config.get('exclude_words', []):
            self.exclude_list.addItem(word)
        card_layout.addWidget(self.exclude_list)

        add_row = QHBoxLayout()
        add_row.addStretch()

        self.exclude_word_edit = QLineEdit()
        self.exclude_word_edit.setPlaceholderText('\u05de\u05d9\u05dc\u05d4/\u05e7\u05d9\u05e6\u05d5\u05e8')
        self.exclude_word_edit.setFixedWidth(180)
        add_row.addWidget(self.exclude_word_edit)

        add_btn = QPushButton('\u2795 \u05d4\u05d5\u05e1\u05e3')
        add_btn.clicked.connect(self._add_exclude)
        add_row.addWidget(add_btn)

        del_btn = QPushButton('\U0001f5d1\ufe0f \u05d4\u05e1\u05e8')
        del_btn.setProperty('cssClass', 'danger')
        del_btn.clicked.connect(self._remove_exclude)
        add_row.addWidget(del_btn)

        card_layout.addLayout(add_row)
        layout.addWidget(card)

    def _add_exclude(self):
        word = self.exclude_word_edit.text().strip()
        if not word:
            return
        existing = [self.exclude_list.item(i).text().lower() for i in range(self.exclude_list.count())]
        if word.lower() not in existing:
            self.exclude_list.addItem(word)
        self.exclude_word_edit.clear()

    def _remove_exclude(self):
        for item in reversed(self.exclude_list.selectedItems()):
            self.exclude_list.takeItem(self.exclude_list.row(item))

    # ─── Tab 5: General Settings ─────────────────────────

    def _build_general_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\u2699\ufe0f \u05db\u05dc\u05dc\u05d9')
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet('QScrollArea { border: none; background: transparent; }')
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(8, 8, 8, 8)
        scroll.setWidget(scroll_content)
        tab_layout.addWidget(scroll)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        lbl = QLabel('\u05d4\u05d2\u05d3\u05e8\u05d5\u05ea \u05db\u05dc\u05dc\u05d9\u05d5\u05ea')
        lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(lbl)

        self.enabled_cb = QCheckBox('\u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4 \u05e4\u05e2\u05d9\u05dc')
        self.enabled_cb.setChecked(self.config.get('enabled', True))
        self.enabled_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.enabled_cb)

        self.debug_cb = QCheckBox('(Debug) \u05d4\u05e6\u05d2 \u05dc\u05d5\u05d2 \u05d1\u05e7\u05d5\u05e0\u05e1\u05d5\u05dc\u05d4')
        self.debug_cb.setChecked(self.config.get('debug', False))
        self.debug_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.debug_cb)

        self.auto_switch_cb = QCheckBox('\u05d4\u05d7\u05dc\u05e4\u05ea \u05e9\u05e4\u05d4 \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9\u05ea \u05d0\u05d7\u05e8\u05d9 \u05ea\u05d9\u05e7\u05d5\u05e0\u05d9\u05dd \u05e8\u05e6\u05d5\u05e4\u05d9\u05dd')
        self.auto_switch_cb.setChecked(self.config.get('auto_switch', True))
        self.auto_switch_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.auto_switch_cb)

        count_row = QHBoxLayout()
        count_row.addStretch()
        self.auto_switch_spin = QSpinBox()
        self.auto_switch_spin.setRange(1, 10)
        self.auto_switch_spin.setValue(self.config.get('auto_switch_count', 2))
        self.auto_switch_spin.setFixedWidth(60)
        count_row.addWidget(self.auto_switch_spin)
        count_label = QLabel('\u05de\u05e1\u05e4\u05e8 \u05ea\u05d9\u05e7\u05d5\u05e0\u05d9\u05dd \u05e8\u05e6\u05d5\u05e4\u05d9\u05dd \u05dc\u05d4\u05d7\u05dc\u05e4\u05ea \u05e9\u05e4\u05d4:')
        count_label.setStyleSheet(f"color: {COLORS['text_primary']}; background: transparent;")
        count_row.addWidget(count_label)
        card_layout.addLayout(count_row)

        # Separator
        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background-color: {COLORS['border']}; border: none;")
        card_layout.addWidget(sep)

        disp_lbl = QLabel('\u05ea\u05e6\u05d5\u05d2\u05d4')
        disp_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 11pt; font-weight: bold; background: transparent;")
        disp_lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(disp_lbl)

        self.show_panel_cb = QCheckBox('\u05d4\u05e6\u05d2 \u05d7\u05dc\u05d5\u05e0\u05d9\u05ea \u05d4\u05e7\u05dc\u05d3\u05d4 \u05d1\u05d4\u05e4\u05e2\u05dc\u05d4')
        self.show_panel_cb.setChecked(self.config.get('show_typing_panel', False))
        self.show_panel_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.show_panel_cb)

        self.hide_scores_cb = QCheckBox('\u05d4\u05e1\u05ea\u05e8 \u05e6\u05d9\u05d5\u05e0\u05d9 NLP \u05dc\u05de\u05d9\u05dc\u05d9\u05dd')
        self.hide_scores_cb.setChecked(self.config.get('hide_scores', False))
        self.hide_scores_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.hide_scores_cb)

        # ── Privacy Guard section ──
        sep_priv = QFrame()
        sep_priv.setFixedHeight(1)
        sep_priv.setStyleSheet(f"background-color: {COLORS['border']}; border: none;")
        card_layout.addWidget(sep_priv)

        priv_lbl = QLabel('\U0001f512 \u05d4\u05d2\u05e0\u05ea \u05e4\u05e8\u05d8\u05d9\u05d5\u05ea')
        priv_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 11pt; font-weight: bold; background: transparent;")
        priv_lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(priv_lbl)

        priv_desc = QLabel(
            '\u05d7\u05e1\u05d9\u05de\u05ea \u05d0\u05d9\u05e1\u05d5\u05e3 \u05d1\u05e9\u05d3\u05d5\u05ea \u05e1\u05d9\u05e1\u05de\u05d4, \u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d5\u05ea \u05e8\u05d2\u05d9\u05e9\u05d5\u05ea,\n'
            '\u05d5\u05d6\u05d9\u05d4\u05d5\u05d9 \u05ea\u05d1\u05e0\u05d9\u05d5\u05ea \u05e8\u05d2\u05d9\u05e9\u05d5\u05ea (\u05db\u05e8\u05d8\u05d9\u05e1 \u05d0\u05e9\u05e8\u05d0\u05d9, \u05ea.\u05d6., OTP)'
        )
        priv_desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        priv_desc.setAlignment(Qt.AlignRight)
        priv_desc.setWordWrap(True)
        card_layout.addWidget(priv_desc)

        self.privacy_guard_cb = QCheckBox('\u05d4\u05d2\u05e0\u05ea \u05e4\u05e8\u05d8\u05d9\u05d5\u05ea \u05e4\u05e2\u05d9\u05dc\u05d4 (Password / UIA / Regex)')
        self.privacy_guard_cb.setChecked(self.config.get('privacy_guard', True))
        self.privacy_guard_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.privacy_guard_cb)

        # Info section
        sep2 = QFrame()
        sep2.setFixedHeight(1)
        sep2.setStyleSheet(f"background-color: {COLORS['border']}; border: none;")
        card_layout.addWidget(sep2)

        info_lbl = QLabel('\u2139\ufe0f  AutoLang \u2014 \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05ea \u05de\u05e7\u05dc\u05d3\u05ea \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9')
        info_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 11pt; font-weight: bold; background: transparent;")
        info_lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(info_lbl)

        info_text = (
            "\u2022 \u05d1\u05d5\u05d3\u05e7 \u05d0\u05ea 2 \u05d4\u05de\u05d9\u05dc\u05d9\u05dd \u05d4\u05e8\u05d0\u05e9\u05d5\u05e0\u05d5\u05ea \u05d0\u05d7\u05e8\u05d9 ENTER/\u05e0\u05e7\u05d5\u05d3\u05d4\n"
            "\u2022 \u05de\u05d9\u05dc\u05d4 1 \u05e7\u05d5\u05d1\u05e2\u05ea \u05d0\u05ea \u05e9\u05e4\u05ea \u05d4\u05de\u05e9\u05e4\u05d8, \u05de\u05d9\u05dc\u05d4 2 \u05d4\u05d5\u05dc\u05db\u05ea \u05d0\u05d7\u05e8\u05d9\u05d4\n"
            "\u2022 \u05d1\u05d3\u05d9\u05e7\u05ea \u05d0\u05d5\u05ea\u05d9\u05d5\u05ea \u05e1\u05d5\u05e4\u05d9\u05d5\u05ea (\u05da,\u05dd,\u05df,\u05e3,\u05e5) \u05de\u05d5\u05e0\u05e2\u05ea \u05d4\u05de\u05e8\u05d5\u05ea \u05e9\u05d2\u05d5\u05d9\u05d5\u05ea\n"
            "\u2022 \u05ea\u05e8\u05d2\u05d5\u05dd character-by-character \u05dc\u05e4\u05d9 \u05de\u05d9\u05e4\u05d5\u05d9 \u05de\u05e7\u05dc\u05d3\u05ea"
        )
        info_body = QLabel(info_text)
        info_body.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
        info_body.setAlignment(Qt.AlignRight)
        info_body.setWordWrap(True)
        card_layout.addWidget(info_body)

        card_layout.addStretch()
        layout.addWidget(card)

    # ─── Tab 6: Spell & Grammar ──────────────────────────

    def _build_spell_grammar_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f4dd \u05db\u05ea\u05d9\u05d1 \u05d5\u05e0\u05d9\u05e1\u05d5\u05d7')
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet('QScrollArea { border: none; background: transparent; }')
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(8, 8, 8, 8)
        scroll.setWidget(scroll_content)
        tab_layout.addWidget(scroll)

        # ── Spell Check Section ──
        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)

        spell_lbl = QLabel('\U0001f4dd \u05ea\u05d9\u05e7\u05d5\u05df \u05db\u05ea\u05d9\u05d1')
        spell_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        spell_lbl.setAlignment(Qt.AlignRight)
        card_layout.addWidget(spell_lbl)

        spell_desc = QLabel('\u05d1\u05d3\u05d9\u05e7\u05ea \u05e9\u05d2\u05d9\u05d0\u05d5\u05ea \u05db\u05ea\u05d9\u05d1 \u05de\u05e7\u05d5\u05de\u05d9\u05ea (\u05d0\u05e0\u05d2\u05dc\u05d9\u05ea + \u05e2\u05d1\u05e8\u05d9\u05ea) \u2014 \u05dc\u05dc\u05d0 \u05d0\u05d9\u05e0\u05d8\u05e8\u05e0\u05d8')
        spell_desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        spell_desc.setAlignment(Qt.AlignRight)
        spell_desc.setWordWrap(True)
        card_layout.addWidget(spell_desc)

        self.spell_enabled_cb = QCheckBox('\u05ea\u05d9\u05e7\u05d5\u05df \u05db\u05ea\u05d9\u05d1 \u05e4\u05e2\u05d9\u05dc')
        self.spell_enabled_cb.setChecked(self.config.get('spell_enabled', True))
        self.spell_enabled_cb.setLayoutDirection(Qt.RightToLeft)
        card_layout.addWidget(self.spell_enabled_cb)

        # Mode selection
        mode_row = QHBoxLayout()
        mode_row.addStretch()
        self.spell_mode_combo = QComboBox()
        self.spell_mode_combo.addItems([
            '\U0001f4ac \u05d1\u05dc\u05d5\u05df \u05d4\u05e6\u05e2\u05d4 (tooltip)',
            '\u26a1 \u05ea\u05d9\u05e7\u05d5\u05df \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9',
            '\U0001f534 \u05e1\u05d9\u05de\u05d5\u05df \u05d5\u05d9\u05d6\u05d5\u05d0\u05dc\u05d9 \u05d1\u05dc\u05d1\u05d3',
        ])
        mode_map = {'tooltip': 0, 'auto': 1, 'visual': 2}
        self.spell_mode_combo.setCurrentIndex(mode_map.get(self.config.get('spell_mode', 'tooltip'), 0))
        self.spell_mode_combo.setFixedWidth(220)
        mode_row.addWidget(self.spell_mode_combo)
        mode_label = QLabel('\u05de\u05e6\u05d1 \u05ea\u05d9\u05e7\u05d5\u05df:')
        mode_label.setStyleSheet(f"color: {COLORS['text_primary']}; background: transparent;")
        mode_row.addWidget(mode_label)
        card_layout.addLayout(mode_row)

        layout.addWidget(card)

        # ── Grammar / Phrasing Section ──
        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background-color: {COLORS['border']}; border: none;")
        layout.addWidget(sep)

        card2 = self._make_card()
        card2_layout = QVBoxLayout(card2)
        card2_layout.setContentsMargins(14, 12, 14, 12)

        gram_lbl = QLabel('\U0001f4ac \u05ea\u05d9\u05e7\u05d5\u05df \u05e0\u05d9\u05e1\u05d5\u05d7 / \u05d3\u05e7\u05d3\u05d5\u05e7 (LLM)')
        gram_lbl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 14pt; font-weight: bold; background: transparent;")
        gram_lbl.setAlignment(Qt.AlignRight)
        card2_layout.addWidget(gram_lbl)

        gram_desc = QLabel(
            '\u05ea\u05d9\u05e7\u05d5\u05df \u05d3\u05e7\u05d3\u05d5\u05e7, \u05e0\u05d9\u05e1\u05d5\u05d7 \u05d5\u05e4\u05d9\u05e1\u05d5\u05e7 \u05d1\u05d0\u05de\u05e6\u05e2\u05d5\u05ea AI.\\n'
            '\u05e1\u05de\u05df \u05d8\u05e7\u05e1\u05d8, \u05dc\u05d7\u05e5 Ctrl+Shift+G \u2014 \u05d4\u05d8\u05e7\u05e1\u05d8 \u05d9\u05d9\u05e9\u05dc\u05d7 \u05dc\u05d1\u05d3\u05d9\u05e7\u05d4 \u05d5\u05d9\u05d5\u05d7\u05d6\u05e8 \u05de\u05ea\u05d5\u05e7\u05df.'
        )
        gram_desc.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; background: transparent;")
        gram_desc.setAlignment(Qt.AlignRight)
        gram_desc.setWordWrap(True)
        card2_layout.addWidget(gram_desc)

        self.grammar_enabled_cb = QCheckBox('\u05ea\u05d9\u05e7\u05d5\u05df \u05e0\u05d9\u05e1\u05d5\u05d7 \u05e4\u05e2\u05d9\u05dc')
        self.grammar_enabled_cb.setChecked(self.config.get('grammar_enabled', False))
        self.grammar_enabled_cb.setLayoutDirection(Qt.RightToLeft)
        card2_layout.addWidget(self.grammar_enabled_cb)

        # Provider
        provider_row = QHBoxLayout()
        provider_row.addStretch()
        self.grammar_provider_combo = QComboBox()
        self.grammar_provider_combo.addItems(['OpenAI (GPT)', 'Anthropic (Claude)', 'Google Gemini'])
        provider_map = {'openai': 0, 'anthropic': 1, 'gemini': 2}
        self.grammar_provider_combo.setCurrentIndex(
            provider_map.get(self.config.get('grammar_provider', 'openai'), 0)
        )
        self.grammar_provider_combo.setFixedWidth(220)
        self.grammar_provider_combo.currentIndexChanged.connect(self._on_grammar_provider_changed)
        provider_row.addWidget(self.grammar_provider_combo)
        provider_label = QLabel('\u05e1\u05e4\u05e7:')
        provider_label.setStyleSheet(f"color: {COLORS['text_primary']}; background: transparent;")
        provider_row.addWidget(provider_label)
        card2_layout.addLayout(provider_row)

        # API Key
        key_row = QHBoxLayout()
        key_row.addStretch()
        self.grammar_api_key_edit = QLineEdit()
        self.grammar_api_key_edit.setEchoMode(QLineEdit.Password)
        self.grammar_api_key_edit.setPlaceholderText('sk-... / claude-... / AI...')
        self.grammar_api_key_edit.setText(self.config.get('grammar_api_key', ''))
        self.grammar_api_key_edit.setFixedWidth(300)
        key_row.addWidget(self.grammar_api_key_edit)
        key_label = QLabel('API Key:')
        key_label.setStyleSheet(f"color: {COLORS['text_primary']}; background: transparent;")
        key_row.addWidget(key_label)
        card2_layout.addLayout(key_row)

        # Model
        model_row = QHBoxLayout()
        model_row.addStretch()
        self.grammar_model_combo = QComboBox()
        self.grammar_model_combo.setFixedWidth(220)
        self._populate_grammar_models()
        # Set saved model
        saved_model = self.config.get('grammar_model', '')
        idx = self.grammar_model_combo.findText(saved_model)
        if idx >= 0:
            self.grammar_model_combo.setCurrentIndex(idx)
        model_row.addWidget(self.grammar_model_combo)
        model_label = QLabel('\u05de\u05d5\u05d3\u05dc:')
        model_label.setStyleSheet(f"color: {COLORS['text_primary']}; background: transparent;")
        model_row.addWidget(model_label)
        card2_layout.addLayout(model_row)

        # Hotkey info
        hotkey_lbl = QLabel('\u2328\ufe0f  \u05e7\u05d9\u05e6\u05d5\u05e8: Ctrl+Shift+G  \u2014  \u05e1\u05de\u05df \u05d8\u05e7\u05e1\u05d8 \u05d5\u05dc\u05d7\u05e5 \u05dc\u05ea\u05d9\u05e7\u05d5\u05df')
        hotkey_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9pt; background: transparent;")
        hotkey_lbl.setAlignment(Qt.AlignRight)
        card2_layout.addWidget(hotkey_lbl)

        card2_layout.addStretch()
        layout.addWidget(card2)
        layout.addStretch()

    def _on_grammar_provider_changed(self, index):
        self._populate_grammar_models()

    def _populate_grammar_models(self):
        self.grammar_model_combo.clear()
        provider_keys = ['openai', 'anthropic', 'gemini']
        idx = self.grammar_provider_combo.currentIndex()
        provider = provider_keys[idx] if 0 <= idx < len(provider_keys) else 'openai'
        try:
            import grammar_module
            models = grammar_module.get_provider_models(provider)
            self.grammar_model_combo.addItems(models)
        except Exception:
            self.grammar_model_combo.addItems(['(default)'])

    # ─── Tab 7: Help Guide ───────────────────────────────

    def _build_help_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, '\U0001f4d6 \u05de\u05d3\u05e8\u05d9\u05da')
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)

        card = self._make_card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(4, 4, 4, 4)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['text_primary']};
                border: none;
                padding: 10px;
            }}
        """)

        sections = [
            ('\u05de\u05d4 \u05d6\u05d4 AutoLang?',
             'AutoLang \u05d4\u05d9\u05d0 \u05ea\u05d5\u05db\u05e0\u05d4 \u05e9\u05de\u05ea\u05e7\u05e0\u05ea \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9\u05ea \u05d8\u05e7\u05e1\u05d8 \u05e9\u05d4\u05d5\u05e7\u05dc\u05d3 \u05d1\u05e9\u05e4\u05d4 \u05d4\u05dc\u05d0 \u05e0\u05db\u05d5\u05e0\u05d4.\n'
             '\u05dc\u05de\u05e9\u05dc \u2014 \u05d0\u05dd \u05d4\u05de\u05e7\u05dc\u05d3\u05ea \u05e9\u05dc\u05da \u05e2\u05dc \u05d0\u05e0\u05d2\u05dc\u05d9\u05ea \u05d5\u05d0\u05ea\u05d4 \u05de\u05e7\u05dc\u05d9\u05d3 \u05e2\u05d1\u05e8\u05d9\u05ea, \u05d4\u05ea\u05d5\u05db\u05e0\u05d4 \u05de\u05d6\u05d4\u05d4 \u05d0\u05ea \u05d4\u05d8\u05e2\u05d5\u05ea \u05d5\u05de\u05ea\u05e7\u05e0\u05ea \u05d0\u05d5\u05ea\u05d4 \u05d1\u05d6\u05de\u05df \u05d0\u05de\u05ea.'),
            ('\U0001f535 \u05db\u05e4\u05ea\u05d5\u05e8\u05d9 \u05d4\u05e2\u05d9\u05d2\u05d5\u05dc',
             '\u23fb  \u05dc\u05de\u05e2\u05dc\u05d4 \u2014 \u05d4\u05e4\u05e2\u05dc\u05d4 / \u05d4\u05e9\u05d1\u05ea\u05d4 \u05e9\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d4\u05e9\u05e4\u05d4\n'
             '\u21a9  \u05e9\u05de\u05d0\u05dc \u2014 \u05d1\u05d9\u05d8\u05d5\u05dc \u05d4\u05ea\u05d9\u05e7\u05d5\u05df \u05d4\u05d0\u05d7\u05e8\u05d5\u05df (\u05d2\u05dd F12)\n'
             '\u270e  \u05dc\u05de\u05d8\u05d4 \u2014 \u05e4\u05ea\u05d9\u05d7\u05ea \u05e2\u05d5\u05e8\u05da \u05ea\u05e8\u05d2\u05d5\u05dd\n'
             '\u2699  \u05d9\u05de\u05d9\u05df \u2014 \u05e4\u05ea\u05d9\u05d7\u05ea \u05d7\u05dc\u05d5\u05df \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea\n'
             '\U0001f534  \u05de\u05e8\u05db\u05d6 \u2014 \u05d4\u05e7\u05dc\u05d8\u05d4 \u05e7\u05d5\u05dc\u05d9\u05ea (Speech-to-Text)'),
            ('\U0001f50d \u05ea\u05e8\u05d2\u05d5\u05dd \u05de\u05d4\u05d9\u05e8',
             '\u05e1\u05de\u05df \u05d8\u05e7\u05e1\u05d8 \u05d1\u05db\u05dc \u05d7\u05dc\u05d5\u05df \u05d5\u05dc\u05d7\u05e5 \u05e7\u05dc\u05d9\u05e7 \u05d9\u05de\u05e0\u05d9 \u2014 \u05d9\u05d5\u05e4\u05d9\u05e2 \u05db\u05e4\u05ea\u05d5\u05e8 "\u05ea\u05e8\u05d2\u05dd".\n'
             '\u05dc\u05d7\u05d9\u05e6\u05d4 \u05e2\u05dc\u05d9\u05d5 \u05ea\u05ea\u05e8\u05d2\u05dd \u05d0\u05ea \u05d4\u05d8\u05e7\u05e1\u05d8 \u05d4\u05de\u05e1\u05d5\u05de\u05df \u05d5\u05ea\u05d3\u05d1\u05d9\u05e7 \u05d0\u05d5\u05ea\u05d5 \u05d1\u05d7\u05dc\u05d5\u05df \u05d4\u05e4\u05e2\u05d9\u05dc.'),
            ('\U0001f3a4 \u05d4\u05e7\u05dc\u05d8\u05d4 \u05e7\u05d5\u05dc\u05d9\u05ea',
             '\u05dc\u05d7\u05d9\u05e6\u05d4 \u05e2\u05dc \u05d4\u05db\u05e4\u05ea\u05d5\u05e8 \u05d4\u05d0\u05d3\u05d5\u05dd \u05d1\u05de\u05e8\u05db\u05d6 \u05d4\u05e2\u05d9\u05d2\u05d5\u05dc \u05de\u05e4\u05e2\u05d9\u05dc\u05d4 \u05d4\u05e7\u05dc\u05d8\u05d4.\n'
             '\u05d4\u05ea\u05d5\u05db\u05e0\u05d4 \u05de\u05e7\u05e9\u05d9\u05d1\u05d4 \u05dc\u05d3\u05d9\u05d1\u05d5\u05e8, \u05de\u05d6\u05d4\u05d4 \u05d8\u05e7\u05e1\u05d8 (Whisper) \u05d5\u05de\u05d3\u05d1\u05d9\u05e7\u05d4 \u05d0\u05d5\u05ea\u05d5 \u05d1\u05d7\u05dc\u05d5\u05df \u05d4\u05e4\u05e2\u05d9\u05dc.\n'
             '\u05dc\u05d7\u05d9\u05e6\u05d4 \u05e0\u05d5\u05e1\u05e4\u05ea \u05e2\u05d5\u05e6\u05e8\u05ea \u05d0\u05ea \u05d4\u05d4\u05e7\u05dc\u05d8\u05d4.'),
            ('\u2699\ufe0f \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea',
             '\u05d0\u05e4\u05dc\u05d9\u05e7\u05e6\u05d9\u05d5\u05ea \u2014 \u05d4\u05d2\u05d3\u05e8\u05ea \u05e9\u05e4\u05ea \u05d1\u05e8\u05d9\u05e8\u05ea \u05de\u05d7\u05d3\u05dc \u05dc\u05db\u05dc \u05ea\u05d5\u05db\u05e0\u05d4 (\u05dc\u05e4\u05d9 \u05e9\u05dd \u05d4-exe).\n'
             '\u05e6\'\u05d0\u05d8\u05d9\u05dd \u2014 \u05e9\u05e4\u05d4 \u05e9\u05d5\u05e0\u05d4 \u05dc\u05e4\u05d9 \u05e9\u05dd \u05e9\u05d9\u05d7\u05d4 (WhatsApp, Teams).\n'
             '\u05d3\u05e4\u05d3\u05e4\u05e0\u05d9\u05dd \u2014 \u05e9\u05e4\u05d4 \u05dc\u05e4\u05d9 \u05de\u05d9\u05dc\u05ea \u05de\u05e4\u05ea\u05d7 \u05d1\u05db\u05d5\u05ea\u05e8\u05ea \u05d4\u05d8\u05d0\u05d1.\n'
             '\u05de\u05d9\u05dc\u05d9\u05dd \u05de\u05d5\u05d7\u05e8\u05d2\u05d5\u05ea \u2014 \u05de\u05d9\u05dc\u05d9\u05dd \u05e9\u05dc\u05d0 \u05d9\u05ea\u05d5\u05e7\u05e0\u05d5 (\u05e9\u05de\u05d5\u05ea, \u05de\u05d5\u05e0\u05d7\u05d9\u05dd \u05d8\u05db\u05e0\u05d9\u05d9\u05dd).\n'
             '\u05db\u05dc\u05dc\u05d9 \u2014 \u05d4\u05e4\u05e2\u05dc\u05d4/\u05d4\u05e9\u05d1\u05ea\u05d4, Debug, \u05d4\u05d7\u05dc\u05e4\u05ea \u05e9\u05e4\u05d4 \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9\u05ea, \u05ea\u05e6\u05d5\u05d2\u05d4.'),
            ('\U0001f9e0 \u05d0\u05d9\u05da \u05d4\u05d6\u05d9\u05d4\u05d5\u05d9 \u05e2\u05d5\u05d1\u05d3?',
             '\u05db\u05dc \u05de\u05d9\u05dc\u05d4 \u05e9\u05de\u05d5\u05e7\u05dc\u05d3\u05ea \u05e0\u05d1\u05d3\u05e7\u05ea \u05de\u05d5\u05dc \u05de\u05d0\u05d2\u05e8 \u05de\u05d9\u05dc\u05d9\u05dd (NLP) \u05d1\u05e2\u05d1\u05e8\u05d9\u05ea \u05d5\u05d1\u05d0\u05e0\u05d2\u05dc\u05d9\u05ea.\n'
             '\u05d4\u05de\u05e2\u05e8\u05db\u05ea \u05de\u05e9\u05d5\u05d5\u05d4 \u05e6\u05d9\u05d5\u05e0\u05d9\u05dd (zipf) \u05d5\u05d1\u05d5\u05d7\u05e8\u05ea \u05d0\u05ea \u05d4\u05e9\u05e4\u05d4 \u05e9\u05d4\u05de\u05d9\u05dc\u05d4 \u05e9\u05d9\u05d9\u05db\u05ea \u05d0\u05dc\u05d9\u05d4.\n'
             '\u05d0\u05d7\u05e8\u05d9 2 \u05de\u05d9\u05dc\u05d9\u05dd \u05d1\u05e8\u05e6\u05e3, \u05d4\u05de\u05e0\u05d5\u05e2 "\u05e0\u05d5\u05e2\u05dc" \u05d0\u05ea \u05e9\u05e4\u05ea \u05d4\u05de\u05e9\u05e4\u05d8.\n'
             'ENTER \u05d0\u05d5 \u05e0\u05e7\u05d5\u05d3\u05d4 \u05de\u05d0\u05e4\u05e1\u05d9\u05dd \u05d0\u05ea \u05d4\u05e0\u05e2\u05d9\u05dc\u05d4 \u05d5\u05de\u05ea\u05d7\u05d9\u05dc\u05d9\u05dd \u05de\u05e9\u05e4\u05d8 \u05d7\u05d3\u05e9.'),
            ('\u2328\ufe0f \u05e7\u05d9\u05e6\u05d5\u05e8\u05d9 \u05de\u05e7\u05e9\u05d9\u05dd',
             'F12 \u2014 \u05d1\u05d9\u05d8\u05d5\u05dc \u05ea\u05d9\u05e7\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df\n'
             'Ctrl+Alt+Q \u2014 \u05d9\u05e6\u05d9\u05d0\u05d4 \u05de\u05d4\u05ea\u05d5\u05db\u05e0\u05d4\n'
             'Ctrl+Alt+I \u2014 \u05d4\u05e6\u05d2\u05ea \u05de\u05d9\u05d3\u05e2 (Debug)'),
        ]

        html = ''
        for title_text, body in sections:
            html += f'<h3 style="color: {COLORS["accent"]}; margin-top: 12px;">{title_text}</h3>'
            body_html = body.replace('\n', '<br>')
            html += f'<p style="color: {COLORS["text_primary"]}; line-height: 1.5;">{body_html}</p>'

        help_text.setHtml(html)
        card_layout.addWidget(help_text)
        layout.addWidget(card)

    # ─── Save ────────────────────────────────────────────

    def _save(self):
        # App defaults
        app_defaults = {}
        for row in range(self.app_table.rowCount()):
            app = self.app_table.item(row, 0).text()
            lang = _lang_code_from_display(self.app_table.item(row, 1).text())
            app_defaults[app] = lang

        # Chat defaults
        chat_defaults = {}
        for row in range(self.chat_table.rowCount()):
            exe = self.chat_table.item(row, 0).text()
            title = self.chat_table.item(row, 1).text()
            lang = _lang_code_from_display(self.chat_table.item(row, 2).text())
            if exe not in chat_defaults:
                chat_defaults[exe] = {}
            chat_defaults[exe][title] = lang

        # Browser defaults
        browser_defaults = {}
        for row in range(self.browser_table.rowCount()):
            keyword = self.browser_table.item(row, 0).text()
            lang = _lang_code_from_display(self.browser_table.item(row, 1).text())
            browser_defaults[keyword] = lang

        # Exclude words
        exclude_words = [self.exclude_list.item(i).text() for i in range(self.exclude_list.count())]

        self.config['app_defaults'] = app_defaults
        self.config['chat_defaults'] = chat_defaults
        self.config['browser_defaults'] = browser_defaults
        self.config['exclude_words'] = exclude_words

        # Auto-sync watch_title_exes
        current_watch = set(self.config.get('watch_title_exes', []))
        current_watch.update(chat_defaults.keys())
        self.config['watch_title_exes'] = sorted(current_watch)

        self.config['enabled'] = self.enabled_cb.isChecked()
        self.config['debug'] = self.debug_cb.isChecked()
        self.config['auto_switch'] = self.auto_switch_cb.isChecked()
        self.config['auto_switch_count'] = self.auto_switch_spin.value()
        self.config['hide_scores'] = self.hide_scores_cb.isChecked()
        self.config['show_typing_panel'] = self.show_panel_cb.isChecked()
        self.config['privacy_guard'] = self.privacy_guard_cb.isChecked()

        # Spell & Grammar
        self.config['spell_enabled'] = self.spell_enabled_cb.isChecked()
        mode_map_rev = {0: 'tooltip', 1: 'auto', 2: 'visual'}
        self.config['spell_mode'] = mode_map_rev.get(self.spell_mode_combo.currentIndex(), 'tooltip')

        self.config['grammar_enabled'] = self.grammar_enabled_cb.isChecked()
        provider_keys = ['openai', 'anthropic', 'gemini']
        p_idx = self.grammar_provider_combo.currentIndex()
        self.config['grammar_provider'] = provider_keys[p_idx] if 0 <= p_idx < len(provider_keys) else 'openai'
        self.config['grammar_api_key'] = self.grammar_api_key_edit.text().strip()
        self.config['grammar_model'] = self.grammar_model_combo.currentText().strip()

        save_config(self.config)
        self.status_label.setText('\u2705 \u05d4\u05d4\u05d2\u05d3\u05e8\u05d5\u05ea \u05e0\u05e9\u05de\u05e8\u05d5 \u05d1\u05d4\u05e6\u05dc\u05d7\u05d4!')

        if self.on_save:
            self.on_save(self.config)

        QTimer.singleShot(2000, lambda: self.status_label.setText(''))


# ──────────────────────────────────────────────────────────
# System Tray Application
# ──────────────────────────────────────────────────────────

class AutoLangTray:
    """System Tray application — replaces pystray with QSystemTrayIcon."""

    def __init__(self):
        self.config = load_config()
        self.enabled = self.config.get('enabled', True)
        self.widget = None
        self.settings_window = None
        self.tray_icon = None
        self.app = None

    def _toggle_enabled(self, *args):
        self.enabled = not self.enabled
        self.config['enabled'] = self.enabled
        save_config(self.config)
        self.tray_icon.setIcon(create_tray_icon(self.enabled))
        if self.widget:
            self.widget.update_state(self.enabled)
        try:
            import auto_lang
            auto_lang.ENGINE_ENABLED = self.enabled
        except Exception:
            pass
        # Update menu text
        if self.enabled:
            self._toggle_action.setText('\u2705 \u05e4\u05e2\u05d9\u05dc')
        else:
            self._toggle_action.setText('\u274c \u05de\u05d5\u05e9\u05d1\u05ea')

    def _open_settings(self, *args):
        if self.settings_window and self.settings_window.isVisible():
            self.settings_window.raise_()
            self.settings_window.activateWindow()
            return
        self.config = load_config()
        self.settings_window = SettingsWindow(self.config, on_save_callback=self._on_settings_saved)
        self.settings_window.show()

    def _on_settings_saved(self, new_config):
        self.config = new_config
        self.enabled = new_config.get('enabled', True)
        if self.tray_icon:
            self.tray_icon.setIcon(create_tray_icon(self.enabled))
        if self.widget:
            self.widget.update_state(self.enabled)
            new_hide = new_config.get('hide_scores', False)
            new_panel = new_config.get('show_typing_panel', False)
            if self.widget._hide_scores != new_hide:
                self.widget._hide_scores = new_hide
                self.widget.update()
            if self.widget._panel_visible != new_panel:
                self.widget._panel_visible = new_panel
                self.widget._relayout()
        try:
            import auto_lang
            self._apply_config_to_engine(auto_lang, new_config)
        except Exception:
            pass

    @staticmethod
    def _apply_config_to_engine(engine, cfg):
        engine.APP_DEFAULT_LANG_BY_EXE.clear()
        engine.APP_DEFAULT_LANG_BY_EXE.update(cfg.get('app_defaults', {}))

        engine.APP_DEFAULT_LANG_BY_TITLE.clear()
        for exe, chats in cfg.get('chat_defaults', {}).items():
            if chats:
                engine.APP_DEFAULT_LANG_BY_TITLE[exe] = chats

        engine.WATCH_TITLE_CHANGES_EXE.clear()
        engine.WATCH_TITLE_CHANGES_EXE.update(cfg.get('watch_title_exes', []))
        engine.WATCH_TITLE_CHANGES_EXE.update(engine.APP_DEFAULT_LANG_BY_TITLE.keys())

        engine.EXCLUDE_WORDS = set(w.lower() for w in cfg.get('exclude_words', []))

        engine.BROWSER_LANG_BY_KEYWORD.clear()
        engine.BROWSER_LANG_BY_KEYWORD.update(cfg.get('browser_defaults', {}))
        if engine.BROWSER_LANG_BY_KEYWORD:
            engine.WATCH_TITLE_CHANGES_EXE.update(engine.BROWSER_EXES)

        engine.ENGINE_ENABLED = cfg.get('enabled', True)
        engine.DEBUG = cfg.get('debug', False)
        engine.AUTO_SWITCH_LAYOUT = cfg.get('auto_switch', True)
        engine.AUTO_SWITCH_AFTER_CONSECUTIVE = cfg.get('auto_switch_count', 2)

        # Privacy Guard
        engine.PRIVACY_GUARD_ENABLED = cfg.get('privacy_guard', True)
        blocked = cfg.get('privacy_blocked_exes', [])
        engine.PRIVACY_BLOCKED_EXE = set(e.lower() for e in blocked)

        # Spell check
        try:
            import spell_module
            spell_module.SPELL_ENABLED = cfg.get('spell_enabled', True)
            spell_module.SPELL_MODE = cfg.get('spell_mode', 'tooltip')
        except Exception:
            pass

        # Grammar / LLM
        try:
            import grammar_module
            grammar_module.GRAMMAR_ENABLED = cfg.get('grammar_enabled', False)
            grammar_module.GRAMMAR_PROVIDER = cfg.get('grammar_provider', 'openai')
            grammar_module.GRAMMAR_API_KEY = cfg.get('grammar_api_key', '')
            grammar_module.GRAMMAR_MODEL = cfg.get('grammar_model', '')
        except Exception:
            pass

    def _quit(self, *args):
        if self.widget:
            self.widget.close()
        try:
            import auto_lang
            auto_lang.stop_event.set()
        except Exception:
            pass
        if self.tray_icon:
            self.tray_icon.hide()
        os._exit(0)

    def _start_engine(self):
        try:
            import auto_lang
            self._apply_config_to_engine(auto_lang, self.config)
            auto_lang.main()
        except Exception as e:
            print(f'Engine start failed: {e}')
            import traceback
            traceback.print_exc()

    def run(self):
        self.app = QApplication.instance() or QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        self.app.setStyleSheet(DARK_QSS)

        # System tray icon
        self.tray_icon = QSystemTrayIcon(create_tray_icon(self.enabled))
        self.tray_icon.setToolTip('AutoLang - \u05ea\u05d9\u05e7\u05d5\u05df \u05e9\u05e4\u05d4 \u05d0\u05d5\u05d8\u05d5\u05de\u05d8\u05d9')

        menu = QMenu()
        menu.setStyleSheet(f"""
            QMenu {{
                background-color: {COLORS['bg_medium']};
                color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']};
                padding: 4px;
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
            QMenu::item {{
                padding: 6px 16px;
            }}
            QMenu::item:selected {{
                background-color: {COLORS['accent_dim']};
                color: #ffffff;
            }}
            QMenu::separator {{
                height: 1px;
                background-color: {COLORS['border']};
                margin: 4px 8px;
            }}
        """)
        self._toggle_action = menu.addAction('\u2705 \u05e4\u05e2\u05d9\u05dc')
        self._toggle_action.triggered.connect(self._toggle_enabled)
        menu.addSeparator()
        menu.addAction('\u2699\ufe0f \u05d4\u05d2\u05d3\u05e8\u05d5\u05ea', self._open_settings)
        menu.addSeparator()
        menu.addAction('\U0001f6aa \u05d9\u05e6\u05d9\u05d0\u05d4', self._quit)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(
            lambda reason: self._toggle_enabled() if reason == QSystemTrayIcon.DoubleClick else None)
        self.tray_icon.show()

        # Start engine in background thread
        engine_thread = threading.Thread(target=self._start_engine, daemon=True)
        engine_thread.start()

        # Create and show floating widget
        self.widget = FloatingWidget(self)
        self.widget.start()
        self.widget.show()

        # Run Qt event loop
        self.app.exec()


def main():
    app = AutoLangTray()
    app.run()


if __name__ == '__main__':
    main()
