"""
AutoLang UI - System Tray + Settings Window
============================================
ממשק משתמש לסקריפט auto_lang.py:
- אייקון ב-System Tray עם חיווי שהתוכנה רצה
- חלון הגדרות עם טאבים:
  1. שפות ברירת מחדל לאפליקציות
  2. שפות ברירת מחדל לצ'אטים
  3. מילים לא להמרה (exclude list)
  4. הגדרות כלליות
"""

import json
import os
import sys

# When running as --noconsole EXE (PyInstaller), sys.stdout/stderr are None.
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

import threading
import tkinter as tk
from tkinter import ttk, messagebox
from io import BytesIO

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    Image = None

try:
    import pystray
except ImportError:
    pystray = None

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
        # exe -> { title_substring: lang }
        'whatsapp.exe': {},
        'ms-teams.exe': {},
        'teams.exe': {},
    },
    'watch_title_exes': [           # אפליקציות נוספות לעקוב אחרי שינוי כותרת (בלי chat defaults)
        'chrome.exe',
        'msedge.exe',
        'outlook.exe',
        'olk.exe',
    ],
    'browser_defaults': {},         # keyword → lang — matches any browser tab title
    'exclude_words': [],  # מילים לא להמיר
    'enabled': True,
    'debug': False,
    'auto_switch': True,
    'auto_switch_count': 2,
    'hide_scores': False,
    'show_typing_panel': False,
}


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            # Merge with defaults
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

# Language display info: code -> (display_name, flag)
LANG_DISPLAY = {
    'en': ('English', '🇺🇸'),
    'he': ('עברית', '🇮🇱'),
    'ru': ('Русский', '🇷🇺'),
    'ar': ('العربية', '🇸🇦'),
    'uk': ('Українська', '🇺🇦'),
    'fr': ('Français', '🇫🇷'),
    'de': ('Deutsch', '🇩🇪'),
    'es': ('Español', '🇪🇸'),
    'el': ('Ελληνικά', '🇬🇷'),
    'fa': ('فارسی', '🇮🇷'),
    'tr': ('Türkçe', '🇹🇷'),
    'th': ('ภาษาไทย', '🇹🇭'),
    'hi': ('हिन्दी', '🇮🇳'),
    'ko': ('한국어', '🇰🇷'),
    'pl': ('Polski', '🇵🇱'),
}


def _lang_display(code: str) -> str:
    """Get display string for a language code."""
    info = LANG_DISPLAY.get(code)
    if info:
        return f'{info[0]} {info[1]}'
    return code


def _lang_code_from_display(display: str) -> str:
    """Reverse: get language code from display string."""
    for code, (name, flag) in LANG_DISPLAY.items():
        if f'{name} {flag}' == display or code == display:
            return code
    return display


def _get_available_lang_codes() -> list[str]:
    """Get list of available language codes (en + whatever is installed)."""
    codes = ['en', 'he']  # defaults always available
    try:
        from keyboard_maps import ALL_PROFILES
        for code in ALL_PROFILES:
            if code not in codes:
                codes.append(code)
    except ImportError:
        pass
    return codes


# ----------------------------
# System Tray Icon
# ----------------------------

def create_tray_icon_image(enabled: bool = True) -> 'Image':
    """Create a modern icon with gradient feel."""
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    if enabled:
        # Outer glow ring — cyan
        draw.ellipse([2, 2, size - 2, size - 2], fill=(15, 60, 120, 80))
        # Main circle — deep blue
        draw.ellipse([5, 5, size - 5, size - 5], fill=(20, 75, 140, 255))
        # Inner highlight — bright cyan
        draw.ellipse([8, 8, size - 8, size - 8], fill=(78, 201, 240, 255))
    else:
        draw.ellipse([2, 2, size - 2, size - 2], fill=(120, 50, 50, 80))
        draw.ellipse([5, 5, size - 5, size - 5], fill=(160, 60, 60, 255))
        draw.ellipse([8, 8, size - 8, size - 8], fill=(248, 113, 113, 255))

    # Letters "אa" in white
    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except Exception:
        font = ImageFont.load_default()

    draw.text((12, 12), "אa", fill=(255, 255, 255, 255), font=font)
    return img


# ──────────────────────────────────────────────────────────
# Floating Widget (always-on-top draggable icon)
# ──────────────────────────────────────────────────────────

class FloatingWidget:
    """
    חלון צף קטן עם האייקון של התוכנה.
    תמיד מעל כל החלונות, ניתן לגרירה, לחיצה פותחת תפריט פעולות.
    מוצג רק כשתיקון השגיאות פעיל.
    """

    WIDGET_SIZE = 68          # px – height of the floating widget (3 lines)
    OPACITY = 0.92            # window opacity
    EDGE_SNAP_MARGIN = 8      # margin from screen edge when snapping
    PANEL_WIDTH = 220         # width of the text preview panel

    def __init__(self, tray_app: 'AutoLangTray'):
        self.tray = tray_app
        self.root: tk.Tk | None = None
        self._drag_data = {'x': 0, 'y': 0}
        self._menu: tk.Menu | None = None
        self._canvas = None
        self._visible = False
        # Load panel/scores prefs from config
        _cfg = tray_app.config if tray_app else {}
        self._panel_visible = _cfg.get('show_typing_panel', False)
        self._hide_scores = _cfg.get('hide_scores', False)
        # Canvas item IDs
        self._word_items_line1 = []     # list of canvas text item IDs for line 1
        self._word_items_line2 = []     # list of canvas text item IDs for line 2
        self._score_items = []          # canvas text item IDs for score line
        self._icon_text_id = None
        # Cached text for redraws
        self._cur_line1 = ' '
        self._cur_line2 = ' '
        # Correction map: corrected_word -> original_word
        self._corrections = {}          # current corrections from engine
        # Per-word NLP scores: word -> {lang_flag: score}
        self._word_scores = {}          # current word scores from engine
        # Edit mode (text input area – drawn on the SAME canvas)
        self._edit_mode = False
        self._editor_height = 0         # 0 when closed, >0 when open
        self._text_widget = None        # tk.Text for input
        self._trans_widget = None       # tk.Text for translation output
        self._status_label = None       # Status label in edit area
        self._send_btn = None           # Send button label
        self._auto_translate_after_id = None  # Debounce timer for auto-translate
        self._send_target_hwnd = None   # HWND to send translation to
        self._edit_canvas_ids = []      # canvas item IDs for the editor area
        self._resize_grip = None        # resize handle widget
        self._resize_drag_y = None      # y at resize drag start
        self._editor_min_h = 148        # minimum editor height
        self._editor_max_h = 500        # maximum editor height
        # Speech-to-text state
        self._speech_recording = False
        self._speech_pulse_id = None    # after() id for pulse animation
        self._last_speech_text = ''     # dedup last recognized text
        self._last_speech_time = 0      # timestamp of last speech
        # Right-click translate popup
        self._translate_popup = None        # Toplevel popup window
        self._translate_popup_after = None  # after() id for auto-hide
        self._pending_selection = ''        # text copied on right-click
        self._send_target_hwnd = None       # HWND of window under cursor on right-click
        # Tooltip state
        self._tooltip_win = None            # Toplevel tooltip window
        self._tooltip_after_id = None       # after() id for delayed show
        self._tooltip_texts = {
            'q_top':    'הפעל / השבת תיקון שפה',
            'q_left':   'בטל תיקון אחרון',
            'q_bottom': 'פתח עורך תרגום',
            'q_right':  'הגדרות',
            'q_center': 'הקלטה קולית',
        }

    # ── public API (called from any thread) ──────────────

    def start(self):
        """Launch the widget's own Tk mainloop in a daemon thread."""
        threading.Thread(target=self._run, daemon=True).start()

    def update_state(self, enabled: bool):
        """Reflect engine enabled/disabled state."""
        if self.root is None:
            return
        try:
            self.root.after_idle(self._apply_state, enabled)
        except Exception:
            pass

    def _on_buffer_update(self, actual: str, translated: str,
                           corrections: dict | None = None,
                           word_scores: dict | None = None):
        """Called from engine thread when buffer changes."""
        if self.root is None:
            return
        try:
            self.root.after_idle(self._set_panel_text, actual, translated,
                                corrections or {}, word_scores or {})
        except Exception:
            pass

    def _set_panel_text(self, actual: str, translated: str,
                        corrections: dict | None = None,
                        word_scores: dict | None = None):
        """Update the panel text items (runs on Tk thread)."""
        if self._canvas is None or not self._panel_visible:
            return
        # Show last ~25 chars to keep it compact
        a = actual[-25:] if len(actual) > 25 else actual
        t = translated[-25:] if len(translated) > 25 else translated
        self._cur_line1 = a if a else ' '
        self._cur_line2 = t if t else ' '
        self._corrections = corrections or {}
        self._word_scores = word_scores or {}
        self._redraw_text_words()

    def destroy(self):
        if self.root:
            try:
                self.root.after_idle(self.root.destroy)
            except Exception:
                pass

    # ── internal ─────────────────────────────────────────

    def _run(self):
        self.root = tk.Tk()
        self.root.withdraw()           # hide until ready

        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH

        # Frameless, topmost, transparent-bg window
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', self.OPACITY)
        self.root.configure(bg='#010101')
        self.root.wm_attributes('-transparentcolor', '#010101')

        # Total window: seamless pill = panel + icon (no gap)
        total_w = pw + sz
        total_h = sz

        # Position: bottom-right, above taskbar
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = screen_w - total_w - self.EDGE_SNAP_MARGIN - 12
        y = screen_h - sz - 60      # above taskbar
        self.root.geometry(f'{total_w}x{total_h}+{x}+{y}')

        # Single unified canvas for the entire widget
        self._canvas = tk.Canvas(self.root, width=total_w, height=total_h,
                                  bg='#010101', highlightthickness=0, bd=0)
        self._canvas.pack(fill='both', expand=True)

        # Draw the full widget
        self._draw_widget()

        # Context menu
        self._build_menu()

        # Bindings (single canvas handles everything)
        self._canvas.bind('<Button-1>', self._on_click_start)
        self._canvas.bind('<B1-Motion>', self._on_drag)
        self._canvas.bind('<ButtonRelease-1>', self._on_click_release)
        self._canvas.bind('<Button-3>', self._show_menu)

        # Hover effects for icon quadrants
        for tag in ('q_top', 'q_left', 'q_right', 'q_bottom', 'q_center'):
            self._canvas.tag_bind(tag, '<Enter>',
                lambda e, t=tag: self._q_hover(t, True))
            self._canvas.tag_bind(tag, '<Leave>',
                lambda e, t=tag: self._q_hover(t, False))

        # Register engine callback + tell engine our HWND
        try:
            import auto_lang
            auto_lang._buffer_callback = self._on_buffer_update
            # Tell the engine to ignore focus changes to our widget
            auto_lang._widget_hwnd = self.root.winfo_id()
        except Exception:
            pass

        # Preload Whisper model in background for faster first speech
        try:
            import speech_module
            speech_module.preload_model()
        except Exception:
            pass

        # Right-click → show translate popup  (use mouse.hook because on_right_click is unreliable)
        try:
            import mouse as _mouse
            def _rc_hook(event):
                if isinstance(event, _mouse.ButtonEvent) and event.event_type == 'up' and event.button == 'right':
                    self._on_right_click_global()
            _mouse.hook(_rc_hook)
            try:
                import auto_lang as _al2
                _al2._dbg('[UI] Right-click hook registered OK (via mouse.hook)')
            except Exception:
                pass
        except Exception as e:
            try:
                import auto_lang as _al3
                _al3._dbg(f'[UI] Failed to register right-click hook: {e}')
            except Exception:
                pass

        # Show
        self._apply_state(self.tray.enabled)
        self.root.deiconify()
        self._visible = True

        self.root.mainloop()

    # ── drawing helpers ──────────────────────────────────

    def _create_pill_shape(self, x1, y1, x2, y2, **kwargs):
        """Draw a true capsule/stadium shape with perfect semicircle ends."""
        h = y2 - y1
        r = h // 2
        fill = kwargs.get('fill', '')
        outline = kwargs.get('outline', '')
        width = kwargs.get('width', 1)

        # Fill: two full circles at ends + connecting rectangle
        self._canvas.create_oval(x1, y1, x1 + h, y2, fill=fill, outline='')
        self._canvas.create_oval(x2 - h, y1, x2, y2, fill=fill, outline='')
        self._canvas.create_rectangle(x1 + r, y1, x2 - r, y2,
                                      fill=fill, outline='')

        # Outline: two semicircle arcs + top/bottom lines
        if outline:
            self._canvas.create_arc(x1, y1, x1 + h, y2,
                                    start=90, extent=180, style='arc',
                                    outline=outline, width=width)
            self._canvas.create_arc(x2 - h, y1, x2, y2,
                                    start=270, extent=180, style='arc',
                                    outline=outline, width=width)
            self._canvas.create_line(x1 + r, y1, x2 - r, y1,
                                     fill=outline, width=width)
            self._canvas.create_line(x1 + r, y2, x2 - r, y2,
                                     fill=outline, width=width)

    def _draw_unified_shape(self, total_w, total_h):
        """Draw ONE seamless shape: pill top + straight sides + rounded bottom corners."""
        import math
        steps = 24
        sz = self.WIDGET_SIZE
        pill_r = (sz - 2) // 2      # semicircle radius of pill = 33
        br = 20                      # bottom corner radius
        x1, y1, x2 = 1, 1, total_w - 1

        # Arc centers
        lcx, lcy = x1 + pill_r, y1 + pill_r          # left pill semicircle
        rcx, rcy = x2 - pill_r, y1 + pill_r          # right pill semicircle
        brcx, brcy = x2 - br, total_h - 1 - br      # bottom-right corner
        blcx, blcy = x1 + br, total_h - 1 - br      # bottom-left corner

        # Build single closed polygon (clockwise in screen coords)
        pts = []
        # 1) Top line between pill arcs
        pts += [lcx, y1, rcx, y1]
        # 2) Top-right pill arc (top → right)
        for i in range(steps + 1):
            t = 3 * math.pi / 2 + (math.pi / 2) * i / steps
            pts.append(rcx + pill_r * math.cos(t))
            pts.append(rcy + pill_r * math.sin(t))
        # 3) Right side down
        pts += [x2, brcy]
        # 4) Bottom-right arc (right → bottom)
        for i in range(steps + 1):
            t = (math.pi / 2) * i / steps
            pts.append(brcx + br * math.cos(t))
            pts.append(brcy + br * math.sin(t))
        # 5) Bottom line
        pts += [blcx, total_h - 1]
        # 6) Bottom-left arc (bottom → left)
        for i in range(steps + 1):
            t = math.pi / 2 + (math.pi / 2) * i / steps
            pts.append(blcx + br * math.cos(t))
            pts.append(blcy + br * math.sin(t))
        # 7) Left side up
        pts += [x1, lcy]
        # 8) Top-left pill arc (left → top)
        for i in range(steps + 1):
            t = math.pi + (math.pi / 2) * i / steps
            pts.append(lcx + pill_r * math.cos(t))
            pts.append(lcy + pill_r * math.sin(t))

        self._canvas.create_polygon(pts, fill=COLORS['bg_medium'],
                                     outline=COLORS['border'], width=1,
                                     smooth=False)

    def _draw_widget(self):
        """Draw (or redraw) the entire floating widget on the canvas."""
        self._canvas.delete('all')
        self._edit_canvas_ids = []  # canvas items were just deleted
        self._word_items_line1 = []
        self._word_items_line2 = []
        self._score_items = []
        sz = self.WIDGET_SIZE
        enabled = self.tray.enabled
        edit_open = self._editor_height > 0

        if self._panel_visible:
            pw = self.PANEL_WIDTH
            total_w = pw + sz
            if edit_open:
                # ONE seamless shape: pill top flowing into rounded-bottom editor
                self._draw_unified_shape(total_w, sz + self._editor_height)
            else:
                # Normal capsule pill shape
                self._create_pill_shape(1, 1, total_w - 1, sz - 1,
                                        fill=COLORS['bg_medium'],
                                        outline=COLORS['border'], width=1)

            # Subtle vertical separator between text and icon
            sep_x = total_w - sz
            self._canvas.create_line(sep_x, 10, sep_x, sz - 10,
                                     fill=COLORS['border'], width=1)

            # Draw per-word text items (created by _redraw_text_words)
            self._redraw_text_words()
        else:
            self._word_items_line1 = []
            self._word_items_line2 = []
            self._score_items = []

        # ── 4-quadrant icon with center circle ──
        icon_x = (self.PANEL_WIDTH if self._panel_visible else 0)
        cx = icon_x + sz // 2
        cy = sz // 2
        r = sz // 2 - 4          # outer radius
        bbox = (cx - r, cy - r, cx + r, cy + r)

        gap = 4                   # degrees gap between quadrants
        extent = 90 - gap         # 86° per quadrant

        q_toggle = COLORS['accent'] if enabled else COLORS['error']
        q_base = '#1a5276'

        # (center_angle, tag, fill, label, dx, dy)
        quadrants = [
            (90,  'q_top',    q_toggle, '⏻', 0, -1),    # Top — Toggle
            (180, 'q_left',   q_base,   '↩', -1, 0),    # Left — Undo
            (270, 'q_bottom', q_base,   '✎', 0,  1),    # Bottom — Translate
            (0,   'q_right',  q_base,   '⚙', 1,  0),    # Right — Settings
        ]

        r_center = 13
        r_text = (r_center + r) // 2 + 1  # label centered in the band

        for angle, tag, color, label, dx, dy in quadrants:
            self._canvas.create_arc(
                *bbox, start=angle - extent / 2, extent=extent,
                style='pieslice', fill=color,
                outline=COLORS['bg_dark'], width=2, tags=(tag,))
            self._canvas.create_text(
                cx + dx * r_text, cy + dy * r_text,
                text=label, fill='#ffffff',
                font=('Segoe UI Symbol', 10), tags=(tag,))

        # Center circle with red dot (mic button)
        self._canvas.create_oval(
            cx - r_center, cy - r_center, cx + r_center, cy + r_center,
            fill='#0d1f38', outline=COLORS['accent'], width=2,
            tags=('q_center',))
        # Red dot indicator inside center circle
        dot_r = 5
        self._canvas.create_oval(
            cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r,
            fill='#e74c3c', outline='', tags=('q_center', 'center_dot'))
        # If currently recording, switch to red square immediately
        if self._speech_recording:
            self._update_center_recording(True)

        # If editor is open, redraw editor content (canvas.delete('all') cleared it)
        if edit_open:
            self._draw_editor_content()

    def _redraw_text_words(self):
        """Draw individual words on the canvas with corrected words highlighted and clickable.
        Line 1: actual typed text
        Line 2: translated text
        Line 3: NLP scores per word"""
        # Remove old word items
        for item_id in self._word_items_line1 + self._word_items_line2 + self._score_items:
            self._canvas.delete(item_id)
        self._word_items_line1 = []
        self._word_items_line2 = []
        self._score_items = []

        if not self._panel_visible or self._canvas is None:
            return

        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH
        text_right = pw - 10   # right edge of text area
        corrections = self._corrections

        # Split into words
        words1 = self._cur_line1.split() if self._cur_line1.strip() else []
        words2 = self._cur_line2.split() if self._cur_line2.strip() else []

        # 3-line vertical layout: y positions spread across sz height
        y1 = 14          # line 1: actual text
        y2 = 34          # line 2: translated
        y3 = 54          # line 3: scores

        font_normal = ('Segoe UI', 10)
        font_corrected = ('Segoe UI', 10, 'underline')

        # Draw line 1 (actual text) — words right-to-left from text_right
        x = text_right
        for word in reversed(words1):
            is_corrected = word in corrections
            fill = '#fbbf24' if is_corrected else COLORS['text_primary']
            font = font_corrected if is_corrected else font_normal
            tag = f'word_{word}' if is_corrected else ''
            item = self._canvas.create_text(
                x, y1, text=word, fill=fill,
                font=font, anchor='e', tags=(tag,) if tag else ())
            self._word_items_line1.append(item)
            if is_corrected:
                self._canvas.tag_bind(tag, '<Enter>', lambda e: self._canvas.configure(cursor='hand2'))
                self._canvas.tag_bind(tag, '<Leave>', lambda e: self._canvas.configure(cursor=''))
            bbox = self._canvas.bbox(item)
            if bbox:
                word_w = bbox[2] - bbox[0]
                x -= word_w + 6

        # Draw line 2 (translated) — same positions
        x = text_right
        font_normal2 = ('Segoe UI', 9)
        for word in reversed(words2):
            item = self._canvas.create_text(
                x, y2, text=word, fill=COLORS['accent'],
                font=font_normal2, anchor='e')
            self._word_items_line2.append(item)
            bbox = self._canvas.bbox(item)
            if bbox:
                word_w = bbox[2] - bbox[0]
                x -= word_w + 6

        # Draw line 3 (NLP scores) — compact summary for the last word
        if not self._hide_scores:
            font_score = ('Segoe UI', 8)
            # Build score text from the last (current) word's scores
            # Show all words' scores compactly: "EN:0.0 עב:5.1 | EN:3.2 עב:0.0"
            score_parts = []
            orig_words = self._cur_line1.split() if self._cur_line1.strip() else []
            for w in orig_words:
                ws = self._word_scores.get(w)
                if ws:
                    pairs = ' '.join(f'{k}:{v}' for k, v in ws.items())
                    score_parts.append(pairs)
            if score_parts:
                score_text = ' | '.join(score_parts)
                # Truncate to fit panel
                if len(score_text) > 40:
                    score_text = score_text[-40:]
                item = self._canvas.create_text(
                    text_right, y3, text=score_text,
                    fill='#8899aa', font=font_score, anchor='e')
                self._score_items.append(item)

    def _on_word_click(self, corrected_word: str):
        """Handle click on a corrected word - revert it to original."""
        try:
            import auto_lang
            threading.Thread(
                target=auto_lang._revert_word,
                args=(corrected_word,),
                daemon=True
            ).start()
        except Exception:
            pass

    def _build_menu(self):
        self._menu = tk.Menu(self.root, tearoff=0,
                              bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                              activebackground=COLORS['accent_dim'],
                              activeforeground='#ffffff',
                              font=('Segoe UI', 10), bd=1,
                              relief='flat')
        self._menu.add_command(label='✅  תיקון שפה פעיל', command=self._toggle)
        self._menu.add_separator()
        self._menu.add_command(label='🔍  הצג/הסתר תיבת הקלדה', command=self._toggle_panel)
        self._menu.add_separator()
        self._menu.add_command(label='📊  הצג/הסתר ציונים', command=self._toggle_scores)
        self._menu.add_separator()
        self._menu.add_command(label='⏪  בטל תיקון אחרון  (F12)', command=self._undo)
        self._menu.add_separator()
        self._menu.add_command(label='⚙️  הגדרות', command=self._open_settings)
        self._menu.add_separator()
        self._menu.add_command(label='❓  מדריך למשתמש', command=self._show_help)
        self._menu.add_separator()
        self._menu.add_command(label='🚪  יציאה', command=self._quit)

    def _refresh_menu_label(self):
        """Update the toggle labels to reflect current state."""
        if self.tray.enabled:
            self._menu.entryconfigure(0, label='✅  תיקון שפה פעיל')
        else:
            self._menu.entryconfigure(0, label='❌  תיקון שפה מושבת')
        if self._panel_visible:
            self._menu.entryconfigure(2, label='🔍  הסתר תיבת הקלדה')
        else:
            self._menu.entryconfigure(2, label='🔍  הצג תיבת הקלדה')
        if self._hide_scores:
            self._menu.entryconfigure(4, label='📊  הצג ציונים')
        else:
            self._menu.entryconfigure(4, label='📊  הסתר ציונים')

    # ── panel toggle ──────────────────────────────────────

    def _toggle_panel(self):
        self._panel_visible = not self._panel_visible
        # Persist to config
        self.tray.config['show_typing_panel'] = self._panel_visible
        save_config(self.tray.config)
        self._apply_panel()

    def _apply_panel(self):
        """Show or hide the text panel and resize the window accordingly."""
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH
        eh = self._editor_height
        cur_x = self.root.winfo_x()
        cur_y = self.root.winfo_y()
        total_h = sz + eh
        if self._panel_visible:
            total_w = pw + sz
            self._canvas.configure(width=total_w, height=total_h)
            self.root.geometry(f'{total_w}x{total_h}+{cur_x}+{cur_y}')
        else:
            # Shift x so the icon stays in the same position
            new_x = cur_x + pw
            self._canvas.configure(width=sz, height=total_h)
            self.root.geometry(f'{sz}x{total_h}+{new_x}+{cur_y}')
        self._draw_widget()
        # Redraw editor content if it's open
        if eh > 0:
            self._draw_editor_content()

    # ── state ────────────────────────────────────────────

    def _apply_state(self, enabled: bool):
        if self.root is None:
            return
        if enabled:
            self.root.deiconify()
            self._visible = True
            # Update toggle quadrant to accent color
            for item_id in self._canvas.find_withtag('q_top'):
                if self._canvas.type(item_id) == 'arc':
                    self._canvas.itemconfigure(item_id, fill=COLORS['accent'])
        else:
            # Close text editor on disable
            if self._edit_mode:
                self._edit_mode = False
                self._hide_edit_window()
            self.root.withdraw()
            self._visible = False

    # ── drag logic ───────────────────────────────────────

    def _on_click_start(self, event):
        self._drag_data['x'] = event.x
        self._drag_data['y'] = event.y
        self._drag_data['moved'] = False

    def _on_drag(self, event):
        dx = event.x - self._drag_data['x']
        dy = event.y - self._drag_data['y']
        x = self.root.winfo_x() + dx
        y = self.root.winfo_y() + dy
        self.root.geometry(f'+{x}+{y}')
        if abs(dx) > 3 or abs(dy) > 3:
            self._drag_data['moved'] = True

    def _on_click_release(self, event):
        if self._drag_data.get('moved', False):
            return  # Was a drag, not a click

        # Click on the dark panel area → toggle text editor
        if self._panel_visible and event.x < self.PANEL_WIDTH:
            self._toggle_edit_mode()
            return

        # Check clicked canvas items
        items = self._canvas.find_overlapping(
            event.x - 2, event.y - 2, event.x + 2, event.y + 2)
        for item in reversed(items):  # topmost first
            tags = self._canvas.gettags(item)
            # Corrected word clicks
            for tag in tags:
                if tag.startswith('word_'):
                    corrected = tag[5:]
                    self._on_word_click(corrected)
                    return
            # Icon quadrant clicks
            if 'q_center' in tags:
                self._toggle_speech()
                return
            if 'q_top' in tags:
                self._toggle()
                return
            if 'q_left' in tags:
                self._undo()
                return
            if 'q_right' in tags:
                self._open_settings()
                return
            if 'q_bottom' in tags:
                self._toggle_edit_mode()
                return

        # Fallback
        self._show_menu(event)

    def _show_tooltip(self, tag):
        """Show a tooltip for the given quadrant tag after a delay."""
        self._hide_tooltip()
        text = self._tooltip_texts.get(tag, '')
        if not text:
            return
        def _do_show():
            self._tooltip_after_id = None
            if self._tooltip_win:
                return
            tw = tk.Toplevel(self.root)
            tw.overrideredirect(True)
            tw.attributes('-topmost', True)
            tw.attributes('-alpha', 0.93)
            tw.configure(bg=COLORS['border'])
            frame = tk.Frame(tw, bg=COLORS['bg_dark'], bd=0)
            frame.pack(padx=1, pady=1)
            lbl = tk.Label(frame, text=text, bg=COLORS['bg_dark'],
                           fg=COLORS['accent'], font=('Segoe UI', 9),
                           padx=8, pady=4)
            lbl.pack()
            # Position below the widget
            x = self.root.winfo_rootx() + self.root.winfo_width() // 2
            y = self.root.winfo_rooty() + self.WIDGET_SIZE + 4
            tw.geometry(f'+{x - 60}+{y}')
            self._tooltip_win = tw
        self._tooltip_after_id = self.root.after(800, _do_show)

    def _hide_tooltip(self):
        """Hide any visible tooltip and cancel pending show."""
        if self._tooltip_after_id:
            try:
                self.root.after_cancel(self._tooltip_after_id)
            except Exception:
                pass
            self._tooltip_after_id = None
        if self._tooltip_win:
            try:
                self._tooltip_win.destroy()
            except Exception:
                pass
            self._tooltip_win = None

    def _q_hover(self, tag, entering):
        """Handle hover effect on an icon quadrant."""
        self._canvas.configure(cursor='hand2' if entering else '')
        if entering:
            self._show_tooltip(tag)
        else:
            self._hide_tooltip()
        if tag == 'q_top':
            base = COLORS['accent'] if self.tray.enabled else COLORS['error']
            hover = '#6dd5f5' if self.tray.enabled else '#f9a3a3'
        elif tag == 'q_center':
            if self._speech_recording:
                base, hover = '#c0392b', '#e74c3c'
            else:
                base, hover = '#0d1f38', '#162d4a'
        else:
            base, hover = '#1a5276', '#2980b9'
        color = hover if entering else base
        for item_id in self._canvas.find_withtag(tag):
            itype = self._canvas.type(item_id)
            if itype in ('arc', 'oval'):
                # Don't recolor the red dot / red square inside the center circle
                tags = self._canvas.gettags(item_id)
                if 'center_dot' in tags or 'center_sq' in tags:
                    continue
                self._canvas.itemconfigure(item_id, fill=color)

    # ── text editor ───────────────────────────────────────────────

    def _toggle_edit_mode(self):
        """Toggle the text editor popup below the widget."""
        self._edit_mode = not self._edit_mode
        if self._edit_mode:
            self._show_edit_window()
        else:
            self._hide_edit_window()

    def _show_edit_window(self):
        """Show the translation editor by expanding the main canvas downward."""
        if self._editor_height > 0:
            return

        # Remember which window had focus BEFORE we open the editor
        try:
            import auto_lang
            self._send_target_hwnd = auto_lang._last_target_hwnd
        except Exception:
            self._send_target_hwnd = None

        # Start loading translation models in background
        try:
            import translator
            translator.ensure_models_loaded(callback=self._on_models_ready)
        except Exception as e:
            print(f'[UI] translator import failed: {e}')
            import traceback; traceback.print_exc()

        # ── Expand the main window downward ──
        sz = self.WIDGET_SIZE
        eh = 250                            # editor area height (bigger for readability)
        self._editor_height = eh
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = (pw + sz) if self._panel_visible else sz
        total_h = sz + eh
        cur_x = self.root.winfo_x()
        cur_y = self.root.winfo_y()

        # If expanding would push below screen → move window up
        screen_h = self.root.winfo_screenheight()
        if cur_y + total_h > screen_h - 40:
            cur_y = max(0, screen_h - 40 - total_h)

        self._canvas.configure(width=total_w, height=total_h)
        self.root.geometry(f'{total_w}x{total_h}+{cur_x}+{cur_y}')

        # Redraw the pill shape without the bottom border (flows into editor)
        self._draw_widget()

        # Force OS-level focus to this window and then to the text widget
        self.root.lift()
        self.root.focus_force()
        if self._text_widget:
            self._text_widget.focus_set()

        # Tell the engine to ignore this window
        try:
            import auto_lang
            auto_lang._widget_edit_hwnd = self.root.winfo_id()
        except Exception:
            pass

    def _draw_editor_content(self):
        """Draw the editor area below the pill on the main canvas."""
        # Clean up previous editor items
        for item_id in self._edit_canvas_ids:
            try:
                self._canvas.delete(item_id)
            except Exception:
                pass
        self._edit_canvas_ids = []
        # Destroy old embedded widgets if any
        if self._text_widget:
            self._text_widget.destroy()
            self._text_widget = None
        if self._trans_widget:
            self._trans_widget.destroy()
            self._trans_widget = None
        if self._status_label:
            self._status_label.destroy()
            self._status_label = None
        if self._send_btn:
            self._send_btn.destroy()
            self._send_btn = None

        c = self._canvas
        sz = self.WIDGET_SIZE
        eh = self._editor_height
        if eh <= 0:
            return

        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = (pw + sz) if self._panel_visible else sz
        total_h = sz + eh

        # Background/border is already drawn by _draw_unified_shape
        # (or _draw_widget). Here we only place the editor widgets.

        # ── Content layout ──
        pad = 12
        cw = total_w - pad * 2
        top = sz + 6  # start just below the pill

        # Dynamic layout: allocate space for text areas based on editor height
        bar_h = 24         # bottom bar height
        sep_h = 8          # separator height
        grip_h = 14        # resize grip height
        available = eh - (top - sz) - bar_h - sep_h - grip_h - 6
        input_h = max(36, available // 2)
        output_h = max(36, available - input_h)

        # Input area
        self._text_widget = tk.Text(
            c, bg=COLORS['bg_dark'], fg='#ffffff',
            insertbackground=COLORS['accent'],
            selectbackground=COLORS['accent_dim'],
            selectforeground='#ffffff',
            font=('Segoe UI', 11), relief='flat', wrap='word',
            padx=6, pady=4, highlightthickness=0, bd=0,
            undo=True,
        )
        self._edit_canvas_ids.append(
            c.create_window(pad, top, anchor='nw',
                            window=self._text_widget, width=cw, height=input_h))

        # Separator
        sep_y = top + input_h + 3
        self._edit_canvas_ids.append(
            c.create_line(pad + 6, sep_y, total_w - pad - 6, sep_y,
                          fill=COLORS['accent'], width=1, dash=(3, 3)))

        # Translation output
        trans_y = sep_y + 5
        self._trans_widget = tk.Text(
            c, bg=COLORS['bg_dark'], fg=COLORS['accent'],
            font=('Segoe UI', 11), relief='flat', wrap='word',
            padx=6, pady=4, highlightthickness=0, bd=0,
            state='disabled',
        )
        self._edit_canvas_ids.append(
            c.create_window(pad, trans_y, anchor='nw',
                            window=self._trans_widget, width=cw, height=output_h))

        # Bottom bar
        bar_y = trans_y + output_h + 4
        self._status_label = tk.Label(
            c, text='\u2705 \u05de\u05d5\u05db\u05df \u05dc\u05ea\u05e8\u05d2\u05d5\u05dd', bg=COLORS['bg_medium'],
            fg=COLORS['text_muted'], font=('Segoe UI', 8), anchor='w',
        )
        self._edit_canvas_ids.append(
            c.create_window(pad, bar_y, anchor='nw',
                            window=self._status_label, width=100, height=20))

        self._send_btn = tk.Label(
            c, text='\U0001f4e4 \u05e9\u05dc\u05d7', bg=COLORS['accent_dim'], fg='#ffffff',
            font=('Segoe UI', 9, 'bold'), cursor='hand2', anchor='center',
        )
        self._send_btn.bind('<Button-1>', lambda e: self._send_translation())
        self._send_btn.bind('<Enter>', lambda e: self._send_btn.configure(bg=COLORS['accent']))
        self._send_btn.bind('<Leave>', lambda e: self._send_btn.configure(bg=COLORS['accent_dim']))
        self._edit_canvas_ids.append(
            c.create_window(total_w - pad, bar_y, anchor='ne',
                            window=self._send_btn, width=58, height=20))

        # ── Resize grip at the very bottom ──
        grip_y = total_h - grip_h
        self._edit_canvas_ids.append(
            c.create_text(total_w // 2, grip_y + grip_h // 2,
                          text='⋯', fill=COLORS['text_muted'],
                          font=('Segoe UI', 9), tags=('resize_grip',)))
        c.tag_bind('resize_grip', '<Enter>',
                   lambda e: c.configure(cursor='sb_v_double_arrow'))
        c.tag_bind('resize_grip', '<Leave>',
                   lambda e: c.configure(cursor=''))
        c.tag_bind('resize_grip', '<Button-1>', self._resize_start)
        c.tag_bind('resize_grip', '<B1-Motion>', self._resize_drag)
        c.tag_bind('resize_grip', '<ButtonRelease-1>', self._resize_end)

        self._auto_translate_after_id = None

        # Bindings
        self._text_widget.bind('<Escape>', lambda e: self._toggle_edit_mode())
        self._text_widget.bind('<Return>', self._on_enter_translate)
        self._text_widget.bind('<KeyRelease>', self._on_edit_key_release)
        self._text_widget.focus_set()

    def _resize_start(self, event):
        """Start resizing the editor area."""
        self._resize_drag_y = event.y

    def _resize_drag(self, event):
        """Handle drag to resize the editor area."""
        if self._resize_drag_y is None or self._editor_height <= 0:
            return
        dy = event.y - self._resize_drag_y
        self._resize_drag_y = event.y
        new_eh = max(self._editor_min_h, min(self._editor_max_h, self._editor_height + dy))
        if new_eh == self._editor_height:
            return
        self._editor_height = new_eh
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = (pw + sz) if self._panel_visible else sz
        total_h = sz + new_eh
        cur_x = self.root.winfo_x()
        cur_y = self.root.winfo_y()
        self._canvas.configure(width=total_w, height=total_h)
        self.root.geometry(f'{total_w}x{total_h}+{cur_x}+{cur_y}')
        self._draw_widget()

    def _resize_end(self, event):
        """Finish resizing."""
        self._resize_drag_y = None
        self._canvas.configure(cursor='')

    def _on_models_ready(self):
        """Called (from background thread) when translation models are loaded."""
        if self._status_label and self._editor_height > 0:
            try:
                self.root.after_idle(self._update_ready_status)
            except Exception:
                pass

    def _update_ready_status(self):
        """Update status label to show translation ready (runs on Tk thread)."""
        if self._status_label:
            try:
                self._status_label.configure(text='✅ מוכן לתרגום')
            except Exception:
                pass

    # ── Right-click translate popup ─────────────────────

    def _on_right_click_global(self):
        """Called from mouse library thread on right-click.
        Always show translate popup near cursor."""
        try:
            from auto_lang import _dbg
        except Exception:
            _dbg = lambda m: None
        _dbg('[UI-RC] right-click detected')
        try:
            import mouse as _mouse
            x, y = _mouse.get_position()
            _dbg(f'[UI-RC] mouse pos: {x},{y}')
        except Exception as e:
            _dbg(f'[UI-RC] get_position failed: {e}')
            return

        # Save target HWND while original app still has focus
        try:
            import auto_lang
            self._send_target_hwnd = auto_lang._last_target_hwnd
            _dbg(f'[UI-RC] target_hwnd={self._send_target_hwnd}')
        except Exception:
            self._send_target_hwnd = None

        try:
            self.root.after_idle(self._show_translate_popup, x, y)
            _dbg('[UI-RC] after_idle scheduled')
        except Exception as e:
            _dbg(f'[UI-RC] after_idle FAILED: {e}')

    def _show_translate_popup(self, x, y):
        """Show a small 'Translate' button near the cursor."""
        try:
            from auto_lang import _dbg
        except Exception:
            _dbg = lambda m: None
        _dbg(f'[UI-RC] _show_translate_popup called x={x} y={y}')
        self._hide_translate_popup()

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.attributes('-alpha', 0.95)
        popup.configure(bg=COLORS['border'])

        frame = tk.Frame(popup, bg=COLORS['bg_medium'], bd=0)
        frame.pack(padx=1, pady=1)

        btn = tk.Label(frame, text='\U0001F50D  תרגם', font=('Segoe UI', 10, 'bold'),
                       bg=COLORS['bg_medium'], fg=COLORS['accent'],
                       padx=10, pady=5, cursor='hand2')
        btn.pack()
        btn.bind('<Button-1>', lambda e: self._on_popup_translate_click())
        btn.bind('<Enter>', lambda e: btn.configure(bg=COLORS['bg_light']))
        btn.bind('<Leave>', lambda e: btn.configure(bg=COLORS['bg_medium']))

        # Position above and to the right of cursor
        popup.geometry(f'+{x + 12}+{y - 42}')

        self._translate_popup = popup
        _dbg(f'[UI-RC] popup created at +{x + 12}+{y - 42}')
        # Auto-hide after 4 seconds
        self._translate_popup_after = self.root.after(4000, self._hide_translate_popup)

    def _hide_translate_popup(self):
        """Destroy the translate popup if it exists."""
        if self._translate_popup_after:
            try:
                self.root.after_cancel(self._translate_popup_after)
            except Exception:
                pass
            self._translate_popup_after = None
        if self._translate_popup:
            try:
                self._translate_popup.destroy()
            except Exception:
                pass
            self._translate_popup = None

    def _on_popup_translate_click(self):
        """User clicked 'Translate' — refocus target, copy, translate."""
        self._hide_translate_popup()
        target_hwnd = self._send_target_hwnd

        def _do_copy():
            import time as _time
            import ctypes
            user32 = ctypes.windll.user32

            # Re-focus original window
            if target_hwnd:
                try:
                    user32.SetForegroundWindow(target_hwnd)
                except Exception:
                    pass
                _time.sleep(0.2)

            # Clear clipboard
            try:
                if user32.OpenClipboard(0):
                    user32.EmptyClipboard()
                    user32.CloseClipboard()
            except Exception:
                pass
            _time.sleep(0.05)

            # Send Ctrl+C using keybd_event (simplest, most reliable)
            VK_CONTROL = 0x11
            VK_C = 0x43
            KEYEVENTF_KEYUP = 0x0002
            user32.keybd_event(VK_CONTROL, 0, 0, 0)
            user32.keybd_event(VK_C, 0, 0, 0)
            user32.keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
            _time.sleep(0.3)

            # Read clipboard
            text = ''
            CF_UNICODETEXT = 13
            try:
                if user32.OpenClipboard(0):
                    handle = user32.GetClipboardData(CF_UNICODETEXT)
                    if handle:
                        kernel32 = ctypes.windll.kernel32
                        kernel32.GlobalLock.restype = ctypes.c_void_p
                        ptr = kernel32.GlobalLock(handle)
                        if ptr:
                            text = ctypes.wstring_at(ptr)
                            kernel32.GlobalUnlock(handle)
                    user32.CloseClipboard()
            except Exception:
                pass

            if text and text.strip():
                self._pending_selection = text.strip()
                try:
                    self.root.after_idle(self._fill_and_translate)
                except Exception:
                    pass

        threading.Thread(target=_do_copy, daemon=True).start()

    def _fill_and_translate(self):
        """Open editor, fill with pending selection, translate."""
        selected = self._pending_selection
        self._pending_selection = ''
        if not selected:
            return

        # Open editor if not already open
        if self._editor_height == 0:
            self._show_edit_window()
            self.root.update_idletasks()

        # Fill the input field
        if self._text_widget:
            self._text_widget.delete('1.0', 'end')
            self._text_widget.insert('1.0', selected)

        # Auto-translate
        self.root.after(100, self._do_translate)

    def _translate_selection(self):
        """Copy selected text from any app, open editor, fill & translate."""
        # Remember target window before we steal focus
        try:
            import auto_lang
            self._send_target_hwnd = auto_lang._last_target_hwnd
        except Exception:
            self._send_target_hwnd = None

        # Clear clipboard so we can detect fresh content
        try:
            self.root.clipboard_clear()
        except Exception:
            pass

        # Send Ctrl+C to copy the selected text (runs in background to not block Tk)
        def _do_copy_and_translate():
            import time as _time
            try:
                import keyboard as _kb
                _kb.send('ctrl+c')
            except Exception:
                return

            # Wait for clipboard to update
            _time.sleep(0.25)

            # Schedule reading clipboard on Tk thread
            try:
                self.root.after_idle(self._read_clipboard_and_translate)
            except Exception:
                pass

        threading.Thread(target=_do_copy_and_translate, daemon=True).start()

    def _read_clipboard_and_translate(self):
        """Read clipboard after Ctrl+C and translate the content."""
        try:
            selected = self.root.clipboard_get()
        except Exception:
            selected = ''

        if not selected or not selected.strip():
            return

        selected = selected.strip()

        # Open editor if not already open
        if self._editor_height == 0:
            self._show_edit_window()
            # Give the editor a moment to render
            self.root.update_idletasks()

        # Fill the input field with the selected text
        if self._text_widget:
            self._text_widget.delete('1.0', 'end')
            self._text_widget.insert('1.0', selected)

        # Auto-translate immediately
        self.root.after(100, self._do_translate)

    def _on_edit_key_release(self, event):
        """Auto-translate after each word boundary (space / Enter / punctuation)."""
        if event.keysym in ('space', 'Return', 'period', 'comma', 'semicolon',
                            'colon', 'exclam', 'question'):
            # Cancel any pending debounce
            if self._auto_translate_after_id is not None:
                self.root.after_cancel(self._auto_translate_after_id)
            # Debounce: translate 150ms after last boundary key
            self._auto_translate_after_id = self.root.after(150, self._do_translate)

    def _on_enter_translate(self, event):
        """Enter key: translate (Shift+Enter for newline)."""
        try:
            import keyboard as _kb
            if _kb.is_pressed('shift'):
                return  # Allow Shift+Enter for newline
        except Exception:
            pass
        self._do_translate()
        return 'break'  # Prevent newline insertion

    def _do_translate(self):
        """Translate the input text and show in the output area."""
        if self._text_widget is None:
            return
        text = self._text_widget.get('1.0', 'end-1c').strip()
        if not text:
            return

        # Update status
        if self._status_label:
            self._status_label.configure(text='⏳ מתרגם...')

        # Run translation in background thread
        threading.Thread(target=self._translate_worker, args=(text,), daemon=True).start()

    def _translate_worker(self, text: str):
        """Background translation worker."""
        try:
            import translator
            result = translator.translate(text)
            if result is None:
                msg = '\u274c \u05ea\u05e8\u05d2\u05d5\u05dd \u05e0\u05db\u05e9\u05dc'
                self.root.after_idle(
                    lambda m=msg: self._show_translation_result(m, error=True))
            else:
                # Show source label (online / offline)
                source = 'Google' if translator._online else 'Argos'
                self.root.after_idle(lambda r=result, s=source: self._show_translation_result(r, source=s))
        except Exception as e:
            self.root.after_idle(
                lambda err=e: self._show_translation_result(f'שגיאה: {err}', error=True))

    def _show_translation_result(self, text: str, error: bool = False, source: str = ''):
        """Display translation result in the output area."""
        if self._trans_widget is None:
            return
        self._trans_widget.configure(state='normal')
        self._trans_widget.delete('1.0', 'end')
        self._trans_widget.insert('1.0', text)
        self._trans_widget.configure(
            state='disabled',
            fg=COLORS['error'] if error else COLORS['accent'],
        )
        if self._status_label:
            if error:
                self._status_label.configure(text='\u274c \u05e9\u05d2\u05d9\u05d0\u05d4')
            elif self._text_widget:
                try:
                    import translator
                    pair = translator.detect_direction(
                        self._text_widget.get('1.0', 'end-1c').strip())
                    direction = 'EN \u2192 \u05e2\u05d1' if pair == 'en-he' else '\u05e2\u05d1 \u2192 EN'
                    src_label = f' ({source})' if source else ''
                    self._status_label.configure(text=f'\u2705 {direction}{src_label}')
                except Exception:
                    self._status_label.configure(text='\u2705 \u05ea\u05d5\u05e8\u05d2\u05dd')

    def _send_translation(self):
        """Send the translated text to the last focused application window."""
        if self._trans_widget is None:
            return
        translated = self._trans_widget.get('1.0', 'end-1c').strip()
        if not translated:
            # No translation yet — translate first, then send
            self._do_translate()
            return

        # Don't send error messages to the target app
        if translated.startswith('❌') or translated.startswith('שגיאה'):
            return

        # Close the editor
        self._edit_mode = False
        self._hide_edit_window()

        # Small delay to let the editor close, then paste in a background thread
        # (avoids blocking the Tk mainloop with time.sleep)
        self.root.after(100, lambda t=translated: threading.Thread(
            target=self._paste_to_target, args=(t,), daemon=True).start())

    def _paste_to_target(self, text: str):
        """Focus the last target window and paste text. Runs in background thread."""
        try:
            import auto_lang
            import ctypes
            send_t = self._send_target_hwnd
            last_t = auto_lang._last_target_hwnd
            hwnd = send_t or last_t
            if not hwnd:
                hwnd = self._find_target_hwnd()
            print(f'[Translator] Pasting to HWND={hwnd}, _send_target={send_t}, _last_target={last_t}')
            if not hwnd:
                print('[Translator] No target HWND!')
                return
            # Use AllowSetForegroundWindow + SetForegroundWindow
            ctypes.windll.user32.AllowSetForegroundWindow(-1)  # ASFW_ANY
            result = ctypes.windll.user32.SetForegroundWindow(hwnd)
            print(f'[Translator] SetForegroundWindow result={result}')
            import time
            time.sleep(0.3)
            auto_lang.injecting.set()
            try:
                auto_lang._paste_text(text)
                print(f'[Translator] Paste done: {text[:50]!r}')
            finally:
                time.sleep(0.05)
                auto_lang.injecting.clear()
        except Exception as e:
            import traceback
            print(f'[Translator] Paste failed: {e}')
            traceback.print_exc()

    def _hide_edit_window(self):
        """Hide the translation editor by shrinking the canvas back."""
        self._edit_mode = False
        # Cancel pending auto-translate
        if self._auto_translate_after_id is not None:
            self.root.after_cancel(self._auto_translate_after_id)
            self._auto_translate_after_id = None
        # Clear engine's edit HWND
        try:
            import auto_lang
            auto_lang._widget_edit_hwnd = 0
        except Exception:
            pass
        # Remove editor canvas items
        for item_id in self._edit_canvas_ids:
            try:
                self._canvas.delete(item_id)
            except Exception:
                pass
        self._edit_canvas_ids = []
        # Destroy embedded widgets
        if self._text_widget:
            self._text_widget.destroy()
            self._text_widget = None
        if self._trans_widget:
            self._trans_widget.destroy()
            self._trans_widget = None
        if self._status_label:
            self._status_label.destroy()
            self._status_label = None
        if self._send_btn:
            self._send_btn.destroy()
            self._send_btn = None
        # Shrink canvas and window back
        self._editor_height = 0
        sz = self.WIDGET_SIZE
        pw = self.PANEL_WIDTH if self._panel_visible else 0
        total_w = (pw + sz) if self._panel_visible else sz
        cur_x = self.root.winfo_x()
        cur_y = self.root.winfo_y()
        self._canvas.configure(width=total_w, height=sz)
        self.root.geometry(f'{total_w}x{sz}+{cur_x}+{cur_y}')
        # Redraw the widget (restores full pill outline with bottom border)
        self._draw_widget()

    def _show_menu(self, event):
        self._refresh_menu_label()
        try:
            abs_x = self.root.winfo_x() + event.x
            abs_y = self.root.winfo_y() + event.y
            self._menu.tk_popup(abs_x, abs_y)
        finally:
            self._menu.grab_release()

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

    def _open_settings(self):
        self.tray._open_settings()

    def _quit(self):
        self.tray._quit()

    def _toggle_scores(self):
        """Toggle NLP score display on/off."""
        self._hide_scores = not self._hide_scores
        # Persist to config
        self.tray.config['hide_scores'] = self._hide_scores
        save_config(self.tray.config)
        self._redraw_text_words()

    def _show_help(self):
        """Show a help/guide window explaining all features."""
        help_win = tk.Toplevel(self.root)
        help_win.title('AutoLang — מדריך למשתמש')
        help_win.geometry('520x620')
        help_win.resizable(True, True)
        help_win.configure(bg=COLORS['bg_dark'])
        help_win.attributes('-topmost', True)

        # Title
        tk.Label(help_win, text='📖  מדריך למשתמש',
                 font=('Segoe UI', 16, 'bold'),
                 bg=COLORS['bg_dark'], fg=COLORS['accent']).pack(pady=(14, 4))
        tk.Frame(help_win, bg=COLORS['accent'], height=2).pack(fill='x', padx=20, pady=(0, 8))

        # Scrollable text
        text_frame = tk.Frame(help_win, bg=COLORS['bg_medium'])
        text_frame.pack(fill='both', expand=True, padx=14, pady=(0, 14))

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side='left', fill='y')

        help_text = tk.Text(text_frame, wrap='word', font=('Segoe UI', 10),
                            bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                            bd=0, highlightthickness=0, padx=12, pady=10,
                            yscrollcommand=scrollbar.set, spacing3=4)
        help_text.pack(fill='both', expand=True)
        scrollbar.config(command=help_text.yview)

        # Define tags for styling
        help_text.tag_configure('title', font=('Segoe UI', 12, 'bold'),
                                foreground=COLORS['accent'], spacing1=10)
        help_text.tag_configure('body', font=('Segoe UI', 10),
                                foreground=COLORS['text_primary'])
        help_text.tag_configure('key', font=('Segoe UI', 10, 'bold'),
                                foreground='#fbbf24')

        sections = [
            ('מה זה AutoLang?',
             'AutoLang היא תוכנה שמתקנת אוטומטית טקסט שהוקלד בשפה הלא נכונה.\n'
             'למשל — אם המקלדת שלך על אנגלית ואתה מקליד עברית, התוכנה מזהה את הטעות ומתקנת אותה בזמן אמת.\n'),
            ('🔵 כפתורי העיגול',
             '⏻  למעלה — הפעלה / השבתה של תיקון השפה\n'
             '↩  שמאל — ביטול התיקון האחרון (גם F12)\n'
             '✎  למטה — פתיחת עורך תרגום\n'
             '⚙  ימין — פתיחת חלון הגדרות\n'
             '🔴  מרכז — הקלטה קולית (Speech-to-Text)\n'),
            ('📝 חלונית ההקלדה',
             'מציגה את המילים שהוקלדו + התרגום שלהן בזמן אמת.\n'
             'שורה 1 — הטקסט שהוקלד\n'
             'שורה 2 — התרגום לשפה השנייה\n'
             'שורה 3 — ציוני NLP (ניתן להסתיר דרך התפריט)\n'
             'מילים שתוקנו מופיעות בצהוב עם קו תחתון — לחיצה עליהן מבטלת את התיקון.\n'
             'כדי להציג/להסתיר את החלונית: קליק ימני → הצג/הסתר תיבת הקלדה.\n'),
            ('🔍 תרגום מהיר',
             'סמן טקסט בכל חלון ולחץ קליק ימני — יופיע כפתור "תרגם".\n'
             'לחיצה עליו תתרגם את הטקסט המסומן ותדביק אותו בחלון הפעיל.\n'),
            ('🎤 הקלטה קולית',
             'לחיצה על הכפתור האדום במרכז העיגול מפעילה הקלטה.\n'
             'התוכנה מקשיבה לדיבור, מזהה טקסט (Whisper) ומדביקה אותו בחלון הפעיל.\n'
             'לחיצה נוספת עוצרת את ההקלטה.\n'),
            ('⚙️ הגדרות',
             'אפליקציות — הגדרת שפת ברירת מחדל לכל תוכנה (לפי שם ה-exe).\n'
             'צ\'אטים — שפה שונה לפי שם שיחה (WhatsApp, Teams).\n'
             'דפדפנים — שפה לפי מילת מפתח בכותרת הטאב.\n'
             'מילים מוחרגות — מילים שלא יתוקנו (שמות, מונחים טכניים).\n'
             'כללי — הפעלה/השבתה, מצב Debug, החלפת שפה אוטומטית.\n'),
            ('🧠 איך הזיהוי עובד?',
             'כל מילה שמוקלדת נבדקת מול מאגר מילים (NLP) בעברית ובאנגלית.\n'
             'המערכת משווה ציונים (zipf) ובוחרת את השפה שהמילה שייכת אליה.\n'
             'אחרי 2 מילים ברצף, המנוע "נועל" את שפת המשפט ומתקן מילים שוטפות ללא בדיקה נוספת\n'
             '(אלא אם המנוע מזהה שהתיקון לא הגיוני — אז הוא מדלג).\n'
             'ENTER או נקודה מאפסים את הנעילה ומתחילים משפט חדש.\n'),
            ('⌨️ קיצורי מקשים',
             'F12 — ביטול תיקון אחרון\n'
             'Ctrl+Alt+Q — יציאה מהתוכנה\n'
             'Ctrl+Alt+I — הצגת מידע (Debug)\n'),
        ]

        for title, body in sections:
            help_text.insert('end', title + '\n', 'title')
            help_text.insert('end', body + '\n', 'body')

        help_text.configure(state='disabled')

        # Close button
        ttk.Button(help_win, text='סגור', command=help_win.destroy).pack(pady=(0, 10))

    # ── speech-to-text ─────────────────────────────────────

    def _toggle_speech(self):
        """Toggle speech recording from the center button."""
        if self._speech_recording:
            self._stop_speech()
        else:
            self._start_speech()

    def _start_speech(self):
        """Start speech-to-text recording."""
        if self._speech_recording:
            return
        self._speech_recording = True

        # Remember target window BEFORE we start
        try:
            import auto_lang
            hwnd = auto_lang._last_target_hwnd
            if not hwnd:
                # Fallback: find last active window via Win32
                hwnd = self._find_target_hwnd()
            self._send_target_hwnd = hwnd
        except Exception:
            pass

        # Update center button visual → red recording indicator
        self._update_center_recording(True)
        self._start_pulse_animation()

        try:
            import speech_module
            speech_module.start_recording(
                callback=self._on_speech_text,
                on_error=self._on_speech_error,
                on_state=self._on_speech_state,
            )
        except Exception as e:
            print(f'[Speech] Failed to start: {e}')
            self._speech_recording = False
            self._update_center_recording(False)

    def _stop_speech(self):
        """Stop speech-to-text recording."""
        self._speech_recording = False
        self._stop_pulse_animation()
        self._update_center_recording(False)
        try:
            import speech_module
            speech_module.stop_recording()
        except Exception:
            pass

    def _update_center_recording(self, recording: bool):
        """Update center button appearance for recording state."""
        if self._canvas is None:
            return
        # Delete old center_dot / center_sq
        self._canvas.delete('center_dot')
        self._canvas.delete('center_sq')

        # Find the center circle to get coordinates
        icon_x = (self.PANEL_WIDTH if self._panel_visible else 0)
        sz = self.WIDGET_SIZE
        cx = icon_x + sz // 2
        cy = sz // 2

        # Update outer circle colors
        for item_id in self._canvas.find_withtag('q_center'):
            itype = self._canvas.type(item_id)
            if itype == 'oval':
                self._canvas.itemconfigure(
                    item_id,
                    fill='#c0392b' if recording else '#0d1f38',
                    outline='#e74c3c' if recording else COLORS['accent'],
                )

        if recording:
            # Draw red square in center
            sq = 5
            self._canvas.create_rectangle(
                cx - sq, cy - sq, cx + sq, cy + sq,
                fill='#e74c3c', outline='', tags=('q_center', 'center_sq'))
        else:
            # Draw red dot in center
            dot_r = 5
            self._canvas.create_oval(
                cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r,
                fill='#e74c3c', outline='', tags=('q_center', 'center_dot'))

    def _start_pulse_animation(self):
        """Animate blinking red square in center while recording."""
        self._pulse_step = 0

        def _pulse():
            if not self._speech_recording or self._canvas is None:
                return
            visible = (self._pulse_step % 2 == 0)
            # Toggle square visibility by changing fill
            for item_id in self._canvas.find_withtag('center_sq'):
                self._canvas.itemconfigure(
                    item_id, fill='#e74c3c' if visible else '#c0392b')
            # Also pulse the outer circle
            outer_colors = ['#c0392b', '#a93226']
            oc = outer_colors[self._pulse_step % 2]
            for item_id in self._canvas.find_withtag('q_center'):
                if self._canvas.type(item_id) == 'oval' and 'center_dot' not in self._canvas.gettags(item_id):
                    self._canvas.itemconfigure(item_id, fill=oc)
            self._pulse_step += 1
            self._speech_pulse_id = self.root.after(500, _pulse)

        self._speech_pulse_id = self.root.after(500, _pulse)

    def _stop_pulse_animation(self):
        """Stop the pulse animation."""
        if self._speech_pulse_id is not None:
            try:
                self.root.after_cancel(self._speech_pulse_id)
            except Exception:
                pass
            self._speech_pulse_id = None

    def _on_speech_text(self, text: str, lang_code: str):
        """Called from speech thread when text is recognized."""
        if not text:
            return

        # Dedup: skip identical text within 3 seconds
        import time as _time
        now = _time.time()
        if text == self._last_speech_text and now - self._last_speech_time < 3.0:
            try:
                import os
                with open(os.path.expanduser('~/speech_debug.log'), 'a', encoding='utf-8') as f:
                    f.write(f'[UI] DEDUP skipped: {text!r}\n')
            except Exception:
                pass
            return
        self._last_speech_text = text
        self._last_speech_time = now
        # File-based debug log (print goes to devnull in --noconsole EXE)
        try:
            import os
            with open(os.path.expanduser('~/speech_debug.log'), 'a', encoding='utf-8') as f:
                f.write(f'[UI] _on_speech_text called: [{lang_code}] {text!r}\n')
                f.write(f'[UI] edit_mode={self._edit_mode}, text_widget={self._text_widget is not None}\n')
        except Exception:
            pass

        # Schedule on Tk thread
        self.root.after_idle(lambda t=text, lc=lang_code: self._handle_speech_result(t, lc))

    def _handle_speech_result(self, text: str, lang_code: str):
        """Handle recognized speech - runs on Tk thread."""
        try:
            import os
            with open(os.path.expanduser('~/speech_debug.log'), 'a', encoding='utf-8') as f:
                f.write(f'[UI] _handle_speech_result: edit_mode={self._edit_mode}\n')
        except Exception:
            pass

        # If the translation editor is open, insert text there
        if self._edit_mode and self._text_widget:
            self._insert_speech_to_editor(text)
        else:
            # Paste directly into the target window (in background thread)
            threading.Thread(
                target=self._paste_speech_to_target,
                args=(text, lang_code),
                daemon=True,
            ).start()

    def _insert_speech_to_editor(self, text: str):
        """Insert recognized speech text into the translation editor."""
        if self._text_widget is None:
            return
        # Insert at cursor with a space
        current = self._text_widget.get('1.0', 'end-1c')
        if current and not current.endswith(' ') and not current.endswith('\n'):
            self._text_widget.insert('end', ' ')
        self._text_widget.insert('end', text)
        # Trigger auto-translate
        if self._auto_translate_after_id is not None:
            self.root.after_cancel(self._auto_translate_after_id)
        self._auto_translate_after_id = self.root.after(300, self._do_translate)

    def _paste_speech_to_target(self, text: str, lang_code: str):
        """Paste recognized speech text directly to the target app."""
        log_path = None
        try:
            import os
            log_path = os.path.expanduser('~/speech_debug.log')
        except Exception:
            pass

        def _slog(msg):
            if log_path:
                try:
                    with open(log_path, 'a', encoding='utf-8') as f:
                        f.write(f'[Paste] {msg}\n')
                except Exception:
                    pass

        try:
            import auto_lang
            import ctypes
            hwnd = self._send_target_hwnd or auto_lang._last_target_hwnd
            if not hwnd:
                hwnd = self._find_target_hwnd()
            _slog(f'hwnd={hwnd}, send_target={self._send_target_hwnd}, last_target={auto_lang._last_target_hwnd}')
            if not hwnd:
                _slog('No target HWND!')
                return
            ctypes.windll.user32.AllowSetForegroundWindow(-1)
            res = ctypes.windll.user32.SetForegroundWindow(hwnd)
            _slog(f'SetForegroundWindow result={res}')
            import time
            time.sleep(0.2)
            auto_lang.injecting.set()
            try:
                auto_lang._paste_text(text)
                _slog(f'Pasted OK [{lang_code}]: {text[:60]!r}')
            finally:
                time.sleep(0.05)
                auto_lang.injecting.clear()
        except Exception as e:
            _slog(f'FAILED: {e}')
            import traceback
            _slog(traceback.format_exc())

    def _on_speech_error(self, msg: str):
        """Called from speech thread on error."""
        print(f'[Speech] Error: {msg}')
        # Stop recording on error
        self.root.after_idle(self._stop_speech)

    def _on_speech_state(self, state: str):
        """Called from speech thread on state changes."""
        # Could update UI status if needed
        pass

    def _find_target_hwnd(self):
        """Find the last active non-own window using Win32 EnumWindows."""
        try:
            import ctypes
            from ctypes import wintypes
            import os

            _user32 = ctypes.windll.user32
            _kernel32 = ctypes.windll.kernel32
            own_pid = os.getpid()

            WNDENUMPROC = ctypes.WINFUNCTYPE(
                wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

            result = [None]

            def _enum_cb(hwnd, _):
                # Skip invisible/minimized windows
                if not _user32.IsWindowVisible(hwnd):
                    return True
                # Get window PID
                pid = wintypes.DWORD(0)
                _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value == own_pid:
                    return True  # Skip our own windows
                # Check it has a title (real window)
                length = _user32.GetWindowTextLengthW(hwnd)
                if length <= 0:
                    return True
                # Found it — first visible non-own window in Z-order
                result[0] = hwnd
                return False  # Stop enumerating

            cb = WNDENUMPROC(_enum_cb)
            _user32.EnumWindows(cb, 0)
            return result[0]
        except Exception:
            return None


# ──────────────────────────────────────────────────────────
# Dark Theme Color Palette
# ──────────────────────────────────────────────────────────
COLORS = {
    'bg_dark':       '#0a1628',     # Main background — deep navy
    'bg_medium':     '#0f2137',     # Card / panel background
    'bg_light':      '#162d4a',     # Input / entry background
    'bg_highlight':  '#1e3d5f',     # Hover / selection background
    'accent':        '#4ec9f0',     # Primary accent — light cyan
    'accent_hover':  '#7dd8f5',     # Accent hover — brighter cyan
    'accent_dim':    '#1a6a8f',     # Accent muted
    'text_primary':  '#e0e8f0',     # Main text
    'text_secondary':'#7a8fa5',     # Secondary text
    'text_muted':    '#4a6580',     # Muted text
    'border':        '#1a3a5c',     # Borders — subtle blue
    'success':       '#4ade80',     # Success green
    'error':         '#f87171',     # Error red
    'warning':       '#fbbf24',     # Warning yellow
    'tab_active':    '#4ec9f0',     # Active tab accent
    'tab_inactive':  '#162d4a',     # Inactive tab bg
    'btn_primary':   '#4ec9f0',     # Primary button
    'btn_danger':    '#f87171',     # Danger button
    'tree_stripe':   '#0c1f35',     # Treeview alternating rows
    'scrollbar':     '#2a4a6a',     # Scrollbar thumb
}


class SettingsWindow:
    """חלון הגדרות עם טאבים."""

    def __init__(self, config: dict, on_save_callback=None):
        self.config = config
        self.on_save = on_save_callback
        self.root = None

    def show(self):
        if self.root and self.root.winfo_exists():
            self.root.lift()
            self.root.focus_force()
            return

        self.root = tk.Tk()
        self.root.title('AutoLang - הגדרות')
        self.root.geometry('740x600')
        self.root.resizable(True, True)
        self.root.configure(bg=COLORS['bg_dark'])

        # RTL support
        self.root.option_add('*TCombobox*Listbox.justify', 'right')
        self.root.option_add('*TCombobox*Listbox.background', COLORS['bg_light'])
        self.root.option_add('*TCombobox*Listbox.foreground', COLORS['text_primary'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', COLORS['accent_dim'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', '#ffffff')

        # ── Dark Theme Styling ──────────────────────────────
        style = ttk.Style()
        style.theme_use('clam')

        # ── Fix clam theme on Windows ──
        # clam's Treeview.field and Heading elements ignore color settings.
        # Replace them with the 'default' theme's elements which DO respect colors.
        try:
            style.element_create('custom.Treeview.field', 'from', 'default')
            style.layout('Treeview', [
                ('custom.Treeview.field', {'sticky': 'nswe', 'border': 0, 'children': [
                    ('Treeview.padding', {'sticky': 'nswe', 'children': [
                        ('Treeview.treearea', {'sticky': 'nswe'})
                    ]})
                ]})
            ])
        except Exception:
            pass
        try:
            style.element_create('custom.Heading.border', 'from', 'default')
            style.layout('Treeview.Heading', [
                ('custom.Heading.border', {'sticky': 'nswe', 'children': [
                    ('Treeview.Heading.padding', {'sticky': 'nswe', 'children': [
                        ('Treeview.Heading.image', {'side': 'right', 'sticky': ''}),
                        ('Treeview.Heading.text', {'sticky': 'we'})
                    ]})
                ]})
            ])
        except Exception:
            pass

        # Global background
        style.configure('.', background=COLORS['bg_dark'], foreground=COLORS['text_primary'],
                        fieldbackground=COLORS['bg_light'], bordercolor=COLORS['border'],
                        troughcolor=COLORS['bg_medium'], selectbackground=COLORS['accent_dim'],
                        selectforeground='#ffffff', font=('Segoe UI', 10))

        # Frames
        style.configure('TFrame', background=COLORS['bg_dark'])
        style.configure('Card.TFrame', background=COLORS['bg_medium'])

        # Labels
        style.configure('TLabel', background=COLORS['bg_medium'], foreground=COLORS['text_primary'],
                        font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 14, 'bold'),
                        foreground=COLORS['accent'], background=COLORS['bg_medium'])
        style.configure('Status.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['success'], background=COLORS['bg_medium'])
        style.configure('Subtitle.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['text_secondary'], background=COLORS['bg_medium'])
        style.configure('Info.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['text_muted'], background=COLORS['bg_medium'])

        # Notebook (tabs)
        style.configure('TNotebook', background=COLORS['bg_dark'], borderwidth=0,
                        tabmargins=[2, 8, 2, 0],
                        lightcolor=COLORS['bg_dark'], darkcolor=COLORS['bg_dark'],
                        bordercolor=COLORS['bg_dark'])
        style.configure('TNotebook.Tab', background=COLORS['tab_inactive'],
                        foreground=COLORS['text_secondary'], padding=[16, 8],
                        font=('Segoe UI', 11), borderwidth=0,
                        lightcolor=COLORS['tab_inactive'], darkcolor=COLORS['tab_inactive'],
                        bordercolor=COLORS['bg_dark'])
        style.map('TNotebook.Tab',
                  background=[('selected', COLORS['bg_dark']), ('!selected', COLORS['tab_inactive'])],
                  foreground=[('selected', COLORS['accent']), ('!selected', COLORS['text_secondary'])],
                  lightcolor=[('selected', COLORS['bg_dark']), ('!selected', COLORS['tab_inactive'])],
                  darkcolor=[('selected', COLORS['bg_dark']), ('!selected', COLORS['tab_inactive'])],
                  bordercolor=[('selected', COLORS['bg_dark']), ('!selected', COLORS['bg_dark'])],
                  expand=[('selected', [0, 0, 0, 2])])

        # Buttons
        style.configure('TButton', background=COLORS['bg_light'], foreground=COLORS['text_primary'],
                        font=('Segoe UI', 10), padding=[10, 5], borderwidth=1,
                        relief='flat', anchor='center')
        style.map('TButton',
                  background=[('active', COLORS['bg_highlight']), ('pressed', COLORS['accent_dim'])],
                  foreground=[('active', COLORS['accent_hover'])])

        style.configure('Accent.TButton', background=COLORS['accent'], foreground=COLORS['bg_dark'],
                        font=('Segoe UI', 10, 'bold'), padding=[12, 5])
        style.map('Accent.TButton',
                  background=[('active', COLORS['accent_hover']), ('pressed', COLORS['accent_dim'])],
                  foreground=[('active', COLORS['bg_dark'])])

        style.configure('Danger.TButton', background=COLORS['error'], foreground='#ffffff',
                        font=('Segoe UI', 10), padding=[10, 5])
        style.map('Danger.TButton',
                  background=[('active', '#ff7070')],
                  foreground=[('active', '#ffffff')])

        # Treeview — colors now work via element_create from 'default'
        style.configure('Treeview', background=COLORS['bg_medium'], foreground=COLORS['text_primary'],
                        fieldbackground=COLORS['bg_medium'], borderwidth=0,
                        font=('Segoe UI', 10), rowheight=28)
        style.configure('Treeview.Heading', background=COLORS['bg_light'],
                        foreground=COLORS['accent'], font=('Segoe UI', 10, 'bold'),
                        borderwidth=1, relief='flat')
        style.map('Treeview',
                  background=[('selected', COLORS['accent_dim'])],
                  foreground=[('selected', '#ffffff')])
        style.map('Treeview.Heading',
                  background=[('active', COLORS['bg_highlight'])])

        # Entry
        style.configure('TEntry', fieldbackground=COLORS['bg_light'],
                        foreground=COLORS['text_primary'], bordercolor=COLORS['border'],
                        insertcolor=COLORS['accent'], padding=[6, 4])
        style.map('TEntry',
                  bordercolor=[('focus', COLORS['accent'])],
                  fieldbackground=[('focus', COLORS['bg_light'])])

        # Combobox
        style.configure('TCombobox', fieldbackground=COLORS['bg_light'],
                        foreground=COLORS['text_primary'], background=COLORS['bg_light'],
                        bordercolor=COLORS['border'], arrowcolor=COLORS['accent'],
                        padding=[4, 2])
        style.map('TCombobox',
                  fieldbackground=[('readonly', COLORS['bg_light']), ('focus', COLORS['bg_light'])],
                  bordercolor=[('focus', COLORS['accent'])],
                  arrowcolor=[('active', COLORS['accent_hover'])])

        # Checkbutton — ttk indicator ignores colors on Windows,
        # so we use tk.Checkbutton directly (see _build_general_tab).

        # Spinbox
        style.configure('TSpinbox', fieldbackground=COLORS['bg_light'],
                        foreground=COLORS['text_primary'], bordercolor=COLORS['border'],
                        arrowcolor=COLORS['accent'], background=COLORS['bg_medium'])
        style.map('TSpinbox', bordercolor=[('focus', COLORS['accent'])])

        # Scrollbar
        style.configure('TScrollbar', background=COLORS['bg_medium'],
                        troughcolor=COLORS['bg_dark'], borderwidth=0,
                        arrowcolor=COLORS['accent'])
        style.map('TScrollbar',
                  background=[('active', COLORS['scrollbar'])])

        # Separator
        style.configure('TSeparator', background=COLORS['border'])

        # Progressbar
        style.configure('TProgressbar', background=COLORS['accent'],
                        troughcolor=COLORS['bg_medium'], borderwidth=0, thickness=8)

        # ── Title bar area ──────────────────────────────
        title_frame = tk.Frame(self.root, bg=COLORS['bg_dark'], height=56)
        title_frame.pack(fill='x', padx=20, pady=(14, 0))
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text='⌨️', font=('Segoe UI Emoji', 20),
                 bg=COLORS['bg_dark'], fg=COLORS['accent']).pack(side='right', padx=(0, 10))

        title_text_frame = tk.Frame(title_frame, bg=COLORS['bg_dark'])
        title_text_frame.pack(side='right')
        tk.Label(title_text_frame, text='AutoLang', font=('Segoe UI', 20, 'bold'),
                 bg=COLORS['bg_dark'], fg='#ffffff').pack(anchor='e')
        tk.Label(title_text_frame, text='הגדרות מתקדמות', font=('Segoe UI', 9),
                 bg=COLORS['bg_dark'], fg=COLORS['accent']).pack(anchor='e')

        # Accent line under title
        tk.Frame(self.root, bg=COLORS['accent'], height=2).pack(fill='x', padx=20, pady=(2, 8))

        # Notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=20, pady=(5, 10))

        # Tab 1: App defaults
        self._build_app_defaults_tab(notebook)

        # Tab 2: Chat defaults
        self._build_chat_defaults_tab(notebook)

        # Tab 3: Browser/website defaults
        self._build_browser_defaults_tab(notebook)

        # Tab 4: Exclude words
        self._build_exclude_words_tab(notebook)

        # Tab 4: General settings
        self._build_general_tab(notebook)

        # Tab 5: Help guide
        self._build_help_tab(notebook)

        # Bottom buttons
        btn_frame = tk.Frame(self.root, bg=COLORS['bg_dark'])
        btn_frame.pack(fill='x', padx=20, pady=(0, 14))

        ttk.Button(btn_frame, text='💾 שמור', style='Accent.TButton', command=self._save).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='❌ בטל', command=self.root.destroy).pack(side='right', padx=5)

        self.status_var = tk.StringVar(value='')
        tk.Label(btn_frame, textvariable=self.status_var, font=('Segoe UI', 9),
                 bg=COLORS['bg_dark'], fg=COLORS['success']).pack(side='left')

        self.root.mainloop()

    # ─── Tab 1: App Defaults ───────────────────────────────────────
    def _build_app_defaults_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='🖥️ אפליקציות')

        # Card container
        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        frame = tk.Frame(card, bg=COLORS['bg_medium'])
        frame.pack(fill='both', expand=True, padx=14, pady=12)

        tk.Label(frame, text='שפת ברירת מחדל לכל אפליקציה',
                 font=('Segoe UI', 14, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(0, 4))
        tk.Label(frame, text='הגדר שפת ברירת מחדל לפי שם התהליך (exe). התוכנה תחליף את השפה אוטומטית כשתעבור לאפליקציה.',
                 font=('Segoe UI', 9), bg=COLORS['bg_medium'],
                 fg=COLORS['text_secondary'], wraplength=660, justify='right').pack(anchor='e')

        # Treeview
        tree_frame = tk.Frame(frame, bg=COLORS['bg_light'], bd=0, highlightthickness=1,
                              highlightbackground=COLORS['border'])
        tree_frame.pack(fill='both', expand=True, pady=10)

        self.app_tree = ttk.Treeview(tree_frame, columns=('app', 'lang'), show='headings', height=10)
        self.app_tree.heading('app', text='אפליקציה (exe)', anchor='e')
        self.app_tree.heading('lang', text='שפה', anchor='e')
        self.app_tree.column('app', width=350, anchor='e')
        self.app_tree.column('lang', width=100, anchor='center')

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.app_tree.yview)
        self.app_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='y')
        self.app_tree.pack(side='left', fill='both', expand=True)

        # Populate with alternating row tags
        for i, (app, lang) in enumerate(sorted(self.config.get('app_defaults', {}).items())):
            tag = 'stripe' if i % 2 else 'even'
            self.app_tree.insert('', 'end', values=(app, _lang_display(lang)), tags=(tag,))
        self.app_tree.tag_configure('stripe', background=COLORS['tree_stripe'])
        self.app_tree.tag_configure('even', background=COLORS['bg_medium'])

        # Add/Remove buttons
        btn_row = tk.Frame(frame, bg=COLORS['bg_medium'])
        btn_row.pack(fill='x', pady=(4, 0))

        self.app_exe_var = tk.StringVar()
        self.app_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.app_exe_var, width=25, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שם תהליך').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.app_lang_var,
                                 values=_get_available_lang_codes(), width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שפה').pack(side='right')

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_app).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton', command=self._remove_app).pack(side='right', padx=3)

    def _add_app(self):
        exe = self.app_exe_var.get().strip()
        lang = self.app_lang_var.get()
        if not exe:
            return
        lang_display = _lang_display(lang)
        # Check if exists
        for item in self.app_tree.get_children():
            if self.app_tree.item(item)['values'][0] == exe:
                self.app_tree.item(item, values=(exe, lang_display))
                self.app_exe_var.set('')
                return
        n = len(self.app_tree.get_children())
        tag = 'stripe' if n % 2 else 'even'
        self.app_tree.insert('', 'end', values=(exe, _lang_display(lang)), tags=(tag,))
        self.app_exe_var.set('')

    def _remove_app(self):
        selected = self.app_tree.selection()
        for item in selected:
            self.app_tree.delete(item)

    # ─── Tab 2: Chat Defaults ─────────────────────────────────────
    def _build_chat_defaults_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='💬 צ\'אטים')

        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        frame = tk.Frame(card, bg=COLORS['bg_medium'])
        frame.pack(fill='both', expand=True, padx=14, pady=12)

        tk.Label(frame, text='שפת ברירת מחדל לכל צ\'אט',
                 font=('Segoe UI', 14, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(0, 4))
        tk.Label(frame, text='הגדר שפה לפי שם צ\'אט (חלק מכותרת החלון). למשל: "מריה" = English',
                 font=('Segoe UI', 9), bg=COLORS['bg_medium'],
                 fg=COLORS['text_secondary']).pack(anchor='e')

        # Treeview
        tree_frame = tk.Frame(frame, bg=COLORS['bg_light'], bd=0, highlightthickness=1,
                              highlightbackground=COLORS['border'])
        tree_frame.pack(fill='both', expand=True, pady=10)

        self.chat_tree = ttk.Treeview(tree_frame, columns=('app', 'chat', 'lang'), show='headings', height=10)
        self.chat_tree.heading('app', text='אפליקציה', anchor='e')
        self.chat_tree.heading('chat', text='שם צ\'אט (מילת מפתח)', anchor='e')
        self.chat_tree.heading('lang', text='שפה', anchor='e')
        self.chat_tree.column('app', width=180, anchor='e')
        self.chat_tree.column('chat', width=250, anchor='e')
        self.chat_tree.column('lang', width=100, anchor='center')

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.chat_tree.yview)
        self.chat_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='y')
        self.chat_tree.pack(side='left', fill='both', expand=True)

        # Populate with alternating rows
        idx = 0
        for exe, chats in self.config.get('chat_defaults', {}).items():
            for title, lang in chats.items():
                tag = 'stripe' if idx % 2 else 'even'
                self.chat_tree.insert('', 'end', values=(exe, title, _lang_display(lang)), tags=(tag,))
                idx += 1
        self.chat_tree.tag_configure('stripe', background=COLORS['tree_stripe'])
        self.chat_tree.tag_configure('even', background=COLORS['bg_medium'])

        # Add/Remove
        btn_row = tk.Frame(frame, bg=COLORS['bg_medium'])
        btn_row.pack(fill='x')

        self.chat_exe_var = tk.StringVar(value='ms-teams.exe')
        self.chat_title_var = tk.StringVar()
        self.chat_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.chat_title_var, width=20, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שם צ\'אט').pack(side='right')

        # Build exe list
        known_exes = set(self.config.get('app_defaults', {}).keys())
        known_exes.update(self.config.get('chat_defaults', {}).keys())
        known_exes.update(self.config.get('watch_title_exes', []))
        known_exes.update(['chrome.exe', 'msedge.exe', 'outlook.exe', 'olk.exe', 'firefox.exe'])
        exe_values = sorted(known_exes)

        exe_combo = ttk.Combobox(btn_row, textvariable=self.chat_exe_var,
                                 values=exe_values, width=18)
        exe_combo.pack(side='right', padx=5)
        ttk.Label(btn_row, text=':אפליקציה').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.chat_lang_var,
                                 values=_get_available_lang_codes(), width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_chat).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton', command=self._remove_chat).pack(side='right', padx=3)

        # Second row for test button
        btn_row2 = tk.Frame(frame, bg=COLORS['bg_medium'])
        btn_row2.pack(fill='x', pady=(8, 0))
        ttk.Button(btn_row2, text='🔍 בדוק אפליקציה — האם תומכת בזיהוי צ\'אטים?',
                   command=self._test_app_title_support).pack(side='right', padx=5)

    def _add_chat(self):
        exe = self.chat_exe_var.get().strip()
        title = self.chat_title_var.get().strip()
        lang = self.chat_lang_var.get()
        if not exe or not title:
            return
        n = len(self.chat_tree.get_children())
        tag = 'stripe' if n % 2 else 'even'
        self.chat_tree.insert('', 'end', values=(exe, title, _lang_display(lang)), tags=(tag,))
        self.chat_title_var.set('')

    def _remove_chat(self):
        for item in self.chat_tree.selection():
            self.chat_tree.delete(item)

    def _test_app_title_support(self):
        """Open a popup that monitors the foreground window for 12s to check if title changes."""
        import ctypes
        import ctypes.wintypes as wintypes
        import time

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        _user32 = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32
        _ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

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
            if exe == 'applicationframehost.exe':
                parent_pid = pid.value
                result = ['']
                def cb(child, _):
                    cpid = wintypes.DWORD(0)
                    _user32.GetWindowThreadProcessId(child, ctypes.byref(cpid))
                    if cpid.value and cpid.value != parent_pid:
                        e = _pid_to_exe(cpid.value)
                        if e and e != 'applicationframehost.exe':
                            result[0] = e
                            return False
                    return True
                cb_func = _ENUM_CHILD_PROC(cb)
                _user32.EnumChildWindows(hwnd, cb_func, 0)
                if result[0]:
                    exe = result[0]
            length = _user32.GetWindowTextLengthW(hwnd)
            title = ''
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                _user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
            return exe, title

        # Shared state between thread and UI
        monitor_data = {'seen': {}, 'step': 0, 'done': False, 'report': '', 'supported_exe': ''}

        # Create popup — dark themed
        popup = tk.Toplevel(self.root)
        popup.title('בדיקת תמיכה בזיהוי צ\'אטים')
        popup.geometry('540x400')
        popup.resizable(False, False)
        popup.attributes('-topmost', True)
        popup.configure(bg=COLORS['bg_dark'])

        tk.Label(popup, text='🔍 בדיקת תמיכה בזיהוי צ\'אטים', font=('Segoe UI', 14, 'bold'),
                 bg=COLORS['bg_dark'], fg=COLORS['accent']).pack(pady=(15, 5))

        info_label = tk.Label(popup, text='⏳ עבור לאפליקציה והחלף בין צ\'אטים/מיילים/טאבים...\n\nהבדיקה תימשך 12 שניות',
                              font=('Segoe UI', 10), justify='center', wraplength=480,
                              bg=COLORS['bg_dark'], fg=COLORS['text_secondary'])
        info_label.pack(pady=10, padx=20)

        progress = ttk.Progressbar(popup, length=440, mode='determinate', maximum=12)
        progress.pack(pady=5)

        result_text = tk.Text(popup, font=('Consolas', 9), height=10, width=62, wrap='word',
                              state='disabled', bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                              insertbackground=COLORS['accent'], selectbackground=COLORS['accent_dim'],
                              relief='flat', bd=0, highlightthickness=1,
                              highlightbackground=COLORS['border'])
        result_text.pack(pady=10, padx=20, fill='both', expand=True)

        def _monitor_thread():
            """Background thread: collect foreground window data."""
            seen = {}
            for i in range(24):  # 12 seconds, check every 0.5s
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

            # Build report
            lines = []
            supported_exe = ''
            for exe, titles in sorted(seen.items()):
                count = len(titles)
                if count >= 2:
                    if not supported_exe:
                        supported_exe = exe
                    lines.append(f'✅ {exe} — נמצאו {count} כותרות שונות — תומך!')
                    for t in list(titles)[:3]:
                        short = t[:55] + '...' if len(t) > 55 else t
                        lines.append(f'    • {short}')
                else:
                    t = list(titles)[0] if titles else '(ריק)'
                    short = t[:45] + '...' if len(t) > 45 else t
                    lines.append(f'❌ {exe} — כותרת קבועה: "{short}"')

            if not lines:
                lines.append('לא זוהו אפליקציות. ודא שעברת לאפליקציה אחרת.')

            monitor_data['report'] = '\n'.join(lines)
            monitor_data['supported_exe'] = supported_exe
            monitor_data['done'] = True

        def _poll_ui():
            """Runs on main thread via after() - updates UI from shared state."""
            try:
                if not popup.winfo_exists():
                    return

                # Update progress
                progress.configure(value=monitor_data['step'] / 2)

                if monitor_data['done']:
                    # Show final report
                    has_supported = bool(monitor_data['supported_exe'])
                    info_label.configure(
                        text='✅ הבדיקה הסתיימה!' if has_supported else '⚠️ הבדיקה הסתיימה — לא נמצאו אפליקציות תומכות',
                        fg=COLORS['success'] if has_supported else COLORS['warning'])
                    progress.configure(value=12)

                    result_text.configure(state='normal')
                    result_text.delete('1.0', 'end')
                    result_text.insert('1.0', monitor_data['report'])
                    result_text.configure(state='disabled')

                    if has_supported:
                        self.chat_exe_var.set(monitor_data['supported_exe'])
                else:
                    # Keep polling
                    popup.after(300, _poll_ui)
            except Exception:
                pass  # popup was closed

        # Start
        thread = threading.Thread(target=_monitor_thread, daemon=True)
        thread.start()
        popup.after(300, _poll_ui)

    # ─── Tab 3: Browser / Website Defaults ────────────────────────
    def _build_browser_defaults_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='🌐 אתרים')

        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        frame = tk.Frame(card, bg=COLORS['bg_medium'])
        frame.pack(fill='both', expand=True, padx=14, pady=12)

        tk.Label(frame, text='שפת ברירת מחדל לפי אתר / דף',
                 font=('Segoe UI', 14, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(0, 4))
        tk.Label(frame, text=(
            'הגדר שפה לפי מילת מפתח בכותרת הדף בדפדפן.\n'
            'עובד עם כל דפדפן (Chrome, Edge, Firefox ועוד).\n'
            'למשל: "YouTube" → עברית, "GitHub" → English'
        ), font=('Segoe UI', 9), bg=COLORS['bg_medium'],
           fg=COLORS['text_secondary'], justify='right').pack(anchor='e')

        # Treeview
        tree_frame = tk.Frame(frame, bg=COLORS['bg_light'], bd=0, highlightthickness=1,
                              highlightbackground=COLORS['border'])
        tree_frame.pack(fill='both', expand=True, pady=10)

        self.browser_tree = ttk.Treeview(tree_frame, columns=('keyword', 'lang'), show='headings', height=10)
        self.browser_tree.heading('keyword', text='אתר / מילת מפתח', anchor='e')
        self.browser_tree.heading('lang', text='שפה', anchor='e')
        self.browser_tree.column('keyword', width=350, anchor='e')
        self.browser_tree.column('lang', width=100, anchor='center')

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.browser_tree.yview)
        self.browser_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='y')
        self.browser_tree.pack(side='left', fill='both', expand=True)

        # Populate
        idx = 0
        for keyword, lang in self.config.get('browser_defaults', {}).items():
            tag = 'stripe' if idx % 2 else 'even'
            self.browser_tree.insert('', 'end', values=(keyword, _lang_display(lang)), tags=(tag,))
            idx += 1
        self.browser_tree.tag_configure('stripe', background=COLORS['tree_stripe'])
        self.browser_tree.tag_configure('even', background=COLORS['bg_medium'])

        # Add / Remove
        btn_row = tk.Frame(frame, bg=COLORS['bg_medium'])
        btn_row.pack(fill='x')

        self.browser_keyword_var = tk.StringVar()
        self.browser_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.browser_keyword_var, width=25,
                  font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':אתר / דומיין / שם דף').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.browser_lang_var,
                                  values=_get_available_lang_codes(), width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_browser_rule).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton',
                   command=self._remove_browser_rule).pack(side='right', padx=3)

        # Hint
        hint_frame = tk.Frame(frame, bg=COLORS['bg_medium'])
        hint_frame.pack(fill='x', pady=(8, 0))
        tk.Label(hint_frame, text=(
            '💡 הכותרת בכרום נראית למשל: "YouTube - Google Chrome"\n'
            '     מספיק לכתוב "YouTube" או "youtube.com" כדי לזהות את האתר.'
        ), font=('Segoe UI', 9), bg=COLORS['bg_medium'],
           fg=COLORS['text_muted'], justify='right').pack(anchor='e')

    def _add_browser_rule(self):
        keyword = self.browser_keyword_var.get().strip()
        lang = self.browser_lang_var.get()
        if not keyword:
            return
        n = len(self.browser_tree.get_children())
        tag = 'stripe' if n % 2 else 'even'
        self.browser_tree.insert('', 'end', values=(keyword, _lang_display(lang)), tags=(tag,))
        self.browser_keyword_var.set('')

    def _remove_browser_rule(self):
        for item in self.browser_tree.selection():
            self.browser_tree.delete(item)

    # ─── Tab 4: Exclude Words ─────────────────────────────────────
    def _build_exclude_words_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='🚫 מילים לא להמרה')

        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        frame = tk.Frame(card, bg=COLORS['bg_medium'])
        frame.pack(fill='both', expand=True, padx=14, pady=12)

        tk.Label(frame, text='מילים/קיצורים שלא להמיר',
                 font=('Segoe UI', 14, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(0, 4))
        tk.Label(frame, text='הוסף מילים או קיצורים שהתוכנה לא תנסה להמיר (למשל: lol, ok, brb)',
                 font=('Segoe UI', 9), bg=COLORS['bg_medium'],
                 fg=COLORS['text_secondary']).pack(anchor='e')

        # Listbox — dark themed
        list_frame = tk.Frame(frame, bg=COLORS['bg_light'], bd=0, highlightthickness=1,
                              highlightbackground=COLORS['border'])
        list_frame.pack(fill='both', expand=True, pady=10)

        self.exclude_listbox = tk.Listbox(list_frame, font=('Segoe UI', 11), justify='right',
                                          selectmode='extended', bg=COLORS['bg_medium'],
                                          fg=COLORS['text_primary'], selectbackground=COLORS['accent_dim'],
                                          selectforeground='#ffffff', relief='flat', bd=0,
                                          highlightthickness=0, activestyle='none')
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.exclude_listbox.yview)
        self.exclude_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='y')
        self.exclude_listbox.pack(side='left', fill='both', expand=True)

        for word in self.config.get('exclude_words', []):
            self.exclude_listbox.insert('end', word)

        # Add/Remove
        btn_row = tk.Frame(frame, bg=COLORS['bg_medium'])
        btn_row.pack(fill='x')

        self.exclude_word_var = tk.StringVar()
        ttk.Entry(btn_row, textvariable=self.exclude_word_var, width=25, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':מילה/קיצור').pack(side='right')
        ttk.Button(btn_row, text='➕ הוסף', command=self._add_exclude).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton', command=self._remove_exclude).pack(side='right', padx=3)

    def _add_exclude(self):
        word = self.exclude_word_var.get().strip()
        if not word:
            return
        items = list(self.exclude_listbox.get(0, 'end'))
        if word.lower() not in [w.lower() for w in items]:
            self.exclude_listbox.insert('end', word)
        self.exclude_word_var.set('')

    def _remove_exclude(self):
        for i in reversed(self.exclude_listbox.curselection()):
            self.exclude_listbox.delete(i)

    # ─── Tab 4: General Settings ──────────────────────────────────
    def _build_general_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='⚙️ כללי')

        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        frame = tk.Frame(card, bg=COLORS['bg_medium'])
        frame.pack(fill='both', expand=True, padx=14, pady=12)

        tk.Label(frame, text='הגדרות כלליות',
                 font=('Segoe UI', 14, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(0, 12))

        settings_frame = tk.Frame(frame, bg=COLORS['bg_medium'])
        settings_frame.pack(fill='x', anchor='e')

        # Enabled
        self.enabled_var = tk.BooleanVar(value=self.config.get('enabled', True))
        tk.Checkbutton(settings_frame, text='תיקון שפה פעיל', variable=self.enabled_var,
                       bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                       selectcolor=COLORS['bg_light'], activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary'], font=('Segoe UI', 10),
                       bd=0, highlightthickness=0).pack(anchor='e', pady=5)

        # Debug
        self.debug_var = tk.BooleanVar(value=self.config.get('debug', False))
        tk.Checkbutton(settings_frame, text='(Debug) הצג לוג בקונסולה', variable=self.debug_var,
                       bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                       selectcolor=COLORS['bg_light'], activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary'], font=('Segoe UI', 10),
                       bd=0, highlightthickness=0).pack(anchor='e', pady=5)

        # Auto switch
        self.auto_switch_var = tk.BooleanVar(value=self.config.get('auto_switch', True))
        tk.Checkbutton(settings_frame, text='החלפת שפה אוטומטית אחרי תיקונים רצופים', variable=self.auto_switch_var,
                       bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                       selectcolor=COLORS['bg_light'], activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary'], font=('Segoe UI', 10),
                       bd=0, highlightthickness=0).pack(anchor='e', pady=5)

        # Auto switch count
        count_frame = tk.Frame(settings_frame, bg=COLORS['bg_medium'])
        count_frame.pack(anchor='e', pady=5)
        self.auto_switch_count_var = tk.IntVar(value=self.config.get('auto_switch_count', 2))
        ttk.Spinbox(count_frame, from_=1, to=10, width=5, textvariable=self.auto_switch_count_var).pack(side='right', padx=5)
        ttk.Label(count_frame, text='מספר תיקונים רצופים להחלפת שפה:').pack(side='right')

        # Separator
        tk.Frame(settings_frame, bg=COLORS['border'], height=1).pack(fill='x', pady=8)

        tk.Label(settings_frame, text='תצוגה',
                 font=('Segoe UI', 11, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent']).pack(anchor='e', pady=(4, 2))

        # Show typing panel
        self.show_panel_var = tk.BooleanVar(value=self.config.get('show_typing_panel', False))
        tk.Checkbutton(settings_frame, text='הצג חלונית הקלדה בהפעלה',
                       variable=self.show_panel_var,
                       bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                       selectcolor=COLORS['bg_light'], activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary'], font=('Segoe UI', 10),
                       bd=0, highlightthickness=0).pack(anchor='e', pady=5)

        # Hide word scores
        self.hide_scores_var = tk.BooleanVar(value=self.config.get('hide_scores', False))
        tk.Checkbutton(settings_frame, text='הסתר ציוני NLP למילים',
                       variable=self.hide_scores_var,
                       bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                       selectcolor=COLORS['bg_light'], activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary'], font=('Segoe UI', 10),
                       bd=0, highlightthickness=0).pack(anchor='e', pady=5)

        # Info section — styled inner card
        tk.Frame(frame, bg=COLORS['border'], height=1).pack(fill='x', pady=18)

        info_card = tk.Frame(frame, bg=COLORS['bg_light'], bd=0, highlightthickness=1,
                             highlightbackground=COLORS['border'])
        info_card.pack(fill='x', padx=5)

        info_inner = tk.Frame(info_card, bg=COLORS['bg_light'])
        info_inner.pack(fill='x', padx=15, pady=12)

        tk.Label(info_inner, text='ℹ️  AutoLang — תיקון שפת מקלדת אוטומטי',
                 font=('Segoe UI', 11, 'bold'), bg=COLORS['bg_light'],
                 fg=COLORS['accent'], anchor='e').pack(anchor='e')

        info_text = (
            "• בודק את 2 המילים הראשונות אחרי ENTER/נקודה\n"
            "• מילה 1 קובעת את שפת המשפט, מילה 2 הולכת אחריה\n"
            "• בדיקת אותיות סופיות (ך,ם,ן,ף,ץ) מונעת המרות שגויות\n"
            "• תרגום character-by-character לפי מיפוי מקלדת"
        )
        tk.Label(info_inner, text=info_text, justify='right', font=('Segoe UI', 9),
                 bg=COLORS['bg_light'], fg=COLORS['text_muted'], anchor='e').pack(anchor='e', pady=(6, 0))

    # ─── Tab 5: Help Guide ────────────────────────────────────────
    def _build_help_tab(self, notebook):
        outer = tk.Frame(notebook, bg=COLORS['bg_dark'])
        notebook.add(outer, text='📖 מדריך')

        card = tk.Frame(outer, bg=COLORS['bg_medium'], highlightthickness=1,
                        highlightbackground=COLORS['border'])
        card.pack(fill='both', expand=True, padx=8, pady=8)

        # Scrollable text
        text_frame = tk.Frame(card, bg=COLORS['bg_medium'])
        text_frame.pack(fill='both', expand=True, padx=4, pady=4)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side='left', fill='y')

        help_text = tk.Text(text_frame, wrap='word', font=('Segoe UI', 10),
                            bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                            bd=0, highlightthickness=0, padx=12, pady=10,
                            yscrollcommand=scrollbar.set, spacing3=4)
        help_text.pack(fill='both', expand=True)
        scrollbar.config(command=help_text.yview)

        help_text.tag_configure('title', font=('Segoe UI', 12, 'bold'),
                                foreground=COLORS['accent'], spacing1=10)
        help_text.tag_configure('body', font=('Segoe UI', 10),
                                foreground=COLORS['text_primary'])

        sections = [
            ('מה זה AutoLang?',
             'AutoLang היא תוכנה שמתקנת אוטומטית טקסט שהוקלד בשפה הלא נכונה.\n'
             'למשל — אם המקלדת שלך על אנגלית ואתה מקליד עברית, התוכנה מזהה את הטעות ומתקנת אותה בזמן אמת.\n'),
            ('🔵 כפתורי העיגול',
             '⏻  למעלה — הפעלה / השבתה של תיקון השפה\n'
             '↩  שמאל — ביטול התיקון האחרון (גם F12)\n'
             '✎  למטה — פתיחת עורך תרגום\n'
             '⚙  ימין — פתיחת חלון הגדרות\n'
             '🔴  מרכז — הקלטה קולית (Speech-to-Text)\n'),
            ('📝 חלונית ההקלדה',
             'מציגה את המילים שהוקלדו + התרגום שלהן בזמן אמת.\n'
             'שורה 1 — הטקסט שהוקלד\n'
             'שורה 2 — התרגום לשפה השנייה\n'
             'שורה 3 — ציוני NLP (ניתן להסתיר דרך התפריט)\n'
             'מילים שתוקנו מופיעות בצהוב עם קו תחתון — לחיצה עליהן מבטלת את התיקון.\n'
             'כדי להציג/להסתיר: קליק ימני → הצג/הסתר תיבת הקלדה.\n'),
            ('🔍 תרגום מהיר',
             'סמן טקסט בכל חלון ולחץ קליק ימני — יופיע כפתור "תרגם".\n'
             'לחיצה עליו תתרגם את הטקסט המסומן ותדביק אותו בחלון הפעיל.\n'),
            ('🎤 הקלטה קולית',
             'לחיצה על הכפתור האדום במרכז העיגול מפעילה הקלטה.\n'
             'התוכנה מקשיבה לדיבור, מזהה טקסט (Whisper) ומדביקה אותו בחלון הפעיל.\n'
             'לחיצה נוספת עוצרת את ההקלטה.\n'),
            ('⚙️ הגדרות',
             'אפליקציות — הגדרת שפת ברירת מחדל לכל תוכנה (לפי שם ה-exe).\n'
             'צ\'אטים — שפה שונה לפי שם שיחה (WhatsApp, Teams).\n'
             'דפדפנים — שפה לפי מילת מפתח בכותרת הטאב.\n'
             'מילים מוחרגות — מילים שלא יתוקנו (שמות, מונחים טכניים).\n'
             'כללי — הפעלה/השבתה, Debug, החלפת שפה אוטומטית, תצוגה.\n'),
            ('🧠 איך הזיהוי עובד?',
             'כל מילה שמוקלדת נבדקת מול מאגר מילים (NLP) בעברית ובאנגלית.\n'
             'המערכת משווה ציונים (zipf) ובוחרת את השפה שהמילה שייכת אליה.\n'
             'אחרי 2 מילים ברצף, המנוע "נועל" את שפת המשפט ומתקן מילים שוטפות ללא בדיקה נוספת\n'
             '(אלא אם המנוע מזהה שהתיקון לא הגיוני — אז הוא מדלג).\n'
             'ENTER או נקודה מאפסים את הנעילה ומתחילים משפט חדש.\n'),
            ('⌨️ קיצורי מקשים',
             'F12 — ביטול תיקון אחרון\n'
             'Ctrl+Alt+Q — יציאה מהתוכנה\n'
             'Ctrl+Alt+I — הצגת מידע (Debug)\n'),
        ]

        for title, body in sections:
            help_text.insert('end', title + '\n', 'title')
            help_text.insert('end', body + '\n', 'body')

        help_text.configure(state='disabled')

    # ─── Save ─────────────────────────────────────────────────────
    def _save(self):
        # Collect app defaults
        app_defaults = {}
        for item in self.app_tree.get_children():
            vals = self.app_tree.item(item)['values']
            app = str(vals[0])
            lang = _lang_code_from_display(str(vals[1]))
            app_defaults[app] = lang

        # Collect chat defaults
        chat_defaults = {}
        for item in self.chat_tree.get_children():
            vals = self.chat_tree.item(item)['values']
            exe = str(vals[0])
            title = str(vals[1])
            lang = _lang_code_from_display(str(vals[2]))
            if exe not in chat_defaults:
                chat_defaults[exe] = {}
            chat_defaults[exe][title] = lang

        # Collect exclude words
        exclude_words = list(self.exclude_listbox.get(0, 'end'))

        # Collect browser defaults
        browser_defaults = {}
        for item in self.browser_tree.get_children():
            vals = self.browser_tree.item(item)['values']
            keyword = str(vals[0])
            lang = _lang_code_from_display(str(vals[1]))
            browser_defaults[keyword] = lang

        self.config['app_defaults'] = app_defaults
        self.config['chat_defaults'] = chat_defaults
        self.config['exclude_words'] = exclude_words
        self.config['browser_defaults'] = browser_defaults

        # Auto-sync watch_title_exes: keep existing extras + all chat_defaults exes
        current_watch = set(self.config.get('watch_title_exes', []))
        current_watch.update(chat_defaults.keys())
        self.config['watch_title_exes'] = sorted(current_watch)

        self.config['enabled'] = self.enabled_var.get()
        self.config['debug'] = self.debug_var.get()
        self.config['auto_switch'] = self.auto_switch_var.get()
        self.config['auto_switch_count'] = self.auto_switch_count_var.get()
        self.config['hide_scores'] = self.hide_scores_var.get()
        self.config['show_typing_panel'] = self.show_panel_var.get()

        save_config(self.config)
        self.status_var.set('✅ ההגדרות נשמרו בהצלחה!')

        if self.on_save:
            self.on_save(self.config)

        self.root.after(2000, lambda: self.status_var.set(''))


class AutoLangTray:
    """System Tray application."""

    def __init__(self):
        self.config = load_config()
        self.enabled = self.config.get('enabled', True)
        self.icon = None
        self.settings_window = None
        self.engine_proc = None
        self.widget: FloatingWidget | None = None

    def _create_menu(self):
        return pystray.Menu(
            pystray.MenuItem(
                lambda item: '✅ פעיל' if self.enabled else '❌ מושבת',
                self._toggle_enabled,
                default=True,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('⚙️ הגדרות', self._open_settings),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('🚪 יציאה', self._quit),
        )

    def _toggle_enabled(self, icon=None, item=None):
        self.enabled = not self.enabled
        self.config['enabled'] = self.enabled
        save_config(self.config)
        self.icon.icon = create_tray_icon_image(self.enabled)
        self.icon.update_menu()
        # Update floating widget
        if self.widget:
            self.widget.update_state(self.enabled)
        # Push to engine
        try:
            import auto_lang
            auto_lang.ENGINE_ENABLED = self.enabled
        except Exception:
            pass

    def _open_settings(self, icon=None, item=None):
        # Guard against opening multiple settings windows
        if self.settings_window and self.settings_window.root:
            try:
                if self.settings_window.root.winfo_exists():
                    self.settings_window.root.lift()
                    self.settings_window.root.focus_force()
                    return
            except Exception:
                pass
        def _show():
            self.config = load_config()
            sw = SettingsWindow(self.config, on_save_callback=self._on_settings_saved)
            self.settings_window = sw
            sw.show()
        threading.Thread(target=_show, daemon=True).start()

    def _on_settings_saved(self, new_config):
        self.config = new_config
        self.enabled = new_config.get('enabled', True)
        if self.icon:
            self.icon.icon = create_tray_icon_image(self.enabled)
        # Sync floating widget
        if self.widget:
            self.widget.update_state(self.enabled)
            # Sync display preferences
            new_hide = new_config.get('hide_scores', False)
            new_panel = new_config.get('show_typing_panel', False)
            if self.widget._hide_scores != new_hide:
                self.widget._hide_scores = new_hide
                if self.widget.root:
                    self.widget.root.after_idle(self.widget._redraw_text_words)
            if self.widget._panel_visible != new_panel:
                self.widget._panel_visible = new_panel
                if self.widget.root:
                    self.widget.root.after_idle(self.widget._apply_panel)

        # Push changes to the running engine
        try:
            import auto_lang
            self._apply_config_to_engine(auto_lang, new_config)
        except Exception:
            pass

    @staticmethod
    def _apply_config_to_engine(engine, cfg):
        """Push UI config into the auto_lang module (v3 compatible)."""
        engine.APP_DEFAULT_LANG_BY_EXE.clear()
        engine.APP_DEFAULT_LANG_BY_EXE.update(cfg.get('app_defaults', {}))

        engine.APP_DEFAULT_LANG_BY_TITLE.clear()
        for exe, chats in cfg.get('chat_defaults', {}).items():
            if chats:
                engine.APP_DEFAULT_LANG_BY_TITLE[exe] = chats

        # Rebuild the watch-title set
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

    def _quit(self, icon=None, item=None):
        # Destroy floating widget
        if self.widget:
            self.widget.destroy()
        # Signal the engine to stop
        try:
            import auto_lang
            auto_lang.stop_event.set()
        except Exception:
            pass
        if self.icon:
            self.icon.stop()
        os._exit(0)

    def _start_engine(self):
        """Start the auto_lang engine in a background thread."""
        try:
            import auto_lang

            # Apply UI config directly to the module
            self._apply_config_to_engine(auto_lang, self.config)

            # Run engine main() - this blocks (stop_event.wait)
            auto_lang.main()

        except Exception as e:
            print(f'Engine start failed: {e}')
            import traceback
            traceback.print_exc()

    def run(self):
        if not pystray or not Image:
            print('pystray or Pillow not available. Running engine only.')
            self._start_engine()
            return

        # Start the correction engine in a background thread
        engine_thread = threading.Thread(target=self._start_engine, daemon=True)
        engine_thread.start()

        # Start the floating widget
        self.widget = FloatingWidget(self)
        self.widget.start()

        # Create and run system tray icon (blocks main thread)
        self.icon = pystray.Icon(
            'AutoLang',
            icon=create_tray_icon_image(self.enabled),
            title='AutoLang - תיקון שפה אוטומטי',
            menu=self._create_menu(),
        )
        self.icon.run()


def main():
    app = AutoLangTray()
    app.run()


if __name__ == '__main__':
    main()
