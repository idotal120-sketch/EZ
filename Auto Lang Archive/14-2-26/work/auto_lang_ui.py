"""
AutoLang UI - System Tray + Settings Window
============================================
ממשק משתמש לסקריפט auto_lang2.py:
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
    'exclude_words': [],  # מילים לא להמיר
    'enabled': True,
    'debug': False,
    'auto_switch': True,
    'auto_switch_count': 2,
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
# System Tray Icon
# ----------------------------

def create_tray_icon_image(enabled: bool = True) -> 'Image':
    """Create a modern icon with gradient feel."""
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    if enabled:
        # Outer glow ring
        draw.ellipse([2, 2, size - 2, size - 2], fill=(0, 200, 180, 80))
        # Main circle — teal gradient feel
        draw.ellipse([5, 5, size - 5, size - 5], fill=(0, 180, 160, 255))
        # Inner highlight
        draw.ellipse([8, 8, size - 8, size - 8], fill=(0, 210, 190, 255))
    else:
        draw.ellipse([2, 2, size - 2, size - 2], fill=(120, 50, 50, 80))
        draw.ellipse([5, 5, size - 5, size - 5], fill=(160, 60, 60, 255))
        draw.ellipse([8, 8, size - 8, size - 8], fill=(180, 70, 70, 255))

    # Letters "אa" in white
    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except Exception:
        font = ImageFont.load_default()

    draw.text((12, 12), "אa", fill=(255, 255, 255, 255), font=font)
    return img


# ──────────────────────────────────────────────────────────
# Dark Theme Color Palette
# ──────────────────────────────────────────────────────────
COLORS = {
    'bg_dark':       '#1a1b2e',     # Main background
    'bg_medium':     '#232440',     # Card / panel background
    'bg_light':      '#2d2f52',     # Input / entry background
    'bg_highlight':  '#353766',     # Hover / selection background
    'accent':        '#00d4aa',     # Primary accent (teal/mint)
    'accent_hover':  '#00f0c0',     # Accent hover
    'accent_dim':    '#008f75',     # Accent muted
    'text_primary':  '#e8e8f0',     # Main text
    'text_secondary':'#9a9ab8',     # Secondary text
    'text_muted':    '#6b6b88',     # Muted text
    'border':        '#3a3c66',     # Borders
    'success':       '#00e676',     # Success green
    'error':         '#ff5252',     # Error red
    'warning':       '#ffc107',     # Warning yellow
    'tab_active':    '#00d4aa',     # Active tab accent
    'tab_inactive':  '#2d2f52',     # Inactive tab bg
    'btn_primary':   '#00d4aa',     # Primary button
    'btn_danger':    '#ff5252',     # Danger button
    'tree_stripe':   '#262848',     # Treeview alternating rows
    'scrollbar':     '#4a4c7a',     # Scrollbar thumb
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

        # Global background
        style.configure('.', background=COLORS['bg_dark'], foreground=COLORS['text_primary'],
                        fieldbackground=COLORS['bg_light'], bordercolor=COLORS['border'],
                        troughcolor=COLORS['bg_medium'], selectbackground=COLORS['accent_dim'],
                        selectforeground='#ffffff', font=('Segoe UI', 10))

        # Frames
        style.configure('TFrame', background=COLORS['bg_dark'])
        style.configure('Card.TFrame', background=COLORS['bg_medium'])

        # Labels
        style.configure('TLabel', background=COLORS['bg_dark'], foreground=COLORS['text_primary'],
                        font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 14, 'bold'),
                        foreground=COLORS['accent'], background=COLORS['bg_dark'])
        style.configure('Status.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['success'], background=COLORS['bg_dark'])
        style.configure('Subtitle.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['text_secondary'], background=COLORS['bg_dark'])
        style.configure('Info.TLabel', font=('Segoe UI', 9),
                        foreground=COLORS['text_muted'], background=COLORS['bg_dark'])

        # Notebook (tabs)
        style.configure('TNotebook', background=COLORS['bg_dark'], borderwidth=0,
                        tabmargins=[2, 8, 2, 0])
        style.configure('TNotebook.Tab', background=COLORS['tab_inactive'],
                        foreground=COLORS['text_secondary'], padding=[16, 8],
                        font=('Segoe UI', 11), borderwidth=0)
        style.map('TNotebook.Tab',
                  background=[('selected', COLORS['bg_dark']), ('!selected', COLORS['tab_inactive'])],
                  foreground=[('selected', COLORS['accent']), ('!selected', COLORS['text_secondary'])],
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

        # Treeview
        style.configure('Treeview', background=COLORS['bg_medium'], foreground=COLORS['text_primary'],
                        fieldbackground=COLORS['bg_medium'], borderwidth=0,
                        font=('Segoe UI', 10), rowheight=28)
        style.configure('Treeview.Heading', background=COLORS['bg_light'],
                        foreground=COLORS['accent'], font=('Segoe UI', 10, 'bold'),
                        borderwidth=0, relief='flat')
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

        # Checkbutton
        style.configure('TCheckbutton', background=COLORS['bg_dark'],
                        foreground=COLORS['text_primary'], font=('Segoe UI', 10),
                        indicatorcolor=COLORS['bg_light'], indicatorrelief='flat')
        style.map('TCheckbutton',
                  background=[('active', COLORS['bg_dark'])],
                  indicatorcolor=[('selected', COLORS['accent']), ('!selected', COLORS['bg_light'])])

        # Spinbox
        style.configure('TSpinbox', fieldbackground=COLORS['bg_light'],
                        foreground=COLORS['text_primary'], bordercolor=COLORS['border'],
                        arrowcolor=COLORS['accent'])
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
        title_frame = tk.Frame(self.root, bg=COLORS['bg_dark'], height=50)
        title_frame.pack(fill='x', padx=15, pady=(10, 0))
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text='⌨️', font=('Segoe UI Emoji', 18),
                 bg=COLORS['bg_dark'], fg=COLORS['accent']).pack(side='right', padx=(0, 8))
        tk.Label(title_frame, text='AutoLang', font=('Segoe UI', 18, 'bold'),
                 bg=COLORS['bg_dark'], fg=COLORS['text_primary']).pack(side='right')
        tk.Label(title_frame, text='הגדרות  |  ', font=('Segoe UI', 12),
                 bg=COLORS['bg_dark'], fg=COLORS['text_secondary']).pack(side='right')

        # Accent line under title
        tk.Frame(self.root, bg=COLORS['accent'], height=2).pack(fill='x', padx=15, pady=(0, 5))

        # Notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=15, pady=(5, 10))

        # Tab 1: App defaults
        self._build_app_defaults_tab(notebook)

        # Tab 2: Chat defaults
        self._build_chat_defaults_tab(notebook)

        # Tab 3: Exclude words
        self._build_exclude_words_tab(notebook)

        # Tab 4: General settings
        self._build_general_tab(notebook)

        # Bottom buttons
        btn_frame = tk.Frame(self.root, bg=COLORS['bg_dark'])
        btn_frame.pack(fill='x', padx=15, pady=(0, 12))

        ttk.Button(btn_frame, text='💾 שמור', style='Accent.TButton', command=self._save).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='❌ בטל', command=self.root.destroy).pack(side='right', padx=5)

        self.status_var = tk.StringVar(value='')
        tk.Label(btn_frame, textvariable=self.status_var, font=('Segoe UI', 9),
                 bg=COLORS['bg_dark'], fg=COLORS['success']).pack(side='left')

        self.root.mainloop()

    # ─── Tab 1: App Defaults ───────────────────────────────────────
    def _build_app_defaults_tab(self, notebook):
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text='🖥️ אפליקציות')

        ttk.Label(frame, text='שפת ברירת מחדל לכל אפליקציה', style='Header.TLabel').pack(anchor='e', pady=(0, 6))
        ttk.Label(frame, text='הגדר שפת ברירת מחדל לפי שם התהליך (exe). התוכנה תחליף את השפה אוטומטית כשתעבור לאפליקציה.',
                  style='Subtitle.TLabel', wraplength=680).pack(anchor='e')

        # Treeview
        tree_frame = tk.Frame(frame, bg=COLORS['bg_medium'], bd=0, highlightthickness=1,
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
            lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
            tag = 'stripe' if i % 2 else ''
            self.app_tree.insert('', 'end', values=(app, lang_display), tags=(tag,))
        self.app_tree.tag_configure('stripe', background=COLORS['tree_stripe'])

        # Add/Remove buttons
        btn_row = tk.Frame(frame, bg=COLORS['bg_dark'])
        btn_row.pack(fill='x', pady=(4, 0))

        self.app_exe_var = tk.StringVar()
        self.app_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.app_exe_var, width=25, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שם תהליך').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.app_lang_var, values=['he', 'en'], width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שפה').pack(side='right')

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_app).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton', command=self._remove_app).pack(side='right', padx=3)

    def _add_app(self):
        exe = self.app_exe_var.get().strip()
        lang = self.app_lang_var.get()
        if not exe:
            return
        lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
        # Check if exists
        for item in self.app_tree.get_children():
            if self.app_tree.item(item)['values'][0] == exe:
                self.app_tree.item(item, values=(exe, lang_display))
                self.app_exe_var.set('')
                return
        self.app_tree.insert('', 'end', values=(exe, lang_display))
        self.app_exe_var.set('')

    def _remove_app(self):
        selected = self.app_tree.selection()
        for item in selected:
            self.app_tree.delete(item)

    # ─── Tab 2: Chat Defaults ─────────────────────────────────────
    def _build_chat_defaults_tab(self, notebook):
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text='💬 צ\'אטים')

        ttk.Label(frame, text='שפת ברירת מחדל לכל צ\'אט', style='Header.TLabel').pack(anchor='e', pady=(0, 6))
        ttk.Label(frame, text='הגדר שפה לפי שם צ\'אט (חלק מכותרת החלון). למשל: "מריה" = English',
                  style='Subtitle.TLabel').pack(anchor='e')

        # Treeview
        tree_frame = tk.Frame(frame, bg=COLORS['bg_medium'], bd=0, highlightthickness=1,
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
                lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
                tag = 'stripe' if idx % 2 else ''
                self.chat_tree.insert('', 'end', values=(exe, title, lang_display), tags=(tag,))
                idx += 1
        self.chat_tree.tag_configure('stripe', background=COLORS['tree_stripe'])

        # Add/Remove
        btn_row = tk.Frame(frame, bg=COLORS['bg_dark'])
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

        lang_combo = ttk.Combobox(btn_row, textvariable=self.chat_lang_var, values=['he', 'en'], width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_chat).pack(side='right', padx=3)
        ttk.Button(btn_row, text='🗑️ הסר', style='Danger.TButton', command=self._remove_chat).pack(side='right', padx=3)

        # Second row for test button
        btn_row2 = tk.Frame(frame, bg=COLORS['bg_dark'])
        btn_row2.pack(fill='x', pady=(8, 0))
        ttk.Button(btn_row2, text='🔍 בדוק אפליקציה — האם תומכת בזיהוי צ\'אטים?',
                   command=self._test_app_title_support).pack(side='right', padx=5)

    def _add_chat(self):
        exe = self.chat_exe_var.get().strip()
        title = self.chat_title_var.get().strip()
        lang = self.chat_lang_var.get()
        if not exe or not title:
            return
        lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
        self.chat_tree.insert('', 'end', values=(exe, title, lang_display))
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
                _user32.EnumChildWindows(hwnd, _ENUM_CHILD_PROC(cb), 0)
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
            """Runs on main thread via after() — updates UI from shared state."""
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

    # ─── Tab 3: Exclude Words ─────────────────────────────────────
    def _build_exclude_words_tab(self, notebook):
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text='🚫 מילים לא להמרה')

        ttk.Label(frame, text='מילים/קיצורים שלא להמיר', style='Header.TLabel').pack(anchor='e', pady=(0, 6))
        ttk.Label(frame, text='הוסף מילים או קיצורים שהתוכנה לא תנסה להמיר (למשל: lol, ok, brb)',
                  style='Subtitle.TLabel').pack(anchor='e')

        # Listbox — dark themed
        list_frame = tk.Frame(frame, bg=COLORS['bg_medium'], bd=0, highlightthickness=1,
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
        btn_row = tk.Frame(frame, bg=COLORS['bg_dark'])
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
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text='⚙️ כללי')

        ttk.Label(frame, text='הגדרות כלליות', style='Header.TLabel').pack(anchor='e', pady=(0, 12))

        settings_frame = ttk.Frame(frame)
        settings_frame.pack(fill='x', anchor='e')

        # Enabled
        self.enabled_var = tk.BooleanVar(value=self.config.get('enabled', True))
        ttk.Checkbutton(settings_frame, text='תיקון שפה פעיל', variable=self.enabled_var).pack(anchor='e', pady=5)

        # Debug
        self.debug_var = tk.BooleanVar(value=self.config.get('debug', False))
        ttk.Checkbutton(settings_frame, text='(Debug) הצג לוג בקונסולה', variable=self.debug_var).pack(anchor='e', pady=5)

        # Auto switch
        self.auto_switch_var = tk.BooleanVar(value=self.config.get('auto_switch', True))
        ttk.Checkbutton(settings_frame, text='החלפת שפה אוטומטית אחרי תיקונים רצופים', variable=self.auto_switch_var).pack(anchor='e', pady=5)

        # Auto switch count
        count_frame = tk.Frame(settings_frame, bg=COLORS['bg_dark'])
        count_frame.pack(anchor='e', pady=5)
        self.auto_switch_count_var = tk.IntVar(value=self.config.get('auto_switch_count', 2))
        ttk.Spinbox(count_frame, from_=1, to=10, width=5, textvariable=self.auto_switch_count_var).pack(side='right', padx=5)
        ttk.Label(count_frame, text='מספר תיקונים רצופים להחלפת שפה:').pack(side='right')

        # Info section — styled card
        tk.Frame(frame, bg=COLORS['border'], height=1).pack(fill='x', pady=18)

        info_card = tk.Frame(frame, bg=COLORS['bg_medium'], bd=0, highlightthickness=1,
                             highlightbackground=COLORS['border'])
        info_card.pack(fill='x', padx=5)

        info_inner = tk.Frame(info_card, bg=COLORS['bg_medium'])
        info_inner.pack(fill='x', padx=15, pady=12)

        tk.Label(info_inner, text='ℹ️  AutoLang — תיקון שפת מקלדת אוטומטי',
                 font=('Segoe UI', 11, 'bold'), bg=COLORS['bg_medium'],
                 fg=COLORS['accent'], anchor='e').pack(anchor='e')

        info_text = (
            "• בודק את 2 המילים הראשונות אחרי ENTER/נקודה\n"
            "• מילה 1 קובעת את שפת המשפט, מילה 2 הולכת אחריה\n"
            "• בדיקת אותיות סופיות (ך,ם,ן,ף,ץ) מונעת המרות שגויות\n"
            "• תרגום character-by-character לפי מיפוי מקלדת"
        )
        tk.Label(info_inner, text=info_text, justify='right', font=('Segoe UI', 9),
                 bg=COLORS['bg_medium'], fg=COLORS['text_muted'], anchor='e').pack(anchor='e', pady=(6, 0))

    # ─── Save ─────────────────────────────────────────────────────
    def _save(self):
        # Collect app defaults
        app_defaults = {}
        for item in self.app_tree.get_children():
            vals = self.app_tree.item(item)['values']
            app = str(vals[0])
            lang = 'he' if 'עברית' in str(vals[1]) else 'en'
            app_defaults[app] = lang

        # Collect chat defaults
        chat_defaults = {}
        for item in self.chat_tree.get_children():
            vals = self.chat_tree.item(item)['values']
            exe = str(vals[0])
            title = str(vals[1])
            lang = 'he' if 'עברית' in str(vals[2]) else 'en'
            if exe not in chat_defaults:
                chat_defaults[exe] = {}
            chat_defaults[exe][title] = lang

        # Collect exclude words
        exclude_words = list(self.exclude_listbox.get(0, 'end'))

        self.config['app_defaults'] = app_defaults
        self.config['chat_defaults'] = chat_defaults
        self.config['exclude_words'] = exclude_words

        # Auto-sync watch_title_exes: keep existing extras + all chat_defaults exes
        current_watch = set(self.config.get('watch_title_exes', []))
        current_watch.update(chat_defaults.keys())
        self.config['watch_title_exes'] = sorted(current_watch)

        self.config['enabled'] = self.enabled_var.get()
        self.config['debug'] = self.debug_var.get()
        self.config['auto_switch'] = self.auto_switch_var.get()
        self.config['auto_switch_count'] = self.auto_switch_count_var.get()

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
        # Push to engine
        try:
            import auto_lang2
            auto_lang2.ENGINE_ENABLED = self.enabled
        except Exception:
            pass

    def _open_settings(self, icon=None, item=None):
        def _show():
            self.config = load_config()
            sw = SettingsWindow(self.config, on_save_callback=self._on_settings_saved)
            sw.show()
        threading.Thread(target=_show, daemon=True).start()

    def _on_settings_saved(self, new_config):
        self.config = new_config
        self.enabled = new_config.get('enabled', True)
        if self.icon:
            self.icon.icon = create_tray_icon_image(self.enabled)

        # Push changes to the running engine
        try:
            import auto_lang2
            self._apply_config_to_engine(auto_lang2, new_config)
        except Exception:
            pass

    @staticmethod
    def _apply_config_to_engine(engine, cfg):
        """Push UI config into the auto_lang2 module."""
        engine.APP_DEFAULT_LANG_BY_EXE.clear()
        engine.APP_DEFAULT_LANG_BY_EXE.update(cfg.get('app_defaults', {}))

        engine.APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING.clear()
        for exe, chats in cfg.get('chat_defaults', {}).items():
            if chats:
                engine.APP_DEFAULT_LANG_BY_EXE_AND_TITLE_SUBSTRING[exe] = chats

        # Extra watch-title exes
        engine._EXTRA_WATCH_TITLE_EXE.clear()
        engine._EXTRA_WATCH_TITLE_EXE.update(cfg.get('watch_title_exes', []))

        # Rebuild the unified watch set
        engine._rebuild_watch_title_set()

        engine.EXCLUDE_WORDS = set(w.lower() for w in cfg.get('exclude_words', []))
        engine.ENGINE_ENABLED = cfg.get('enabled', True)
        engine.DEBUG = cfg.get('debug', False)
        engine.AUTO_SWITCH_LAYOUT = cfg.get('auto_switch', True)
        engine.AUTO_SWITCH_AFTER_CONSECUTIVE = cfg.get('auto_switch_count', 2)

    def _quit(self, icon=None, item=None):
        # Signal the engine to stop
        try:
            import auto_lang2
            auto_lang2.stop_event.set()
        except Exception:
            pass
        if self.icon:
            self.icon.stop()
        os._exit(0)

    def _start_engine(self):
        """Start the auto_lang2 engine in a background thread."""
        try:
            import auto_lang2

            # Apply UI config directly to the module
            self._apply_config_to_engine(auto_lang2, self.config)

            # Run engine main() - this blocks (stop_event.wait)
            auto_lang2.main()

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
