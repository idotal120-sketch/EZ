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
    """Create a simple icon: green circle if enabled, red if disabled."""
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Background circle
    bg_color = (34, 139, 34, 255) if enabled else (180, 50, 50, 255)
    draw.ellipse([4, 4, size - 4, size - 4], fill=bg_color)

    # Letters "אa" in white
    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except Exception:
        font = ImageFont.load_default()

    draw.text((12, 12), "אa", fill=(255, 255, 255, 255), font=font)
    return img


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
        self.root.geometry('700x550')
        self.root.resizable(True, True)

        # RTL support
        self.root.option_add('*TCombobox*Listbox.justify', 'right')

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook.Tab', font=('Segoe UI', 11), padding=[12, 6])
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=[8, 4])
        style.configure('Header.TLabel', font=('Segoe UI', 13, 'bold'))
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#666')

        # Notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 1: App defaults
        self._build_app_defaults_tab(notebook)

        # Tab 2: Chat defaults
        self._build_chat_defaults_tab(notebook)

        # Tab 3: Exclude words
        self._build_exclude_words_tab(notebook)

        # Tab 4: General settings
        self._build_general_tab(notebook)

        # Bottom buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', padx=10, pady=(0, 10))

        ttk.Button(btn_frame, text='💾 שמור', command=self._save).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='❌ בטל', command=self.root.destroy).pack(side='right', padx=5)

        self.status_var = tk.StringVar(value='')
        ttk.Label(btn_frame, textvariable=self.status_var, style='Status.TLabel').pack(side='left')

        self.root.mainloop()

    # ─── Tab 1: App Defaults ───────────────────────────────────────
    def _build_app_defaults_tab(self, notebook):
        frame = ttk.Frame(notebook, padding=10)
        notebook.add(frame, text='🖥️ אפליקציות')

        ttk.Label(frame, text='שפת ברירת מחדל לכל אפליקציה', style='Header.TLabel').pack(anchor='e', pady=(0, 10))
        ttk.Label(frame, text='הגדר שפת ברירת מחדל לפי שם התהליך (exe). התוכנה תחליף את השפה אוטומטית כשתעבור לאפליקציה.').pack(anchor='e')

        # Treeview
        tree_frame = ttk.Frame(frame)
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

        # Populate
        for app, lang in sorted(self.config.get('app_defaults', {}).items()):
            lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
            self.app_tree.insert('', 'end', values=(app, lang_display))

        # Add/Remove buttons
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill='x')

        self.app_exe_var = tk.StringVar()
        self.app_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.app_exe_var, width=25, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שם תהליך').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.app_lang_var, values=['he', 'en'], width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שפה').pack(side='right')

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_app).pack(side='right', padx=5)
        ttk.Button(btn_row, text='🗑️ הסר', command=self._remove_app).pack(side='right', padx=5)

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
        frame = ttk.Frame(notebook, padding=10)
        notebook.add(frame, text='💬 צ\'אטים')

        ttk.Label(frame, text='שפת ברירת מחדל לכל צ\'אט', style='Header.TLabel').pack(anchor='e', pady=(0, 10))
        ttk.Label(frame, text='הגדר שפה לפי שם צ\'אט (חלק מכותרת החלון). למשל: "מריה" = English').pack(anchor='e')

        # Treeview
        tree_frame = ttk.Frame(frame)
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

        # Populate
        for exe, chats in self.config.get('chat_defaults', {}).items():
            for title, lang in chats.items():
                lang_display = 'עברית 🇮🇱' if lang == 'he' else 'English 🇺🇸'
                self.chat_tree.insert('', 'end', values=(exe, title, lang_display))

        # Add/Remove
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill='x')

        self.chat_exe_var = tk.StringVar(value='ms-teams.exe')
        self.chat_title_var = tk.StringVar()
        self.chat_lang_var = tk.StringVar(value='he')

        ttk.Entry(btn_row, textvariable=self.chat_title_var, width=20, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':שם צ\'אט').pack(side='right')

        exe_combo = ttk.Combobox(btn_row, textvariable=self.chat_exe_var,
                                 values=['ms-teams.exe', 'teams.exe', 'whatsapp.exe', 'whatsapp.root.exe', 'telegram.exe'],
                                 width=18, state='readonly')
        exe_combo.pack(side='right', padx=5)
        ttk.Label(btn_row, text=':אפליקציה').pack(side='right')

        lang_combo = ttk.Combobox(btn_row, textvariable=self.chat_lang_var, values=['he', 'en'], width=5, state='readonly')
        lang_combo.pack(side='right', padx=5)

        ttk.Button(btn_row, text='➕ הוסף', command=self._add_chat).pack(side='right', padx=5)
        ttk.Button(btn_row, text='🗑️ הסר', command=self._remove_chat).pack(side='right', padx=5)

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

    # ─── Tab 3: Exclude Words ─────────────────────────────────────
    def _build_exclude_words_tab(self, notebook):
        frame = ttk.Frame(notebook, padding=10)
        notebook.add(frame, text='🚫 מילים לא להמרה')

        ttk.Label(frame, text='מילים/קיצורים שלא להמיר', style='Header.TLabel').pack(anchor='e', pady=(0, 10))
        ttk.Label(frame, text='הוסף מילים או קיצורים שהתוכנה לא תנסה להמיר (למשל: lol, ok, brb)').pack(anchor='e')

        # Listbox
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, pady=10)

        self.exclude_listbox = tk.Listbox(list_frame, font=('Segoe UI', 11), justify='right', selectmode='extended')
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.exclude_listbox.yview)
        self.exclude_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='y')
        self.exclude_listbox.pack(side='left', fill='both', expand=True)

        for word in self.config.get('exclude_words', []):
            self.exclude_listbox.insert('end', word)

        # Add/Remove
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill='x')

        self.exclude_word_var = tk.StringVar()
        ttk.Entry(btn_row, textvariable=self.exclude_word_var, width=25, font=('Segoe UI', 10)).pack(side='right', padx=5)
        ttk.Label(btn_row, text=':מילה/קיצור').pack(side='right')
        ttk.Button(btn_row, text='➕ הוסף', command=self._add_exclude).pack(side='right', padx=5)
        ttk.Button(btn_row, text='🗑️ הסר', command=self._remove_exclude).pack(side='right', padx=5)

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
        frame = ttk.Frame(notebook, padding=10)
        notebook.add(frame, text='⚙️ כללי')

        ttk.Label(frame, text='הגדרות כלליות', style='Header.TLabel').pack(anchor='e', pady=(0, 15))

        settings_frame = ttk.Frame(frame)
        settings_frame.pack(fill='x', anchor='e')

        # Enabled
        self.enabled_var = tk.BooleanVar(value=self.config.get('enabled', True))
        ttk.Checkbutton(settings_frame, text='תיקון שפה פעיל', variable=self.enabled_var).pack(anchor='e', pady=3)

        # Debug
        self.debug_var = tk.BooleanVar(value=self.config.get('debug', False))
        ttk.Checkbutton(settings_frame, text='(Debug) הצג לוג בקונסולה', variable=self.debug_var).pack(anchor='e', pady=3)

        # Auto switch
        self.auto_switch_var = tk.BooleanVar(value=self.config.get('auto_switch', True))
        ttk.Checkbutton(settings_frame, text='החלפת שפה אוטומטית אחרי תיקונים רצופים', variable=self.auto_switch_var).pack(anchor='e', pady=3)

        # Auto switch count
        count_frame = ttk.Frame(settings_frame)
        count_frame.pack(anchor='e', pady=3)
        self.auto_switch_count_var = tk.IntVar(value=self.config.get('auto_switch_count', 2))
        ttk.Spinbox(count_frame, from_=1, to=10, width=5, textvariable=self.auto_switch_count_var).pack(side='right', padx=5)
        ttk.Label(count_frame, text='מספר תיקונים רצופים להחלפת שפה:').pack(side='right')

        # Info
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)
        info_text = (
            "AutoLang - תיקון שפת מקלדת אוטומטי\n"
            "────────────────────────────\n"
            "הלוגיקה:\n"
            "• בודק את 2 המילים הראשונות אחרי ENTER/נקודה\n"
            "• מילה 1 קובעת את שפת המשפט, מילה 2 הולכת אחריה\n"
            "• בדיקת אותיות סופיות (ך,ם,ן,ף,ץ) מונעת המרות שגויות\n"
            "• תרגום character-by-character לפי מיפוי מקלדת"
        )
        info_label = ttk.Label(frame, text=info_text, justify='right', font=('Segoe UI', 9), foreground='#555')
        info_label.pack(anchor='e')

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
