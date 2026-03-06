"""
Shared config store for AutoLang (engine + UI).

Important behavioral detail:
- The UI expects a *merged* config (defaults filled in).
- The engine expects the *raw saved* JSON so missing keys do not override
  engine-side defaults.
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, Optional


CONFIG_FILE = os.path.expanduser('~/.auto_lang2_config.json')

DEFAULT_CONFIG: Dict[str, Any] = {
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
    # Speech-to-text (Whisper)
    'speech_enabled': False,
    # Grammar / LLM
    'grammar_enabled': False,
    'grammar_provider': 'openai',  # 'openai' | 'anthropic' | 'gemini'
    'grammar_api_key': '',
    'grammar_model': '',           # empty → use provider default
}


def read_saved_config(path: str = CONFIG_FILE) -> Optional[Dict[str, Any]]:
    """Read saved JSON config as-is (no merging). Returns None if missing/invalid."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            cfg = json.load(f)
        return cfg if isinstance(cfg, dict) else None
    except Exception:
        return None


def load_config(path: str = CONFIG_FILE) -> Dict[str, Any]:
    """Load config merged with DEFAULT_CONFIG (UI-friendly)."""
    saved = read_saved_config(path)
    if saved is None:
        return DEFAULT_CONFIG.copy()

    config = DEFAULT_CONFIG.copy()
    for key in DEFAULT_CONFIG:
        if key in saved:
            config[key] = saved[key]
    return config


def save_config(config: Dict[str, Any], path: str = CONFIG_FILE) -> None:
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f'Failed to save config: {e}')

