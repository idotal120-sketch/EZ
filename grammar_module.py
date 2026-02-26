"""
Grammar / phrasing correction module for AutoLang.

Supports multiple LLM providers:
  - OpenAI  (GPT-4o-mini, GPT-4o, etc.)
  - Anthropic  (Claude 3.5 Haiku, Sonnet, etc.)
  - Google Gemini  (Gemini 2.0 Flash, 1.5 Pro, etc.)

All API calls are async (threaded) so they never block the keyboard hook.
"""

import json
import threading
import urllib.request
import urllib.error
from typing import Callable, Optional

# ╔═══════════════════════════════════════════════════════════════╗
# ║ Provider definitions                                         ║
# ╚═══════════════════════════════════════════════════════════════╝

PROVIDERS = {
    'openai': {
        'name': 'OpenAI',
        'models': ['gpt-4o-mini', 'gpt-4o', 'gpt-4-turbo'],
        'default': 'gpt-4o-mini',
        'url': 'https://api.openai.com/v1/chat/completions',
    },
    'anthropic': {
        'name': 'Anthropic',
        'models': ['claude-3-5-haiku-latest', 'claude-3-5-sonnet-latest', 'claude-sonnet-4-20250514'],
        'default': 'claude-3-5-haiku-latest',
        'url': 'https://api.anthropic.com/v1/messages',
    },
    'gemini': {
        'name': 'Google Gemini',
        'models': ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro'],
        'default': 'gemini-2.0-flash',
        'url': 'https://generativelanguage.googleapis.com/v1beta/models',
    },
}

SYSTEM_PROMPT = (
    "You are a spelling, grammar, and phrasing correction assistant.\n"
    "Fix any spelling, grammar, punctuation, and phrasing errors in the text.\n"
    "Preserve the original language — if the text is in Hebrew, respond in Hebrew; "
    "if in English, respond in English.\n"
    "Return ONLY the corrected text, nothing else — no explanations, no quotes, "
    "no markdown formatting.\n"
    "If the text is already correct, return it unchanged."
)

# ╔═══════════════════════════════════════════════════════════════╗
# ║ Public config (set by engine/UI)                             ║
# ╚═══════════════════════════════════════════════════════════════╝

GRAMMAR_ENABLED = False
GRAMMAR_PROVIDER = 'openai'   # 'openai' | 'anthropic' | 'gemini'
GRAMMAR_API_KEY = ''
GRAMMAR_MODEL = ''            # empty → use provider default
GRAMMAR_CALLBACK = None       # type: callable | None
                              # Called with (original, corrected, error) from bg thread


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Public API                                                   ║
# ╚═══════════════════════════════════════════════════════════════╝

def correct_text_async(text: str, callback: Optional[Callable] = None):
    """Send text to LLM for grammar/phrasing correction (async).

    callback(original_text, corrected_text, error_string_or_None)
    """
    cb = callback or GRAMMAR_CALLBACK
    if not text or not text.strip():
        if cb:
            cb(text, text, None)
        return

    if not GRAMMAR_ENABLED:
        if cb:
            cb(text, text, 'Grammar check is disabled')
        return

    if not GRAMMAR_API_KEY:
        if cb:
            cb(text, None, 'No API key configured — set one in Settings → Spell & Grammar')
        return

    provider_info = PROVIDERS.get(GRAMMAR_PROVIDER)
    if not provider_info:
        if cb:
            cb(text, None, f'Unknown provider: {GRAMMAR_PROVIDER}')
        return

    model = GRAMMAR_MODEL or provider_info['default']

    threading.Thread(
        target=_do_correct,
        args=(text, GRAMMAR_PROVIDER, GRAMMAR_API_KEY, model, provider_info, cb),
        name='grammar-llm',
        daemon=True,
    ).start()


def correct_text_sync(text: str, timeout: float = 15.0) -> tuple:
    """Synchronous version — returns (corrected_text, error_or_None)."""
    if not text or not text.strip():
        return (text, None)
    if not GRAMMAR_ENABLED or not GRAMMAR_API_KEY:
        return (None, 'Grammar not configured')
    provider_info = PROVIDERS.get(GRAMMAR_PROVIDER)
    if not provider_info:
        return (None, f'Unknown provider: {GRAMMAR_PROVIDER}')
    model = GRAMMAR_MODEL or provider_info['default']
    try:
        result = _call_provider(text, GRAMMAR_PROVIDER, GRAMMAR_API_KEY, model, provider_info, timeout)
        return (result, None)
    except Exception as e:
        return (None, str(e))


def get_provider_models(provider: str) -> list:
    """Return list of model names for a provider."""
    info = PROVIDERS.get(provider)
    return info['models'] if info else []


def get_default_model(provider: str) -> str:
    """Return the default model for a provider."""
    info = PROVIDERS.get(provider)
    return info['default'] if info else ''


# ╔═══════════════════════════════════════════════════════════════╗
# ║ Internal                                                     ║
# ╚═══════════════════════════════════════════════════════════════╝

def _do_correct(text, provider, api_key, model, provider_info, callback):
    """Background thread target."""
    try:
        result = _call_provider(text, provider, api_key, model, provider_info, 15.0)
        if callback:
            callback(text, result, None)
    except Exception as e:
        if callback:
            callback(text, None, str(e))


def _call_provider(text, provider, api_key, model, provider_info, timeout):
    if provider == 'openai':
        return _call_openai(text, api_key, model, provider_info['url'], timeout)
    elif provider == 'anthropic':
        return _call_anthropic(text, api_key, model, provider_info['url'], timeout)
    elif provider == 'gemini':
        return _call_gemini(text, api_key, model, provider_info['url'], timeout)
    else:
        raise ValueError(f'Unknown provider: {provider}')


# ── OpenAI ────────────────────────────────────────────────

def _call_openai(text, api_key, model, url, timeout):
    payload = {
        'model': model,
        'messages': [
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user', 'content': text},
        ],
        'temperature': 0.1,
        'max_tokens': max(256, len(text) * 3),
    }
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}',
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = json.loads(resp.read().decode('utf-8'))
    return body['choices'][0]['message']['content'].strip()


# ── Anthropic ─────────────────────────────────────────────

def _call_anthropic(text, api_key, model, url, timeout):
    payload = {
        'model': model,
        'max_tokens': max(256, len(text) * 3),
        'system': SYSTEM_PROMPT,
        'messages': [
            {'role': 'user', 'content': text},
        ],
    }
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            'Content-Type': 'application/json',
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = json.loads(resp.read().decode('utf-8'))
    return body['content'][0]['text'].strip()


# ── Gemini ────────────────────────────────────────────────

def _call_gemini(text, api_key, model, base_url, timeout):
    url = f'{base_url}/{model}:generateContent?key={api_key}'
    payload = {
        'contents': [{'parts': [{'text': f'{SYSTEM_PROMPT}\n\n{text}'}]}],
        'generationConfig': {
            'temperature': 0.1,
            'maxOutputTokens': max(256, len(text) * 3),
        },
    }
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=data,
        headers={'Content-Type': 'application/json'},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = json.loads(resp.read().decode('utf-8'))
    return body['candidates'][0]['content']['parts'][0]['text'].strip()
