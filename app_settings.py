"""
Per-user app settings (theme, last-used folders, recent runs) kept in a small JSON file.
The app never depends on it: a missing / broken / read-only file just means defaults are used.

Location: %APPDATA%\\CAMS Commission Automation\\settings.json  (override with env CAMS_SETTINGS_PATH).
"""
import json
import os
import threading

PATH = os.environ.get('CAMS_SETTINGS_PATH') or os.path.join(
    os.environ.get('APPDATA') or os.path.expanduser('~'), 'CAMS Commission Automation', 'settings.json')
MAX_HISTORY = 25

_lock = threading.Lock()
_data = None


def _load():
    global _data
    if _data is None:
        try:
            with open(PATH, encoding='utf-8') as f:
                _data = json.load(f)
            if not isinstance(_data, dict):
                _data = {}
        except (OSError, ValueError):
            _data = {}
    return _data


def get(key, default=None):
    with _lock:
        return _load().get(key, default)


def put(key, value):
    with _lock:
        _load()[key] = value
        _save()


def _save():
    try:
        os.makedirs(os.path.dirname(PATH), exist_ok=True)
        tmp = PATH + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(_data, f, indent=1, ensure_ascii=False)
        os.replace(tmp, PATH)
    except OSError:
        pass


def add_history(entry):
    history = [h for h in get('history', []) if isinstance(h, dict)]
    put('history', [entry] + history[:MAX_HISTORY - 1])
