"""Thread-local log/progress hooks, so two engines never share state."""
import threading
from contextlib import contextmanager

_local = threading.local()


def set_reporter(log_fn=None, progress_fn=None):
    _local.log = log_fn or print
    _local.progress = progress_fn or (lambda pct, text='': None)


def log(msg=''):
    getattr(_local, 'log', print)(msg)


def progress(pct, text=''):
    fn = getattr(_local, 'progress', None)
    if fn:
        fn(int(max(0, min(100, pct))), text)


@contextmanager
def progress_range(lo, hi):
    """Map an engine's 0-100 progress onto lo..hi of the overall bar (used by the one-click pipeline)."""
    base = getattr(_local, 'progress', None)
    if base:
        _local.progress = lambda pct, text='': base(int(lo + (hi - lo) * pct / 100), text)
    try:
        yield
    finally:
        _local.progress = base


class UserError(Exception):
    """An error whose message is safe and meaningful to show to the user."""
