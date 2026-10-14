"""
CAMS Commission Automation System — desktop GUI (CustomTkinter).

  * Sidebar: Dashboard | One-Click | Step 1 | Step 2 | Update Sheets | Help   (+ Light / Dark / System theme)
  * Every task page = settings card on the left (files, options, Run) and results + activity log on the right,
    sized to fit a 1366 x 768 laptop screen.
  * Each task runs in its own background thread; workers talk to the window ONLY through a queue,
    so the interface never freezes and Tk is only touched from the main thread.
  * Theme, last-used folders and recent runs are remembered per Windows user (app_settings.py).
  * If customtkinter is not installed the plain ttk window (gui_classic.py) is used automatically.

Keyboard: Ctrl+1 … Ctrl+6 switch pages, Ctrl+Enter runs the task on the open page.
"""
import datetime
import os
import queue
import subprocess
import sys
import threading
import time
import traceback
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError:                                      # graceful fallback
    ctk = None

import app_settings
import config as C
import reporting

# ───────────── palette  (light, dark) ─────────────
ACCENT = ('#2563EB', '#3B82F6')
ACCENT_HOVER = ('#1D4ED8', '#2563EB')
ACCENT_SOFT = ('#EAF1FF', '#16213D')
BG = ('#F4F6FA', '#0B1120')
CARD = ('#FFFFFF', '#111827')
BORDER = ('#E2E8F0', '#1F2937')
FIELD = ('#F8FAFC', '#0F172A')
TEXT = ('#0F172A', '#E5E7EB')
MUTED = ('#64748B', '#94A3B8')
HOVER = ('#EEF2F7', '#1E293B')
TRACK = ('#E2E8F0', '#1F2937')
SUCCESS, SUCCESS_SOFT = ('#15803D', '#4ADE80'), ('#E9F9EF', '#0B2A1A')
WARN, WARN_SOFT = ('#B45309', '#FBBF24'), ('#FFF6E5', '#2E2208')
DANGER, DANGER_SOFT = ('#B91C1C', '#F87171'), ('#FDECEC', '#341214')
NEUTRAL_SOFT = ('#EEF2F7', '#1E293B')

SIDEBAR = '#0F172A'
SIDE_TEXT = '#94A3B8'
SIDE_ACTIVE = '#1E293B'
SIDE_HOVER = '#162033'
SIDE_SECTION = '#526079'

LOG_TAGS = {
    'Light': {'hdr': '#1D4ED8', 'ok': '#15803D', 'err': '#B91C1C', 'warn': '#B45309', 'dim': '#94A3B8'},
    'Dark': {'hdr': '#7FB2FF', 'ok': '#5BD38A', 'err': '#FF8080', 'warn': '#F6C35B', 'dim': '#6B7A90'},
}

# Segoe MDL2 Assets glyphs (Windows 10+); the second value is the fallback when that font is missing.
ICONS = {
    'home': ('', '⌂'), 'bolt': ('', '⚡'), 'search': ('', '①'), 'people': ('', '②'),
    'sync': ('', '↻'), 'help': ('', '?'), 'check': ('', '✓'), 'error': ('', '✗'),
    'doc': ('', '▤'), 'folder': ('', '▣'), 'clock': ('', '◷'), 'chart': ('', '▥'),
}
_icon_font_ok = False


def icon(name):
    glyph, fallback = ICONS[name]
    return glyph if _icon_font_ok else fallback


def _font(size=13, weight='normal'):
    return ctk.CTkFont(family='Segoe UI', size=size, weight=weight)


def _semibold(size=13):
    return ctk.CTkFont(family='Segoe UI Semibold', size=size)


def _icon_font(size=16):
    return ctk.CTkFont(family='Segoe MDL2 Assets' if _icon_font_ok else 'Segoe UI Symbol', size=size)


def _open_path(p):
    try:
        if os.name == 'nt':
            os.startfile(p)                                   # noqa
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', p])
        else:
            subprocess.Popen(['xdg-open', p])
    except Exception as e:
        messagebox.showerror('Cannot open', str(e))


def _short(path, n=64):
    return path if len(path) <= n else '…' + path[-(n - 1):]


def _size(n):
    for unit in ('bytes', 'KB', 'MB', 'GB'):
        if n < 1024 or unit == 'GB':
            return f'{n:,.0f} {unit}' if unit == 'bytes' else f'{n:,.1f} {unit}'
        n /= 1024


def _duration(sec):
    sec = int(sec)
    return f'{sec // 60}:{sec % 60:02d}' if sec >= 60 else f'{sec} s'


def _card(master, **kw):
    return ctk.CTkFrame(master, fg_color=CARD, corner_radius=12, border_width=1, border_color=BORDER, **kw)


def _ghost_button(master, text, command, width=120, height=32, **kw):
    return ctk.CTkButton(master, text=text, command=command, width=width, height=height, corner_radius=8,
                         font=_semibold(12), fg_color='transparent', border_width=1, border_color=BORDER,
                         text_color=TEXT, hover_color=HOVER, **kw)


def _page_header(master, crumb, title, subtitle):
    """Breadcrumb + title + subtitle on the left; returns (frame, right-hand slot)."""
    head = ctk.CTkFrame(master, fg_color='transparent')
    head.columnconfigure(0, weight=1)
    ctk.CTkLabel(head, text=crumb.upper(), font=_semibold(11), text_color=ACCENT, anchor='w').grid(
        row=0, column=0, sticky='w')
    ctk.CTkLabel(head, text=title, font=_font(22, 'bold'), text_color=TEXT, anchor='w').grid(
        row=1, column=0, sticky='w')
    ctk.CTkLabel(head, text=subtitle, font=_font(12), text_color=MUTED, anchor='w', justify='left',
                 wraplength=820).grid(row=2, column=0, sticky='w', pady=(1, 0))
    slot = ctk.CTkFrame(head, fg_color='transparent')
    slot.grid(row=0, column=1, rowspan=3, sticky='ne')
    return head, slot


# ═════════════════════════════════════════════════════════════════════
class StatusPill(ctk.CTkLabel if ctk else object):
    STYLES = {'ready': ('Ready', NEUTRAL_SOFT, MUTED), 'running': ('Running…', ACCENT_SOFT, ACCENT),
              'done': ('Completed', SUCCESS_SOFT, SUCCESS), 'error': ('Failed', DANGER_SOFT, DANGER)}

    def __init__(self, master):
        super().__init__(master, text='', height=26, corner_radius=13, font=_semibold(12))
        self.set('ready')

    def set(self, state):
        text, bg, fg = self.STYLES[state]
        self.configure(text=f'  ●  {text}   ', fg_color=bg, text_color=fg)


class FileField(ctk.CTkFrame if ctk else object):
    """Label row (+ optional action), path entry + Browse button, and a live check line under it.
    kind: 'file' | 'folder' (inputs)   'save' | 'outdir' (outputs)"""

    def __init__(self, master, label, kind, browse, optional=False):
        super().__init__(master, fg_color='transparent')
        self.kind, self.optional = kind, optional
        self.var = ctk.StringVar()
        self.columnconfigure(0, weight=1)
        top = ctk.CTkFrame(self, fg_color='transparent')
        top.grid(row=0, column=0, columnspan=2, sticky='ew')
        top.columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text=label, font=_semibold(13), text_color=TEXT, anchor='w').grid(row=0, column=0, sticky='w')
        self.action_slot = top
        self.entry = ctk.CTkEntry(self, textvariable=self.var, height=36, corner_radius=8, border_width=1,
                                  fg_color=FIELD, border_color=BORDER, text_color=TEXT, font=_font(12))
        self.entry.grid(row=1, column=0, sticky='ew', pady=(4, 0))
        self.button = ctk.CTkButton(self, text={'file': 'Browse…', 'folder': 'Browse…', 'save': 'Save as…',
                                                'outdir': 'Choose…'}[kind],
                                    width=96, height=36, corner_radius=8, font=_semibold(12),
                                    fg_color=ACCENT_SOFT, hover_color=HOVER, text_color=ACCENT, command=browse)
        self.button.grid(row=1, column=1, padx=(8, 0), pady=(4, 0))
        self.meta = ctk.CTkLabel(self, text='', font=_font(11), text_color=MUTED, anchor='w', height=18)
        self.meta.grid(row=2, column=0, columnspan=2, sticky='w', pady=(2, 0))
        self.var.trace_add('write', lambda *_: self.refresh())
        self.refresh()

    def add_action(self, text, command):
        b = ctk.CTkButton(self.action_slot, text=text, command=command, height=22, width=10, corner_radius=6,
                          font=_semibold(11), fg_color='transparent', hover_color=HOVER, text_color=ACCENT)
        b.grid(row=0, column=1, sticky='e')
        return b

    def set_state(self, st):
        self.entry.configure(state=st)
        self.button.configure(state=st)

    def refresh(self):
        text, color = self._describe(self.var.get().strip())
        self.meta.configure(text=text, text_color=color)

    def _describe(self, p):
        if not p:
            return ('Optional — leave blank to use the first file' if self.optional else
                    'No folder selected' if self.kind in ('folder', 'outdir') else 'No file selected'), MUTED
        try:
            if self.kind == 'file':
                if not os.path.isfile(p):
                    return '⚠  File not found', WARN
                st = os.stat(p)
                when = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%d %b %Y, %H:%M')
                return f'✓  {os.path.basename(p)}   ·   {_size(st.st_size)}   ·   modified {when}', SUCCESS
            if self.kind == 'folder':
                if not os.path.isdir(p):
                    return '⚠  Folder not found', WARN
                n = sum(1 for f in os.listdir(p) if f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'))
                return f'✓  Folder found   ·   {n} Excel file{"s" if n != 1 else ""} inside', (SUCCESS if n else WARN)
            if self.kind == 'save':
                if os.path.isfile(p):
                    return '⚠  A file with this name exists and will be replaced', WARN
                d = os.path.dirname(p) or '.'
                return (f'Will be saved in  {_short(d, 70)}', MUTED) if os.path.isdir(d) else \
                       ('⚠  That folder does not exist', WARN)
            return ('Results are saved inside this folder', MUTED) if os.path.isdir(p) else \
                   ('The folder will be created', MUTED)
        except OSError as e:
            return f'⚠  {e}', WARN


# ═════════════════════════════════════════════════════════════════════
class StepPage(ctk.CTkFrame if ctk else object):
    """One task: files + options + Run on the left; status, result tiles and live log on the right."""

    def __init__(self, master, app, key, crumb, title, subtitle, inputs, out_label, out_name,
                 run_label, worker_fn, stats_fn, step_no, optional=(), out_is_folder=False, option_label=None,
                 in_is_folder=False, default_out=None):
        super().__init__(master, fg_color='transparent')
        self.app, self.key, self.title = app, key, title
        self.worker_fn, self.stats_fn, self.out_name, self.step_no = worker_fn, stats_fn, out_name, step_no
        self.optional, self.out_is_folder = set(optional), out_is_folder
        self.in_is_folder, self.default_out = in_is_folder, default_out
        self.run_label = run_label
        self.q = queue.Queue()
        self.running = False
        self.t0 = 0.0
        self.last_output = ''
        self.columnconfigure(0, weight=5, uniform='cols')
        self.columnconfigure(1, weight=6, uniform='cols')
        self.rowconfigure(1, weight=1)

        head, slot = _page_header(self, crumb, title, subtitle)
        head.grid(row=0, column=0, columnspan=2, sticky='ew', padx=24, pady=(18, 12))
        self.pill = StatusPill(slot)
        self.pill.pack(anchor='e', pady=(4, 0))

        # ── left: settings card
        card = _card(self)
        card.grid(row=1, column=0, sticky='nsew', padx=(24, 8), pady=(0, 18))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        self.card = card
        form = ctk.CTkFrame(card, fg_color='transparent')
        form.grid(row=0, column=0, sticky='new', padx=20, pady=(16, 0))
        form.columnconfigure(0, weight=1)
        ctk.CTkLabel(form, text='INPUT', font=_semibold(11), text_color=MUTED, anchor='w').pack(fill='x')
        self.fields = []
        for n, label in enumerate(inputs):
            f = FileField(form, label, 'folder' if in_is_folder else 'file',
                          lambda n=n, l=label: self._pick_in(n, l), optional=n in self.optional)
            f.pack(fill='x', pady=(6, 4))
            self.fields.append(f)
        self.in_vars = [f.var for f in self.fields]
        ctk.CTkFrame(form, height=1, fg_color=BORDER).pack(fill='x', pady=(8, 10))
        ctk.CTkLabel(form, text='OUTPUT', font=_semibold(11), text_color=MUTED, anchor='w').pack(fill='x')
        self.out_field = FileField(form, ('Output folder' if out_is_folder else 'Output Excel') + f'  ·  {out_label}',
                                   'outdir' if out_is_folder else 'save', self._pick_out)
        self.out_field.pack(fill='x', pady=(6, 4))
        self.out_var = self.out_field.var
        self._lockable = [*self.fields, self.out_field]
        self.opt_var = None
        if option_label:
            self.opt_var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(form, text=option_label, variable=self.opt_var, font=_font(12), text_color=TEXT,
                                 checkbox_width=20, checkbox_height=20, corner_radius=5, border_width=2,
                                 fg_color=ACCENT, hover_color=ACCENT_HOVER)
            cb.pack(anchor='w', pady=(6, 0))
            self._lockable.append(cb)

        run = ctk.CTkFrame(card, fg_color='transparent')
        run.grid(row=2, column=0, sticky='ew', padx=20, pady=(12, 18))
        run.columnconfigure(0, weight=1)
        self.status = ctk.StringVar(value='Ready — select the files to begin.')
        self.elapsed = ctk.StringVar()
        ctk.CTkLabel(run, textvariable=self.status, font=_font(12), text_color=MUTED, anchor='w').grid(
            row=0, column=0, sticky='w')
        self.pct = ctk.CTkLabel(run, text='', font=_semibold(12), text_color=TEXT)
        self.pct.grid(row=0, column=1, sticky='e')
        self.bar = ctk.CTkProgressBar(run, height=8, corner_radius=4, progress_color=ACCENT, fg_color=TRACK)
        self.bar.set(0)
        self.bar.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(6, 12))
        self.run_btn = ctk.CTkButton(run, text=run_label, height=42, corner_radius=9, font=_semibold(14),
                                     fg_color=ACCENT, hover_color=ACCENT_HOVER, command=self.start)
        self.run_btn.grid(row=2, column=0, columnspan=2, sticky='ew')
        ctk.CTkLabel(run, text='Ctrl + Enter', font=_font(10), text_color=MUTED).grid(
            row=3, column=0, columnspan=2, pady=(4, 0))

        # ── right: results + log
        right = ctk.CTkFrame(self, fg_color='transparent')
        right.grid(row=1, column=1, sticky='nsew', padx=(8, 24), pady=(0, 18))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        res = _card(right)
        res.grid(row=0, column=0, sticky='ew')
        res.columnconfigure(0, weight=1)
        rh = ctk.CTkFrame(res, fg_color='transparent')
        rh.grid(row=0, column=0, sticky='ew', padx=18, pady=(14, 0))
        rh.columnconfigure(0, weight=1)
        ctk.CTkLabel(rh, text='Results', font=_semibold(14), text_color=TEXT, anchor='w').grid(row=0, column=0, sticky='w')
        ctk.CTkLabel(rh, textvariable=self.elapsed, font=_font(12), text_color=MUTED).grid(row=0, column=1, sticky='e')
        self.banner = ctk.CTkLabel(res, text='', font=_font(12), corner_radius=8, anchor='w', justify='left',
                                   wraplength=560)
        self.banner.grid(row=1, column=0, sticky='ew', padx=18, pady=(10, 0))
        self.banner.grid_remove()
        self.stats = ctk.CTkFrame(res, fg_color='transparent')
        self.stats.grid(row=2, column=0, sticky='ew', padx=18, pady=(10, 0))
        self.stats.grid_remove()
        self.empty = ctk.CTkLabel(res, text='Numbers and output files appear here when the task finishes.',
                                  font=_font(12), text_color=MUTED, anchor='w')
        self.empty.grid(row=3, column=0, sticky='ew', padx=18, pady=(8, 0))
        acts = ctk.CTkFrame(res, fg_color='transparent')
        acts.grid(row=4, column=0, sticky='ew', padx=18, pady=(12, 16))
        self.open_btn = _ghost_button(acts, 'Open output folder', width=150, state='disabled',
                                      command=lambda: _open_path(self.last_output if self.out_is_folder else
                                                                 os.path.dirname(os.path.abspath(self.last_output))))
        self.open_btn.pack(side='left')
        self.file_btn = _ghost_button(acts, 'Open output file', width=140, state='disabled',
                                      command=lambda: _open_path(self.last_output))
        if not out_is_folder:
            self.file_btn.pack(side='left', padx=(8, 0))

        logc = _card(right)
        logc.grid(row=1, column=0, sticky='nsew', pady=(14, 0))
        logc.columnconfigure(0, weight=1)
        logc.rowconfigure(1, weight=1)
        lh = ctk.CTkFrame(logc, fg_color='transparent')
        lh.grid(row=0, column=0, sticky='ew', padx=18, pady=(12, 6))
        lh.columnconfigure(0, weight=1)
        ctk.CTkLabel(lh, text='Activity log', font=_semibold(14), text_color=TEXT, anchor='w').grid(row=0, column=0, sticky='w')
        _ghost_button(lh, 'Copy', self._copy_log, width=64, height=26).grid(row=0, column=1, padx=(0, 6))
        _ghost_button(lh, 'Clear', self._clear_log, width=64, height=26).grid(row=0, column=2)
        self._log_bold = ctk.CTkFont(family='Consolas', size=12, weight='bold')
        self.log = ctk.CTkTextbox(logc, fg_color=FIELD, text_color=TEXT, corner_radius=8, border_width=1,
                                  border_color=BORDER, font=ctk.CTkFont(family='Consolas', size=12), wrap='none',
                                  activate_scrollbars=True)
        self.log.grid(row=1, column=0, sticky='nsew', padx=12, pady=(0, 12))
        self.apply_theme(ctk.get_appearance_mode())
        self.log.configure(state='disabled')

    # ───────────── helpers used by tests / extras
    def get_progress(self):
        return self.bar.get() * 100

    def make_output(self, p):
        if self.default_out:
            return self.default_out(p)
        return os.path.dirname(p) if self.out_is_folder else os.path.join(os.path.dirname(p), self.out_name)

    def apply_theme(self, mode):
        tags = LOG_TAGS.get(mode, LOG_TAGS['Light'])
        for tag, color in tags.items():
            self.log._textbox.tag_config(tag, foreground=color)
        self.log._textbox.tag_config('hdr', font=self._log_bold)

    # ───────────── pickers
    def _last_dir(self):
        return app_settings.get('dirs', {}).get(self.key) or os.path.expanduser('~')

    def _remember_dir(self, p):
        d = p if os.path.isdir(p) else os.path.dirname(p)
        dirs = dict(app_settings.get('dirs', {}))
        dirs[self.key] = d
        app_settings.put('dirs', dirs)

    def _pick_in(self, n, label):
        if self.in_is_folder:
            p = filedialog.askdirectory(title=label, initialdir=self._last_dir())
        else:
            p = filedialog.askopenfilename(title=label, initialdir=self._last_dir(),
                                           filetypes=[('Excel files', '*.xlsx *.xlsm'), ('All files', '*.*')])
        if p:
            p = os.path.normpath(p)
            self.in_vars[n].set(p)
            self._remember_dir(p)
            if not self.out_var.get().strip():
                self.out_var.set(os.path.normpath(self.make_output(p)))

    def _pick_out(self):
        if self.out_is_folder:
            p = filedialog.askdirectory(title='Choose output folder', initialdir=self.out_var.get() or self._last_dir())
            if p:
                self.out_var.set(os.path.normpath(p))
            return
        cur = self.out_var.get().strip()
        p = filedialog.asksaveasfilename(title='Save output Excel as', defaultextension='.xlsx',
                                         initialdir=os.path.dirname(cur) if cur else self._last_dir(),
                                         initialfile=os.path.basename(cur or self.out_name),
                                         filetypes=[('Excel workbook', '*.xlsx')])
        if p:
            self.out_var.set(os.path.normpath(p))

    # ───────────── log
    def _log(self, text):
        t = text.strip()
        tag = ('hdr' if t.startswith('[Step') or t.startswith('══') else 'err' if t.startswith('ERROR') or '✗' in t else
               'warn' if '⚠' in t else 'ok' if ('✓' in t or t.startswith('Finished')) else
               'dim' if t.startswith('...') else '')
        self.log.configure(state='normal')
        self.log._textbox.insert('end', text + '\n', tag or ())
        self.log._textbox.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    def _copy_log(self):
        self.clipboard_clear()
        self.clipboard_append(self.log.get('1.0', 'end').strip())
        self.app.toast('info', 'Log copied', 'The activity log is on the clipboard.')

    # ───────────── result area
    def _set_banner(self, kind, text):
        fg, bg = {'done': (SUCCESS, SUCCESS_SOFT), 'error': (DANGER, DANGER_SOFT)}[kind]
        self.banner.configure(text=f'  {text}', text_color=fg, fg_color=bg)
        self.banner.grid()

    def _show_stats(self, res):
        for w in self.stats.winfo_children():
            w.destroy()
        cards = self.stats_fn(res) if res else []
        cols = 3 if len(cards) > 4 else max(len(cards), 1)
        for i in range(cols):
            self.stats.columnconfigure(i, weight=1, uniform='s')
        for i, (value, caption) in enumerate(cards):
            c = ctk.CTkFrame(self.stats, fg_color=FIELD, corner_radius=10, border_width=1, border_color=BORDER)
            c.grid(row=i // cols, column=i % cols, sticky='ew', padx=(0 if i % cols == 0 else 8, 0),
                   pady=(0 if i < cols else 8, 0))
            ctk.CTkLabel(c, text=value, font=_font(20, 'bold'), text_color=TEXT, anchor='w').pack(
                fill='x', padx=12, pady=(8, 0))
            ctk.CTkLabel(c, text=caption, font=_font(11), text_color=MUTED, anchor='w').pack(
                fill='x', padx=12, pady=(0, 8))
        if cards:
            self.empty.grid_remove()
            self.stats.grid()

    # ───────────── run
    def _set_running(self, flag):
        self.running = flag
        st = 'disabled' if flag else 'normal'
        self.run_btn.configure(state=st, text='Running…' if flag else self.run_label)
        for w in self._lockable:
            (w.set_state if isinstance(w, FileField) else lambda s, w=w: w.configure(state=s))(st)
        if flag:
            self.open_btn.configure(state='disabled')
            self.file_btn.configure(state='disabled')
        self.app.refresh_nav()

    def start(self):
        if self.running:
            return
        ins = [v.get().strip() for v in self.in_vars]
        out = self.out_var.get().strip()
        need = [p for i, p in enumerate(ins) if i not in self.optional or p]
        exists = os.path.isdir if self.in_is_folder else os.path.isfile
        if not all(need) or not all(exists(p) for p in need):
            messagebox.showwarning('Input needed', 'Please select the folder.' if self.in_is_folder else
                                   'Please select all the required input Excel files.')
            return
        if not out:
            messagebox.showwarning('Output needed', 'Please choose where to save the output.')
            return
        if self.opt_var is not None:
            C.SEPARATE_FILES_PER_INVESTOR = bool(self.opt_var.get())
        self._clear_log()
        for w in self.stats.winfo_children():
            w.destroy()
        self.stats.grid_remove()
        self.banner.grid_remove()
        self.empty.configure(text='Working… results will appear here.')
        self.empty.grid()
        self.bar.set(0)
        self.pct.configure(text='0 %')
        self.t0 = time.time()
        self.last_output = out
        self.status.set('Starting…')
        self.pill.set('running')
        self._set_running(True)
        threading.Thread(target=self._worker, args=(ins, out), daemon=True).start()

    def _worker(self, ins, out):
        try:
            reporting.set_reporter(lambda m: self.q.put(('log', m)),
                                   lambda p, t='': self.q.put(('prog', p, t)))
            res = self.worker_fn(*ins, out)
            self.q.put(('done', out, res))
        except reporting.UserError as e:
            self.q.put(('error', str(e), ''))
        except MemoryError:
            self.q.put(('error', 'Ran out of memory. Close other programs and use 64-bit Python.', ''))
        except PermissionError as e:
            self.q.put(('error', f'Permission denied: {e}. Is the file open in Excel?', ''))
        except Exception as e:
            self.q.put(('error', f'Unexpected error: {e}', traceback.format_exc()))

    def pump(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == 'log':
                    self._log(m[1])
                elif m[0] == 'prog':
                    self.bar.set(m[1] / 100)
                    self.pct.configure(text=f'{m[1]} %')
                    if m[2]:
                        self.status.set(m[2])
                elif m[0] == 'done':
                    took = time.time() - self.t0
                    self.bar.set(1)
                    self.pct.configure(text='100 %')
                    self.status.set('✓ Finished successfully.')
                    self.elapsed.set(f'Took {_duration(took)}')
                    self.pill.set('done')
                    self._set_running(False)
                    self.open_btn.configure(state='normal')
                    self.file_btn.configure(state='disabled' if self.out_is_folder else 'normal')
                    self._show_stats(m[2])
                    for f in (*self.fields, self.out_field):
                        f.refresh()
                    self._set_banner('done', f'✓  Completed in {_duration(took)}  —  saved to  {_short(m[1], 70)}')
                    self.app.on_done(self, m[1], m[2], took)
                elif m[0] == 'error':
                    took = time.time() - self.t0
                    self.status.set('✗ Failed — see the activity log.')
                    self.pill.set('error')
                    self._log('\nERROR: ' + m[1])
                    if m[2]:
                        self._log(m[2])
                    self._set_running(False)
                    self.empty.configure(text='No results — the task stopped with an error.')
                    self._set_banner('error', f'✗  {m[1]}')
                    self.app.on_error(self, m[1], took)
                    messagebox.showerror('Error', m[1])
        except queue.Empty:
            pass
        if self.running:
            self.elapsed.set(f'Elapsed  {_duration(time.time() - self.t0)}')
        self.after(150, self.pump)


# ═════════════════════════════════════════════════════════════════════
TASKS = [  # key, icon, nav label, card title, card text
    ('both', 'bolt', 'One-Click', 'One-Click Automation',
     'Run the Discrepancy Checker and the Investor-wise sheets in one go.'),
    ('s1', 'search', 'Discrepancy Checker', 'Discrepancy Checker',
     'Find temporary and intra-quarter brokerage-rate drops in the CAMS file.'),
    ('s2', 'people', 'Investor-wise Sheets', 'Investor-wise Sheets',
     'One Excel per investor with Brokerage, Rate Change and Rate Summary.'),
    ('update', 'sync', 'Update Sheets', 'Update Existing Sheets',
     'Add Rate Change + Rate Summary to investor files you already have.'),
]
TASK_NAMES = {k: title for k, _, _, title, _ in TASKS}


class HomePage(ctk.CTkFrame if ctk else object):
    def __init__(self, master, app):
        super().__init__(master, fg_color='transparent')
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        hour = datetime.datetime.now().hour
        greet = 'Good morning' if hour < 12 else 'Good afternoon' if hour < 17 else 'Good evening'
        head, slot = _page_header(self, 'Dashboard', greet,
                                  datetime.date.today().strftime('%A, %d %B %Y') + '   ·   choose a task to get started.')
        head.grid(row=0, column=0, sticky='ew', padx=24, pady=(18, 12))
        ctk.CTkLabel(slot, text='  ●  Works offline — your data stays on this PC   ', font=_semibold(12),
                     text_color=SUCCESS, fg_color=SUCCESS_SOFT, corner_radius=13, height=26).pack(pady=(4, 0))

        grid = ctk.CTkFrame(self, fg_color='transparent')
        grid.grid(row=1, column=0, sticky='ew', padx=24)
        for i, (key, ic, _, title, text) in enumerate(TASKS):
            grid.columnconfigure(i, weight=1, uniform='t')
            self._task_card(grid, key, ic, title, text, recommended=(key == 'both')).grid(
                row=0, column=i, sticky='nsew', padx=(0 if i == 0 else 12, 0))

        rec = _card(self)
        rec.grid(row=2, column=0, sticky='nsew', padx=24, pady=(16, 18))
        rec.columnconfigure(0, weight=1)
        rec.rowconfigure(1, weight=1)
        rh = ctk.CTkFrame(rec, fg_color='transparent')
        rh.grid(row=0, column=0, sticky='ew', padx=18, pady=(14, 6))
        rh.columnconfigure(0, weight=1)
        ctk.CTkLabel(rh, text='Recent activity', font=_semibold(14), text_color=TEXT, anchor='w').grid(row=0, column=0, sticky='w')
        _ghost_button(rh, 'Clear history', self._clear, width=104, height=26).grid(row=0, column=1)
        self.rows = ctk.CTkScrollableFrame(rec, fg_color='transparent', scrollbar_button_color=BORDER,
                                           scrollbar_button_hover_color=MUTED)
        self.rows.grid(row=1, column=0, sticky='nsew', padx=8, pady=(0, 10))
        self.rows.columnconfigure(1, weight=1)
        self.refresh()

    def _task_card(self, master, key, ic, title, text, recommended=False):
        c = _card(master)
        badge = ctk.CTkLabel(c, text=icon(ic), font=_icon_font(18), text_color=ACCENT, fg_color=ACCENT_SOFT,
                             corner_radius=10, width=40, height=40)
        badge.pack(anchor='w', padx=16, pady=(16, 10))
        if recommended:
            ctk.CTkLabel(c, text=' RECOMMENDED ', font=_semibold(9), text_color=ACCENT, fg_color=ACCENT_SOFT,
                         corner_radius=6, height=18).place(relx=1, x=-14, y=16, anchor='ne')
        t = ctk.CTkLabel(c, text=title, font=_semibold(14), text_color=TEXT, anchor='w')
        t.pack(fill='x', padx=16)
        d = ctk.CTkLabel(c, text=text, font=_font(12), text_color=MUTED, anchor='w', justify='left', wraplength=210)
        d.pack(fill='x', padx=16, pady=(2, 8))
        o = ctk.CTkLabel(c, text='Open  →', font=_semibold(12), text_color=ACCENT, anchor='w')
        o.pack(fill='x', padx=16, pady=(0, 14))
        for w in (c, badge, t, d, o):
            w.bind('<Button-1>', lambda e, k=key: self.app.show(k))
            w.bind('<Enter>', lambda e: c.configure(border_color=ACCENT))
            w.bind('<Leave>', lambda e: c.configure(border_color=BORDER))
        return c

    def _clear(self):
        app_settings.put('history', [])
        self.refresh()

    def refresh(self):
        for w in self.rows.winfo_children():
            w.destroy()
        history = app_settings.get('history', [])
        if not history:
            ctk.CTkLabel(self.rows, text='Nothing yet — runs you complete will be listed here, with a quick link '
                                         'to their output.', font=_font(12), text_color=MUTED, anchor='w').grid(
                row=0, column=0, columnspan=4, sticky='w', padx=10, pady=8)
            return
        for r, h in enumerate(history[:12]):
            ok = h.get('status') == 'done'
            ctk.CTkLabel(self.rows, text=icon('check' if ok else 'error'), font=_icon_font(13),
                         text_color=SUCCESS if ok else DANGER, width=26).grid(row=r, column=0, padx=(10, 6), pady=5)
            info = ctk.CTkFrame(self.rows, fg_color='transparent')
            info.grid(row=r, column=1, sticky='ew')
            ctk.CTkLabel(info, text=TASK_NAMES.get(h.get('page'), h.get('page', '')), font=_semibold(12),
                         text_color=TEXT, anchor='w').pack(fill='x')
            detail = _short(h.get('output', ''), 90) if ok else h.get('message', '')[:110]
            ctk.CTkLabel(info, text=detail, font=_font(11), text_color=MUTED, anchor='w').pack(fill='x')
            ctk.CTkLabel(self.rows, text=f"{h.get('when', '')}   ·   {_duration(h.get('seconds', 0))}",
                         font=_font(11), text_color=MUTED).grid(row=r, column=2, padx=12)
            out = h.get('output', '')
            target = out if os.path.isdir(out) else os.path.dirname(out)
            b = _ghost_button(self.rows, 'Open', lambda t=target: _open_path(t), width=60, height=26)
            b.grid(row=r, column=3, padx=(0, 10))
            if not (ok and target and os.path.isdir(target)):
                b.configure(state='disabled')


# ═════════════════════════════════════════════════════════════════════
HELP_TEXT = [
    ('Workflow', None),
    ('One-Click', 'Pick the CAMS Excel + an output folder and press Run Everything. Step 1 then Step 2 run '
     'automatically. (If your Step 2 source file is a different file, select it in the optional box.)'),
    ('Step 1 — Discrepancy Checker', 'Select the original CAMS commission Excel. The app finds temporary and intra-quarter '
     'brokerage-rate discrepancies and writes Discrepancy_Result.xlsx (Raw Data highlighted + Discrepancy Report).'),
    ('Step 2 — Investor-wise Sheets', 'Select Discrepancy_Result.xlsx and the large original transaction Excel. '
     'One sheet per investor / Scheme + Folio + Transaction Type is created from the SOURCE file, sorted oldest → newest, '
     'with a Brokerage column (previous − current BRKAGE_RAT).'),
    ('Update Existing Sheets', 'Already have investor Excel files (for example your final investor sheets)? Put them in one '
     'folder, select it and press Update. New copies with the Rate Change column and the Rate Summary sheet are saved in '
     'a different folder; your originals stay untouched.'),
    ('Both steps are independent', 'You can close the app after Step 1 and run Step 2 later.'),
    ('Colour guide', None),
    ('Light red  #FFCCCC', 'Step 1: discrepancy rows.  Step 2: rows where Brokerage is negative, highlighted from '
     'TRADE_DATE to the Brokerage column.'),
    ('Tips', None),
    ('Keep output closed', 'Close the output Excel before running, otherwise Windows blocks saving.'),
    ('Big files', 'Files above 60 MB are read in low-memory streaming mode; it is slower but safe. '
     'All limits are in config.py.'),
    ('Keyboard shortcuts', 'Ctrl+1 … Ctrl+6 switch pages.   Ctrl+Enter runs the task on the open page.'),
]


class HelpPage(ctk.CTkFrame if ctk else object):
    def __init__(self, master):
        super().__init__(master, fg_color='transparent')
        head, _ = _page_header(self, 'Support', 'How it works', 'Everything the app does, in plain words.')
        head.pack(fill='x', padx=24, pady=(18, 12))
        box = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=12, border_width=1, border_color=BORDER,
                                     scrollbar_button_color=BORDER, scrollbar_button_hover_color=MUTED)
        box.pack(fill='both', expand=True, padx=24, pady=(0, 18))
        for head, body in HELP_TEXT:
            if body is None:
                ctk.CTkLabel(box, text=head.upper(), font=_semibold(11), text_color=ACCENT).pack(
                    anchor='w', padx=16, pady=(16, 2))
                continue
            row = ctk.CTkFrame(box, fg_color='transparent')
            row.pack(fill='x', padx=16, pady=(6, 2))
            if head.startswith('Light red'):
                ctk.CTkFrame(row, width=14, height=14, corner_radius=3, fg_color='#FFCCCC', border_width=1,
                             border_color='#E8A0A0').pack(side='left', padx=(0, 8))
            ctk.CTkLabel(row, text=head, font=_semibold(13), text_color=TEXT).pack(side='left')
            ctk.CTkLabel(box, text=body, font=_font(12), text_color=MUTED, justify='left',
                         wraplength=820).pack(anchor='w', padx=16, pady=(0, 4))


# ═════════════════════════════════════════════════════════════════════
class NavItem(ctk.CTkFrame if ctk else object):
    def __init__(self, master, glyph, text, command, shortcut=''):
        super().__init__(master, fg_color='transparent', corner_radius=8, height=38)
        self.pack_propagate(False)
        self.active = False
        self.bar = ctk.CTkFrame(self, width=3, height=18, corner_radius=2, fg_color=SIDEBAR)
        self.bar.place(x=0, rely=0.5, anchor='w')
        self.ic = ctk.CTkLabel(self, text=glyph, font=_icon_font(14), text_color=SIDE_TEXT, width=22)
        self.ic.pack(side='left', padx=(14, 10))
        self.lb = ctk.CTkLabel(self, text=text, font=_font(13), text_color=SIDE_TEXT, anchor='w')
        self.lb.pack(side='left', fill='x', expand=True)
        self.dot = ctk.CTkLabel(self, text='', font=_font(10), text_color='#60A5FA', width=14)
        self.dot.pack(side='right', padx=(0, 10))
        for w in (self, self.ic, self.lb, self.dot):
            w.bind('<Button-1>', lambda e: command())
            w.bind('<Enter>', lambda e: self._hover(True))
            w.bind('<Leave>', lambda e: self._hover(False))

    def _hover(self, on):
        if not self.active:
            self.configure(fg_color=SIDE_HOVER if on else 'transparent')

    def set_active(self, on):
        self.active = on
        fg = 'white' if on else SIDE_TEXT
        self.configure(fg_color=SIDE_ACTIVE if on else 'transparent')
        self.bar.configure(fg_color='#3B82F6' if on else SIDEBAR)
        self.ic.configure(text_color='#60A5FA' if on else SIDE_TEXT)
        self.lb.configure(text_color=fg, font=_semibold(13) if on else _font(13))

    def set_busy(self, on):
        self.dot.configure(text='●' if on else '')


class App(ctk.CTk if ctk else object):
    def __init__(self):
        global _icon_font_ok
        import discrepancy_engine
        import investor_engine
        import pipeline
        import sheet_updater
        theme = app_settings.get('theme', 'Light')
        if theme not in ('Light', 'Dark', 'System'):
            theme = 'Light'
        ctk.set_appearance_mode(theme.lower())
        ctk.set_default_color_theme('blue')
        super().__init__(fg_color=BG)
        try:
            import tkinter.font as tkfont
            _icon_font_ok = 'Segoe MDL2 Assets' in tkfont.families(self)
        except Exception:
            _icon_font_ok = False
        self.title(C.APP_NAME)
        self._set_icon()
        self.geometry('1240x720')
        self.minsize(1040, 620)
        if os.name == 'nt':
            self.after(0, lambda: self.state('zoomed'))
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.last_step1_output = ''
        self.current = None

        # ── sidebar
        side = ctk.CTkFrame(self, width=228, corner_radius=0, fg_color=SIDEBAR)
        side.grid(row=0, column=0, sticky='ns')
        side.pack_propagate(False)
        brand = ctk.CTkFrame(side, fg_color='transparent')
        brand.pack(fill='x', padx=18, pady=(22, 22))
        ctk.CTkLabel(brand, text=icon('chart'), font=_icon_font(18), text_color='white', fg_color='#2563EB',
                     corner_radius=9, width=38, height=38).pack(side='left')
        names = ctk.CTkFrame(brand, fg_color='transparent')
        names.pack(side='left', padx=(10, 0))
        ctk.CTkLabel(names, text='CAMS Automation', font=_font(15, 'bold'), text_color='white', anchor='w',
                     height=20).pack(anchor='w')
        ctk.CTkLabel(names, text='Commission suite', font=_font(11), text_color=SIDE_TEXT, anchor='w',
                     height=16).pack(anchor='w')

        self.nav = {}
        groups = [('', [('home', 'home', 'Dashboard')]),
                  ('Workflow', [(k, ic, lbl) for k, ic, lbl, _, _ in TASKS if k != 'update']),
                  ('Tools', [('update', 'sync', 'Update Sheets')]),
                  ('Support', [('help', 'help', 'How it works')])]
        for title, items in groups:
            if title:
                ctk.CTkLabel(side, text=title.upper(), font=_semibold(10), text_color=SIDE_SECTION, anchor='w').pack(
                    fill='x', padx=26, pady=(14, 4))
            for key, ic, label in items:
                item = NavItem(side, icon(ic), label, lambda k=key: self.show(k))
                item.pack(fill='x', padx=12, pady=1)
                self.nav[key] = item

        foot = ctk.CTkFrame(side, fg_color='transparent')
        foot.pack(side='bottom', fill='x', padx=18, pady=(0, 18))
        ctk.CTkLabel(foot, text='APPEARANCE', font=_semibold(10), text_color=SIDE_SECTION, anchor='w').pack(fill='x')
        self.theme_var = ctk.StringVar(value=theme)
        self.mode = ctk.CTkSegmentedButton(foot, values=['Light', 'Dark', 'System'], variable=self.theme_var,
                                           command=self.set_theme, height=28, font=_font(11),
                                           fg_color='#1E293B', selected_color='#2563EB', selected_hover_color='#1D4ED8',
                                           unselected_color='#1E293B', unselected_hover_color='#273449',
                                           text_color='white')
        self.mode.pack(fill='x', pady=(6, 12))
        ctk.CTkLabel(foot, text='Offline · data stays on this PC', font=_font(10), text_color=SIDE_SECTION,
                     anchor='w').pack(fill='x')

        # ── pages
        host = ctk.CTkFrame(self, fg_color='transparent', corner_radius=0)
        host.grid(row=0, column=1, sticky='nsew')
        host.columnconfigure(0, weight=1)
        host.rowconfigure(0, weight=1)
        self.host = host
        self.pages = {
            'home': HomePage(host, self),
            'both': StepPage(
                host, self, 'both', 'Workflow  ·  Both steps', 'One-Click Automation',
                'Select the CAMS Excel and an output folder. Step 1 and Step 2 run one after the other and the results '
                'are saved in that folder (one Excel file per investor inside Investor_Excels).',
                ['Original CAMS Commission Excel', 'Source Transaction Excel  (optional)'],
                'both result files are saved here', 'folder', 'Run Everything', pipeline.run_both,
                lambda r: [(f"{r['rows']:,}", 'Rows checked'), (f"{r['report_rows']:,}", 'Discrepancy entries'),
                           (f"{r['matched']:,}", 'Investor sheets'), (f"{r['transactions']:,}", 'Transactions'),
                           (f"{len(r['files'])}", 'Excel files created')], 0, optional=(1,), out_is_folder=True,
                option_label='Create a separate Excel file for each investor  (recommended)'),
            's1': StepPage(
                host, self, 's1', 'Workflow  ·  Step 1 of 2', 'Discrepancy Checker',
                'Finds temporary and intra-quarter brokerage-rate discrepancies in the original CAMS commission file.',
                ['Original CAMS Commission Excel'], 'Discrepancy_Result.xlsx', 'Discrepancy_Result.xlsx',
                'Generate Discrepancy Report', discrepancy_engine.run,
                lambda r: [(f"{r['rows']:,}", 'Rows checked'), (f"{r['temporary_rows']:,}", 'Temporary rows'),
                           (f"{r['intra_rows']:,}", 'Intra-quarter rows'), (f"{r['report_rows']:,}", 'Report entries')], 1),
            's2': StepPage(
                host, self, 's2', 'Workflow  ·  Step 2 of 2', 'Investor-wise Sheets',
                'Creates one Excel file per investor (in Investor_Excels) from the large source file, sorted oldest → '
                'newest, with the Brokerage column and negative rows highlighted. The output file is the investor index.',
                ['Discrepancy Report Excel  (from Step 1)', 'Source CAMS Transaction Excel  (large file)'],
                'investor index', 'Investor_Wise_Final.xlsx',
                'Generate Investor-wise Sheets', investor_engine.run,
                lambda r: [(f"{r['groups']:,}", 'Combinations'), (f"{r['matched']:,}", 'Matched'),
                           (f"{r['transactions']:,}", 'Transactions'), (f"{len(r['files'])}", 'Excel files created')], 2,
                option_label='Create a separate Excel file for each investor  (recommended)'),
            'update': StepPage(
                host, self, 'update', 'Tools', 'Update Existing Sheets',
                'Select a folder of investor Excel files. You get NEW copies with the Rate Change column and the '
                'Rate Summary sheet. Your original files are never changed.',
                ['Folder with the investor Excel files'], 'updated copies go here', 'Updated_Sheets',
                'Update All Sheets', sheet_updater.run_update,
                lambda r: [(f"{r['files']}", 'Files updated'), (f"{r['sheets']}", 'Investor sheets'),
                           (f"{r['skipped']}", 'Skipped'), (f"{r['failed']}", 'Failed')], 4,
                out_is_folder=True, in_is_folder=True,
                default_out=lambda p: os.path.join(p, 'Updated_Sheets')),
            'help': HelpPage(host),
        }
        for p in self.pages.values():
            p.grid(row=0, column=0, sticky='nsew')
        # shortcut on Step 2: reuse the file Step 1 just produced
        self.use_btn = self.pages['s2'].fields[0].add_action('Use Step 1 output', self._use_step1_output)
        for p in self.step_pages():
            p.pump()

        # ── toast (in-window notification, bottom-right)
        self._toast = _card(self, width=360)
        self._toast_icon = ctk.CTkLabel(self._toast, text='', font=_icon_font(16), width=30)
        self._toast_icon.grid(row=0, column=0, rowspan=2, padx=(14, 6), pady=12, sticky='n')
        self._toast_title = ctk.CTkLabel(self._toast, text='', font=_semibold(13), text_color=TEXT, anchor='w')
        self._toast_title.grid(row=0, column=1, sticky='w', padx=(0, 16), pady=(10, 0))
        self._toast_msg = ctk.CTkLabel(self._toast, text='', font=_font(12), text_color=MUTED, anchor='w',
                                       justify='left', wraplength=300)
        self._toast_msg.grid(row=1, column=1, sticky='w', padx=(0, 16), pady=(0, 12))
        self._toast.bind('<Button-1>', lambda e: self._toast.place_forget())
        self._toast_job = None

        for i, key in enumerate(('home', 'both', 's1', 's2', 'update', 'help'), 1):
            self.bind_all(f'<Control-Key-{i}>', lambda e, k=key: self.show(k))
        self.bind_all('<Control-Return>', lambda e: self._run_current())
        self.protocol('WM_DELETE_WINDOW', self._on_close)
        self.show('home')

    # ───────────── window chrome
    def _set_icon(self):
        try:
            import app_icon
            ico = app_icon.ensure_ico()
            if os.name == 'nt':
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('CAMS.CommissionAutomation')
                if ico:
                    self.iconbitmap(default=ico)
                    self.after(250, lambda: self.iconbitmap(ico))      # CTk resets the icon shortly after start
        except Exception:
            pass

    def _flash(self):
        try:
            if os.name == 'nt' and self.focus_displayof() is None:
                import ctypes
                ctypes.windll.user32.FlashWindow(int(self.wm_frame(), 16), True)
        except Exception:
            pass

    def _on_close(self):
        busy = [p.title for p in self.step_pages() if p.running]
        if busy and not messagebox.askyesno('Task still running',
                                            f'"{busy[0]}" is still running. If you close now its output will be '
                                            'incomplete.\n\nClose anyway?', icon='warning'):
            return
        self.destroy()

    # ───────────── pages
    def step_pages(self):
        return [p for p in self.pages.values() if isinstance(p, StepPage)]

    def show(self, key):
        self.current = key
        page = self.pages[key]
        if isinstance(page, HomePage):
            page.refresh()
        page.tkraise()
        for k, b in self.nav.items():
            b.set_active(k == key)

    def refresh_nav(self):
        for k, p in self.pages.items():
            if isinstance(p, StepPage) and k in self.nav:
                self.nav[k].set_busy(p.running)

    def _run_current(self):
        page = self.pages.get(self.current)
        if isinstance(page, StepPage):
            page.start()

    def set_theme(self, value):
        self.theme_var.set(value)
        ctk.set_appearance_mode(value.lower())
        app_settings.put('theme', value)
        self.after(50, lambda: [p.apply_theme(ctk.get_appearance_mode()) for p in self.step_pages()])

    # ───────────── notifications
    def toast(self, kind, title, msg, seconds=6):
        glyph, color = {'done': ('check', SUCCESS), 'error': ('error', DANGER), 'info': ('check', ACCENT)}[kind]
        self._toast_icon.configure(text=icon(glyph), text_color=color)
        self._toast_title.configure(text=title)
        self._toast_msg.configure(text=msg)
        self._toast.place(relx=1, rely=1, x=-24, y=-24, anchor='se')
        self._toast.lift()
        if self._toast_job:
            self.after_cancel(self._toast_job)
        self._toast_job = self.after(seconds * 1000, self._toast.place_forget)

    def _record(self, page, status, out, took, message=''):
        app_settings.add_history({'page': page.key, 'status': status, 'output': out, 'seconds': int(took),
                                  'message': message, 'when': datetime.datetime.now().strftime('%d %b, %H:%M')})

    # ───────────── callbacks from StepPage
    def _use_step1_output(self):
        if self.last_step1_output and os.path.isfile(self.last_step1_output):
            self.pages['s2'].in_vars[0].set(self.last_step1_output)
        else:
            messagebox.showinfo('Step 1 output', 'Run Step 1 first, or Browse to an existing Discrepancy_Result.xlsx.')

    def on_done(self, page, out, res, took=0):
        if page.step_no == 1:
            self.last_step1_output = out
            if not self.pages['s2'].in_vars[0].get().strip():
                self.pages['s2'].in_vars[0].set(out)
        self._record(page, 'done', out, took)
        self.toast('done', f'{page.title} finished', f'Completed in {_duration(took)}. Output saved to\n{_short(out, 60)}')
        self._flash()

    def on_error(self, page, message, took=0):
        self._record(page, 'error', page.last_output, took, message)
        self._flash()


def launch():
    if ctk is None:
        import gui_classic
        return gui_classic.launch()
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == '__main__':
    launch()
