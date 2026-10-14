"""
Excel output for Step 2 — sheet naming, workbook planning, and the streaming writers.
All workbooks are written with XlsxWriter in constant-memory mode.
"""
import math
import os
import re
from datetime import date, datetime, time, timedelta

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

import config as C

# ───────────────────────── sheet names ─────────────────────────
_BAD = re.compile(r'[\[\]:*?/\\\x00-\x1f]')
_RESERVED = {'history'}


def sanitize_sheet_name(name, fallback='Investor'):
    s = _BAD.sub(' ', str(name or ''))
    s = ' '.join(s.split()).strip("' ")
    if not s or s.casefold() in _RESERVED:
        s = fallback if not s else s + '_'
    return s[:31].rstrip("' ") or fallback


class SheetNamer:
    """Gives every worksheet a valid, unique (case-insensitive) name, <= 31 chars."""

    def __init__(self, reserved=()):
        self.used = {r.casefold() for r in reserved}

    def make(self, base, suffix=''):
        base = sanitize_sheet_name(base)
        k = 1
        while True:
            tag = (' ' + suffix if suffix else '') + ('' if k == 1 else f' ({k})')
            name = base[:31 - len(tag)].rstrip("' ") + tag
            if name.casefold() not in self.used:
                self.used.add(name.casefold())
                return name
            k += 1


def abort_workbook(wb):
    """After a failure: close and delete XlsxWriter's per-sheet temp files so nothing is left on disk."""
    if wb is None:
        return
    try:
        for ws in wb.worksheets():
            fh, fn = getattr(ws, 'row_data_fh', None), getattr(ws, 'row_data_filename', None)
            try:
                if fh:
                    fh.close()
            except Exception:
                pass
            try:
                if fn and os.path.exists(fn):
                    os.remove(fn)
            except Exception:
                pass
    except Exception:
        pass


# ───────────────────────── file names (separate-file mode) ─────────────────────────
_BAD_FILE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WIN_RESERVED = {'con', 'prn', 'aux', 'nul'} | {f'com{i}' for i in range(1, 10)} | {f'lpt{i}' for i in range(1, 10)}


def sanitize_file_name(name, fallback='Investor'):
    s = _BAD_FILE.sub(' ', str(name or ''))
    s = ' '.join(s.split()).strip(' .')
    if not s:
        s = fallback
    if s.split('.')[0].casefold() in _WIN_RESERVED:
        s += '_'
    return s


class FileNamer:
    """Valid, unique (case-insensitive) Windows file names that keep the full path under ~240 chars."""

    def __init__(self, folder):
        self.max_len = max(20, 235 - len(os.path.abspath(folder)) - len('.xlsx') - 12)
        self.used = set()

    def make(self, base):
        base = sanitize_file_name(base)[:self.max_len].rstrip(' .') or 'Investor'
        k = 1
        while True:
            tag = '' if k == 1 else f' ({k})'
            name = base[:self.max_len - len(tag)].rstrip(' .') + tag
            if name.casefold() not in self.used:
                self.used.add(name.casefold())
                return name
            k += 1


# ───────────────────────── brokerage calc (single source of truth) ─────────────────────────
def brokerage_value(prev_rate, cur_rate):
    """Previous BRKAGE_RAT − Current BRKAGE_RAT, or None if either is unavailable."""
    if prev_rate is None or cur_rate is None:
        return None
    return round(prev_rate - cur_rate, 10)


# ───────────────────────── planning ─────────────────────────
def plan_workbooks(groups, n_cols, reserved=3, extra=0):
    """Deterministic bin-packing in the given (sorted) group order.
    A group is never split across workbooks. Returns list of lists of groups."""
    RES = reserved                                     # sheets kept free for Index / File Index / Excluded Log
    books, cur, sheets, cells = [], [], RES, 0
    for g in groups:
        if not g.n_sheets:
            continue
        g_cells = g.count * n_cols
        if cur and (sheets + g.n_sheets + extra > C.MAX_SHEETS_PER_WORKBOOK or cells + g_cells > C.MAX_CELLS_PER_WORKBOOK):
            books.append(cur)
            cur, sheets, cells = [], RES, 0
        cur.append(g)
        sheets += g.n_sheets + extra
        cells += g_cells
    if cur or not books:
        books.append(cur)
    return books


def part_path(base_path, k):
    if k == 1:
        return base_path
    stem, ext = os.path.splitext(base_path)
    return f'{stem}_Part{k}{ext or ".xlsx"}'


# ───────────────────────── formats ─────────────────────────
class Formats:
    def __init__(self, wb):
        base = dict(font_name='Arial', font_size=9, align='left')
        red = dict(bg_color=C.NEG_FILL)
        self.hdr = wb.add_format(dict(bold=True, font_color='#FFFFFF', bg_color='#1F4E79', font_name='Arial',
                                      font_size=10, align='center', valign='vcenter', text_wrap=True, border=1,
                                      border_color='#FFFFFF'))
        self.hdr_new = wb.add_format(dict(bold=True, font_color='#FFFFFF', bg_color='#2E75B6', font_name='Arial',
                                          font_size=10, align='center', valign='vcenter', text_wrap=True,
                                          border=1, border_color='#FFFFFF'))
        self.cf_red = wb.add_format(dict(bg_color=C.NEG_FILL))
        self.cf_inc_font = wb.add_format(dict(font_color='#B71C1C', bold=True))
        self.cf_dec_font = wb.add_format(dict(font_color='#1F5F99', bold=True))
        self.rc = wb.add_format(dict(base, align='center', bold=True))
        self.link = wb.add_format(dict(base, font_color='#0563C1', underline=1))
        self.plain = wb.add_format(base)
        self.title = wb.add_format(dict(bold=True, font_name='Arial', font_size=12, font_color='#1F4E79'))
        self.cells = {}
        for is_red in (False, True):
            extra = red if is_red else {}
            self.cells[('text', is_red)] = wb.add_format(dict(base, **extra))
            self.cells[('num', is_red)] = wb.add_format(dict(base, **extra))
            self.cells[('rate', is_red)] = wb.add_format(dict(base, num_format='0.0000##', **extra))
            self.cells[('brk', is_red)] = wb.add_format(dict(base, num_format='+0.0000;-0.0000;0.0000', **extra))
            self.cells[('date', is_red)] = wb.add_format(dict(base, num_format='dd-mm-yyyy', **extra))
            self.cells[('datetime', is_red)] = wb.add_format(dict(base, num_format='dd-mm-yyyy hh:mm:ss', **extra))
        self.idx_text = wb.add_format(dict(base, border=1, border_color='#D0D0D0'))
        self.idx_num = wb.add_format(dict(base, border=1, border_color='#D0D0D0', num_format='#,##0'))
        self.idx_bad = wb.add_format(dict(base, border=1, border_color='#D0D0D0', bg_color='#FFEB9C'))
        self.idx_link = wb.add_format(dict(base, border=1, border_color='#D0D0D0',
                                           font_color='#0563C1', underline=1))


# ───────────────────────── index sheets ─────────────────────────
INDEX_HEADERS = ['Investor Name', 'AMC', 'Scheme Code', 'Folio Number', 'Transaction Type', 'Worksheet Name',
                 'Matching Transactions Count', 'Negative Brokerage Difference Count', 'Matching Status',
                 'Workbook File', 'Notes']


def write_index(ws, fm, groups, this_book_idx, book_files, separate=False):
    widths = [30, 8, 14, 18, 16, 34, 16, 18, 34, 34, 70]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    for i, h in enumerate(INDEX_HEADERS):
        ws.write_string(0, i, h, fm.hdr)
    ws.set_row(0, 32)
    for r, g in enumerate(groups, start=1):
        ws.write_string(r, 0, g.investor, fm.idx_text)
        ws.write_string(r, 1, g.amc, fm.idx_text)
        ws.write_string(r, 2, g.scheme, fm.idx_text)
        ws.write_string(r, 3, g.folio, fm.idx_text)
        ws.write_string(r, 4, g.txn, fm.idx_text)
        if separate:
            ws.write_string(r, 5, '; '.join(g.sheets), fm.idx_text)
        elif g.sheets and g.book == this_book_idx:
            ws.write_url(r, 5, f"internal:'{g.sheets[0]}'!A1", fm.idx_link, '; '.join(g.sheets))
        else:
            ws.write_string(r, 5, '; '.join(g.sheets), fm.idx_text)
        ws.write_number(r, 6, g.count, fm.idx_num)
        ws.write_number(r, 7, g.neg, fm.idx_num)
        ws.write_string(r, 8, g.status, fm.idx_bad if not g.count else fm.idx_text)
        if separate:
            if g.count and g.file:
                ws.write_url(r, 9, 'external:' + g.file, fm.idx_link, g.file)      # clickable link to the investor's file
            else:
                ws.write_string(r, 9, '', fm.idx_text)
        else:
            ws.write_string(r, 9, os.path.basename(book_files[g.book]) if g.count else '', fm.idx_text)
        ws.write_string(r, 10, g.notes, fm.idx_text)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(groups), 1), len(INDEX_HEADERS) - 1)


def write_file_index(ws, fm, books, book_files):
    for i, w in enumerate([44, 22, 22, 22]):
        ws.set_column(i, i, w)
    for i, h in enumerate(['Workbook File', 'Investor Groups', 'Worksheets (incl. continuation)', 'Source rows']):
        ws.write_string(0, i, h, fm.hdr)
    for r, (grp, f) in enumerate(zip(books, book_files), start=1):
        ws.write_string(r, 0, os.path.basename(f), fm.idx_text)
        ws.write_number(r, 1, len(grp), fm.idx_num)
        ws.write_number(r, 2, sum(g.n_sheets for g in grp), fm.idx_num)
        ws.write_number(r, 3, sum(g.count for g in grp), fm.idx_num)
    ws.freeze_panes(1, 0)


def write_excluded_log(ws, fm, excluded):
    for i, w in enumerate([16, 18, 16, 10, 34, 14, 80]):
        ws.set_column(i, i, w)
    for i, h in enumerate(['Scheme Code', 'Folio Number', 'Transaction Type', 'Source AMC',
                           'Source Investor', 'Rows Excluded', 'Reason']):
        ws.write_string(0, i, h, fm.hdr)
    for r, (k, (cnt, reason)) in enumerate(sorted(excluded.items()), start=1):
        for c, v in enumerate(k):
            ws.write_string(r, c, v, fm.idx_text)
        ws.write_number(r, 5, cnt, fm.idx_num)
        ws.write_string(r, 6, reason, fm.idx_text)
    ws.freeze_panes(1, 0)


# ───────────────────────── investor sheet writer ─────────────────────────
def _write_cell(ws, r, c, v, fm, red):
    t = type(v)
    if v is None:
        ws.write_blank(r, c, None, fm.cells[('text', red)])
    elif t is str:
        ws.write_string(r, c, v, fm.cells[('text', red)])
    elif t is float or t is int:
        if t is float and (v != v or v in (float('inf'), float('-inf'))):
            ws.write_blank(r, c, None, fm.cells[('text', red)])
        else:
            ws.write_number(r, c, v, fm.cells[('num', red)])
    elif t is bool:
        ws.write_boolean(r, c, v, fm.cells[('text', red)])
    elif t is datetime:
        ws.write_datetime(r, c, v, fm.cells[('datetime' if (v.hour or v.minute or v.second) else 'date', red)])
    elif t is date:
        ws.write_datetime(r, c, datetime(v.year, v.month, v.day), fm.cells[('date', red)])
    else:
        ws.write_string(r, c, str(v), fm.cells[('text', red)])


def write_group_sheets(wb, fm, group, store, header, date_idx, brk_idx):
    """Write the (optional) Rate Summary sheet and all data worksheets of one investor/group.
    Returns the number of negative-Brokerage rows (for verification)."""
    add_rc = C.ADD_RATE_CHANGE_COLUMN
    extra = [C.BROKERAGE_COL] + ([C.RATE_CHANGE_COL] if add_rc else [])
    out_header = header[:brk_idx + 1] + extra + header[brk_idx + 1:]
    n_out = len(out_header)
    brk_out = brk_idx + 1
    rc_out = brk_idx + 2 if add_rc else None
    date_out = date_idx if date_idx <= brk_idx else date_idx + len(extra)
    lo, hi = min(date_out, brk_out), max(date_out, brk_out)       # red band: TRADE_DATE ... Brokerage
    rate_letter = xl_col_to_name(brk_idx)
    brk_letter = xl_col_to_name(brk_out)
    cap = C.STEP2_ROWS_PER_SHEET
    n_in = len(header)
    note = ('Brokerage = Previous BRKAGE_RAT - Current BRKAGE_RAT (trade-date order).\n'
            'MINUS (-)  = rate INCREASED.\nPLUS (+)  = rate DECREASED.\n0 = unchanged.')

    ws_sum = wb.add_worksheet(group.summary) if getattr(group, 'summary', '') else None   # sits before the data
    points = []
    sheet_ix = -1
    ws = None
    row_in_sheet = 0
    prev_rate = None
    prev_valid = False
    negatives = 0

    def open_sheet(ix):
        nonlocal ws, row_in_sheet
        ws = wb.add_worksheet(group.sheets[ix])
        n_data = min(cap, group.count - ix * cap)
        for c, h in enumerate(out_header):
            ws.write_string(0, c, h, fm.hdr_new if c in (brk_out, rc_out) else fm.hdr)
            if h == C.DATE_COL:
                w = 13
            elif h == C.INV_COL:
                w = 30
            elif h in (C.BROKERAGE_COL,):
                w = 13
            elif h == C.RATE_CHANGE_COL:
                w = 17
            else:
                w = max(12, min(len(h) + 3, 28))
            ws.set_column(c, c, w)
        ws.set_row(0, 30)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(n_data, 1), n_out - 1)
        ws.write_comment(0, brk_out, note, {'x_scale': 2.4, 'y_scale': 1.6, 'font_size': 9})
        if n_data:
            ws.conditional_format(1, lo, n_data, hi, {          # live rule, only over TRADE_DATE..Brokerage
                'type': 'formula',
                'criteria': f'=AND(ISNUMBER(${brk_letter}2),${brk_letter}2<0)',
                'format': fm.cf_red})
            if add_rc:
                ws.conditional_format(1, rc_out, n_data, rc_out, {
                    'type': 'formula', 'criteria': f'=AND(ISNUMBER(${brk_letter}2),${brk_letter}2<0)',
                    'format': fm.cf_inc_font})
                ws.conditional_format(1, rc_out, n_data, rc_out, {
                    'type': 'formula', 'criteria': f'=AND(ISNUMBER(${brk_letter}2),${brk_letter}2>0)',
                    'format': fm.cf_dec_font})
        row_in_sheet = 0

    for dkey, dok, rate, rstat, row in store.iter_rows(group.gid):
        if row_in_sheet == 0 or row_in_sheet >= cap:
            sheet_ix += 1
            open_sheet(sheet_ix)
            first_in_sheet = True
        else:
            first_in_sheet = False
        r = row_in_sheet + 1                       # 0-based row index (header = 0)
        excel_row = r + 1

        if dok and prev_valid:
            val = brokerage_value(prev_rate, rate)
        else:
            val = None
        red = val is not None and val < 0
        if red:
            negatives += 1
        if dok and rstat == 0:
            points.append((dkey, rate))

        row = row + [None] * (n_in - len(row)) if len(row) < n_in else row[:n_in]
        c_out = 0
        for c_in in range(n_in):
            v = row[c_in]
            is_red = red and lo <= c_out <= hi
            if c_in == date_idx and dok:
                fmt_kind = 'datetime' if dkey != int(dkey) else 'date'
                ws.write_number(r, c_out, dkey, fm.cells[(fmt_kind, is_red)])
            elif c_in == brk_idx:
                if rstat == 0:
                    ws.write_number(r, c_out, rate, fm.cells[('rate', is_red)])
                else:
                    _write_cell(ws, r, c_out, v, fm, is_red)
            else:
                _write_cell(ws, r, c_out, v, fm, is_red)
            c_out += 1
            if c_in == brk_idx:                    # insert Brokerage (and Rate Change) right after BRKAGE_RAT
                is_red_b = red and lo <= c_out <= hi
                fmt_b = fm.cells[('brk', is_red_b)]
                if val is None:
                    ws.write_blank(r, c_out, None, fmt_b)
                elif first_in_sheet:               # continuation sheet: no previous row on this sheet
                    ws.write_number(r, c_out, val, fmt_b)
                else:
                    f = (f'=IF(AND(ISNUMBER({rate_letter}{excel_row - 1}),ISNUMBER({rate_letter}{excel_row})),'
                         f'ROUND({rate_letter}{excel_row - 1}-{rate_letter}{excel_row},10),"")')
                    ws.write_formula(r, c_out, f, fmt_b, val)
                c_out += 1
                if add_rc:
                    cached = '' if val is None else ('▲ Increased' if val < 0 else '▼ Decreased' if val > 0 else 'Same')
                    f = (f'=IF({brk_letter}{excel_row}="","",IF({brk_letter}{excel_row}<0,"▲ Increased",'
                         f'IF({brk_letter}{excel_row}>0,"▼ Decreased","Same")))')
                    ws.write_formula(r, c_out, f, fm.rc, cached)
                    c_out += 1

        if dok:
            prev_rate = rate if rstat == 0 else None
            prev_valid = True
        else:
            prev_rate, prev_valid = None, False
        row_in_sheet += 1

    if ws_sum is not None:
        from rate_summary import write_rate_summary
        write_rate_summary(wb, ws_sum, group, points)
    return negatives
