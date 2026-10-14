"""
STEP 2 — Investor-wise brokerage automation (engine)

Inputs
  EXCEL A : Discrepancy_Result.xlsx  (only its "Discrepancy Report" sheet is read)
  EXCEL B : the large original CAMS transaction Excel (ALL transactions are taken from here)

Flow
  1. Read the unique (Investor, AMC, Scheme, Folio, Transaction Type) combinations from A.
  2. Stream B once (read-only).  A row is kept only if SCHEME_COD + FOLIO_NO + TRXN_TYPE equal a
     combination exactly (identifiers normalised: spaces trimmed, 123.0 -> 123, leading zeros KEPT).
     AMC_CODE and INV_NAME are used to validate/separate records that share the same triple.
  3. Matched rows go to a SQLite temp DB (disk) tagged with their group.
  4. Per group: sorted TRADE_DATE oldest->newest (stable on original row order),
     Brokerage = previous BRKAGE_RAT − current BRKAGE_RAT inserted right after BRKAGE_RAT,
     negative rows highlighted light red from TRADE_DATE through Brokerage.
  5. Output via XlsxWriter constant-memory; one sheet per group (+ continuation sheets),
     index sheets, automatic multi-workbook split.
"""
import gc
import os
import shutil
import time
from datetime import date, datetime, time as dtime, timedelta

import config as C
from database import TempStore
from excel_reader import norm_amc, norm_id, norm_name, open_sheet
from excel_writer import (FileNamer, Formats, abort_workbook, SheetNamer, brokerage_value, part_path, plan_workbooks,
                          write_excluded_log, write_file_index, write_group_sheets, write_index)
from reporting import UserError, log, progress

_EXCEL_EPOCH = datetime(1899, 12, 30)
_DATE_FORMATS = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d-%m-%Y %H:%M:%S',
                 '%d/%m/%Y %H:%M:%S', '%d-%b-%Y', '%d-%b-%y', '%d-%B-%Y', '%d/%m/%y', '%d-%m-%y',
                 '%Y/%m/%d', '%d %b %Y', '%d.%m.%Y', '%d-%m-%Y %H:%M', '%d/%m/%Y %H:%M']


# ───────────────────────── value parsing ─────────────────────────
class DateParser:
    def __init__(self):
        self._fmts = list(_DATE_FORMATS)

    def __call__(self, v):
        """-> Excel serial (float) or None if missing / invalid."""
        if v is None:
            return None
        if isinstance(v, datetime):
            d = v
        elif isinstance(v, date):
            d = datetime(v.year, v.month, v.day)
        elif isinstance(v, bool):
            return None
        elif isinstance(v, (int, float)):
            if v != v or not (1 <= v < 2958466):
                return None
            return float(v)
        elif isinstance(v, str):
            s = v.strip()
            if not s:
                return None
            d = None
            for i, f in enumerate(self._fmts):
                try:
                    d = datetime.strptime(s, f)
                except ValueError:
                    continue
                if i:                                  # remember the format that worked
                    self._fmts.insert(0, self._fmts.pop(i))
                break
            if d is None:
                return None
        else:
            return None
        if d.year < 1900:
            return None
        return (d - _EXCEL_EPOCH).total_seconds() / 86400.0


def parse_rate(v):
    """-> (rate, status)  status: 0 ok, 1 missing, 2 invalid"""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, 1
    if isinstance(v, bool):
        return None, 2
    if isinstance(v, (int, float)):
        return (None, 2) if v != v or v in (float('inf'), float('-inf')) else (float(v), 0)
    if isinstance(v, str):
        try:
            f = float(v.strip().replace(',', ''))
        except ValueError:
            return None, 2
        return (None, 2) if f != f or f in (float('inf'), float('-inf')) else (f, 0)
    return None, 2


# ───────────────────────── group object ─────────────────────────
class Group:
    __slots__ = ('gid', 'investor', 'amc', 'scheme', 'folio', 'txn', 'amc_n', 'inv_n', 'count', 'neg',
                 'n_sheets', 'sheets', 'book', 'status', 'notes', 'invalid_dates', 'missing_rates',
                 'invalid_rates', 'blank_assigned', 'file', 'summary')

    def __init__(self, gid, investor, amc, scheme, folio, txn):
        self.gid, self.investor, self.amc, self.scheme, self.folio, self.txn = gid, investor, amc, scheme, folio, txn
        self.amc_n, self.inv_n = norm_amc(amc), norm_name(investor)
        self.count = self.neg = self.n_sheets = self.book = 0
        self.sheets, self.status, self.notes = [], '', ''
        self.invalid_dates = self.missing_rates = self.invalid_rates = self.blank_assigned = 0
        self.file = ''
        self.summary = ''


# ───────────────────────── read the discrepancy report ─────────────────────────
def read_combinations(report_path):
    log('\n[Step 2A] Reading discrepancy combinations...')
    rd = open_sheet(report_path, sheet_name=C.REPORT_SHEET_NAME, engine='openpyxl')
    try:
        hdr = rd.header()
        need = {'Investor': None, 'AMC': None, 'Scheme': None, 'Folio': None, 'Transaction Type': None}
        for k in need:
            if k not in hdr:
                raise UserError(f'The "{C.REPORT_SHEET_NAME}" sheet is missing the column "{k}". '
                                'Please select the file produced by Step 1.')
            need[k] = hdr.index(k)
        seen = {}
        for row in rd.rows():
            def g(k):
                i = need[k]
                return row[i] if i < len(row) else None
            if all(v is None for v in row):
                continue
            scheme, folio, txn = norm_id(g('Scheme')), norm_id(g('Folio')), norm_id(g('Transaction Type'))
            if not (scheme and folio and txn):
                continue
            inv, amc = norm_id(g('Investor')), norm_id(g('AMC'))
            key = (scheme, folio, txn, norm_amc(amc), norm_name(inv))
            if key not in seen:
                seen[key] = (inv, amc, scheme, folio, txn)
    finally:
        rd.close()
    combos = sorted(seen.values(), key=lambda t: (t[0].casefold(), t[2], t[3], t[4], t[1]))
    if not combos:
        raise UserError('The Discrepancy Report sheet contains no combinations to process.')
    log(f'         {len(combos):,} unique Scheme + Folio + Transaction Type combinations '
        f'for {len({norm_name(c[0]) for c in combos}):,} investor name(s).')
    return [Group(i, *c) for i, c in enumerate(combos)]


# ───────────────────────── scan the big source ─────────────────────────
REQUIRED_SRC = [C.DATE_COL, C.BROK_COL, C.SCHEME_COL, C.FOLIO_COL, C.TXN_TYPE_COL, C.AMC_COL, C.INV_COL]


def scan_source(source_path, groups, store):
    log('\n[Step 2B] Scanning the source transaction Excel (single pass)...')
    wanted = {}
    for g in groups:
        wanted.setdefault((g.scheme, g.folio, g.txn), []).append(g)

    rd = open_sheet(source_path)
    parse_date = DateParser()
    excluded = {}            # (scheme, folio, txn, amc, inv) -> [count, reason]
    try:
        header = rd.header()
        pos = {}
        for i, h in enumerate(header):
            pos.setdefault(h, i)
        missing = [c for c in REQUIRED_SRC if c not in pos]
        if missing:
            raise UserError(f'The source transaction Excel is missing required column(s): {missing}')
        n_in = len(header)
        i_s, i_f, i_t = pos[C.SCHEME_COL], pos[C.FOLIO_COL], pos[C.TXN_TYPE_COL]
        i_a, i_n, i_d, i_r = pos[C.AMC_COL], pos[C.INV_COL], pos[C.DATE_COL], pos[C.BROK_COL]
        need_len = max(i_s, i_f, i_t, i_a, i_n, i_d, i_r) + 1
        total = rd.total_rows or 0
        scanned = kept = 0
        for seq, row in enumerate(rd.rows(), start=1):
            scanned += 1
            if scanned % 100_000 == 0:
                progress(5 + 50 * min(scanned / total, 1) if total else 30, f'Scanning source… {scanned:,} rows')
                log(f'         ... {scanned:,} rows scanned, {kept:,} matched')
            if len(row) < need_len:
                row = tuple(row) + (None,) * (need_len - len(row))
            combos = wanted.get((norm_id(row[i_s]), norm_id(row[i_f]), norm_id(row[i_t])))
            if combos is None:
                continue
            amc_n, inv_n = norm_amc(row[i_a]), norm_name(row[i_n])
            target = None
            for g in combos:                                   # exact AMC + investor match
                if g.amc_n == amc_n and g.inv_n == inv_n:
                    target = g
                    break
            if target is None and C.STRICT_INVESTOR_VALIDATION:
                # a blank AMC / investor in the source cannot contradict the report -> compatible
                compat = [g for g in combos
                          if (not amc_n or g.amc_n == amc_n) and (not inv_n or g.inv_n == inv_n)]
                if (not amc_n or not inv_n) and len(compat) == 1:
                    target = compat[0]
                    target.blank_assigned += 1
            elif target is None:
                target = combos[0]
            if target is None:
                k = (norm_id(row[i_s]), norm_id(row[i_f]), norm_id(row[i_t]),
                     norm_id(row[i_a]), norm_id(row[i_n]))
                e = excluded.setdefault(k, [0, 'Same Scheme+Folio+Transaction Type but different AMC / investor '
                                               'than the discrepancy report — not mixed in'])
                e[0] += 1
                continue
            rate, rstat = parse_rate(row[i_r])
            r = list(row[:n_in]) if len(row) >= n_in else list(row) + [None] * (n_in - len(row))
            store.add(target.gid, seq, parse_date(row[i_d]), rate, rstat, r)
            kept += 1
    finally:
        rd.close()
    store.finalize()
    log(f'         Source rows scanned: {scanned:,} | matched transactions: {kept:,}')
    if excluded:
        log(f'         ⚠ {sum(v[0] for v in excluded.values()):,} same-key rows belong to a different AMC/investor '
            'and were kept out (see "Excluded Rows Log").')
    return header, pos, {k: (v[0], v[1]) for k, v in excluded.items()}


# ───────────────────────── pass 1: statistics (so the index is complete up-front) ─────────────────────────
def compute_group_stats(groups, store):
    counts = store.group_counts()
    for g in groups:
        n, bad_d, miss_r, inv_r = counts.get(g.gid, (0, 0, 0, 0))
        g.count, g.invalid_dates, g.missing_rates, g.invalid_rates = n, bad_d, miss_r, inv_r
        g.n_sheets = -(-n // C.STEP2_ROWS_PER_SHEET) if n else 0
        neg, prev, prev_valid = 0, None, False
        if n:
            for dok, rate in store.iter_light(g.gid):
                if dok:
                    if prev_valid:
                        v = brokerage_value(prev, rate)
                        if v is not None and v < 0:
                            neg += 1
                    prev, prev_valid = rate, True
                else:
                    prev, prev_valid = None, False
        g.neg = neg
        notes = []
        if n == 0:
            g.status = 'No matching transactions found'
        else:
            g.status = 'Matched' if g.n_sheets == 1 else f'Matched (continued over {g.n_sheets} sheets)'
        if g.invalid_dates:
            notes.append(f'{g.invalid_dates} row(s) with invalid/missing TRADE_DATE placed at the end, no Brokerage')
        if g.missing_rates:
            notes.append(f'{g.missing_rates} row(s) with missing BRKAGE_RAT')
        if g.invalid_rates:
            notes.append(f'{g.invalid_rates} row(s) with invalid BRKAGE_RAT')
        if g.blank_assigned:
            notes.append(f'{g.blank_assigned} row(s) had blank AMC/investor in source (assigned by exact key)')
        if g.n_sheets > 1:
            notes.append('first Brokerage value of each continuation sheet is a static number')
        g.notes = '; '.join(notes)


# ───────────────────────── main entry ─────────────────────────
def _free_space_check(path, source_path):
    folder = os.path.dirname(os.path.abspath(path)) or '.'
    need = int(os.path.getsize(source_path) * 1.5) + 50 * 1_048_576
    free = shutil.disk_usage(folder).free
    if free < need:
        raise UserError(f'Not enough free disk space in "{folder}": about {need // 1_048_576:,} MB needed, '
                        f'{free // 1_048_576:,} MB free.')


def _check_writable(path):
    from discrepancy_engine import check_writable
    check_writable(path)


def run(report_path, source_path, output_path):
    t0 = time.time()
    for p, label in ((report_path, 'Discrepancy report Excel'), (source_path, 'Source CAMS transaction Excel')):
        if not os.path.isfile(p):
            raise UserError(f'{label} not found: {p}')
    if os.path.abspath(output_path) in (os.path.abspath(report_path), os.path.abspath(source_path)):
        raise UserError('Output file must be different from both input files.')
    if not output_path.lower().endswith('.xlsx'):
        output_path += '.xlsx'
    _check_writable(output_path)
    _free_space_check(output_path, source_path)

    progress(2, 'Reading discrepancy report…')
    groups = read_combinations(report_path)

    out_dir = os.path.dirname(os.path.abspath(output_path)) or None
    store = TempStore(out_dir)
    try:
        header, pos, excluded = scan_source(source_path, groups, store)
        progress(58, 'Calculating brokerage differences…')
        log('\n[Step 2C] Sorting and calculating Brokerage per group...')
        compute_group_stats(groups, store)

        import xlsxwriter
        n_out_cols = len(header) + 1
        matched = sum(1 for g in groups if g.count)
        log(f'         {matched:,} of {len(groups):,} combinations matched; '
            f'{sum(g.count for g in groups):,} transactions; negative-brokerage rows: {sum(g.neg for g in groups):,}')
        total_rows = max(sum(g.count for g in groups), 1)
        done = [0]
        date_idx, brk_idx = pos[C.DATE_COL], pos[C.BROK_COL]

        def write_book(fpath, grp, pre=None):
            """Write one workbook containing the given groups (optionally with index sheets first)."""
            _check_writable(fpath)
            log(f'\n[Step 2D] Writing {os.path.basename(fpath)} ({len(grp)} sheet group(s))...')
            wb = xlsxwriter.Workbook(fpath, {'constant_memory': True, 'strings_to_numbers': False,
                                             'use_zip64': True, 'tmpdir': out_dir})
            try:
                fm = Formats(wb)
                if pre:
                    pre(wb, fm)
                for g in grp:
                    written_neg = write_group_sheets(wb, fm, g, store, header, date_idx, brk_idx)
                    if written_neg != g.neg:                       # internal consistency guard
                        raise RuntimeError(f'internal check failed for {g.sheets[0]}: {written_neg} != {g.neg}')
                    done[0] += g.count
                    progress(60 + 38 * done[0] / total_rows, f'Writing… {done[0]:,} / {total_rows:,} rows')
                wb.close()
            except BaseException:
                abort_workbook(wb)
                try:
                    os.remove(fpath)
                except OSError:
                    pass
                raise

        if C.SEPARATE_FILES_PER_INVESTOR:
            # ───── one Excel file per investor (an investor with several combinations gets several sheets)
            folder_abs = os.path.join(out_dir or os.getcwd(), C.INVESTOR_FOLDER_NAME)
            os.makedirs(folder_abs, exist_ok=True)
            existing = [f for f in os.listdir(folder_abs) if f.lower().endswith('.xlsx')]
            if existing:
                log(f'         ⚠ Folder "{C.INVESTOR_FOLDER_NAME}" already has {len(existing)} Excel file(s); '
                    'files with the same name will be replaced.')
            by_inv = {}
            for g in groups:
                if g.count:
                    by_inv.setdefault(g.inv_n, []).append(g)
            fnamer = FileNamer(folder_abs)
            jobs = []
            for inv_n, gl in sorted(by_inv.items(), key=lambda kv: (kv[1][0].investor.casefold(), kv[0])):
                display = gl[0].investor
                base = fnamer.make(display)
                snamer = SheetNamer()
                multi = len(gl) > 1
                for g in gl:
                    label = f'{g.scheme}-{g.folio}-{g.txn}' if multi else display
                    if C.ADD_RATE_SUMMARY:
                        g.summary = snamer.make('Summary ' + label if multi else 'Rate Summary')
                    first = snamer.make(label)
                    g.sheets = [first] + [snamer.make(first, suffix=f'p{i}') for i in range(2, g.n_sheets + 1)]
                for k, grp in enumerate(plan_workbooks(gl, n_out_cols, reserved=1,
                                                       extra=1 if C.ADD_RATE_SUMMARY else 0), start=1):
                    fname = (base if k == 1 else fnamer.make(f'{display}_Part{k}')) + '.xlsx'
                    for g in grp:
                        g.file = os.path.join(C.INVESTOR_FOLDER_NAME, fname)
                    jobs.append((os.path.join(folder_abs, fname), grp))
            log(f'         {len(jobs):,} investor Excel file(s) will be created in "{folder_abs}".')
            for fpath, grp in jobs:
                write_book(fpath, grp)
            files = [f for f, _ in jobs]

            # ───── index workbook (the file the user chose): list of everything + links
            _check_writable(output_path)
            log(f'\n[Step 2E] Writing the index: {os.path.basename(output_path)}')
            wb = xlsxwriter.Workbook(output_path, {'constant_memory': True, 'strings_to_numbers': False,
                                                   'tmpdir': out_dir})
            try:
                fm = Formats(wb)
                ws_idx = wb.add_worksheet(C.INDEX_SHEET_NAME)
                ws_log = wb.add_worksheet(C.EXCLUDED_LOG_SHEET_NAME) if excluded else None
                write_index(ws_idx, fm, groups, -1, [], separate=True)
                if ws_log is not None:
                    write_excluded_log(ws_log, fm, excluded)
                wb.close()
            except BaseException:
                abort_workbook(wb)
                raise
            log(f'         ✓ Saved: {output_path}')
        else:
            # ───── everything in one workbook (split only when too large)
            namer = SheetNamer(reserved=[C.INDEX_SHEET_NAME, C.FILE_INDEX_SHEET_NAME, C.EXCLUDED_LOG_SHEET_NAME])
            for g in groups:
                if g.count:
                    first = namer.make(g.investor)
                    g.sheets = [first] + [namer.make(first, suffix=f'p{i}') for i in range(2, g.n_sheets + 1)]
            books = plan_workbooks(groups, n_out_cols)
            files = [part_path(output_path, k) for k in range(1, len(books) + 1)]
            for bi, grp in enumerate(books):
                for g in grp:
                    g.book = bi
            log(f'         Output will be written to {len(files)} workbook(s).')
            for bi, (grp, fpath) in enumerate(zip(books, files)):
                def pre(wb, fm, bi=bi):
                    ws_idx = wb.add_worksheet(C.INDEX_SHEET_NAME)
                    ws_files = wb.add_worksheet(C.FILE_INDEX_SHEET_NAME) if len(books) > 1 else None
                    ws_log = wb.add_worksheet(C.EXCLUDED_LOG_SHEET_NAME) if (excluded and bi == 0) else None
                    write_index(ws_idx, fm, groups, bi, files)
                    if ws_files is not None:
                        write_file_index(ws_files, fm, books, files)
                    if ws_log is not None:
                        write_excluded_log(ws_log, fm, excluded)
                write_book(fpath, grp, pre)
                log(f'         ✓ Saved: {fpath}')
        progress(100, 'Done')
    except OSError as e:
        if getattr(e, 'errno', None) == 28:
            raise UserError('Not enough disk space to finish writing the output.')
        raise
    finally:
        store.close()                                          # always removes the temp DB folder
        gc.collect()
    log(f'\nFinished in {time.time() - t0:,.1f} s')
    return {'files': files, 'index_file': output_path, 'groups': len(groups), 'matched': matched,
            'transactions': sum(g.count for g in groups)}
