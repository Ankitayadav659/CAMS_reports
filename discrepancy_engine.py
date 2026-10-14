"""
STEP 1 — CAMS Commission Discrepancy Checker (engine)   [v5.0 business logic preserved]

Business rules (unchanged from v5.0):
  * Group = AMC_CODE + SCHEME_COD + FOLIO_NO + INV_NAME + TRXN_TYPE (types never mixed).
  * Quarterly mode of BRKAGE_RAT (rounded to 2 dp); quarters with < 2 valid rates ignored.
  * TEMPORARY : drop vs previous eligible quarter < -0.2  AND  any later eligible quarter
                recovers >= 80 % of the drop.  Permanent drops are NOT flagged.
  * INTRA     : row differs from the quarter mode (4 dp) by >= 0.2 AND its rate is not in the
                group's "normal pool" (rate seen in >= 80 % of the group's quarters, or in >= 2 quarters).
  * TEMPORARY takes precedence over INTRA.
  * Sort: TRADE_DATE -> AMC_CODE -> TRXN_TYPE -> INV_NAME ascending.

NOTE ON THE v5.0 "DAILY SPLIT" COMMENT
  The v5.0 header comment describes a rule "multiple same-day rates that sum to the scheme's
  expected daily total are valid splits".  That rule is NOT implemented in the v5.0 code.
  What the code really does is the "normal pool" rule above, and that is what is kept here.
  No daily-sum rule has been added (it would change the business rules).

Performance-only changes vs v5.0 (results identical, verified by tests against the v5.0 code):
  streaming reader of only the 7 needed columns, vectorised grouping, integer ids, categoricals,
  hash-join annotation, one streaming xlsxwriter pass (no openpyxl reload of the big output),
  automatic continuation sheets above Excel's row limit.
Bug fix: rows dropped for unparseable dates no longer shift the index (v5.0 could attach INTRA
  flags to the wrong rows in that situation).
"""
import gc
import math
import os
import time
import warnings

import numpy as np
import pandas as pd

import config as C
from excel_reader import cell_to_str, open_sheet
from excel_writer import abort_workbook
from reporting import UserError, log, progress

DATE_COL, BROK_COL, AMC_COL = C.DATE_COL, C.BROK_COL, C.AMC_COL
SCHEME_COL, TXN_TYPE_COL, INV_COL, FOLIO_COL = C.SCHEME_COL, C.TXN_TYPE_COL, C.INV_COL, C.FOLIO_COL
KEEP_COLS, UNIQUE_ID_COLS = C.KEEP_COLS, C.UNIQUE_ID_COLS
THRESHOLD, RECOVERY_RATIO, MIN_TXN = C.THRESHOLD, C.RECOVERY_RATIO, C.MIN_TXN


# ═══════════════════════════════════════════════════════
# STEP 1: LOAD — stream only the 7 needed columns
# ═══════════════════════════════════════════════════════
def load_and_clean(filepath):
    log('\n[Step 1] Loading file...')
    rd = open_sheet(filepath)
    try:
        header = rd.header()
        pos = {}
        for i, h in enumerate(header):
            if h in KEEP_COLS and h not in pos:
                pos[h] = i
        missing = [c for c in KEEP_COLS if c not in pos]
        if missing:
            raise UserError('These required columns were NOT found in the first sheet: '
                            f'{missing}. Please check the header names.')
        order = [pos[c] for c in KEEP_COLS]
        cols = [[] for _ in KEEP_COLS]
        total = rd.total_rows or 0
        n = 0
        for row in rd.rows():
            L = len(row)
            for k, i in enumerate(order):
                cols[k].append(row[i] if i < L else None)
            n += 1
            if n % 200_000 == 0:
                progress(2 + 28 * min(n / total, 1) if total else 10, f'Reading Excel… {n:,} rows')
    finally:
        rd.close()
    data = {}
    for c, col in zip(KEEP_COLS, cols):
        if c == DATE_COL:
            data[c] = pd.Series(col, dtype=object)            # keep real date cells as real dates
        else:
            data[c] = [cell_to_str(v) for v in col]
    df = pd.DataFrame(data)
    del cols, data
    log(f'         OK: {len(df):,} rows, keeping {len(KEEP_COLS)} columns: {KEEP_COLS}')
    return df


# ═══════════════════════════════════════════════════════
# STEP 2: PARSE TRADE_DATE
# ═══════════════════════════════════════════════════════
def _parse_text_dates(txt):
    """Text dates: ISO (yyyy-mm-dd...) year-first, everything else day-first (as v5.0 intended)."""
    out = pd.Series(pd.NaT, index=txt.index, dtype='datetime64[ns]')
    iso = txt.str.match(r'^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}').fillna(False).astype(bool)
    if iso.any():
        s = txt[iso].str.replace('/', '-', regex=False).str.replace('.', '-', regex=False)
        p = pd.to_datetime(s, errors='coerce', format='ISO8601')
        bad = p.isna()
        if bad.any():
            p.loc[bad] = pd.to_datetime(s[bad], errors='coerce', format='mixed', dayfirst=False)
        out.loc[iso] = p
    rest = ~iso
    if rest.any():
        s = txt[rest]
        p = pd.to_datetime(s, dayfirst=True, errors='coerce')
        bad = p.isna()
        if bad.any():                                    # mixed formats in one column
            try:
                p.loc[bad] = pd.to_datetime(s[bad], dayfirst=True, errors='coerce', format='mixed')
            except Exception:
                pass
        out.loc[rest] = p
    return out


def _parse_dates_robust(raw):
    import datetime as _dt
    vals = raw.tolist()
    n = len(vals)
    parts = []
    d_idx, d_val, t_idx, t_val, n_idx, n_val = [], [], [], [], [], []
    for i, v in enumerate(vals):
        if v is None:
            continue
        if isinstance(v, _dt.datetime):
            d_idx.append(i); d_val.append(v)
        elif isinstance(v, _dt.date):
            d_idx.append(i); d_val.append(_dt.datetime(v.year, v.month, v.day))
        elif isinstance(v, str):
            s = v.strip()
            if s:
                t_idx.append(i); t_val.append(s)
        elif isinstance(v, (int, float)) and not isinstance(v, bool) and v == v and 1 <= v < 2958466:
            n_idx.append(i); n_val.append(float(v))
    if d_idx:
        parts.append(pd.Series(pd.to_datetime(d_val), index=d_idx).astype('datetime64[ns]'))
    if t_idx:
        parts.append(_parse_text_dates(pd.Series(t_val, index=t_idx, dtype=object)).astype('datetime64[ns]'))
    if n_idx:
        parts.append(pd.Series(pd.to_datetime(n_val, unit='D', origin='1899-12-30'), index=n_idx)
                     .astype('datetime64[ns]'))
    if not parts:
        return pd.Series(pd.NaT, index=range(n), dtype='datetime64[ns]')
    return pd.concat(parts).reindex(range(n)).astype('datetime64[ns]')


def parse_dates(df):
    log('\n[Step 2] Parsing trade dates...')
    if C.LEGACY_V5_DATE_PARSING:
        log('         (LEGACY v5.0 date parsing is ON)')
        raw = df[DATE_COL].map(cell_to_str)
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            parsed = pd.to_datetime(raw, dayfirst=True, errors='coerce')
            retry = parsed.isna() & raw.notna()
            if retry.any():
                try:
                    parsed.loc[retry] = pd.to_datetime(
                        raw.loc[retry], dayfirst=True, errors='coerce', format='mixed')
                except Exception:
                    pass
    else:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            parsed = _parse_dates_robust(df[DATE_COL])
    bad = int(parsed.isna().sum())
    if bad:
        log(f'         Dropped {bad:,} rows with unparseable dates.')
    df[DATE_COL] = parsed
    df = df.loc[parsed.notna()].reset_index(drop=True)
    log(f'         OK: {len(df):,} rows with valid dates.')
    if df.empty:
        raise UserError('No rows with a valid TRADE_DATE were found.')
    return df


# ═══════════════════════════════════════════════════════
# STEP 3: QUARTER + UNIQUE ID  (compact integer codes)
# ═══════════════════════════════════════════════════════
def add_quarter_and_id(df):
    log('\n[Step 3] Adding Quarter and Unique ID...')
    dt = df[DATE_COL].dt
    key = dt.year.to_numpy(dtype=np.int64) * 4 + (dt.month.to_numpy(dtype=np.int64) - 1) // 3
    ukeys, qcode = np.unique(key, return_inverse=True)
    qlabels = [f'{k // 4}-Q{k % 4 + 1}' for k in ukeys.tolist()]
    df['_q'] = qcode.ravel().astype(np.int16)

    s = df[UNIQUE_ID_COLS[0]].fillna('').astype(str)
    for c in UNIQUE_ID_COLS[1:]:
        s = s + '||' + df[c].fillna('').astype(str)
    codes, uniques = pd.factorize(s, sort=True)
    df['_uid'] = codes.astype(np.int32)
    n_uid = len(uniques)
    del s, codes, uniques

    log(f'         Quarters : {qlabels[0]} → {qlabels[-1]} ({len(qlabels)} total)')
    log(f'         Unique IDs: {n_uid:,}')

    df[BROK_COL] = pd.to_numeric(df[BROK_COL], errors='coerce')
    rates = df[BROK_COL].dropna()
    if len(rates):
        log(f'\n[Diagnostic] BRKAGE_RAT — min={rates.min():.4f}  '
            f'max={rates.max():.4f}  mean={rates.mean():.4f}')
        if rates.max() > 100:
            log('         ⚠ WARNING: values look like amounts, not percentages. Check column mapping!')
        else:
            log('         ✓ Values look like percentages (≤ 100).')

    for c in (AMC_COL, SCHEME_COL, FOLIO_COL, TXN_TYPE_COL, INV_COL):
        df[c] = df[c].astype('category')
    return df, qlabels


# MODE per (uid, quarter): ties -> smallest value (same as Series.mode().iloc[0])
def _mode_table(u, q, vals, decimals):
    t = pd.DataFrame({'u': u, 'q': q, 'r': np.round(vals, decimals)}).dropna(subset=['r'])
    c = t.groupby(['u', 'q', 'r'], sort=False).size().rename('n').reset_index()
    c = c.sort_values(['u', 'q', 'n', 'r'], ascending=[True, True, False, True])
    c = c.drop_duplicates(['u', 'q'], keep='first')
    return c[['u', 'q', 'r']].rename(columns={'r': 'mode'}).reset_index(drop=True)


# ═══════════════════════════════════════════════════════
# STEP 4: TEMPORARY vs PERMANENT
# ═══════════════════════════════════════════════════════
def classify_discrepancies(df):
    log('\n[Step 4] Classifying rate changes...')
    u    = df['_uid'].to_numpy()
    q    = df['_q'].to_numpy()
    rate = df[BROK_COL].to_numpy(dtype=float)

    cnt = (pd.DataFrame({'u': u, 'q': q, 'rate': rate})
           .groupby(['u', 'q'], sort=False)['rate'].count()
           .rename('Count').reset_index())
    mode = _mode_table(u, q, rate, 2)

    gs = cnt.merge(mode, on=['u', 'q'], how='left')
    gs = gs[gs['Count'] >= MIN_TXN].rename(columns={'mode': 'Rate'})
    gs = gs.sort_values(['u', 'q']).reset_index(drop=True)
    del cnt, mode

    empty = pd.DataFrame(columns=['u', 'q', 'prev_q', 'prev_rate', 'curr_rate', 'drop', 'target'])
    if gs.empty:
        log('         Temporary discrepancies (will highlight) : 0')
        log('         Permanent rate changes   (skipped)       : 0')
        return empty

    g = gs.groupby('u', sort=False)
    prev_rate = g['Rate'].shift(1)
    prev_q    = g['q'].shift(1)
    rev_max = gs['Rate'][::-1].groupby(gs['u'][::-1]).cummax()[::-1]
    gs['_fm'] = rev_max
    future_max = gs.groupby('u', sort=False)['_fm'].shift(-1)

    drop = gs['Rate'] - prev_rate
    c_idx = np.flatnonzero((drop < THRESHOLD).to_numpy())

    cur = gs['Rate'].to_numpy()[c_idx]
    prv = prev_rate.to_numpy()[c_idx]
    dp  = drop.to_numpy()[c_idx]
    tgt = cur + RECOVERY_RATIO * np.abs(dp)
    fut = future_max.to_numpy()[c_idx]
    with np.errstate(invalid='ignore'):
        rec = fut >= tgt

    temp = pd.DataFrame({
        'u'        : gs['u'].to_numpy()[c_idx][rec],
        'q'        : gs['q'].to_numpy()[c_idx][rec],
        'prev_q'   : prev_q.to_numpy()[c_idx][rec].astype(np.int64),
        'prev_rate': np.round(prv[rec], 4),
        'curr_rate': np.round(cur[rec], 4),
        'drop'     : np.round(dp[rec], 4),
        'target'   : np.round(tgt[rec], 4),
    })
    log(f'         Temporary discrepancies (will highlight) : {len(temp):,}')
    log(f'         Permanent rate changes   (skipped)       : {int((~rec).sum()):,}')
    return temp


# ═══════════════════════════════════════════════════════
# STEP 4b: INTRA-QUARTER DISCREPANCIES
# ═══════════════════════════════════════════════════════
def classify_intra_quarter(df, n_quarters):
    log('\n[Step 4b] Detecting intra-quarter rate inconsistencies...')
    u    = df['_uid'].to_numpy()
    q    = df['_q'].to_numpy()
    rate = df[BROK_COL].to_numpy(dtype=float)
    r4   = np.round(rate, 4)

    t = pd.DataFrame({'u': u, 'q': q, 'r': r4})
    tot_q = t[['u', 'q']].drop_duplicates().groupby('u').size()
    pr = t.dropna(subset=['r']).drop_duplicates(['u', 'q', 'r'])
    qc = pr.groupby(['u', 'r'], sort=False).size().rename('qc').reset_index()
    qc['tot'] = qc['u'].map(tot_q)
    is_normal = ((qc['qc'] / qc['tot']) >= 0.80) | (qc['qc'] >= 2)
    normal = qc.loc[is_normal, ['u', 'r']].assign(_in=1)
    del t, pr, qc

    mode = _mode_table(u, q, rate, 4)
    nq = max(int(n_quarters), 1)
    row_key  = u.astype(np.int64) * nq + q.astype(np.int64)
    mode_key = mode['u'].to_numpy(np.int64) * nq + mode['q'].to_numpy(np.int64)
    mi = pd.Index(mode_key).get_indexer(row_key)
    mode_vals = mode['mode'].to_numpy(dtype=float)
    mode_row = np.where(mi >= 0, mode_vals[np.where(mi >= 0, mi, 0)], np.nan)

    with np.errstate(invalid='ignore'):
        diff = r4 - np.round(mode_row, 4)
        cand = np.flatnonzero(np.abs(diff) >= abs(THRESHOLD))

    if len(cand):
        cdf = pd.DataFrame({'u': u[cand], 'r': r4[cand]}).merge(normal, on=['u', 'r'], how='left')
        in_pool = cdf['_in'].notna().to_numpy()
    else:
        in_pool = np.zeros(0, dtype=bool)

    pos = cand[~in_pool]
    flags = pd.DataFrame({
        'pos'      : pos,
        'row_rate' : r4[pos],
        'mode_rate': np.round(mode_row[pos], 4),
        'diff'     : np.round(rate[pos] - mode_row[pos], 4),
        'q'        : q[pos].astype(np.int64),
    })
    log(f'         Valid split rows skipped (in normal pool) : {len(cand) - len(pos):,}')
    log(f'         Intra-quarter inconsistent rows flagged   : {len(flags):,}')
    return flags


# ═══════════════════════════════════════════════════════
# STEP 5: DISCREPANCY SUMMARY REPORT
# ═══════════════════════════════════════════════════════
def build_discrepancy_report(df, temp, qlabels):
    log('\n[Step 5] Building discrepancy report...')
    if temp.empty:
        log('         No temporary discrepancies found.')
        return pd.DataFrame()

    meta = (df.drop_duplicates('_uid')
              [['_uid', INV_COL, AMC_COL, SCHEME_COL, FOLIO_COL, TXN_TYPE_COL]]
              .set_index('_uid'))
    m = meta.loc[temp['u'].to_numpy()]
    ql = np.array(qlabels, dtype=object)

    def txt(col):
        return m[col].astype(object).fillna('').to_numpy()

    report = pd.DataFrame({
        'Investor'            : txt(INV_COL),
        'AMC'                 : txt(AMC_COL),
        'Scheme'              : txt(SCHEME_COL),
        'Folio'               : txt(FOLIO_COL),
        'Transaction Type'    : txt(TXN_TYPE_COL),
        'Previous Quarter'    : ql[temp['prev_q'].to_numpy()],
        'Discrepancy Quarter' : ql[temp['q'].to_numpy(np.int64)],
        'Prev Rate (%)'       : temp['prev_rate'].to_numpy(),
        'Disc Rate (%)'       : temp['curr_rate'].to_numpy(),
        'Difference (%)'      : temp['drop'].to_numpy(),
        'Recovery Target (%)' : temp['target'].to_numpy(),
    })
    report = report.sort_values(
        ['AMC', 'Investor', 'Scheme', 'Transaction Type', 'Discrepancy Quarter'],
        ascending=True).reset_index(drop=True)
    log(f'         Report rows: {len(report):,}')
    return report


# ═══════════════════════════════════════════════════════
# STEP 6: ANNOTATE RAW DATA ROWS
# ═══════════════════════════════════════════════════════
def annotate_raw_data(df, temp, intra, qlabels):
    log('\n[Step 6] Annotating raw data rows...')
    n = len(df)
    disc = np.full(n, '', dtype=object)
    det  = np.full(n, '', dtype=object)
    rdf  = np.full(n, '', dtype=object)
    nq = max(len(qlabels), 1)

    if len(temp):
        row_key  = df['_uid'].to_numpy(np.int64) * nq + df['_q'].to_numpy(np.int64)
        flag_key = temp['u'].to_numpy(np.int64) * nq + temp['q'].to_numpy(np.int64)
        idx  = pd.Index(flag_key).get_indexer(row_key)
        rows = np.flatnonzero(idx >= 0)
        fi   = idx[rows]
        f_det = [
            f"TEMPORAL DISCREPANCY: Rate dropped from {pr}% to {cr}% (diff={d}%) "
            f"in {qlabels[q]}, then recovered. Prev quarter: {qlabels[pq]}"
            for pr, cr, d, q, pq in zip(
                temp['prev_rate'].tolist(), temp['curr_rate'].tolist(),
                temp['drop'].tolist(), temp['q'].tolist(), temp['prev_q'].tolist())
        ]
        f_rd = [str(d) for d in temp['drop'].tolist()]
        disc[rows] = 'TEMPORARY'
        det[rows]  = np.array(f_det, dtype=object)[fi]
        rdf[rows]  = np.array(f_rd, dtype=object)[fi]

    if len(intra):
        pos = intra['pos'].to_numpy()
        keep = disc[pos] == ''
        sub = intra[keep]
        p = sub['pos'].to_numpy()
        disc[p] = 'INTRA'
        det[p] = np.array([
            f"INTRA-QUARTER DISCREPANCY: This row rate={rr}% differs from quarter mode={mr}% "
            f"(diff={d}%) in {qlabels[q]}"
            for rr, mr, d, q in zip(sub['row_rate'].tolist(), sub['mode_rate'].tolist(),
                                    sub['diff'].tolist(), sub['q'].tolist())
        ], dtype=object)
        rdf[p] = np.array([str(d) for d in sub['diff'].tolist()], dtype=object)

    df['_Discrepancy'] = disc
    df['_ChangeType']  = det
    df['_RateDiff']    = rdf
    log(f'         Temporal discrepancy rows : {int((disc == "TEMPORARY").sum()):,}')
    log(f'         Intra-quarter anomaly rows: {int((disc == "INTRA").sum()):,}')
    return df


# ═══════════════════════════════════════════════════════
# STEP 7: SORT — row ORDER only (no big copy)
# ═══════════════════════════════════════════════════════
def sort_order(df):
    sort_cols = [c for c in [DATE_COL, AMC_COL, TXN_TYPE_COL, INV_COL] if c in df.columns]
    log(f'\n[Step 7] Sorting by: {sort_cols}')
    return df[sort_cols].sort_values(sort_cols, ascending=True).index.to_numpy()


# ═══════════════════════════════════════════════════════
# STEP 8: EXPORT — streamed with xlsxwriter (constant memory)
# ═══════════════════════════════════════════════════════
_EPOCH = pd.Timestamp('1899-12-30')


def export(df, order, report, output_path, _holder=None):
    import xlsxwriter

    log(f'\n[Step 8] Writing output: {output_path}')
    src_cols = KEEP_COLS + ['_Discrepancy', '_ChangeType', '_RateDiff']
    out_cols = KEEP_COLS + ['Change Type', 'Change Details', 'Rate Diff (%)']
    n_cols = len(out_cols)
    n_rows = len(order)
    disc_i = out_cols.index('Change Type')
    kinds = [2 if c == DATE_COL else 1 if c == BROK_COL else 0 for c in src_cols]   # 0 text 1 num 2 date

    cap = C.STEP1_ROWS_PER_SHEET
    hdr_rows = C.STEP1_HEADER_ROWS
    n_sheets = max(1, math.ceil(n_rows / cap))
    log(f'         Writing {n_rows:,} rows × {n_cols} cols (streaming)'
        + (f' across {n_sheets} sheets' if n_sheets > 1 else '') + '...')

    tmpdir = os.path.dirname(os.path.abspath(output_path)) or None
    wb = xlsxwriter.Workbook(output_path, {'constant_memory': True, 'strings_to_numbers': False,
                                           'use_zip64': True, 'tmpdir': tmpdir})
    if _holder is not None:
        _holder.append(wb)
    names = [C.RAW_SHEET_NAME] + [f'{C.RAW_SHEET_NAME} {i}' for i in range(2, n_sheets + 1)]
    sheets = [wb.add_worksheet(nm) for nm in names]
    ws2 = wb.add_worksheet(C.REPORT_SHEET_NAME) if not report.empty else None

    fmt = lambda **k: wb.add_format(k)
    hdr_f = fmt(bold=True, font_color='#FFFFFF', bg_color='#1F4E79', font_name='Arial',
                font_size=10, align='center', valign='vcenter', text_wrap=True)
    leg_f = fmt(bold=True, font_color='#1F4E79', font_name='Arial', font_size=10, align='left')
    red_f = fmt(bold=True, font_color='#000000', bg_color='#FFCCCC', font_name='Arial', font_size=9, align='left')
    red_d = fmt(bold=True, font_color='#000000', bg_color='#FFCCCC', font_name='Arial', font_size=9,
                align='left', num_format='dd-mm-yyyy')
    alt_f = fmt(bg_color='#F2F7FF', font_name='Arial', font_size=9, align='left')
    alt_d = fmt(bg_color='#F2F7FF', font_name='Arial', font_size=9, align='left', num_format='dd-mm-yyyy')
    nrm_f = fmt(font_name='Arial', font_size=9, align='left')
    nrm_d = fmt(font_name='Arial', font_size=9, align='left', num_format='dd-mm-yyyy')

    wide_cols   = {'Change Details', INV_COL, SCHEME_COL}
    medium_cols = {DATE_COL, AMC_COL, FOLIO_COL, TXN_TYPE_COL}
    legend = ('LIGHT RED = Discrepancy (temporal rate drop/recovery  OR  intra-quarter rate anomaly)   '
              'WHITE/BLUE = Normal')
    for ws in sheets:
        for c_i, col in enumerate(out_cols):
            ws.set_column(c_i, c_i, 35 if col in wide_cols else 22 if col in medium_cols else 16)
        ws.merge_range(0, 0, 0, n_cols - 1, legend, leg_f)
        ws.set_row(0, 18)
        for c_i, col in enumerate(out_cols):
            ws.write_string(1, c_i, col, hdr_f)
        ws.set_row(1, 18)
        ws.freeze_panes(2, 0)
        ws.autofilter(1, 0, 1, n_cols - 1)

    g = 0
    for s in range(0, n_rows, C.WRITE_CHUNK):
        sel = order[s:s + C.WRITE_CHUNK]
        cols = []
        for col, kind in zip(src_cols, kinds):
            ser = df[col].iloc[sel]
            if kind == 2:
                cols.append(((ser - _EPOCH).dt.total_seconds() / 86400.0).tolist())
            elif kind == 1:
                cols.append(ser.tolist())
            else:
                cols.append(ser.astype(object).tolist())
        for row in zip(*cols):
            si, rr = divmod(g, cap)
            r_i = rr + hdr_rows
            ws = sheets[si]
            if row[disc_i] in ('TEMPORARY', 'INTRA'):
                fmt_t, fmt_d = red_f, red_d
            elif r_i % 2 == 0:
                fmt_t, fmt_d = alt_f, alt_d
            else:
                fmt_t, fmt_d = nrm_f, nrm_d
            for c_i, v in enumerate(row):
                k = kinds[c_i]
                if k == 2:
                    ws.write_number(r_i, c_i, v, fmt_d)
                elif v is None or v != v:
                    ws.write_blank(r_i, c_i, None, fmt_t)
                elif k == 1:
                    ws.write_number(r_i, c_i, v, fmt_t)
                else:
                    ws.write_string(r_i, c_i, str(v), fmt_t)
            g += 1
        progress(75 + 22 * g / max(n_rows, 1), f'Writing rows… {g:,} / {n_rows:,}')
        log(f'         ... {g:,} rows written')
    log('         Raw Data sheet done.')

    if ws2 is not None:
        log(f'         Writing Discrepancy Report ({len(report):,} rows)...')
        rhdr = fmt(bold=True, font_color='#FFFFFF', bg_color='#1F4E79', font_name='Arial',
                   font_size=10, align='center', valign='vcenter', text_wrap=True)
        rcell = fmt(bold=True, font_color='#000000', bg_color='#FFCCCC', font_name='Arial', font_size=9,
                    align='left', valign='vcenter', bottom=1, right=1,
                    bottom_color='#D0D0D0', right_color='#D0D0D0')
        for c_i, col in enumerate(report.columns):
            longest = max([len(str(col))] + report[col].astype(str).str.len().tolist())
            ws2.set_column(c_i, c_i, min(longest + 3, 50))
        for c_i, col in enumerate(report.columns):
            ws2.write_string(0, c_i, col, rhdr)
        for r_i, row_vals in enumerate(report.itertuples(index=False), start=1):
            for c_i, val in enumerate(row_vals):
                if isinstance(val, str):
                    ws2.write_string(r_i, c_i, val, rcell)      # keeps leading zeros / never a formula
                else:
                    ws2.write_number(r_i, c_i, val, rcell)
        ws2.freeze_panes(1, 0)
        ws2.autofilter(0, 0, len(report), len(report.columns) - 1)
        log('         Discrepancy Report done.')

    wb.close()
    progress(100, 'Done')
    log(f'\n         ✓ Saved: {output_path}')
    log(f'           Sheet 1 "{C.RAW_SHEET_NAME}" — LIGHT RED rows = discrepancies')
    if ws2 is not None:
        log(f'           Last sheet "{C.REPORT_SHEET_NAME}" — sorted AMC → Investor → Scheme → Quarter')


def check_writable(path):
    existed = os.path.exists(path)
    try:
        with open(path, 'ab'):
            pass
    except PermissionError:
        raise UserError(f'Cannot write to "{path}". If it is open in Excel, close it and try again.')
    except OSError as e:
        raise UserError(f'Cannot write to "{path}": {e}')
    if not existed:
        try:
            os.remove(path)
        except OSError:
            pass


def run(input_path, output_path):
    t0 = time.time()
    if not os.path.isfile(input_path):
        raise UserError(f'Input file not found: {input_path}')
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise UserError('Output file must be different from the input file.')
    check_writable(output_path)

    progress(2, 'Reading Excel…')
    df = load_and_clean(input_path)
    progress(35, 'Parsing dates…')
    df = parse_dates(df)
    progress(40, 'Preparing groups…')
    df, qlabels = add_quarter_and_id(df)
    gc.collect()
    progress(48, 'Detecting temporary discrepancies…')
    temp = classify_discrepancies(df)
    progress(58, 'Detecting intra-quarter discrepancies…')
    intra = classify_intra_quarter(df, len(qlabels))
    progress(66, 'Building report…')
    report = build_discrepancy_report(df, temp, qlabels)
    progress(70, 'Annotating rows…')
    df = annotate_raw_data(df, temp, intra, qlabels)
    progress(73, 'Sorting…')
    order = sort_order(df)
    progress(75, 'Writing Excel…')
    holder = []
    try:
        export(df, order, report, output_path, holder)
    except BaseException as e:
        abort_workbook(holder[0] if holder else None)          # no half-written file / temp files left
        try:
            os.remove(output_path)
        except OSError:
            pass
        if isinstance(e, OSError) and getattr(e, 'errno', None) == 28:
            raise UserError('Not enough disk space to write the output file.')
        raise
    stats = {'rows': int(len(order)), 'temporary_rows': int((df['_Discrepancy'] == 'TEMPORARY').sum()),
             'intra_rows': int((df['_Discrepancy'] == 'INTRA').sum()), 'report_rows': len(report),
             'output': output_path}
    del df, order, temp, intra
    gc.collect()
    log(f'\nFinished in {time.time() - t0:,.1f} s')
    return stats
