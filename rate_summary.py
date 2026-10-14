"""
"Rate Summary" sheet — a simple one-page explanation of how the brokerage rate moved, per client (no charts).

  * How-to-read box  : Brokerage = Previous − Current;  MINUS (−) = rate INCREASED, PLUS (+) = DECREASED.
  * KPI block        : transactions, first/last date, current rate, number of increases / decreases, highest / lowest.
  * Rate-period table: FROM date – TO date for every rate, how it changed vs the previous period,
                       plain-language explanation, date of the next change.

Rows with a missing / invalid date or rate are left out (they are reported in the investor index notes).
The sheet is written in constant-memory mode, so cells are buffered and written strictly row by row.
"""
from collections import defaultdict
from datetime import datetime, timedelta

import config as C

_EPOCH = datetime(1899, 12, 30)
RED, BLUE, NAVY = '#E53935', '#2E75B6', '#1F4E79'
INC_FILL, DEC_FILL, START_FILL = '#FFCCCC', '#DDEBF7', '#F2F2F2'


# ───────────────────────── pure logic (unit-tested) ─────────────────────────
def serial_to_text(serial):
    return (_EPOCH + timedelta(days=float(serial))).strftime('%d-%m-%Y')


def compute_periods(points):
    """points: iterable of (excel_serial_date, rate) in trade-date order (valid rows only).
    A period = consecutive transactions with the same rate.
    change = previous period rate − this period rate  (same sign convention as the Brokerage column):
        negative -> rate INCREASED,  positive -> rate DECREASED,  None -> first period."""
    periods = []
    for d, r in points:
        key = round(r, 8)
        if periods and periods[-1]['key'] == key:
            periods[-1]['to'] = d
            periods[-1]['n'] += 1
        else:
            periods.append({'key': key, 'rate': r, 'from': d, 'to': d, 'n': 1})
    for i, p in enumerate(periods):                  # first pass: change vs previous period
        p['change'] = None if i == 0 else round(periods[i - 1]['rate'] - p['rate'], 10)
    for i, p in enumerate(periods):                  # second pass: needs the NEXT period's change
        nxt = periods[i + 1] if i + 1 < len(periods) else None
        p['next_from'] = nxt['from'] if nxt else None
        p['next_change'] = nxt['change'] if nxt else None
        p['next_rate'] = nxt['rate'] if nxt else None
        p['status'] = ('START' if p['change'] is None else 'INCREASED' if p['change'] < 0 else
                       'DECREASED' if p['change'] > 0 else 'SAME')
    return periods


def what_happened(p, prev_rate):
    """One short, plain phrase for the table."""
    if p['status'] == 'START':
        t = 'Starting rate'
    elif p['status'] == 'INCREASED':
        t = f"▲  Rate went UP by {abs(p['change']):.4f}   (was {prev_rate:.4f})"
    elif p['status'] == 'DECREASED':
        t = f"▼  Rate went DOWN by {abs(p['change']):.4f}   (was {prev_rate:.4f})"
    else:
        t = 'No change'
    return t + ('     ← current rate' if p['next_from'] is None else '')


def summarize(periods):
    inc = sum(1 for p in periods if p['status'] == 'INCREASED')
    dec = sum(1 for p in periods if p['status'] == 'DECREASED')
    hi = max(periods, key=lambda p: (p['rate'], -p['from']))
    lo = min(periods, key=lambda p: (p['rate'], p['from']))
    return {'transactions': sum(p['n'] for p in periods), 'first': periods[0]['from'], 'last': periods[-1]['to'],
            'current': periods[-1]['rate'], 'increases': inc, 'decreases': dec, 'high': hi, 'low': lo}


# ───────────────────────── writer ─────────────────────────
class _Buffer:
    """constant_memory needs rows in ascending order -> collect everything, then flush sorted."""

    def __init__(self, ws):
        self.ws, self.ops, self.heights = ws, defaultdict(list), {}

    def s(self, r, c, text, fmt):
        self.ops[r].append((c, 's', (text, fmt)))

    def n(self, r, c, val, fmt):
        self.ops[r].append((c, 'n', (val, fmt)))

    def b(self, r, c, fmt):
        self.ops[r].append((c, 'b', (fmt,)))

    def merge(self, r, c1, c2, text, fmt):
        self.ops[r].append((c1, 'm', (c2, text, fmt)))

    def height(self, r, h):
        self.heights[r] = h

    def flush(self):
        ws = self.ws
        for r in sorted(set(self.ops) | set(self.heights)):
            if r in self.heights:
                ws.set_row(r, self.heights[r])
            for c, kind, a in sorted(self.ops.get(r, []), key=lambda x: x[0]):
                if kind == 's':
                    ws.write_string(r, c, a[0], a[1])
                elif kind == 'n':
                    ws.write_number(r, c, a[0], a[1])
                elif kind == 'b':
                    ws.write_blank(r, c, None, a[0])
                else:
                    ws.merge_range(r, c, r, a[0], a[1], a[2])


def _formats(wb):
    cache = getattr(wb, '_rs_fmt', None)
    if cache:
        return cache
    base = dict(font_name='Arial', font_size=10, valign='vcenter')
    f = {}
    A = wb.add_format
    f['title'] = A(dict(base, bold=True, font_size=18, font_color=NAVY))
    f['sub'] = A(dict(base, font_size=11, font_color='#475569'))
    f['how_h'] = A(dict(base, bold=True, font_color='#FFFFFF', bg_color=NAVY))
    f['how'] = A(dict(base, bg_color='#F8FAFC', text_wrap=True, border=1, border_color='#E2E8F0'))
    f['how_inc'] = A(dict(base, bg_color=INC_FILL, bold=True, border=1, border_color='#E2E8F0'))
    f['how_dec'] = A(dict(base, bg_color=DEC_FILL, bold=True, border=1, border_color='#E2E8F0'))
    f['kpi_l'] = A(dict(base, bold=True, bg_color='#EEF3F9', border=1, border_color='#D0D7E2'))
    f['kpi_v'] = A(dict(base, bold=True, font_size=11, font_color=NAVY, border=1, border_color='#D0D7E2', align='left'))
    f['kpi_d'] = A(dict(base, bold=True, font_size=11, font_color=NAVY, border=1, border_color='#D0D7E2',
                        align='left', num_format='dd-mm-yyyy'))
    f['kpi_r'] = A(dict(base, bold=True, font_size=11, font_color=NAVY, border=1, border_color='#D0D7E2',
                        align='left', num_format='0.0000'))
    f['sec'] = A(dict(base, bold=True, font_size=12, font_color=NAVY, bottom=2, bottom_color=NAVY))
    f['th'] = A(dict(base, bold=True, font_color='#FFFFFF', bg_color=NAVY, align='center', text_wrap=True,
                     border=1, border_color='#FFFFFF'))
    for tag, fill in (('inc', INC_FILL), ('dec', DEC_FILL), ('start', START_FILL)):
        b = dict(base, bg_color=fill, border=1, border_color='#D0D7E2')
        f[tag + '_t'] = A(dict(b, align='left'))
        f[tag + '_c'] = A(dict(b, align='center'))
        f[tag + '_d'] = A(dict(b, align='center', num_format='dd-mm-yyyy'))
        f[tag + '_r'] = A(dict(b, align='center', num_format='0.0000', bold=True))
        f[tag + '_x'] = A(dict(b, align='center', num_format='+0.0000;-0.0000;0.0000', bold=True))
        f[tag + '_s'] = A(dict(b, align='left', bold=True,
                               font_color={'inc': '#B71C1C', 'dec': '#1F5F99', 'start': '#475569'}[tag]))
        f[tag + '_w'] = A(dict(b, align='left', text_wrap=True))
    f['line'] = A(dict(base, font_size=11, bold=True, font_color=NAVY, bg_color='#EEF3F9', border=1, border_color='#D0D7E2'))
    f['note'] = A(dict(base, italic=True, font_color='#64748B'))
    wb._rs_fmt = f
    return f


def write_rate_summary(wb, ws, group, points):
    """ws: a worksheet already added to wb (so it sits before the data sheet). points: [(serial, rate), ...]"""
    F = _formats(wb)
    buf = _Buffer(ws)
    LAST = 4                                           # last used column (E)
    ws.hide_gridlines(2)
    ws.set_zoom(110)
    ws.set_paper(9)                                    # A4
    ws.fit_to_pages(1, 0)
    for c, w in enumerate([6, 16, 16, 12, 62]):
        ws.set_column(c, c, w)

    buf.height(0, 28)
    buf.merge(0, 0, LAST, f'Brokerage Rate Summary — {group.investor}', F['title'])
    buf.merge(1, 0, LAST, f'AMC {group.amc}   |   Scheme {group.scheme}   |   Folio {group.folio}   |   '
                          f'Transaction Type {group.txn}', F['sub'])

    buf.merge(3, 0, LAST, '  HOW TO READ', F['how_h'])
    buf.merge(4, 0, LAST, '  ▲  MINUS ( − ) in the Brokerage column  =  the rate went UP (increased).', F['how_inc'])
    buf.merge(5, 0, LAST, '  ▼  PLUS ( + )  =  the rate went DOWN (decreased).        0.0000  =  no change.', F['how_dec'])
    for r_ in (4, 5):
        buf.height(r_, 21)

    periods = compute_periods(points)
    if not periods:
        buf.merge(7, 0, LAST, 'No transactions with a valid TRADE_DATE and BRKAGE_RAT were found for this client.', F['note'])
        buf.flush()
        return
    K = summarize(periods)

    buf.merge(7, 0, LAST, '  SUMMARY', F['sec'])
    buf.merge(8, 0, LAST, f"  Total transactions: {K['transactions']}        From {serial_to_text(K['first'])}"
                          f"  to  {serial_to_text(K['last'])}", F['line'])
    buf.merge(9, 0, LAST, f"  The rate went UP {K['increases']} time{'s' if K['increases'] != 1 else ''} and "
                          f"DOWN {K['decreases']} time{'s' if K['decreases'] != 1 else ''}.        "
                          f"Current rate: {K['current']:.4f}", F['line'])
    for r_ in (8, 9):
        buf.height(r_, 21)

    T0 = 12                                            # table header row
    buf.merge(T0 - 1, 0, LAST, '  RATE PERIODS  —  which rate applied from which date to which date', F['sec'])
    buf.height(T0, 24)
    for c, h in enumerate(['No.', 'From Date', 'To Date', 'Rate', 'What happened']):
        buf.s(T0, c, h, F['th'])
    for i, p in enumerate(periods):
        r = T0 + 1 + i
        tag = {'INCREASED': 'inc', 'DECREASED': 'dec'}.get(p['status'], 'start')
        buf.height(r, 22)
        buf.n(r, 0, i + 1, F[tag + '_c'])
        buf.n(r, 1, p['from'], F[tag + '_d'])
        buf.n(r, 2, p['to'], F[tag + '_d'])
        buf.n(r, 3, p['rate'], F[tag + '_r'])
        buf.s(r, 4, what_happened(p, periods[i - 1]['rate'] if i else None), F[tag + '_s'])
    last = T0 + len(periods)
    ws.autofilter(T0, 0, last, LAST)
    buf.s(last + 2, 0, 'Red row = rate went up.   Blue row = rate went down.   Grey = first rate.', F['note'])
    buf.s(last + 3, 0, 'Brokerage = previous BRKAGE_RAT − current BRKAGE_RAT (oldest date first). '
                       'Rows without a valid date or rate are not included.', F['note'])
    ws.print_area(0, 0, last + 3, LAST)
    buf.flush()
