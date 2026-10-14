"""
Disk-backed temporary store (SQLite) for the matched source transactions.

Only rows that match a discrepancy combination are stored, one pickled row per record,
tagged with their group id.  Sorting (TRADE_DATE, then original row sequence) is done by
SQLite through an index, so Python never holds a whole group in memory.
"""
import os
import pickle
import shutil
import sqlite3
import tempfile

import config as C

_INVALID_DATE_KEY = 1.0e12        # sentinel -> rows with invalid dates sort LAST


class TempStore:
    def __init__(self, base_dir=None):
        self.dir = tempfile.mkdtemp(prefix='cams_tmp_', dir=base_dir or None)
        self.path = os.path.join(self.dir, 'tx.sqlite')
        self.conn = sqlite3.connect(self.path)
        c = self.conn
        c.execute('PRAGMA journal_mode=OFF')
        c.execute('PRAGMA synchronous=OFF')
        c.execute('PRAGMA temp_store=FILE')
        c.execute('PRAGMA cache_size=-65536')
        c.execute('''CREATE TABLE tx(
                        gid   INTEGER NOT NULL,
                        seq   INTEGER NOT NULL,   -- original row number in the source sheet
                        dkey  REAL    NOT NULL,   -- Excel serial date (sentinel if invalid)
                        dok   INTEGER NOT NULL,   -- 1 valid date, 0 invalid/missing
                        rate  REAL,               -- parsed BRKAGE_RAT (NULL unless rstat = 0)
                        rstat INTEGER NOT NULL,   -- 0 ok, 1 missing, 2 invalid
                        row   BLOB    NOT NULL)''')
        self._buf = []

    # ------------------------------------------------------------------
    def add(self, gid, seq, dkey, rate, rstat, row):
        dok = 0 if dkey is None else 1
        self._buf.append((gid, seq, _INVALID_DATE_KEY if dkey is None else dkey, dok, rate, rstat,
                          pickle.dumps(row, protocol=4)))
        if len(self._buf) >= C.SQLITE_BATCH:
            self.flush()

    def flush(self):
        if self._buf:
            self.conn.executemany('INSERT INTO tx VALUES (?,?,?,?,?,?,?)', self._buf)
            self.conn.commit()
            self._buf = []

    def finalize(self):
        self.flush()
        self.conn.execute('CREATE INDEX ix_g ON tx(gid, dkey, seq)')
        self.conn.commit()

    # ------------------------------------------------------------------
    def group_counts(self):
        """gid -> (rows, invalid_dates, missing_rates, invalid_rates)"""
        q = ('SELECT gid, COUNT(*), SUM(1-dok), SUM(rstat=1), SUM(rstat=2) FROM tx GROUP BY gid')
        return {g: (n, d or 0, m or 0, i or 0) for g, n, d, m, i in self.conn.execute(q)}

    def iter_light(self, gid):
        """Ordered (dok, rate) pairs without touching the big BLOB column."""
        cur = self.conn.execute(
            'SELECT dok, rate FROM tx WHERE gid=? ORDER BY dkey, seq', (gid,))
        while True:
            chunk = cur.fetchmany(20000)
            if not chunk:
                return
            yield from chunk

    def iter_rows(self, gid):
        """Ordered (dkey, dok, rate, rstat, row_list) for one group."""
        cur = self.conn.execute(
            'SELECT dkey, dok, rate, rstat, row FROM tx WHERE gid=? ORDER BY dkey, seq', (gid,))
        while True:
            chunk = cur.fetchmany(5000)
            if not chunk:
                return
            for dkey, dok, rate, rstat, blob in chunk:
                yield dkey, dok, rate, rstat, pickle.loads(blob)

    # ------------------------------------------------------------------
    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass
        shutil.rmtree(self.dir, ignore_errors=True)
