"""
Tkinter GUI — two independent tabs.
  STEP 1 — CAMS DISCREPANCY CHECKER
  STEP 2 — INVESTOR-WISE BROKERAGE AUTOMATION
Each step runs in its own background thread; the GUI is only updated from the Tk main thread
(workers talk to it through a queue), so the window never freezes.
"""
import os
import queue
import subprocess
import sys
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import config as C
import reporting


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


class StepPanel(ttk.Frame):
    def get_progress(self):
        return float(self.bar['value'])

    """Generic panel: N file pickers + output picker + run button + progress + log."""

    def __init__(self, master, title, inputs, out_default, out_label, run_label, worker_fn, make_output):
        super().__init__(master, padding=12)
        self.worker_fn = worker_fn
        self.make_output = make_output
        self.q = queue.Queue()
        self.running = False
        self.t0 = 0.0
        self.last_output = ''
        self.in_vars = []
        self.columnconfigure(1, weight=1)
        self.rowconfigure(len(inputs) + 4, weight=1)

        ttk.Label(self, text=title, font=('Segoe UI', 12, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))
        self.buttons = []
        for i, (label, ftypes) in enumerate(inputs, start=1):
            v = tk.StringVar()
            self.in_vars.append(v)
            ttk.Label(self, text=label).grid(row=i, column=0, sticky='w', pady=4)
            ttk.Entry(self, textvariable=v).grid(row=i, column=1, sticky='ew', padx=8, pady=4)
            b = ttk.Button(self, text='Browse…', command=lambda v=v, l=label: self._pick_in(v, l))
            b.grid(row=i, column=2, pady=4)
            self.buttons.append(b)

        r = len(inputs) + 1
        self.out_var = tk.StringVar(value=out_default)
        ttk.Label(self, text=out_label).grid(row=r, column=0, sticky='w', pady=4)
        ttk.Entry(self, textvariable=self.out_var).grid(row=r, column=1, sticky='ew', padx=8, pady=4)
        b = ttk.Button(self, text='Save as…', command=self._pick_out)
        b.grid(row=r, column=2, pady=4)
        self.buttons.append(b)

        row = ttk.Frame(self)
        row.grid(row=r + 1, column=0, columnspan=3, sticky='ew', pady=6)
        row.columnconfigure(1, weight=1)
        self.run_btn = ttk.Button(row, text=run_label, command=self.start)
        self.run_btn.grid(row=0, column=0, padx=(0, 10))
        self.bar = ttk.Progressbar(row, mode='determinate', maximum=100)
        self.bar.grid(row=0, column=1, sticky='ew')
        self.status = tk.StringVar(value='Ready.')
        self.elapsed = tk.StringVar()
        ttk.Label(self, textvariable=self.status).grid(row=r + 2, column=0, columnspan=2, sticky='w')
        ttk.Label(self, textvariable=self.elapsed).grid(row=r + 2, column=2, sticky='e')

        self.logbox = tk.Text(self, height=14, wrap='none', font=('Consolas', 9), state='disabled')
        self.logbox.grid(row=r + 3, column=0, columnspan=3, sticky='nsew', pady=6)
        sb = ttk.Scrollbar(self, command=self.logbox.yview)
        sb.grid(row=r + 3, column=3, sticky='ns')
        self.logbox.configure(yscrollcommand=sb.set)

        bottom = ttk.Frame(self)
        bottom.grid(row=r + 4, column=0, columnspan=3, sticky='e')
        self.open_btn = ttk.Button(bottom, text='Open Output Folder', state='disabled',
                                   command=lambda: _open_path(os.path.dirname(os.path.abspath(self.last_output))))
        self.open_btn.pack(side='left', padx=4)
        self.rowconfigure(r + 3, weight=1)

    # ---------------------------------------------------------------- pickers
    def _pick_in(self, var, label):
        p = filedialog.askopenfilename(title=label, filetypes=[('Excel files', '*.xlsx *.xlsm'), ('All', '*.*')])
        if p:
            var.set(os.path.normpath(p))
            if not self.out_var.get().strip() or self.out_var.get().startswith('<auto>'):
                self.out_var.set(os.path.normpath(self.make_output(p)))

    def _pick_out(self):
        p = filedialog.asksaveasfilename(title='Save output Excel as', defaultextension='.xlsx',
                                         initialfile=os.path.basename(self.out_var.get() or 'output.xlsx'),
                                         filetypes=[('Excel workbook', '*.xlsx')])
        if p:
            self.out_var.set(os.path.normpath(p))

    # ---------------------------------------------------------------- run
    def _log(self, text):
        self.logbox.configure(state='normal')
        self.logbox.insert('end', text + '\n')
        self.logbox.see('end')
        self.logbox.configure(state='disabled')

    def _set_running(self, flag):
        self.running = flag
        st = 'disabled' if flag else 'normal'
        self.run_btn.configure(state=st)
        for b in self.buttons:
            b.configure(state=st)
        if flag:
            self.open_btn.configure(state='disabled')

    def start(self):
        ins = [v.get().strip() for v in self.in_vars]
        out = self.out_var.get().strip()
        if not all(ins) or not all(os.path.isfile(p) for p in ins):
            messagebox.showwarning('Input needed', 'Please select all the required input Excel files.')
            return
        if not out:
            messagebox.showwarning('Output needed', 'Please choose where to save the output Excel.')
            return
        self.logbox.configure(state='normal')
        self.logbox.delete('1.0', 'end')
        self.logbox.configure(state='disabled')
        self.bar['value'] = 0
        self.t0 = time.time()
        self.last_output = out
        self.status.set('Starting…')
        self._set_running(True)
        threading.Thread(target=self._worker, args=(ins, out), daemon=True).start()

    def _worker(self, ins, out):
        try:
            reporting.set_reporter(lambda m: self.q.put(('log', m)),
                                   lambda p, t='': self.q.put(('prog', p, t)))
            self.worker_fn(*ins, out)
            self.q.put(('done', out))
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
                    self.bar['value'] = m[1]
                    if m[2]:
                        self.status.set(m[2])
                elif m[0] == 'done':
                    self.bar['value'] = 100
                    self.status.set('✓ Finished.')
                    self._set_running(False)
                    self.open_btn.configure(state='normal')
                    messagebox.showinfo('Done', f'Finished. Output saved to:\n{m[1]}')
                elif m[0] == 'error':
                    self.status.set('✗ Failed — see log.')
                    self._log('\nERROR: ' + m[1])
                    if m[2]:
                        self._log(m[2])
                    self._set_running(False)
                    messagebox.showerror('Error', m[1])
        except queue.Empty:
            pass
        if self.running:
            self.elapsed.set(f'Elapsed: {int(time.time() - self.t0)} s')
        self.after(150, self.pump)


def launch():
    """Plain ttk fallback used automatically when customtkinter is not installed."""
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    import discrepancy_engine
    import investor_engine

    root = tk.Tk()
    root.title(C.APP_NAME)
    root.geometry('900x640')
    root.minsize(780, 560)
    ttk.Label(root, text=C.APP_NAME, font=('Segoe UI', 15, 'bold')).pack(anchor='w', padx=14, pady=(10, 0))
    nb = ttk.Notebook(root)
    nb.pack(fill='both', expand=True, padx=10, pady=10)

    p1 = StepPanel(
        nb, 'STEP 1 — CAMS DISCREPANCY CHECKER',
        inputs=[('Original CAMS Commission Excel:', None)],
        out_default='', out_label='Output Excel (Discrepancy_Result):', run_label='▶  Generate Discrepancy Report',
        worker_fn=discrepancy_engine.run,
        make_output=lambda p: os.path.join(os.path.dirname(p), 'Discrepancy_Result.xlsx'))
    p2 = StepPanel(
        nb, 'STEP 2 — INVESTOR-WISE BROKERAGE AUTOMATION',
        inputs=[('Discrepancy Report Excel (from Step 1):', None),
                ('Source CAMS Transaction Excel (large):', None)],
        out_default='', out_label='Output Excel (Investor_Wise_Final):', run_label='▶  Generate Investor-wise Sheets',
        worker_fn=investor_engine.run,
        make_output=lambda p: os.path.join(os.path.dirname(p), 'Investor_Wise_Final.xlsx'))
    nb.add(p1, text='  Step 1 — Discrepancy Checker  ')
    nb.add(p2, text='  Step 2 — Investor-wise Automation  ')
    p1.pump()
    p2.pump()
    root.mainloop()


if __name__ == '__main__':
    launch()
