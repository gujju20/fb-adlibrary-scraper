"""Local desktop interface for the Meta Ad Library exporter."""
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import time
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

ROOT = Path(__file__).resolve().parent


def prepare_tcl():
    # Some Windows Python builds ship Tcl scripts as archives only.
    bundled = Path(sys.base_prefix) / 'tcl'
    for name, variable, folder, marker in (
        ('libtcl*.zip', 'TCL_LIBRARY', 'tcl_library', 'init.tcl'),
        ('libtk*.zip', 'TK_LIBRARY', 'tk_library', 'tk.tcl'),
    ):
        archives = sorted(bundled.glob(name))
        if archives and variable not in os.environ:
            target = ROOT / '.runtime' / archives[-1].stem
            if not (target / folder / marker).exists():
                with zipfile.ZipFile(archives[-1]) as archive:
                    archive.extractall(target)
            os.environ[variable] = str(target / folder)



from progress import duration, estimate


class App(tk.Tk):
    def __init__(self):
        prepare_tcl()
        super().__init__()
        self.title('Ad Library Studio')
        self.geometry('980x780')
        self.minsize(760, 650)
        self.configure(bg='#eef2f6')
        self.messages = queue.Queue()
        self.running = False
        self.output = None
        self.cancel_file = None
        self.started = 0
        self.last_event = {}
        self.terminal = None
        self.closing = False
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background='#eef2f6')
        style.configure('TLabel', background='#eef2f6', foreground='#172b4d', font=('Segoe UI', 10))
        style.configure('Title.TLabel', font=('Segoe UI', 25, 'bold'))
        style.configure('TButton', font=('Segoe UI', 10), padding=9)
        style.configure('TEntry', padding=8)
        style.configure('Horizontal.TProgressbar', background='#177e73', troughcolor='#dce4ed')
        body = ttk.Frame(self, padding=28)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='Ad Library Studio', style='Title.TLabel').pack(anchor='w')
        ttk.Label(body, text='Facebook Ad Library → Excel   •   Headless browser automation').pack(anchor='w', pady=(3, 22))
        ttk.Label(body, text='FACEBOOK AD LIBRARY LINK').pack(anchor='w')
        self.url = tk.StringVar()
        self.entry = ttk.Entry(body, textvariable=self.url)
        self.entry.pack(fill='x', pady=(6, 5))
        ttk.Label(body, text='Exports active ads, creative versions and available location transparency.').pack(anchor='w')
        settings = ttk.Frame(body)
        settings.pack(fill='x', pady=16)
        ttk.Label(settings, text='Ad limit (0 = all)').grid(row=0, column=0, sticky='w')
        self.limit = tk.StringVar(value='0')
        self.limit_entry = ttk.Entry(settings, textvariable=self.limit, width=12)
        self.limit_entry.grid(row=1, column=0, sticky='w', padx=(0, 20), pady=5)
        ttk.Label(settings, text='Load wait (seconds)').grid(row=0, column=1, sticky='w')
        self.wait = tk.StringVar(value='30')
        self.wait_entry = ttk.Entry(settings, textvariable=self.wait, width=12)
        self.wait_entry.grid(row=1, column=1, sticky='w', pady=5)
        ttk.Label(body, text='OUTPUT WORKBOOK').pack(anchor='w')
        output_row = ttk.Frame(body)
        output_row.pack(fill='x', pady=6)
        self.destination = tk.StringVar(value=str(ROOT / 'exports' / f'ads_{datetime.now():%Y%m%d_%H%M%S}.xlsx'))
        self.path_entry = ttk.Entry(output_row, textvariable=self.destination)
        self.path_entry.pack(side='left', fill='x', expand=True)
        self.browse = ttk.Button(output_row, text='Browse…', command=self.choose)
        self.browse.pack(side='left', padx=(8, 0))
        actions = ttk.Frame(body)
        actions.pack(fill='x', pady=12)
        self.start_button = ttk.Button(actions, text='Start export', command=self.start)
        self.start_button.pack(side='left')
        self.stop_button = ttk.Button(actions, text='Stop & save', command=self.stop, state='disabled')
        self.stop_button.pack(side='left', padx=8)
        self.open_button = ttk.Button(actions, text='Open Excel', command=self.open_output, state='disabled')
        self.open_button.pack(side='right')
        self.status = tk.StringVar(value='Ready to export')
        ttk.Label(body, textvariable=self.status, font=('Segoe UI', 13, 'bold')).pack(anchor='w', pady=(10, 8))
        self.bar = ttk.Progressbar(body, maximum=100)
        self.bar.pack(fill='x')
        self.metrics = tk.StringVar(value='Progress and time estimates will appear here.')
        ttk.Label(body, textvariable=self.metrics).pack(anchor='w', pady=8)
        ttk.Label(body, text='LIVE ACTIVITY').pack(anchor='w', pady=(10, 5))
        self.log = ScrolledText(body, height=12, bg='#142238', fg='#dce8f5', insertbackground='white', font=('Consolas', 10), relief='flat', wrap='word', state='disabled')
        self.log.pack(fill='both', expand=True)
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.after(150, self.poll)

    def choose(self):
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel workbook', '*.xlsx')])
        if path:
            self.destination.set(path)

    def append(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', f'[{datetime.now():%H:%M:%S}] {text}\n')
        if int(self.log.index('end-1c').split('.')[0]) > 2500:
            self.log.delete('1.0', '251.0')
        self.log.see('end')
        self.log.configure(state='disabled')

    def start(self):
        if self.running:
            return
        try:
            from meta_ads_export import source_settings
            url, _, _ = source_settings(self.url.get().strip())
            limit, wait = int(self.limit.get()), int(self.wait.get())
            if limit < 0 or wait < 1:
                raise ValueError('Use a nonnegative limit and a positive loading wait.')
            output = Path(self.destination.get()).expanduser().resolve()
            if output.suffix.lower() != '.xlsx':
                raise ValueError('Select an .xlsx output file.')
            output.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror('Check export settings', str(exc))
            return
        self.output = output
        self.cancel_file = output.with_name(f'.cancel-{time.time_ns()}')
        self.started = time.monotonic()
        self.running = True
        self.terminal = None
        self.last_event = {'phase': 'starting'}
        self.status.set('Starting browser…')
        self.bar.configure(mode='indeterminate')
        self.bar.start(15)
        for widget in (self.start_button, self.entry, self.limit_entry, self.wait_entry, self.path_entry, self.browse, self.open_button):
            widget.configure(state='disabled')
        self.stop_button.configure(state='normal')
        self.append('Starting export. Existing workbook data is preserved; a backup is created before updates.')
        command = [sys.executable, '-u', str(ROOT / 'meta_ads_export.py'), '--url', url,
                   '--output', str(output), '--max-ads', str(limit), '--load-timeout', str(wait),
                   '--headless', '--events', '--cancel-file', str(self.cancel_file)]
        threading.Thread(target=self.worker, args=(command,), daemon=True).start()

    def worker(self, command):
        try:
            env = dict(os.environ, PYTHONIOENCODING='utf-8')
            with subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                  text=True, encoding='utf-8', errors='replace', env=env,
                                  creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)) as process:
                for line in process.stdout:
                    self.messages.put(('line', line.rstrip()))
                code = process.wait()
            self.messages.put(('exit', code))
        except Exception as exc:
            self.messages.put(('line', str(exc)))
            self.messages.put(('exit', 1))

    def stop(self):
        if self.running:
            try:
                self.cancel_file.touch()
                self.stop_button.configure(state='disabled')
                self.append('Stop requested. Waiting for the current browser action, then saving available data.')
                self.status.set('Stopping after current browser action…')
            except OSError as exc:
                messagebox.showerror('Unable to stop', str(exc))

    def open_output(self):
        try:
            os.startfile(self.output)
        except OSError as exc:
            messagebox.showerror('Cannot open workbook', str(exc))

    def handle_event(self, event):
        self.last_event = event
        phase = event['phase']
        if phase == 'finished':
            self.terminal = event
        if phase in ('error', 'cancelled', 'saving', 'finished'):
            self.append(event.get('message', phase))
        if phase == 'scraping':
            self.bar.stop()
            self.bar.configure(mode='determinate', value=100 * event['completed'] / max(1, event['total']))
            self.status.set(f"Reading ad details · {event['completed']} / {event['total']} listing cards")
        elif phase == 'loading':
            self.status.set(f"Loading all results · {event['loaded']} ad IDs discovered")
        elif phase == 'saving':
            self.bar.stop()
            self.bar.configure(mode='indeterminate')
            self.bar.start(15)
            self.status.set('Saving workbook…')

    def finish(self, code):
        self.running = False
        self.bar.stop()
        success = code == 0 and self.terminal and self.terminal['success']
        self.bar.configure(mode='determinate', value=100 if success else 0)
        if self.terminal:
            e = self.terminal
            self.status.set(f"{'Export complete' if success else 'Partial export saved'} · {e['ads']} ads · {e['warnings']} with incomplete details")
            self.open_button.configure(state='normal')
        else:
            self.status.set('Export failed — see activity log')
        self.metrics.set(f'Total elapsed: {duration(time.monotonic() - self.started)}')
        for widget in (self.start_button, self.entry, self.limit_entry, self.wait_entry, self.path_entry, self.browse):
            widget.configure(state='normal')
        self.stop_button.configure(state='disabled')
        self.cancel_file.unlink(missing_ok=True)
        if self.closing:
            self.destroy()

    def poll(self):
        for _ in range(200):
            try:
                kind, payload = self.messages.get_nowait()
            except queue.Empty:
                break
            if kind == 'exit':
                self.finish(payload)
                if self.closing:
                    return
            elif payload.startswith('@event '):
                try:
                    self.handle_event(json.loads(payload[7:]))
                except (ValueError, KeyError):
                    self.append(payload)
            else:
                self.append(payload)
        if self.running:
            e = self.last_event
            eta = None
            if e['phase'] == 'scraping':
                eta = estimate(e['completed'], e['total'], e['elapsed'])
            note = f'Estimated detail time remaining: ~{duration(eta)}' if eta is not None else 'ETA: calculating / total work not yet known'
            self.metrics.set(f'Elapsed: {duration(time.monotonic() - self.started)}    •    {note}')
        self.after(150, self.poll)

    def close(self):
        if self.running:
            self.closing = True
            self.stop()
        else:
            self.destroy()


if __name__ == '__main__':
    App().mainloop()
