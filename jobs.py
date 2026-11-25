"""Single-instance durable job manager. Browser and ZIP work run in child processes."""
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import json
import os
from pathlib import Path
import queue
import re
import shutil
import signal
import subprocess
import sys
import threading
import time
import uuid

ROOT = Path(__file__).resolve().parent
BUSY = {'queued', 'running', 'stopping'}


class JobManager:
    def __init__(self, data_dir, workers=1):
        self.root = Path(data_dir).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.lock = threading.RLock()
        self.jobs = {}
        self.executor = ThreadPoolExecutor(max_workers=workers, thread_name_prefix='exports')
        self.shutdown_event = threading.Event()
        self.max_jobs = int(os.getenv('MAX_PENDING_JOBS', '10'))
        self.timeout = int(os.getenv('JOB_TIMEOUT_SECONDS', '21600'))
        self.retention = int(os.getenv('RETENTION_HOURS', '24')) * 3600
        for path in self.root.glob('*/job.json'):
            try:
                job = json.loads(path.read_text(encoding='utf-8'))
                if not re.fullmatch(r'[a-f0-9]{32}', job['id']) or path.parent.name != job['id']:
                    continue
                for kind in ('scrape', 'assets'):
                    if job[kind]['status'] in BUSY:
                        job[kind].update(status='interrupted', message='Server restarted. Start a new export or rebuild the ZIP.')
                self.jobs[job['id']] = job
                self._save(job)
            except (OSError, ValueError, KeyError):
                continue
        self.cleanup()

    def directory(self, job_id):
        return self.root / job_id

    def _save(self, job):
        path = self.directory(job['id']) / 'job.json'
        temporary = path.with_suffix('.tmp')
        temporary.write_text(json.dumps(job), encoding='utf-8')
        temporary.replace(path)

    def _log(self, job, message):
        job['logs'].append({'time': time.time(), 'text': str(message)[:2000]})
        job['logs'] = job['logs'][-300:]

    def cleanup(self):
        # Only remove this application's expired, terminal job directories.
        with self.lock:
            now = time.time()
            for key, job in list(self.jobs.items()):
                if (now - job['updated'] > self.retention
                        and all(job[k]['status'] not in BUSY for k in ('scrape', 'assets'))):
                    path = self.directory(key).resolve()
                    if path.parent == self.root and re.fullmatch(r'[a-f0-9]{32}', path.name):
                        shutil.rmtree(path)
                        del self.jobs[key]

    def _capacity(self, owner):
        pending = [job for job in self.jobs.values() if any(job[k]['status'] in BUSY for k in ('scrape', 'assets'))]
        if len(pending) >= self.max_jobs or sum(job['owner'] == owner for job in pending) >= 2:
            raise ValueError('The export queue is full. Wait for an existing job to finish.')
        if shutil.disk_usage(self.root).free < 512 * 1024**2:
            raise ValueError('Server storage is low. Remove old exports or increase storage.')

    def create(self, owner, url, code, country, limit):
        with self.lock:
            self.cleanup()
            self._capacity(owner)
            key = uuid.uuid4().hex
            self.directory(key).mkdir()
            now = time.time()
            job = dict(id=key, owner=owner, url=url, country=country, code=code, limit=limit,
                       created=now, updated=now, logs=[],
                       scrape={'status': 'queued', 'phase': 'queued', 'message': 'Waiting for a browser worker'},
                       assets={'status': 'idle', 'phase': 'idle'})
            self.jobs[key] = job
            self._log(job, f'Queued export for {country}; limit: {limit or "all matching ads"}')
            self._save(job)
            self.executor.submit(self._run, key, 'scrape')
            return key

    def _owned(self, key, owner):
        job = self.jobs.get(key)
        if not job or job['owner'] != owner:
            raise KeyError('Export not found')
        return job

    def get(self, key, owner):
        with self.lock:
            job = deepcopy(self._owned(key, owner))
        job.pop('owner')
        folder = self.directory(key)
        job['xlsx_ready'] = job['scrape']['status'] not in BUSY and (folder / 'ads.xlsx').exists()
        job['zip_ready'] = job['assets']['status'] not in BUSY and (folder / 'assets.zip').exists()
        job['server_time'] = time.time()
        job['expires_at'] = job['updated'] + self.retention
        for kind in ('scrape', 'assets'):
            state = job[kind]
            started = state.get('started')
            state['elapsed_seconds'] = max(0, state.get('ended', time.time()) - started) if started else 0
            complete, total = state.get('completed', 0), state.get('total', 0)
            state['percent'] = round(100 * complete / total, 1) if total else None
            if state['status'] == 'complete':
                state['percent'] = 100
            # Estimate only a measured phase; freeze no stale countdown during stalls.
            elapsed = state.get('elapsed', 0)
            sample_age = time.time() - state.get('sample_at', time.time())
            eta = elapsed / complete * (total - complete) if complete >= 3 and total > complete else None
            state['eta_seconds'] = max(0, eta - sample_age) if eta and sample_age < min(eta, 60) and state['status'] == 'running' else None
        return job

    def list(self, owner):
        with self.lock:
            keys = [job['id'] for job in sorted(self.jobs.values(), key=lambda j: j['created'], reverse=True) if job['owner'] == owner]
        return [self.get(key, owner) for key in keys[:30]]

    def start_assets(self, key, owner):
        with self.lock:
            job = self._owned(key, owner)
            if job['assets']['status'] in BUSY:
                return
            if job['scrape']['status'] in BUSY or not (self.directory(key) / 'ads.xlsx').exists():
                raise ValueError('Wait for the workbook to finish before downloading assets')
            if job['assets']['status'] in ('complete', 'partial') and (self.directory(key) / 'assets.zip').exists():
                return
            self._capacity(owner)
            (self.directory(key) / 'assets.cancel').unlink(missing_ok=True)
            job['assets'] = {'status': 'queued', 'phase': 'queued', 'message': 'Waiting to prepare ZIP'}
            job['updated'] = time.time()
            self._log(job, 'Preparing all available ad and version media for ZIP download')
            self._save(job)
            self.executor.submit(self._run, key, 'assets')

    def cancel(self, key, owner):
        with self.lock:
            job = self._owned(key, owner)
            for kind in ('scrape', 'assets'):
                if job[kind]['status'] in BUSY:
                    (self.directory(key) / f'{kind}.cancel').touch()
                    if job[kind]['status'] == 'queued':
                        job[kind].update(status='cancelled', phase='cancelled', ended=time.time())
                    else:
                        job[kind].update(status='stopping', message='Saving after the current network action')
            self._log(job, 'Stop requested. Available output will be saved when possible.')
            job['updated'] = time.time()
            self._save(job)

    def _event(self, key, kind, event):
        with self.lock:
            job = self.jobs[key]
            state = job[kind]
            phase = event.get('phase', '')
            if phase == 'log':
                self._log(job, event.get('message', ''))
            else:
                # Never publish local paths from child-process events.
                state.update({k: v for k, v in event.items() if k not in ('path', 'status')})
                state['sample_at'] = time.time()
                if event.get('message'):
                    self._log(job, event['message'])
            job['updated'] = time.time()
            self._save(job)

    @staticmethod
    def _kill(process):
        if process.poll() is not None:
            return
        if os.name == 'nt':
            subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'], capture_output=True)
        else:
            try:
                os.killpg(process.pid, signal.SIGTERM)
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass

    def _run(self, key, kind):
        folder = self.directory(key)
        flag = folder / f'{kind}.cancel'
        with self.lock:
            job = self.jobs[key]
            if job[kind]['status'] == 'cancelled' or self.shutdown_event.is_set():
                return
            job[kind].update(status='running', phase='starting', started=time.time())
            self._save(job)
            if kind == 'scrape':
                command = [sys.executable, '-u', str(ROOT / 'meta_ads_export.py'), '--headless', '--events',
                           '--url', job['url'], '--max-ads', str(job['limit']), '--output', str(folder / 'ads.xlsx'),
                           '--load-timeout', '30', '--load-retries', '6', '--cancel-file', str(flag)]
            else:
                command = [sys.executable, '-u', str(ROOT / 'asset_zip.py'), '--xlsx', str(folder / 'ads.xlsx'),
                           '--output', str(folder / 'assets.zip'), '--country', 'ALL', '--cancel-file', str(flag)]
        process = None
        code = 1
        forced = False
        try:
            options = {'creationflags': subprocess.CREATE_NO_WINDOW} if os.name == 'nt' else {'start_new_session': True}
            process = subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       text=True, encoding='utf-8', errors='replace',
                                       env=dict(os.environ, PYTHONIOENCODING='utf-8'), **options)
            lines = queue.Queue(maxsize=1000)
            def reader():
                try:
                    for line in process.stdout:
                        lines.put(line)
                finally:
                    lines.put(None)
            thread = threading.Thread(target=reader, daemon=True)
            thread.start()
            deadline = time.monotonic() + self.timeout
            stop_started = None
            while True:
                if (self.shutdown_event.is_set() or time.monotonic() > deadline) and not flag.exists():
                    flag.touch()
                    self._event(key, kind, {'phase': 'stopping', 'message': 'Server shutdown or worker time limit reached; saving available output.'})
                if flag.exists():
                    stop_started = stop_started or time.monotonic()
                    if time.monotonic() - stop_started > 45:
                        forced = True
                        self._kill(process)
                try:
                    line = lines.get(timeout=.5)
                except queue.Empty:
                    if process.poll() is not None and not thread.is_alive():
                        break
                    continue
                if line is None:
                    break
                line = line.strip()
                if line.startswith('@event '):
                    try:
                        self._event(key, kind, json.loads(line[7:]))
                        continue
                    except (ValueError, TypeError):
                        pass
                # Paths and full source URLs need not appear in the browser activity log.
                if line and not str(folder) in line:
                    with self.lock:
                        self._log(self.jobs[key], line)
                        self._save(self.jobs[key])
            code = process.wait(timeout=10)
        except Exception as exc:
            with self.lock:
                self._log(self.jobs[key], f'{kind} failed: {exc}')
            if process:
                self._kill(process)
        finally:
            if process and process.stdout:
                process.stdout.close()
            with self.lock:
                job = self.jobs[key]
                state = job[kind]
                output = folder / ('ads.xlsx' if kind == 'scrape' else 'assets.zip')
                good = code == 0 and state.get('success') is True and output.exists()
                status = 'complete' if good else ('partial' if output.exists() else 'failed')
                if flag.exists():
                    status = 'cancelled'
                message = state.get('message', '')
                if forced:
                    message = 'Worker stopped after cancellation timeout. Only saved checkpoints are available.'
                elif code != 0 and not message:
                    message = 'Worker failed; see activity log.'
                state.update(status=status, ended=time.time(), message=message)
                job['updated'] = time.time()
                self._log(job, f'{kind.title()} {status}. {message}')
                self._save(job)

    def close(self):
        self.shutdown_event.set()
        self.executor.shutdown(wait=True, cancel_futures=True)
