"""Optional publisher updates. The supplied configuration disables network requests."""

import base64
from contextlib import contextmanager
from dataclasses import asdict, dataclass
import hashlib
import json
import msvcrt
import os
from pathlib import Path
from queue import Empty, Queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk
from urllib.error import URLError
from urllib.parse import unquote, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

REPO = 'WadaaDev/Wadaa-PDF-Manager'
RELEASE_URL = f'https://api.github.com/repos/{REPO}/releases/latest'
ASSET_PATH = f'/{REPO}/releases/download/'
MAX_SIZE = 512 * 1024 * 1024
NUMBER = r'(0|[1-9][0-9]*)'
VERSION = re.compile(rf'v?{NUMBER}\.{NUMBER}(?:\.{NUMBER})?'
                     r'(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?'
                     r'(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?')


def version_key(value):
    # This application also uses 1.0 as an alias for 1.0.0.
    match = VERSION.fullmatch(value)
    if not match:
        raise ValueError('Invalid release version.')
    major, minor, patch, prerelease = match.groups()
    parts = []
    for part in prerelease.split('.') if prerelease else ():
        if part.isdigit() and len(part) > 1 and part.startswith('0'):
            raise ValueError('Invalid numeric prerelease identifier.')
        parts.append((0, int(part)) if part.isdigit() else (1, part))
    return int(major), int(minor), int(patch or 0), prerelease is None, tuple(parts)


def allowed_url(url, asset):
    parsed = urlsplit(url)
    if (parsed.scheme != 'https' or parsed.username or parsed.password or parsed.fragment
            or parsed.port not in (None, 443)):
        return False
    if not asset:
        return url == RELEASE_URL
    if parsed.hostname == 'github.com':
        parts = unquote(parsed.path[len(ASSET_PATH):]).split('/')
        return (parsed.path.startswith(ASSET_PATH) and not parsed.query and len(parts) == 2
                and all(part not in ('', '.', '..') and '\\' not in part for part in parts))
    return parsed.hostname == 'release-assets.githubusercontent.com'


class ReleaseRedirect(HTTPRedirectHandler):
    def __init__(self, asset):
        self.asset = asset

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not allowed_url(newurl, self.asset):
            raise ValueError('Unexpected redirect from the official release source.')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def open_https(url, *, asset=False, headers=None):
    if not allowed_url(url, asset):
        raise ValueError('Unexpected update source.')
    request = Request(url, headers={'User-Agent': 'Wadaa-PDF-Manager-Updater',
                                   'Accept-Encoding': 'identity', **(headers or {})})
    return build_opener(ReleaseRedirect(asset)).open(request, timeout=12)


@dataclass(frozen=True)
class ReleaseAsset:
    version: str
    asset_id: int
    url: str
    size: int
    sha256: str


def select_release(data, current):
    tag = data['tag_name']
    if data.get('draft') or data.get('prerelease') or not version_key(tag)[3]:
        return None
    if version_key(tag) <= version_key(current):
        return None
    name = f'Wadaa PDF Manager Setup v{tag.removeprefix("v")}.exe'
    # GitHub stores spaces in asset names as dots, so accept both spellings.
    names = {name, name.replace(' ', '.')}
    matches = [item for item in data['assets'] if item.get('name') in names]
    if len(matches) != 1:
        raise ValueError('The release must contain exactly one matching versioned installer.')
    item = matches[0]
    url, aid, size, digest = (item.get(key) for key in ('browser_download_url', 'id', 'size', 'digest'))
    if (type(aid) is not int or aid <= 0 or type(size) is not int or not 0 < size <= MAX_SIZE
            or not isinstance(url, str) or not allowed_url(url, True)
            or urlsplit(url).hostname != 'github.com'
            or unquote(urlsplit(url).path) not in {f'{ASSET_PATH}{tag}/{n}' for n in names}
            or not isinstance(digest, str) or not re.fullmatch(r'sha256:[a-fA-F0-9]{64}', digest)):
        raise ValueError('The official installer lacks a valid identity, size or SHA-256 digest.')
    return ReleaseAsset(tag, aid, url, size, digest[7:].lower())


@contextmanager
def update_lock(directory):
    directory.mkdir(parents=True, exist_ok=True)
    with (directory / 'updater.lock').open('a+b') as lock:
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise ValueError('Another Wadaa PDF Manager update is active.') from exc
        try:
            yield
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


class Cancelled(Exception):
    pass


def sha256(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def download(asset, directory, cancel, progress, opener=open_https):
    partial = directory / 'installer.part'
    manifest = directory / 'partial.json'
    target = directory / f'setup-{asset.asset_id}-{asset.sha256}.exe'
    if target.is_file() and target.stat().st_size == asset.size and sha256(target) == asset.sha256:
        return target
    try:
        previous = json.loads(manifest.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        previous = None
    if previous != asdict(asset):
        partial.unlink(missing_ok=True)
        manifest.write_text(json.dumps(asdict(asset)), encoding='utf-8')
    offset = partial.stat().st_size if partial.exists() else 0
    if offset > asset.size:
        partial.unlink()
        offset = 0

    def check_cancel():
        if cancel.is_set():
            raise Cancelled

    check_cancel()
    if offset < asset.size:
        headers = {'Range': f'bytes={offset}-'} if offset else {}
        with opener(asset.url, asset=True, headers=headers) as response:
            if response.status == 200:
                offset = 0
            elif not (offset and response.status == 206 and response.headers.get('Content-Range')
                      == f'bytes {offset}-{asset.size - 1}/{asset.size}'):
                partial.unlink(missing_ok=True)
                raise ValueError('The server returned an invalid range. Retry to start over.')
            if response.headers.get('Content-Encoding', 'identity').lower() != 'identity':
                raise ValueError('Unexpected installer response encoding.')
            length = response.headers.get('Content-Length')
            if length is not None and length != str(asset.size - offset):
                partial.unlink(missing_ok=True)
                raise ValueError('Installer response size does not match the release.')
            with partial.open('ab' if offset else 'wb') as stream:
                last_percent = -1
                while chunk := response.read(128 * 1024):
                    check_cancel()
                    offset += len(chunk)
                    if offset > asset.size:
                        break
                    stream.write(chunk)
                    percent = offset * 100 // asset.size
                    if percent != last_percent:
                        progress(percent)
                        last_percent = percent
        if offset > asset.size:
            partial.unlink(missing_ok=True)
    if offset != asset.size:
        raise ValueError('Download incomplete. Retry to resume.')
    check_cancel()
    if sha256(partial) != asset.sha256:
        partial.unlink()
        raise ValueError('The installer SHA-256 did not match. It was discarded.')
    check_cancel()
    os.replace(partial, target)
    return target


# Wait outside the app directory so the installed executable can be replaced.
HANDOFF = r'''
$ErrorActionPreference = 'Stop'
$gate = [Threading.Mutex]::new($false, 'Local\WadaaPDFManagerInstaller')
$held = $false
try {
    try { $held = $gate.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held = $true }
    if (-not $held) { exit 2 }
    $parent = Get-Process -Id ([int]$env:WADAA_PARENT) -ErrorAction SilentlyContinue
    if ($parent) {
        $null = $parent.Handle
        if (-not $parent.WaitForExit(60000)) { exit 3 }
    }
    if ((Get-FileHash -LiteralPath $env:WADAA_INSTALLER -Algorithm SHA256).Hash -ne $env:WADAA_SHA256) {
        throw 'The installer failed the final SHA-256 check.'
    }
    $arguments = '/SILENT /NORESTART /WADAAUPDATE=1 /DIR="' + $env:WADAA_DEST + '"'
    $setup = Start-Process -FilePath $env:WADAA_INSTALLER -PassThru -ArgumentList $arguments
    $setup.WaitForExit()
    if ($setup.ExitCode -ne 0) { throw "Setup exited with code $($setup.ExitCode)." }
    Remove-Item -LiteralPath $env:WADAA_INSTALLER -ErrorAction SilentlyContinue
} catch {
    try { Start-Process -FilePath $env:WADAA_APP } catch {}
    Add-Type -AssemblyName System.Windows.Forms
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Wadaa PDF Manager update')
    exit 1
} finally {
    if ($held) { $gate.ReleaseMutex() }
    $gate.Dispose()
}
'''


def handoff(installer, asset):
    app = Path(sys.executable).resolve()
    shell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT='1',
               WADAA_PARENT=str(os.getpid()), WADAA_INSTALLER=str(installer),
               WADAA_SHA256=asset.sha256, WADAA_APP=str(app), WADAA_DEST=str(app.parent))
    encoded = base64.b64encode(HANDOFF.encode('utf-16-le')).decode('ascii')
    import ctypes
    ctypes.windll.kernel32.SetDllDirectoryW(None)
    try:
        subprocess.Popen([str(shell), '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
                         cwd=shell.parent, env=env, creationflags=subprocess.CREATE_NO_WINDOW)
    finally:
        if getattr(sys, 'frozen', False):
            ctypes.windll.kernel32.SetDllDirectoryW(sys._MEIPASS)


class UpdateController:
    def __init__(self, app, config, data_directory, version):
        self.app = app
        self.directory = data_directory / 'updates'
        try:
            data = json.loads(config.read_text(encoding='utf-8'))
            self.version = version
            version_key(self.version)
            self.enabled = data.get('enabled') is True
        except (OSError, ValueError, TypeError, KeyError):
            self.enabled = False
        self.events = Queue()
        self.cancel = threading.Event()
        self.active = False
        self.retry = None
        self.poll = None
        self.dialog = None
        self.pending = None

    def apply_pending(self):
        # Called when the app closes: hand the verified installer to the silent installer.
        if self.pending is None:
            return
        path, asset = self.pending
        self.pending = None
        try:
            handoff(path, asset)
        except (OSError, KeyError):
            pass

    def check(self, manual=False):
        if not self.enabled:
            if manual:
                self.app.info('Updates', 'Online updates are disabled in this build.')
            return
        if self.active:
            if manual:
                self.app.status.set(self.app.tr('An update check or download is already running.'))
            return
        if self.retry is not None:
            self.app.after_cancel(self.retry)
            self.retry = None

        def query():
            with open_https(RELEASE_URL, headers={'Accept': 'application/vnd.github+json'}) as response:
                data = response.read(2 * 1024 * 1024 + 1)
            if len(data) > 2 * 1024 * 1024:
                raise ValueError('Unexpected release metadata size.')
            return select_release(json.loads(data), self.version)

        self.start(query, lambda asset: self.offer(asset, manual), manual, retry=True)

    def start(self, operation, complete, manual, retry=False):
        self.active = True
        self.cancel.clear()

        def worker():
            try:
                with update_lock(self.directory):
                    result = operation()
                self.events.put(('result', result))
            except Cancelled:
                self.events.put(('cancel', None))
            except Exception as exc:
                self.events.put(('error', (str(exc), isinstance(exc, (URLError, TimeoutError)))))

        def poll():
            self.poll = None
            try:
                while True:
                    kind, value = self.events.get_nowait()
                    if kind == 'progress':
                        bar = getattr(self.app, 'update_bar', None)
                        if bar is not None:
                            bar['value'] = value
                        if self.dialog is not None:
                            self.progress['value'] = value
                        continue
                    bar = getattr(self.app, 'update_bar', None)
                    if bar is not None:
                        bar['value'] = 0
                    self.active = False
                    if self.dialog:
                        self.dialog.destroy()
                        self.dialog = None
                    if kind == 'error':
                        error, transient = value
                        if retry and transient:
                            self.retry = self.app.after(60_000, self.check)
                        if manual:
                            self.app.themed_dialog('Updates', error)
                    elif kind == 'result' and not self.cancel.is_set():
                        complete(value)
                    return
            except Empty:
                self.poll = self.app.after(100, poll)

        threading.Thread(target=worker, daemon=True).start()
        self.poll = self.app.after(100, poll)

    def offer(self, asset, manual):
        if asset is None:
            if manual:
                self.app.info('Updates', 'You are using the latest version.')
            return
        if not getattr(sys, 'frozen', False) or not any(Path(sys.executable).parent.glob('unins*.dat')):
            if manual:
                self.app.info('Updates', self.app.tr(
                    'Version {version} is available. Install the app with Setup to use automatic upgrades.',
                    version=asset.version))
            return
        if not manual:
            # Automatic mode: download quietly in the background, install when the app closes.
            self.start(lambda: download(asset, self.directory, self.cancel,
                                       lambda p: self.events.put(('progress', p))),
                       lambda path: setattr(self, 'pending', (path, asset)), manual=False)
            return
        if not self.app.themed_dialog('Update available', self.app.tr(
            'Download version {version}?', version=asset.version), confirm=True):
            return
        self.dialog = tk.Toplevel(self.app)
        def refresh_dialog():
            self.dialog.title(self.app.tr('Downloading update'))
            self.dialog.configure(bg=self.app.cget('background'))

        self.dialog._refresh_appearance = refresh_dialog
        self.dialog.transient(self.app)
        self.progress = ttk.Progressbar(self.dialog, maximum=100, length=240,
                                        style='Dark.Horizontal.TProgressbar')
        self.progress.pack(padx=18, pady=12)
        self.app.button(self.dialog, 'Cancel', self.cancel.set).pack(pady=(0, 12))
        self.dialog.protocol('WM_DELETE_WINDOW', self.cancel.set)
        self.app.style_widgets(self.dialog)
        self.start(lambda: download(asset, self.directory, self.cancel,
                                   lambda p: self.events.put(('progress', p))),
                   lambda path: self.install(path, asset), manual=True)

    def install(self, path, asset):
        if self.app.busy:
            self.app.info('Updates', 'Wait for the PDF merge to finish before installing.')
            return
        if not self.app.themed_dialog('Install update', 'Close Wadaa PDF Manager and install the update?',
                                      confirm=True):
            return
        if not self.app.save_session():
            self.app.themed_dialog('Updates', 'The session could not be saved. The update was postponed.')
            return
        try:
            handoff(path, asset)
        except (OSError, KeyError) as exc:
            self.app.themed_dialog('Updates', str(exc))
            return
        self.close()
        self.app.destroy()

    def close(self):
        self.cancel.set()
        for callback in (self.retry, self.poll):
            if callback is not None:
                self.app.after_cancel(callback)
