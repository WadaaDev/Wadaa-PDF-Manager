"""Wadaa PDF Manager — dark Windows GUI."""
from __future__ import annotations

import io
import json
import shutil
import subprocess
import tempfile
import os
import platform
import re
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlencode
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageOps, ImageSequence, ImageDraw, ImageTk
import tkinter.font as tkfont
import fitz  # PyMuPDF: PDF page preview rendering
from pypdf import PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
from updater import REPO, UpdateController

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp'}
OFFICE_EXTS = {'.doc', '.docx', '.docm', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx', '.pptm'}
EXTS = {'.pdf'} | IMAGE_EXTS | OFFICE_EXTS
BG = '#212121'
SURFACE = '#2f2f2f'
HOVER = '#3b3b3b'
TEXT = '#f4f4f4'
MUTED = '#b4b4b4'
ACCENT = '#d4d4d4'
SELECT = '#444b53'
DONATE = 'https://linktr.ee/WadaaTips'
SUPPORT_EMAIL = 'wadaadev.support@gmail.com'
REPOSITORY_URL = f'https://github.com/{REPO}'
BUG_REPORT_URL = f'{REPOSITORY_URL}/issues'


def resource_path(relative: str) -> Path:
    return Path(getattr(sys, '_MEIPASS', Path(__file__).parent)) / relative


def support_mailto(version: str) -> str:
    release, build, *_ = platform.win32_ver()
    windows_version = f'Windows {release} ({build})' if release and build else build or 'Unknown'
    body = (f'Application Version: {version}\r\n'
            f'Windows Version: {windows_version}\r\n\r\n'
            'Problem description:\r\n\r\n')
    fields = {'subject': 'Wadaa PDF Manager - Support Request', 'body': body}
    return f'mailto:{SUPPORT_EMAIL}?{urlencode(fields, quote_via=quote)}'


def sort_key(path: Path):
    m = re.search(r'(?i)(?<!\w)u(\d+)(?!\w)', path.stem)
    n = int(m.group(1)) if m else float('inf')
    parts = tuple((0, int(s)) if s.isdigit() else (1, s.casefold())
                  for s in re.split(r'(\d+)', path.name))
    return n, parts


def parse_selection(expr: str, count: int) -> list[int]:
    if not expr.strip():
        return list(range(count))
    selected = []
    for token in re.split(r'[\s,،;؛]+', expr.strip()):
        if not token:
            continue
        m = re.fullmatch(r'(\d+)\s*[-–]\s*(\d+)', token)
        if m:
            a, b = map(int, m.groups())
            if not (1 <= a <= count and 1 <= b <= count):
                raise ValueError(f'Range {token} is outside 1–{count}.')
            step = 1 if a <= b else -1
            selected.extend(range(a - 1, b - 1 + step, step))
        elif token.isdecimal():
            n = int(token)
            if not 1 <= n <= count:
                raise ValueError(f'File number {n} is outside 1–{count}.')
            selected.append(n - 1)
        else:
            raise ValueError(f'Invalid number/range: {token}. Example: 1-9 12 15-18')
    return selected


def image_pdf_bytes(path: Path) -> io.BytesIO:
    frames = []
    with Image.open(path) as image:
        for frame in ImageSequence.Iterator(image):
            oriented = ImageOps.exif_transpose(frame.copy())
            if oriented.mode in ('RGBA', 'LA') or 'transparency' in oriented.info:
                rgba = oriented.convert('RGBA')
                rgb = Image.new('RGB', rgba.size, 'white')
                rgb.paste(rgba, mask=rgba.getchannel('A'))
            else:
                rgb = oriented.convert('RGB')
            frames.append(rgb)
    if not frames:
        raise ValueError(f'Image has no pages: {path.name}')
    buffer = io.BytesIO()
    frames[0].save(buffer, 'PDF', save_all=len(frames) > 1,
                   append_images=frames[1:], resolution=96.0)
    buffer.seek(0)
    return buffer


def convert_office_to_pdf(path: Path, folder: Path) -> Path:
    """Convert an Office file to PDF using installed Microsoft Office, or LibreOffice.

    Conversion runs in the merge worker thread. Microsoft Office automation needs
    COM initialized in that thread. Conversion uses a unique temporary destination.
    """
    output = folder / (path.stem + '.pdf')
    # Avoid collisions when two files have the same stem but different extensions.
    if output.exists():
        output = folder / (path.stem + '_' + path.suffix[1:] + '.pdf')
    ext = path.suffix.lower()
    office_error = None
    if sys.platform == 'win32':
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = None
            doc = None
            try:
                if ext in {'.doc', '.docx', '.docm'}:
                    app = win32com.client.DispatchEx('Word.Application')
                    app.Visible = False
                    app.DisplayAlerts = 0
                    doc = app.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
                    doc.ExportAsFixedFormat(str(output.resolve()), 17)
                elif ext in {'.xls', '.xlsx', '.xlsm'}:
                    app = win32com.client.DispatchEx('Excel.Application')
                    app.Visible = False
                    app.DisplayAlerts = False
                    doc = app.Workbooks.Open(str(path.resolve()), ReadOnly=True, UpdateLinks=0)
                    doc.ExportAsFixedFormat(0, str(output.resolve()))
                else:
                    app = win32com.client.DispatchEx('PowerPoint.Application')
                    doc = app.Presentations.Open(str(path.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
                    doc.SaveAs(str(output.resolve()), 32)
                if output.is_file() and output.stat().st_size:
                    return output
                raise RuntimeError('Microsoft Office did not create a PDF.')
            finally:
                if doc is not None:
                    try:
                        doc.Close(False) if ext not in {'.ppt', '.pptx', '.pptm'} else doc.Close()
                    except Exception:
                        pass
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except Exception as exc:
            office_error = exc
    # LibreOffice is an optional alternative to Microsoft Office.
    candidates = [shutil.which('soffice'), shutil.which('libreoffice')]
    if sys.platform == 'win32':
        candidates.extend([
            Path(os.environ.get('PROGRAMFILES', 'C:/Program Files')) / 'LibreOffice/program/soffice.exe',
            Path(os.environ.get('PROGRAMFILES(X86)', 'C:/Program Files (x86)')) / 'LibreOffice/program/soffice.exe',
        ])
    executable = next((str(x) for x in candidates if x and Path(x).is_file()), None)
    if executable:
        # Isolate the LO user profile to avoid issues with other open instances.
        profile = (folder / 'lo_profile').resolve().as_uri()
        result = subprocess.run(
            [executable, f'-env:UserInstallation={profile}', '--headless', '--convert-to',
             'pdf', '--outdir', str(folder), str(path.resolve())],
            capture_output=True, text=True, timeout=180,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
        )
        generated = folder / (path.stem + '.pdf')
        if result.returncode == 0 and generated.is_file():
            if generated != output:
                generated.replace(output)
            return output
        raise RuntimeError(f'LibreOffice could not convert {path.name}: {result.stderr or result.stdout}')
    details = f' (Microsoft Office error: {office_error})' if office_error else ''
    raise RuntimeError(f'Cannot convert {path.name}. Install Microsoft Office with the pywin32 '
                       f'Python package, or LibreOffice.{details}')


def selected_page_indices(expr: str, count: int) -> list[int]:
    """1-based visible page ranges, with deduplication preserving requested order."""
    pages = parse_selection(expr, count)
    result, seen = [], set()
    for p in pages:
        if p not in seen:
            seen.add(p)
            result.append(p)
    if not result:
        raise ValueError('At least one page is required.')
    return result


def merge_files(paths: list, destination: Path, progress):
    writer = PdfWriter()
    streams = []
    temporary = destination.with_name(destination.name + '.wada-part')
    try:
        with tempfile.TemporaryDirectory(prefix='wada_pdf_') as temp_dir:
            workdir = Path(temp_dir)
            for i, job in enumerate(paths, 1):
                path, page_choice = job if isinstance(job, tuple) else (job, None)
                progress(i - 1, len(paths), path.name)
                ext = path.suffix.lower()
                if ext in OFFICE_EXTS:
                    # Each file gets its own directory to avoid same-name collisions.
                    office_dir = workdir / str(i)
                    office_dir.mkdir()
                    converted = convert_office_to_pdf(path, office_dir)
                    writer.append(PdfReader(str(converted), strict=False), import_outline=False)
                elif ext == '.pdf':
                    reader = PdfReader(str(path), strict=False)
                    if reader.is_encrypted and reader.decrypt('') == 0:
                        raise ValueError(f'Password-protected PDF: {path.name}')
                    if page_choice is None:
                        writer.append(reader, import_outline=False)
                    else:
                        # Explicit selection (including a user-defined page order).
                        for number in page_choice:
                            if number < 0 or number >= len(reader.pages):
                                raise ValueError(f'Page selection is no longer valid for {path.name}.')
                            writer.add_page(reader.pages[number])
                else:
                    stream = image_pdf_bytes(path)
                    streams.append(stream)
                    writer.append(PdfReader(stream), import_outline=False)
            progress(len(paths), len(paths), 'Writing final PDF...')
            with temporary.open('wb') as out:
                writer.write(out)
            os.replace(temporary, destination)
    finally:
        writer.close()
        for stream in streams:
            stream.close()
        if temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass



class DarkScrollbar(tk.Canvas):
    """Slim fully dark scrollbar: no light Windows arrow buttons."""
    def __init__(self, parent, command, **kwargs):
        super().__init__(parent, width=13, bg=SURFACE, highlightthickness=0,
                         bd=0, relief='flat', cursor='arrow', **kwargs)
        self.command = command
        self.first, self.last = 0.0, 1.0
        self._grab_offset = None
        self.bind('<Configure>', lambda e: self.redraw())
        self.bind('<ButtonPress-1>', self._press)
        self.bind('<B1-Motion>', self._drag)
        self.bind('<ButtonRelease-1>', self._release)

    def set(self, first, last):
        self.first, self.last = float(first), float(last)
        self.redraw()

    def _bounds(self):
        height = max(self.winfo_height(), 1)
        track = max(height - 8, 1)
        thumb_height = min(track, max(26, track * (self.last - self.first)))
        top = 4 + (track - thumb_height) * self.first / max(1e-8, 1.0 - (self.last - self.first))
        return top, top + thumb_height, track, thumb_height

    def redraw(self):
        self.delete('all')
        width = max(self.winfo_width(), 13)
        height = self.winfo_height()
        self.create_rectangle(0, 0, width, height, fill='#272727', outline='')
        if self.last - self.first >= 0.99999:
            return
        top, bottom, _, _ = self._bounds()
        color = '#717875' if self._grab_offset is not None else '#565d5a'
        # Rounded thumb with 3px corner radius (native Canvas round rectangle).
        x1, x2, r = 3, width - 3, 3
        self.create_polygon(x1+r, top, x2-r, top, x2, top+r,
                            x2, bottom-r, x2-r, bottom, x1+r, bottom,
                            x1, bottom-r, x1, top+r, fill=color, smooth=True)

    def _press(self, event):
        top, bottom, _, thumb_height = self._bounds()
        if top <= event.y <= bottom:
            self._grab_offset = event.y - top
        else:
            self._grab_offset = thumb_height / 2
            self._move(event.y)
        self.redraw()

    def _move(self, y):
        _, _, track, thumb_height = self._bounds()
        travel = max(track - thumb_height, 1)
        fraction = max(0.0, min(1.0, (y - self._grab_offset - 4) / travel))
        self.command('moveto', fraction * max(0, 1 - (self.last - self.first)))

    def _drag(self, event):
        if self._grab_offset is not None:
            self._move(event.y)

    def _release(self, event):
        self._grab_offset = None
        self.redraw()


class RoundedButton(tk.Canvas):
    """Antialiased button, drawn directly onto the actual dark app background.

    Unlike a ttk 9-slice image element, this does not let the Windows ttk
    renderer place an opaque rectangular button surface behind round corners.
    """
    PALETTES = {
        'normal': ('#333639', '#42474a', '#282d2f', '#474b4e', '#f4f4f4'),
        'primary': ('#e9eceb', '#d4e0db', '#c6d4ce', '#e9eceb', '#202625'),
        'donate': ('#303a35', '#394a40', '#29382f', '#476452', '#82cda0'),
    }

    def __init__(self, parent, text, command, kind='normal', **kwargs):
        self.font_obj = tkfont.Font(family='Segoe UI', size=10,
                                    weight='bold' if kind in ('primary','donate') else 'normal')
        self._label = text
        self._command = command
        self._kind = kind
        self._state = 'normal'
        self._hover = False
        self._pressed = False
        self._photo = None
        width = max(104, self.font_obj.measure(text) + 30)
        super().__init__(parent, bg=BG, width=width, height=36,
                         bd=0, highlightthickness=0, relief='flat',
                         cursor='hand2', takefocus=1, **kwargs)
        self.bind('<Configure>', self._redraw)
        self.bind('<Enter>', self._enter)
        self.bind('<Leave>', self._leave)
        self.bind('<ButtonPress-1>', self._down)
        self.bind('<ButtonRelease-1>', self._up)
        self.bind('<Key-space>', self._invoke_key)
        self.bind('<Key-Return>', self._invoke_key)
        self._redraw()

    def configure(self, cnf=None, **kw):
        if 'state' in kw:
            self._state = kw.pop('state')
            self._redraw()
        if 'text' in kw:
            self._label = kw.pop('text')
            self._redraw()
        if cnf is None and not kw:
            return None
        return super().configure(cnf, **kw)

    config = configure

    def _redraw(self, event=None):
        width = max(2, self.winfo_width())
        height = max(2, self.winfo_height())
        if width <= 2: width = int(self.cget('width'))
        if height <= 2: height = int(self.cget('height'))
        fill, hover, pressed, edge, text_color = self.PALETTES[self._kind]
        if self._state == 'disabled':
            fill, edge, text_color = '#292b2c', '#353739', '#818686'
        elif self._pressed:
            fill = pressed
        elif self._hover:
            fill = hover
        # Premultiply the rounded shape onto the exact parent background.
        # Oversampling gives genuinely round anti-aliased corner pixels.
        k = 3
        art = Image.new('RGB', (width*k, height*k), BG)
        draw = ImageDraw.Draw(art)
        draw.rounded_rectangle((k, k, (width-2)*k, (height-2)*k),
                               radius=10*k, fill=fill, outline=edge, width=k)
        self._photo = ImageTk.PhotoImage(art.resize((width,height), Image.Resampling.LANCZOS), master=self)
        self.delete('all')
        self.create_image(0, 0, image=self._photo, anchor='nw')
        self.create_text(width//2, height//2, text=self._label,
                         font=self.font_obj, fill=text_color, anchor='center')

    def _enter(self, event):
        self._hover = True
        self._redraw()

    def _leave(self, event):
        self._hover = False
        self._pressed = False
        self._redraw()

    def _down(self, event):
        if self._state == 'disabled': return
        self.focus_set()
        self._pressed = True
        self._redraw()

    def _up(self, event):
        if self._state == 'disabled': return
        inside = 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height()
        activated = self._pressed and inside
        self._pressed = False
        self._redraw()
        if activated: self._command()

    def _invoke_key(self, event):
        if self._state != 'disabled': self._command()
        return 'break'


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title('Wadaa PDF Manager')
        self.version = resource_path('VERSION.txt').read_text(encoding='ascii').strip()
        self.about_window = None
        # Set a real content-aware minimum after building the complete layout.
        # Do not let Windows resize the app until the bottom controls disappear.
        self.geometry('960x680')
        self.configure(bg=BG)
        self.paths: list[Path] = []
        self.selection_text = tk.StringVar()
        self.status = tk.StringVar(value='Drop PDF, images or Office documents here, or choose a folder.')
        self.location = tk.StringVar(value='No folder selected')
        self.busy = False
        self.drag_index = None
        self.drag_started = False
        self.drag_start_y = 0
        self.drag_item = None
        self.output_file: Path | None = None
        self.undo_stack = []
        self.redo_stack = []
        self.page_cache = {}
        self.page_choices: dict[Path, str] = {}
        self.search_text = tk.StringVar()
        self.search_text.trace_add('write', lambda *_: self.refresh_view())
        self.last_session_enabled = self.load_session_preference()
        self.total_text = tk.StringVar(value='0 files  •  0 pages  •  0 B')
        self.open_after_merge = tk.BooleanVar(value=False)
        self.configure_styles()
        self.build_ui()
        self.update_idletasks()
        minimum_height = max(565, self.winfo_reqheight() + 5)
        minimum_width = max(800, self._min_content_width())
        self.minsize(minimum_width, minimum_height)
        # The window may grow freely; it cannot shrink below the visible content.
        usable_height = max(minimum_height, self.winfo_screenheight() - 85)
        self.geometry(f'{max(960, minimum_width)}x{max(minimum_height, min(690, usable_height))}')
        self.load_icon()
        self.register_drop(self)
        self.register_drop(self.tree)
        self.register_drop(self.wrapper)
        self.bind('<Control-z>', lambda e: self.undo() if self.focus_get() != self.entry else None)
        self.bind('<Control-y>', lambda e: self.redo() if self.focus_get() != self.entry else None)
        self.protocol('WM_DELETE_WINDOW', self.close_app)
        if self.last_session_enabled:
            self.restore_session()
        self.updater = UpdateController(self, resource_path('update_settings.json'),
                                        self._session_path().parent, self.version)
        self.after(1500, self.updater.check)

    def _min_content_width(self):
        # Tk pack doesn't wrap; enforce the widest required toolbar width.
        rows = (self.top_buttons, self.order_buttons, self.action_buttons)
        content_width = max(row.winfo_reqwidth() for row in rows)
        return content_width + 2*12 + 18

    def load_icon(self):
        ico = resource_path('assets/icon.ico')
        png = resource_path('assets/icon.png')
        try:
            if ico.exists():
                self.iconbitmap(str(ico))
            elif png.exists():
                self._icon_image = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._icon_image)
        except (tk.TclError, OSError):
            pass

    def configure_styles(self):
        s = ttk.Style(self)
        s.theme_use('clam')
        s.configure('Dark.TFrame', background=BG)
        s.configure('Surface.TFrame', background=SURFACE)
        s.configure('Dark.TLabel', background=BG, foreground=TEXT, font=('Segoe UI', 10))
        s.configure('Muted.TLabel', background=BG, foreground=MUTED, font=('Segoe UI', 9))
        s.configure('Title.TLabel', background=BG, foreground=TEXT, font=('Segoe UI Semibold', 23))
        # Calm modern controls: generous spacing, subtle edges, modest hover feedback.
        s.configure('Dark.TButton', background='#333639', foreground=TEXT,
                    bordercolor='#474b4e', lightcolor='#333639', darkcolor='#333639',
                    relief='flat', borderwidth=1, padding=(10, 0),
                    font=('Segoe UI', 10), focuscolor='#333639')
        s.map('Dark.TButton', background=[('active', '#42474a'), ('pressed', '#282d2f'), ('disabled', BG)],
              foreground=[('disabled', MUTED)], bordercolor=[('active', '#5a6062')])
        s.configure('Primary.TButton', background='#e9eceb', foreground='#202625',
                    bordercolor='#e9eceb', lightcolor='#e9eceb', darkcolor='#e9eceb',
                    relief='flat', borderwidth=1, padding=(10, 0),
                    font=('Segoe UI Semibold', 10), focuscolor='#e9eceb')
        s.map('Primary.TButton', background=[('active', '#d4e0db'), ('pressed', '#c6d4ce'), ('disabled', '#999999')])
        s.configure('Donate.TButton', background='#303a35', foreground='#82cda0',
                    bordercolor='#476452', lightcolor='#303a35', darkcolor='#303a35',
                    relief='flat', borderwidth=1, padding=(10, 0),
                    font=('Segoe UI Semibold', 10), focuscolor='#303a35')
        s.map('Donate.TButton', background=[('active', '#394a40'), ('pressed', '#29382f')],
              foreground=[('active', '#a0e5b9')], bordercolor=[('active', '#679777')])
        s.configure('Dark.Treeview', background=SURFACE, fieldbackground=SURFACE,
                    foreground=TEXT, rowheight=31, font=('Segoe UI', 10), borderwidth=0,
                    relief='flat')
        s.configure('Dark.Treeview.Heading', background='#3a3a3a', foreground=TEXT,
                    borderwidth=0, font=('Segoe UI Semibold', 10), padding=(8, 8))
        s.map('Dark.Treeview', background=[('selected', SELECT)], foreground=[('selected', '#ffffff')])
        s.map('Dark.Treeview.Heading', background=[('active', HOVER)])
        s.configure('Dark.Horizontal.TProgressbar', background='#6fb68b', troughcolor=SURFACE,
                    lightcolor='#6fb68b', darkcolor='#6fb68b', borderwidth=0, thickness=10)


    def button(self, parent, label, command, primary=False, donate=False, **kwargs):
        kind = 'donate' if donate else ('primary' if primary else 'normal')
        return RoundedButton(parent, text=label, command=command, kind=kind, **kwargs)

    def build_ui(self):
        self.wrapper = ttk.Frame(self, style='Dark.TFrame', padding=12)
        self.wrapper.pack(fill='both', expand=True)
        ttk.Label(self.wrapper, text='Wadaa PDF Manager', style='Title.TLabel').pack(anchor='w')
        ttk.Label(self.wrapper, text='PDF + images + Office  •  Drag & drop  •  Reorder  •  Merge',
                  style='Muted.TLabel').pack(anchor='w', pady=(0, 9))

        buttons = ttk.Frame(self.wrapper, style='Dark.TFrame')
        self.top_buttons = buttons
        buttons.pack(fill='x', pady=(0, 8))
        self.button(buttons, '+ Add Files', self.add_files).pack(side='left', padx=(0, 8))
        self.button(buttons, 'Open Folder', self.open_folder).pack(side='left', padx=(0, 8))
        self.button(buttons, 'Sort by Name (A-Z)', self.sort_naturally).pack(side='left', padx=(0, 8))
        self.button(buttons, 'Clear All', self.clear).pack(side='left')
        self.last_session_button = self.button(buttons, self.session_button_label(),
                                               self.toggle_last_session)
        self.last_session_button.pack(side='left', padx=(8, 0))
        self.button(buttons, 'Support / Donate', lambda: webbrowser.open(DONATE), donate=True).pack(side='right')

        ttk.Label(self.wrapper, textvariable=self.location, style='Muted.TLabel').pack(anchor='w', pady=(0, 5))
        ttk.Label(self.wrapper, text='Drop files into the list  •  Drag a row to move it  •  Ctrl/Shift to select multiple',
                  style='Muted.TLabel').pack(anchor='w', pady=(0, 8))

        searchbar = ttk.Frame(self.wrapper, style='Dark.TFrame')
        searchbar.pack(fill='x', pady=(0, 6))
        ttk.Label(searchbar, text='Search files:', style='Muted.TLabel').pack(side='left', padx=(0, 8))
        self.search_entry = tk.Entry(searchbar, textvariable=self.search_text, bg=SURFACE, fg=TEXT,
            insertbackground=TEXT, highlightthickness=1, highlightcolor=HOVER,
            highlightbackground=HOVER, relief='flat', font=('Segoe UI', 10))
        self.search_entry.pack(side='left', fill='x', expand=True, ipady=5)
        self.button(searchbar, 'Clear Search', lambda: self.search_text.set('')).pack(side='left', padx=(8, 0))

        tablewrap = ttk.Frame(self.wrapper, style='Surface.TFrame', height=145)
        tablewrap.pack(fill='both', expand=True)
        tablewrap.pack_propagate(False)
        self.tree = ttk.Treeview(tablewrap, columns=('number', 'name', 'type', 'pages'), show='headings',
                                 selectmode='extended', style='Dark.Treeview', height=5)
        self.tree.heading('number', text='#')
        self.tree.heading('name', text='File name (current merge order)')
        self.tree.heading('type', text='Type')
        self.tree.heading('pages', text='Pages')
        self.tree.column('number', width=58, stretch=False, anchor='center')
        self.tree.column('name', minwidth=280, anchor='w')
        self.tree.column('type', width=65, stretch=False, anchor='center')
        self.tree.column('pages', width=65, stretch=False, anchor='center')
        scrollbar = DarkScrollbar(tablewrap, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.tree.bind('<ButtonPress-1>', self.mouse_press, add='+')
        self.tree.bind('<B1-Motion>', self.mouse_drag, add='+')
        self.tree.bind('<ButtonRelease-1>', self.mouse_release, add='+')
        self.tree.bind('<Double-1>', lambda e: self.preview_pages(), add='+')
        self.tree.bind('<Delete>', self._delete_selected_key)

        orderbuttons = ttk.Frame(self.wrapper, style='Dark.TFrame')
        self.order_buttons = orderbuttons
        orderbuttons.pack(fill='x', pady=(7, 8))
        self.button(orderbuttons, 'Move Up', lambda: self.move_selected(-1)).pack(side='left', padx=(0, 8))
        self.button(orderbuttons, 'Move Down', lambda: self.move_selected(1)).pack(side='left', padx=(0, 8))
        self.button(orderbuttons, 'Remove Selected', self.remove_selected).pack(side='left', padx=(0, 8))
        self.undo_button = self.button(orderbuttons, 'Undo', self.undo)
        self.undo_button.pack(side='left', padx=(0, 6))
        self.redo_button = self.button(orderbuttons, 'Redo', self.redo)
        self.redo_button.pack(side='left')
        self.button(orderbuttons, 'Open Output Folder', self.open_output_folder).pack(side='right')

        ttk.Label(self.wrapper, text='Optional custom selection: 1-9 12 15-18  •  Empty = all in list order',
                  style='Dark.TLabel').pack(anchor='w')
        self.entry = tk.Entry(self.wrapper, textvariable=self.selection_text, bg=SURFACE, fg=TEXT,
                              insertbackground=TEXT, highlightthickness=1, highlightcolor=HOVER,
                              highlightbackground=HOVER, relief='flat', font=('Consolas', 12))
        self.entry.pack(fill='x', ipady=7, pady=(7, 9))
        actionbar = ttk.Frame(self.wrapper, style='Dark.TFrame')
        self.action_buttons = actionbar
        actionbar.pack(fill='x')
        self.button(actionbar, 'Use Highlighted', self.use_highlighted).pack(side='left', padx=(0, 8))
        self.button(actionbar, 'Preview Merge Order', self.preview).pack(side='left', padx=(0, 8))
        self.button(actionbar, 'Preview Pages', self.preview_pages).pack(side='left', padx=(0, 8))
        self.button(actionbar, 'Select PDF Pages', self.choose_pdf_pages).pack(side='left')
        self.merge_button = self.button(actionbar, 'Merge to PDF...', self.start_merge, primary=True)
        self.merge_button.pack(side='right')
        self.open_after_check = tk.Checkbutton(actionbar, text='Open PDF after merge',
            variable=self.open_after_merge, bg=BG, fg=MUTED, selectcolor=SURFACE,
            activebackground=BG, activeforeground=TEXT, font=('Segoe UI', 9),
            highlightthickness=0, bd=0)
        self.open_after_check.pack(side='right', padx=(0, 10))
        ttk.Label(self.wrapper, textvariable=self.total_text, style='Muted.TLabel').pack(anchor='w', pady=(6, 0))
        self.progress = ttk.Progressbar(self.wrapper, maximum=100, mode='determinate',
                                        style='Dark.Horizontal.TProgressbar')
        self.progress.pack(fill='x', pady=(12, 6))
        # Keep the status message on the left and the attribution anchored right.
        footer = ttk.Frame(self.wrapper, style='Dark.TFrame')
        footer.pack(fill='x')
        ttk.Label(footer, text='Powered by WadaaDev', style='Muted.TLabel',
                  font=('Segoe UI Semibold', 9)).pack(side='right', padx=(12, 0))
        ttk.Label(footer, textvariable=self.status, style='Muted.TLabel',
                  wraplength=650).pack(side='left', fill='x', expand=True, anchor='w')
        self.button(footer, 'Check for Updates', lambda: self.updater.check(manual=True)
                    ).pack(side='right', padx=(8, 0))

        self.button(footer, 'About & Support', self.show_about).pack(side='right', padx=(8, 0))

    def show_about(self):
        if self.about_window is not None and self.about_window.winfo_exists():
            self.about_window.deiconify()
            self.about_window.lift()
            self.about_window.focus_set()
            return

        dialog = tk.Toplevel(self)
        self.about_window = dialog
        dialog.title('About & Support')
        dialog.transient(self)
        dialog.configure(bg=BG)
        dialog.resizable(False, False)
        try:
            dialog.iconbitmap(str(resource_path('assets/icon.ico')))
        except (tk.TclError, OSError):
            pass
        dialog.bind('<Escape>', lambda event: dialog.destroy())

        content = ttk.Frame(dialog, padding=(24, 20), style='Dark.TFrame')
        content.pack(fill='both', expand=True)
        ttk.Label(content, text='Wadaa PDF Manager', style='Title.TLabel').pack(anchor='w')
        ttk.Label(content, text=f'Version: {self.version}', style='Muted.TLabel'
                  ).pack(anchor='w', pady=(1, 12))
        ttk.Label(content, text='Developed by WadaaDev', style='Dark.TLabel').pack(anchor='w')
        ttk.Label(content, text='Powered by WadaaDev', style='Muted.TLabel').pack(anchor='w')
        ttk.Label(content, text=SUPPORT_EMAIL, style='Muted.TLabel').pack(anchor='w', pady=(2, 14))

        ttk.Button(content, text='Contact Support', style='Dark.TButton',
                   command=lambda: self.open_support_link(dialog, support_mailto(self.version),
                                                          mail=True)).pack(fill='x', ipady=5)
        ttk.Button(content, text='Report a Bug', style='Dark.TButton',
                   command=lambda: self.open_support_link(dialog, BUG_REPORT_URL)
                   ).pack(fill='x', ipady=5, pady=5)
        ttk.Button(content, text='GitHub Repository', style='Dark.TButton',
                   command=lambda: self.open_support_link(dialog, REPOSITORY_URL)
                   ).pack(fill='x', ipady=5)
        ttk.Label(content, text='© 2026 WadaaDev. All rights reserved.', style='Muted.TLabel'
                  ).pack(anchor='w', pady=(15, 0))

        dialog.update_idletasks()
        x = max(0, self.winfo_rootx() + (self.winfo_width() - dialog.winfo_reqwidth()) // 2)
        y = max(0, self.winfo_rooty() + (self.winfo_height() - dialog.winfo_reqheight()) // 2)
        dialog.geometry(f'+{x}+{y}')
        dialog.focus_set()

    def open_support_link(self, parent, url, *, mail=False):
        try:
            os.startfile(url)
        except OSError:
            if mail:
                hint = ('Windows could not open the default email app. Set a default mail app '
                        f'or write to {SUPPORT_EMAIL} manually.')
            else:
                hint = f'Windows could not open the default browser. Copy this address into a browser:\n{url}'
            messagebox.showerror('Unable to open link', hint, parent=parent)

    def register_drop(self, widget):
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind('<<Drop>>', self.on_external_drop)

    def on_external_drop(self, event):
        if self.busy:
            return
        # tk.splitlist handles spaces and Windows braces around paths.
        names = self.tk.splitlist(event.data)
        paths = []
        for raw in names:
            p = Path(raw)
            if p.is_dir():
                try:
                    paths.extend(x for x in p.iterdir() if x.is_file() and x.suffix.lower() in EXTS)
                except OSError:
                    continue
            elif p.is_file() and p.suffix.lower() in EXTS:
                paths.append(p)
        self.add_paths(paths, sort_new=False)

    def add_paths(self, paths, sort_new=False):
        existing = {p.resolve() for p in self.paths}
        unique = []
        for p in paths:
            p = Path(p)
            try:
                key = p.resolve()
            except OSError:
                continue
            if key not in existing and p.is_file() and p.suffix.lower() in EXTS:
                unique.append(p)
                existing.add(key)
        if sort_new:
            unique.sort(key=sort_key)
        if unique:
            self._remember()
        self.paths.extend(unique)
        self.refresh()
        duplicates = sum(1 for p in paths if Path(p).is_file() and Path(p).suffix.lower() in EXTS) - len(unique)
        self.status.set(f'{len(unique)} new file(s) added; {max(duplicates, 0)} duplicate(s) skipped. {len(self.paths)} total.')

    def refresh(self, focus_indices=None):
        self.refresh_view(focus_indices)
        self._count_pages_async()
        self._update_history_buttons()

    def refresh_view(self, focus_indices=None):
        if not hasattr(self, 'tree'):
            return
        # Keep file indexes stable; filtering never changes the underlying merge order.
        selected = set(self.tree.selection()) if focus_indices is None else {str(i) for i in focus_indices}
        self.tree.delete(*self.tree.get_children())
        query = self.search_text.get().strip().casefold()
        for i, p in enumerate(self.paths):
            if query and query not in p.name.casefold():
                continue
            count = self.page_cache.get(p, '…')
            if p.suffix.lower() == '.pdf' and p in self.page_choices and isinstance(count, int):
                try:
                    count = f'{len(selected_page_indices(self.page_choices[p], count))}/{count}'
                except ValueError:
                    count = '!'
            self.tree.insert('', 'end', iid=str(i), values=(i + 1, p.name, p.suffix[1:].upper(), count))
        valid = [i for i in selected if self.tree.exists(i)]
        if valid:
            self.tree.selection_set(valid)
            if focus_indices is not None:
                self.tree.see(valid[0])
        self._update_totals()

    def _update_totals(self):
        total_bytes, known_pages, unknown = 0, 0, 0
        for p in self.paths:
            try:
                total_bytes += p.stat().st_size
            except OSError:
                pass
            count = self.page_cache.get(p)
            if isinstance(count, int):
                if p in self.page_choices and p.suffix.lower() == '.pdf':
                    try:
                        count = len(selected_page_indices(self.page_choices[p], count))
                    except ValueError:
                        count = 0
                known_pages += count
            else:
                unknown += 1
        unit = 'B'
        size = float(total_bytes)
        for u in ('KB', 'MB', 'GB', 'TB'):
            if size < 1024:
                break
            size /= 1024
            unit = u
        label = f'{len(self.paths)} files  •  {known_pages} pages'
        if unknown:
            label += f' (+ {unknown} uncounted files)'
        label += f'  •  {size:.1f} {unit}' if unit != 'B' else f'  •  {total_bytes} B'
        self.total_text.set(label)

    def _snapshot(self):
        return (self.paths.copy(), self.location.get(), self.page_choices.copy())

    def _session_path(self):
        home = Path(os.environ.get('APPDATA') or (Path.home() / '.config'))
        return home / 'WadaaPDFManager' / 'session.json'

    def _session_preference_path(self):
        return self._session_path().with_name('session-preference.json')

    def load_session_preference(self):
        try:
            data = json.loads(self._session_preference_path().read_text(encoding='utf-8'))
            return not (isinstance(data, dict) and data.get('enabled') is False)
        except (OSError, ValueError):
            return True

    @staticmethod
    def write_session_json(target, data):
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=target.parent,
                                         delete=False) as output:
            tmp = Path(output.name)
        try:
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
            os.replace(tmp, target)
        finally:
            tmp.unlink(missing_ok=True)

    def session_button_label(self):
        return 'Last Session: ON' if self.last_session_enabled else 'Last Session: OFF'

    def toggle_last_session(self):
        enabled = not self.last_session_enabled
        try:
            self.write_session_json(self._session_preference_path(), {'enabled': enabled})
        except OSError as exc:
            self.status.set(f'Unable to change Last Session: {exc}')
            return
        self.last_session_enabled = enabled
        self.last_session_button.configure(text=self.session_button_label())
        if enabled and not self.paths and self._session_path().is_file():
            if messagebox.askyesno('Last Session', 'Restore the previously saved list now?',
                                   parent=self):
                self.restore_session()

    def save_session(self):
        if not self.last_session_enabled:
            return True
        try:
            target = self._session_path()
            target.parent.mkdir(parents=True, exist_ok=True)
            data = {'version': 1, 'paths': [str(p) for p in self.paths],
                    'location': self.location.get(),
                    'page_choices': {str(p): expr for p, expr in self.page_choices.items()},
                    'open_after_merge': bool(self.open_after_merge.get())}
            self.write_session_json(target, data)
            return True
        except (OSError, ValueError) as exc:
            self.status.set(f'Unable to save session: {exc}')
            return False

    def restore_session(self):
        try:
            target = self._session_path()
            if not target.is_file():
                return
            data = json.loads(target.read_text(encoding='utf-8'))
            if not isinstance(data, dict) or data.get('version') != 1:
                return
            paths, seen = [], set()
            for name in data.get('paths', []):
                p = Path(name)
                if p.suffix.lower() in EXTS and p.is_file() and p.resolve() not in seen:
                    paths.append(p)
                    seen.add(p.resolve())
            self.paths = paths
            self.location.set(str(data.get('location', 'No folder selected')))
            self.open_after_merge.set(bool(data.get('open_after_merge', False)))
            saved = data.get('page_choices', {})
            self.page_choices = {p: saved[str(p)] for p in paths
                                 if p.suffix.lower() == '.pdf' and isinstance(saved, dict)
                                 and isinstance(saved.get(str(p)), str)}
            self.refresh()
            self.status.set(f'Restored {len(paths)} file(s) from the previous session.')
        except (ValueError, OSError, TypeError, KeyError) as exc:
            self.status.set(f'Previous session could not be restored: {exc}')

    def close_app(self):
        self.save_session()
        self.updater.close()
        self.destroy()

    def choose_pdf_pages(self):
        if self.busy:
            return
        selection = self.selected_indices()
        if len(selection) != 1:
            messagebox.showinfo('Select PDF Pages', 'Select one PDF file in the list first.')
            return
        path = self.paths[selection[0]]
        if path.suffix.lower() != '.pdf':
            messagebox.showinfo('Select PDF Pages', 'Page selection is available for PDFs only.')
            return
        try:
            reader = PdfReader(str(path), strict=False)
            if reader.is_encrypted and reader.decrypt('') == 0:
                raise ValueError('This PDF is password protected.')
            total = len(reader.pages)
        except Exception as exc:
            messagebox.showerror('Select PDF Pages', f'Cannot read {path.name}: {exc}')
            return
        dialog = tk.Toplevel(self, bg=BG)
        dialog.title(f'Select pages — {path.name}')
        dialog.geometry('530x220')
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        frame = tk.Frame(dialog, bg=BG, padx=18, pady=16)
        frame.pack(fill='both', expand=True)
        tk.Label(frame, bg=BG, fg=TEXT, text=f'{path.name}  •  {total} pages',
                 font=('Segoe UI', 10)).pack(anchor='w')
        tk.Label(frame, bg=BG, fg=MUTED,
                 text='Examples: 1-5 8 10-12  |  Blank = all pages',
                 font=('Segoe UI', 9)).pack(anchor='w', pady=(9, 5))
        value = tk.StringVar(value=self.page_choices.get(path, ''))
        entry = tk.Entry(frame, textvariable=value, bg=SURFACE, fg=TEXT,
                         insertbackground=TEXT, font=('Consolas', 12), relief='flat')
        entry.pack(fill='x', ipady=7)
        entry.focus_set()
        buttons = tk.Frame(frame, bg=BG)
        buttons.pack(fill='x', pady=(14, 0))
        def apply():
            expr = value.get().strip()
            if expr:
                try:
                    selected_page_indices(expr, total)
                except ValueError as exc:
                    messagebox.showerror('Invalid pages', str(exc), parent=dialog)
                    return
            if expr != self.page_choices.get(path, ''):
                self._remember()
                if expr:
                    self.page_choices[path] = expr
                else:
                    self.page_choices.pop(path, None)
                self.refresh()
                self.status.set(f'Page selection saved for {path.name}: {expr or "all pages"}')
            dialog.destroy()
        self.button(buttons, 'Apply', apply, primary=True).pack(side='right')
        self.button(buttons, 'Cancel', dialog.destroy).pack(side='right', padx=(0, 8))
        dialog.bind('<Return>', lambda e: apply())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _remember(self):
        if self.busy:
            return
        self.undo_stack.append(self._snapshot())
        if len(self.undo_stack) > 60:
            self.undo_stack.pop(0)
        self.redo_stack.clear()
        self._update_history_buttons()

    def _update_history_buttons(self):
        if hasattr(self, 'undo_button'):
            self.undo_button.configure(state=('normal' if self.undo_stack and not self.busy else 'disabled'))
            self.redo_button.configure(state=('normal' if self.redo_stack and not self.busy else 'disabled'))

    def undo(self):
        if self.busy or not self.undo_stack:
            return
        self.redo_stack.append(self._snapshot())
        self.paths, where, choices = self.undo_stack.pop()
        self.page_choices = choices
        self.location.set(where)
        self.refresh()
        self.status.set('Undo completed.')

    def redo(self):
        if self.busy or not self.redo_stack:
            return
        self.undo_stack.append(self._snapshot())
        self.paths, where, choices = self.redo_stack.pop()
        self.page_choices = choices
        self.location.set(where)
        self.refresh()
        self.status.set('Redo completed.')

    def _count_pages_async(self):
        if not hasattr(self, '_page_pending'):
            self._page_pending = set()
        pending = [p for p in self.paths if p not in self.page_cache and p not in self._page_pending]
        if not pending:
            return
        # Mark office page counts unknown until converted, don't display invented totals.
        office = [p for p in pending if p.suffix.lower() in OFFICE_EXTS]
        for p in office:
            self.page_cache[p] = '—'
        pending = [p for p in pending if p.suffix.lower() not in OFFICE_EXTS]
        if not pending:
            return
        self._page_pending.update(pending)
        def worker(files):
            for path in files:
                try:
                    if path.suffix.lower() == '.pdf':
                        reader = PdfReader(str(path), strict=False)
                        count = len(reader.pages) if not reader.is_encrypted else 'Lock'
                    else:
                        with Image.open(path) as img:
                            count = getattr(img, 'n_frames', 1)
                except Exception:
                    count = '?'
                self.after(0, self._set_page_count, path, count)
        threading.Thread(target=worker, args=(pending,), daemon=True).start()

    def _set_page_count(self, path, count):
        self.page_cache[path] = count
        self._page_pending.discard(path)
        # Update only currently matching rows; never corrupt reordered row indexes.
        self.refresh_view()

    def preview_pages(self):
        if self.busy:
            return
        selected = self.selected_indices()
        if not selected:
            messagebox.showinfo('Preview Pages', 'Select a PDF or image in the file list first.')
            return
        path = self.paths[selected[0]]
        if path.suffix.lower() in OFFICE_EXTS:
            messagebox.showinfo('Preview Pages', 'Office page preview is not available before conversion. Select a PDF or image.')
            return
        pop = tk.Toplevel(self, bg=BG)
        pop.title(f'Page preview — {path.name}')
        pop.geometry('680x670')
        pop.minsize(440, 390)
        current = [0]
        render_handle = None
        try:
            if path.suffix.lower() == '.pdf':
                render_handle = fitz.open(str(path))
                length = len(render_handle)
                if render_handle.needs_pass:
                    raise ValueError('This PDF is password protected.')
            else:
                with Image.open(path) as img:
                    length = getattr(img, 'n_frames', 1)
        except Exception as exc:
            if render_handle is not None:
                render_handle.close()
            pop.destroy()
            messagebox.showerror('Preview error', str(exc))
            return
        if length < 1:
            if render_handle is not None:
                render_handle.close()
            pop.destroy()
            messagebox.showinfo('Preview', 'File has no pages.')
            return
        outer = tk.Frame(pop, bg=BG)
        outer.pack(fill='both', expand=True, padx=14, pady=12)
        label = tk.Label(outer, bg=SURFACE, fg=TEXT, text='Rendering...', relief='flat')
        label.pack(fill='both', expand=True)
        footer = tk.Frame(outer, bg=BG)
        footer.pack(fill='x', pady=(9, 0))
        indicator = tk.StringVar()
        picture = [None]
        def render():
            if not pop.winfo_exists():
                return
            try:
                if render_handle is not None:
                    page = render_handle.load_page(current[0])
                    # Limit raster size for smooth navigation and controlled memory.
                    scale = min(1.6, 850 / max(page.rect.width, 1), 1050 / max(page.rect.height, 1))
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                else:
                    with Image.open(path) as img:
                        img.seek(current[0])
                        image = ImageOps.exif_transpose(img.convert('RGB'))
                w = max(label.winfo_width() - 20, 330)
                h = max(label.winfo_height() - 20, 260)
                image.thumbnail((w, h), Image.Resampling.LANCZOS)
                picture[0] = ImageTk.PhotoImage(image, master=pop)
                label.configure(image=picture[0], text='')
            except Exception as exc:
                label.configure(image='', text=f'Cannot preview page: {exc}')
            indicator.set(f'Page {current[0] + 1} of {length}  •  {path.name}')
        def step(n):
            current[0] = max(0, min(length - 1, current[0] + n))
            render()
        self.button(footer, '◀ Previous', lambda: step(-1)).pack(side='left')
        tk.Label(footer, textvariable=indicator, bg=BG, fg=MUTED,
                 font=('Segoe UI', 9)).pack(side='left', expand=True)
        self.button(footer, 'Next ▶', lambda: step(1)).pack(side='right')
        def close():
            if render_handle is not None:
                render_handle.close()
            pop.destroy()
        pop.protocol('WM_DELETE_WINDOW', close)
        pop.after(120, render)

    def open_folder(self):
        if self.busy:
            return
        name = filedialog.askdirectory(title='Choose folder with PDFs / images / Office files')
        if not name:
            return
        self._remember()
        self.paths = sorted((p for p in Path(name).iterdir() if p.is_file() and p.suffix.lower() in EXTS),
                            key=sort_key)
        self.page_choices = {p: expr for p, expr in self.page_choices.items() if p in self.paths}
        self.location.set(name)
        self.refresh()
        self.status.set(f'{len(self.paths)} file(s) loaded in natural order.')

    def add_files(self):
        if self.busy:
            return
        names = filedialog.askopenfilenames(title='Add PDFs / images / Office files',
            filetypes=[('PDF, images and Office', '*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.doc *.docx *.docm *.xls *.xlsx *.xlsm *.ppt *.pptx *.pptm'), ('All files', '*.*')])
        if names:
            self.location.set('Files from one or more locations')
            self.add_paths([Path(n) for n in names], sort_new=True)

    def sort_naturally(self):
        if self.busy:
            return
        self._remember()
        self.paths.sort(key=sort_key)
        self.refresh()
        self.status.set('List sorted naturally.')

    def clear(self):
        if self.busy:
            return
        if self.paths:
            self._remember()
        self.paths.clear()
        self.page_choices.clear()
        self.location.set('No folder selected')
        self.refresh()
        self.status.set('List cleared.')

    def selected_indices(self):
        return sorted(int(i) for i in self.tree.selection())

    def _delete_selected_key(self, event=None):
        # Delete affects files only while the file list has keyboard focus.
        # Other text boxes retain their normal editing keys.
        if not self.busy:
            self.remove_selected()
        return 'break'

    def remove_selected(self):
        if self.busy:
            return
        indices = set(self.selected_indices())
        if not indices:
            return
        self._remember()
        self.paths = [p for i, p in enumerate(self.paths) if i not in indices]
        self.page_choices = {p: v for p, v in self.page_choices.items() if p in self.paths}
        # Do not select rows that slid into the removed indexes.
        # A second press of Delete must be a no-op until the user reselects.
        self.tree.selection_remove(self.tree.selection())
        self.tree.focus('')
        self.selection_text.set('')
        self.refresh()
        self.status.set(f'Removed {len(indices)} row(s).')

    def move_selected(self, direction):
        if self.busy:
            return
        selected = self.selected_indices()
        if not selected:
            return
        self._remember()
        marks = set(selected)
        order = range(1, len(self.paths)) if direction == -1 else range(len(self.paths)-2, -1, -1)
        for i in order:
            if direction == -1 and i in marks and i-1 not in marks:
                self.paths[i-1], self.paths[i] = self.paths[i], self.paths[i-1]
                marks.remove(i)
                marks.add(i-1)
            if direction == 1 and i in marks and i+1 not in marks:
                self.paths[i+1], self.paths[i] = self.paths[i], self.paths[i+1]
                marks.remove(i)
                marks.add(i+1)
        self.refresh(sorted(marks))

    def mouse_press(self, event):
        if self.search_text.get().strip():
            self.drag_index = None
            self.drag_item = None
            return
        if self.busy or self.tree.identify_region(event.x, event.y) != 'cell':
            self.drag_index = None
            self.drag_item = None
            return
        item = self.tree.identify_row(event.y)
        self.drag_index = int(item) if item else None
        self.drag_item = item or None
        self.drag_start_y = event.y
        self.drag_started = False

    def mouse_drag(self, event):
        if self.busy or self.drag_item is None:
            return
        if not self.drag_started and abs(event.y - self.drag_start_y) < 7:
            return
        if not self.drag_started:
            self.drag_started = True
            self.tree.configure(cursor='fleur')
        target = self.tree.identify_row(event.y)
        if target and target != self.drag_item:
            self.tree.move(self.drag_item, '', self.tree.index(target))
            # Immediate visual movement with live row numbering.
            for number, item in enumerate(self.tree.get_children(), 1):
                values = list(self.tree.item(item, 'values'))
                values[0] = number
                self.tree.item(item, values=values)
        self.tree.selection_set(self.drag_item)
        self.tree.focus(self.drag_item)
        self.tree.see(self.drag_item)

    def mouse_release(self, event):
        self.tree.configure(cursor='')
        if self.drag_item is None or not self.drag_started or self.busy:
            self.drag_item = None
            self.drag_index = None
            self.drag_started = False
            return
        source = self.drag_index
        order = [int(item) for item in self.tree.get_children()]
        dest = order.index(source)
        if dest != source:
            self._remember()
        self.paths = [self.paths[i] for i in order]
        moved = self.paths[dest]
        self.drag_item = None
        self.drag_index = None
        self.drag_started = False
        self.refresh([dest])
        if dest != source:
            self.status.set(f'Moved {moved.name} to position {dest + 1}.')

    def use_highlighted(self):
        self.selection_text.set(' '.join(str(i + 1) for i in self.selected_indices()))

    def selected_jobs(self):
        paths = self.selected_paths()
        jobs = []
        for p in paths:
            expression = self.page_choices.get(p)
            if p.suffix.lower() == '.pdf' and expression:
                reader = PdfReader(str(p), strict=False)
                if reader.is_encrypted and reader.decrypt('') == 0:
                    raise ValueError(f'Password-protected PDF: {p.name}')
                jobs.append((p, selected_page_indices(expression, len(reader.pages))))
            else:
                jobs.append((p, None))
        return jobs

    def selected_paths(self):
        if not self.paths:
            raise ValueError('Add PDF, image or Office files first.')
        indices = parse_selection(self.selection_text.get(), len(self.paths))
        if not indices:
            raise ValueError('No valid files selected.')
        return [self.paths[i] for i in indices]

    def preview(self):
        try:
            files = self.selected_paths()
        except ValueError as exc:
            messagebox.showerror('Selection error', str(exc))
            return
        pop = tk.Toplevel(self, bg=BG)
        pop.title(f'Merge order — {len(files)} files')
        pop.geometry('690x470')
        preview = tk.Text(pop, bg=SURFACE, fg=TEXT, insertbackground=TEXT, borderwidth=0,
                          font=('Consolas', 11), wrap='none', padx=14, pady=14)
        preview.pack(fill='both', expand=True, padx=12, pady=12)
        preview.insert('1.0', '\n'.join(f'{i}. {p.name}' for i, p in enumerate(files, 1)))
        preview.config(state='disabled')

    def start_merge(self):
        if self.busy:
            return
        try:
            files = self.selected_jobs()
        except (ValueError, OSError) as exc:
            messagebox.showerror('Selection error', str(exc))
            return
        target = filedialog.asksaveasfilename(title='Save merged PDF', defaultextension='.pdf',
                initialfile=f'Wadaa_PDF_{datetime.now():%Y-%m-%d_%H-%M-%S}.pdf',
                filetypes=[('PDF files', '*.pdf')])
        if not target:
            return
        destination = Path(target)
        if any(destination.resolve() == p.resolve() for p, _ in files):
            messagebox.showerror('Unsafe output', 'Choose a different file name; cannot overwrite a source PDF.')
            return
        if destination.exists() and not messagebox.askyesno('Overwrite?', f'Replace {destination.name}?'):
            return
        self.busy = True
        self.merge_button.configure(state='disabled')
        self.progress['value'] = 0
        self.status.set('Merging...')
        self._update_history_buttons()
        threading.Thread(target=self.worker, args=(files, destination), daemon=True).start()

    def worker(self, files, destination):
        try:
            merge_files(files, destination, lambda done, total, name:
                self.after(0, self.update_progress, done, total, name))
        except Exception as exc:
            self.after(0, self.finish, False, str(exc), destination)
        else:
            self.after(0, self.finish, True, f'Merged {len(files)} file(s).', destination)

    def update_progress(self, done, total, name):
        self.progress['value'] = done * 100 / total if total else 0
        self.status.set(f'{done}/{total} — {name}')

    def finish(self, ok, text, destination):
        self.busy = False
        self.merge_button.configure(state='normal')
        self._update_history_buttons()
        if ok:
            self.output_file = destination
            self.progress['value'] = 100
            self.status.set(f'Completed: {destination}')
            if self.open_after_merge.get():
                try:
                    if sys.platform == 'win32':
                        os.startfile(str(destination))
                    else:
                        webbrowser.open(destination.as_uri())
                except OSError as exc:
                    messagebox.showwarning('Open PDF', f'PDF saved but could not open it: {exc}')
            elif messagebox.askyesno('Merge completed', f'{text}\n\n{destination}\n\nOpen output folder?'):
                self.open_output_folder()
        else:
            self.status.set(f'Error: {text}')
            messagebox.showerror('Merge failed', text)

    def open_output_folder(self):
        if not self.output_file:
            messagebox.showinfo('Output folder', 'Merge a PDF first.')
            return
        try:
            if sys.platform == 'win32':
                import subprocess
                subprocess.Popen(['explorer', '/select,', str(self.output_file)])
            else:
                webbrowser.open(self.output_file.parent.as_uri())
        except OSError as exc:
            messagebox.showerror('Open folder', str(exc))


if __name__ == '__main__':
    App().mainloop()
