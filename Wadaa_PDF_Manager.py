"""Wadaa PDF Manager — dark Windows GUI."""
from __future__ import annotations

import ctypes
import io
import json
import math
import shutil
import subprocess
import tempfile
import unicodedata
import os
import platform
import re
import sys
import threading
import webbrowser
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZIP_DEFLATED, ZipFile
from collections import OrderedDict
from datetime import datetime
from itertools import groupby
from statistics import median
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
BG = '#0D1117'
SURFACE = '#151B23'
BORDER = '#26313D'
TEXT = '#F8F5EF'
MUTED = '#98A6B8'
ACCENT = '#62D6C0'
HOVER = '#1D2834'
SELECT = '#294154'
ACCENT_TEXT = '#10221F'
ACCENT_HOVER = '#78E0CC'
ACCENT_PRESSED = '#4ABCA6'
LIGHT_COLORS = {
    BG: '#F8F5EF', SURFACE: '#F2EBDF', BORDER: '#D7C8B6',
    TEXT: '#4B3528', MUTED: '#755B49', ACCENT: '#287D74',
    HOVER: '#EAE0D1', SELECT: '#E0D4C3', ACCENT_TEXT: '#FFFFFF',
    ACCENT_HOVER: '#266F67', ACCENT_PRESSED: '#205F58',
}
UI_AR = {
    'File': 'ملف', 'Edits': 'تعديلات', 'Add Files': 'إضافة ملفات',
    '+ Add Files': '+ إضافة ملفات', 'Open Folder': 'فتح مجلد',
    'Export File to PDF': 'تصدير إلى PDF', 'Export File to Word': 'تصدير إلى Word',
    'Export File to PowerPoint': 'تصدير إلى PowerPoint',
    'Export File to Images': 'تصدير إلى صور', 'Open Output Folder': 'فتح مجلد الإخراج',
    'Exit': 'خروج', 'Theme': 'المظهر', 'Dark': 'داكن', 'Light': 'فاتح',
    'Language': 'اللغة', 'English': 'الإنجليزية', 'Arabic': 'العربية',
    'Save Last Session': 'حفظ آخر جلسة', 'ON': 'تشغيل', 'OFF': 'إيقاف',
    'Options...': 'الإعدادات…', 'Options': 'الإعدادات', 'Close': 'إغلاق',
    'Check for Updates': 'التحقق من التحديثات', 'Donate and Support': 'التبرع والدعم',
    'Support / Donate': 'الدعم / التبرع', 'About': 'عن البرنامج',
    'About & Support': 'عن البرنامج والدعم', 'Contact and Report': 'التواصل والإبلاغ',
    'Contact Support': 'التواصل مع الدعم', 'Report a Bug': 'الإبلاغ عن مشكلة',
    'GitHub Repository': 'مستودع GitHub', 'Sort by Name (A-Z)': 'ترتيب بالاسم',
    'Clear All': 'مسح الكل', 'Search files:': 'بحث عن ملفات',
    'Clear Search': 'مسح البحث', 'Move Up': 'تحريك لأعلى', 'Move Down': 'تحريك لأسفل',
    'Remove Selected': 'إزالة المحدد', 'Undo': 'تراجع', 'Redo': 'إعادة',
    'Use Highlighted': 'استخدام الملفات المظللة', 'Preview Merge Order': 'معاينة ترتيب الدمج',
    'Preview Pages': 'معاينة الصفحات', 'Select PDF Pages': 'تحديد صفحات PDF',
    'Merge to PDF...': 'دمج إلى PDF…', 'Open PDF after merge': 'فتح PDF بعد الدمج',
    'File name (current merge order)': 'اسم الملف (ترتيب الدمج الحالي)',
    'Type': 'النوع', 'Pages': 'الصفحات', 'Last Session: ON': 'آخر جلسة: تشغيل',
    'Last Session: OFF': 'آخر جلسة: إيقاف', 'Apply': 'تطبيق', 'Cancel': 'إلغاء',
    'Select All': 'تحديد الكل', 'Clear Selection': 'مسح التحديد',
    'Apply Selection': 'اعتماد التحديد', 'Previous': 'السابق', 'Next': 'التالي',
    'Rendering...': 'جارٍ عرض الصفحة…', 'Select pages - {name}': 'تحديد صفحات - {name}',
    'Page preview - {name}': 'معاينة الصفحات - {name}',
    'Page {page} of {total}  •  {name}': 'صفحة {page} من {total}  •  {name}',
    '{count} selected': 'عدد الصفحات المحددة: {count}',
    'No folder selected': 'لم يتم اختيار مجلد',
    'Files from one or more locations': 'ملفات من مجلد واحد أو أكثر',
    'Drop PDF, images or Office documents here, or choose a folder.':
        'اسحب ملفات PDF أو الصور أو مستندات Office هنا، أو اختر مجلدًا.',
    'PDF + images + Office  •  Drag & drop  •  Reorder  •  Merge':
        'PDF وصور وOffice  •  سحب وإفلات  •  ترتيب  •  دمج',
    'Drop files into the list  •  Drag a row to move it  •  Ctrl/Shift to select multiple':
        'أسقط الملفات في القائمة  •  اسحب الصف لتحريكه  •  Ctrl/Shift للتحديد المتعدد',
    'Optional custom selection: 1-9 12 15-18  •  Empty = all in list order':
        'تحديد اختياري للملفات: 1-9 12 15-18  •  اتركه فارغًا لدمج الكل بترتيب القائمة',
    'Examples: 1-5 8 10-12  |  Blank = all pages':
        'أمثلة: 1-5 8 10-12  |  اتركه فارغًا لاستخدام جميع الصفحات',
    '{name}  •  {total} pages': '{name}  •  {total} صفحة',
    'Click a page to include or remove it.': 'اضغط على الصفحة لتحديدها أو إلغاء تحديدها.',
    'No usable text on output page(s) {pages}. OCR is required to recover text from scans.':
        'لم يُعثر على نص صالح للاستخراج في صفحات الإخراج {pages}. يلزم OCR لاستخراج نص من الصفحات الممسوحة.',
    'Page-sized raster images omitted on output page(s) {pages}.':
        'تم تجاوز صور بحجم الصفحة في صفحات الإخراج {pages}.',
    'Some images could not be extracted on output page(s) {pages}.':
        'تعذر استخراج بعض الصور في صفحات الإخراج {pages}.',
    'Table analysis failed on output page(s) {pages}. Extracted text was retained.':
        'تعذر تحليل الجداول في صفحات الإخراج {pages}، وتم الاحتفاظ بالنص المستخرج.',
    'Some vector drawings were omitted on output page(s) {pages}.':
        'تعذر نقل بعض الرسومات المتجهة في صفحات الإخراج {pages}.',
    'Choose a parent folder for the JPG export': 'اختر المجلد الذي سيحتوي على مجلد صور JPG',
    'Save merged PDF': 'حفظ ملف PDF المدمج', 'Save Word document': 'حفظ مستند Word',
    'Save PowerPoint presentation': 'حفظ عرض PowerPoint',
    'Export': 'تصدير', 'Export completed': 'اكتمل التصدير', 'Export failed': 'تعذر التصدير',
    'Wait for the export to finish before closing.': 'انتظر انتهاء التصدير قبل إغلاق البرنامج.',
    'Completed: {path}': 'اكتمل الحفظ: {path}', 'Merging...': 'جارٍ الدمج…',
    'Page selection saved for {name}: {pages}': 'تم حفظ تحديد صفحات {name}: {pages}',
    'all pages': 'كل الصفحات',
    'Version: {version}': 'الإصدار: {version}',
    'Email: {email}': 'البريد الإلكتروني: {email}',
    'Developed by WadaaDev': 'تطوير WadaaDev',
    'Powered by WadaaDev': 'بدعم من WadaaDev',
    'Selection error': 'خطأ في التحديد',
    'Unsafe output': 'مسار إخراج غير صالح',
    'Use the .{kind} extension.': 'استخدم امتداد .{kind}.',
    'Choose a different file name; cannot overwrite a source file.': 'اختر اسمًا مختلفًا لحماية الملف الأصلي.',
    'Overwrite?': 'استبدال الملف؟',
    'Replace {name}?': 'هل تريد استبدال {name}؟',
    'PDF files': 'ملفات PDF',
    'Word documents': 'مستندات Word',
    'PowerPoint presentations': 'عروض PowerPoint',
    'Open PDF': 'فتح PDF',
    'Merge completed': 'اكتمل الدمج',
    'Merge failed': 'تعذر الدمج',
    'Open output folder?': 'هل تريد فتح مجلد الإخراج؟',
    'Open folder': 'فتح المجلد',
    'Output folder': 'مجلد الإخراج',
    'Export a file first.': 'صدّر ملفًا أولًا.',
    'Error: {error}': 'خطأ: {error}',
    'Exported {count} page(s).': 'تم تصدير {count} صفحة.',
    'Merged {count} file(s).': 'تم دمج {count} ملف.',
    'Wait for the export to finish before changing pages.': 'انتظر انتهاء التصدير قبل تغيير الصفحات.',
    'Select a PDF or image in the file list first.': 'اختر ملف PDF أو صورة من القائمة أولًا.',
    'Office page preview is not available before conversion. Select a PDF or image.': 'معاينة ملفات Office متاحة بعد تحويلها. اختر ملف PDF أو صورة.',
    'This PDF is password protected.': 'ملف PDF محمي بكلمة مرور.',
    'File has no pages.': 'الملف لا يحتوي على صفحات.',
    'Preview error': 'خطأ في المعاينة',
    'Cannot preview page:': 'تعذرت معاينة الصفحة:',
    'Select one PDF file in the list first.': 'اختر ملف PDF واحدًا من القائمة أولًا.',
    'Page selection is available for PDFs only.': 'تحديد الصفحات متاح لملفات PDF فقط.',
    'Cannot read {name}: {error}': 'تعذرت قراءة {name}: {error}',
    'Invalid pages': 'تحديد صفحات غير صالح',
    'Merge order - {count} files': 'ترتيب الدمج - {count} ملف',
    'Choose folder with PDFs / images / Office files': 'اختر مجلدًا يحتوي على ملفات PDF أو صور أو ملفات Office',
    'Add PDFs / images / Office files': 'إضافة ملفات PDF أو صور أو ملفات Office',
    'PDF, images and Office': 'ملفات PDF والصور وOffice',
    'All files': 'كل الملفات',
    'Last Session': 'آخر جلسة',
    'Restore the previously saved list now?': 'هل تريد استعادة القائمة المحفوظة؟',
    'Undo completed.': 'تم التراجع.',
    'Redo completed.': 'تمت الإعادة.',
    'List sorted naturally.': 'تم ترتيب القائمة بالاسم والأرقام.',
    'List cleared.': 'تم مسح القائمة.',
    'Removed {count} row(s).': 'تم حذف {count} صف.',
    'Moved {name} to position {position}.': 'تم نقل {name} إلى الموضع {position}.',
    '{count} file(s) loaded in natural order.': 'تم تحميل {count} ملف بالترتيب الطبيعي.',
    '{count} new file(s) added; {duplicates} duplicate(s) skipped. {total} total.': 'تمت إضافة {count} ملف. تم تجاوز {duplicates} ملف مكرر. الإجمالي {total}.',
    'Restored {count} file(s) from the previous session.': 'تمت استعادة {count} ملف من الجلسة السابقة.',
    'Unable to save settings: {error}': 'تعذر حفظ الإعدادات: {error}',
    'Unable to save session: {error}': 'تعذر حفظ الجلسة: {error}',
    'Previous session could not be restored: {error}': 'تعذرت استعادة الجلسة السابقة: {error}',
    'Unable to open link': 'تعذر فتح الرابط',
    '{files} files  •  {pages} pages': '{files} ملف  •  {pages} صفحة',
    ' (+ {count} uncounted files)': ' (+ {count} ملف دون عدد صفحات معروف)',
    'Yes': 'نعم', 'No': 'لا', 'OK': 'موافق', 'Updates': 'التحديثات',
    'Online updates are disabled in this build.': 'التحقق من التحديثات عبر الإنترنت معطّل في هذه النسخة.',
    'An update check or download is already running.': 'التحقق من التحديثات أو تنزيلها جارٍ بالفعل.',
    'You are using the latest version.': 'أنت تستخدم أحدث إصدار.',
    'Update available': 'يتوفر تحديث', 'Download version {version}?': 'هل تريد تنزيل الإصدار {version}؟',
    'Version {version} is available. Install the app with Setup to use automatic upgrades.':
        'الإصدار {version} متاح. ثبّت البرنامج باستخدام ملف التثبيت لتفعيل التحديث التلقائي.',
    'Downloading update': 'جارٍ تنزيل التحديث', 'Install update': 'تثبيت التحديث',
    'Wait for the PDF merge to finish before installing.': 'انتظر اكتمال دمج PDF قبل التثبيت.',
    'Close Wadaa PDF Manager and install the update?': 'هل تريد إغلاق Wadaa PDF Manager وتثبيت التحديث؟',
    'The session could not be saved. The update was postponed.': 'تعذر حفظ الجلسة، فتأجّل التحديث.',
}
DONATE = 'https://linktr.ee/WadaaTips'
SUPPORT_EMAIL = 'wadaadev.support@gmail.com'
REPOSITORY_URL = f'https://github.com/{REPO}'
BUG_REPORT_URL = f'{REPOSITORY_URL}/issues'


LTR_UI_TEXT = re.compile(r'[A-Za-z0-9][A-Za-z0-9_.:/\\+%–@()-]*(?:[ \t]+[A-Za-z0-9][A-Za-z0-9_.:/\\+%–@()-]*)*')
ARABIC_UI_TEXT = re.compile(r'[\u0600-\u06ff]')


def rtl_display(text):
    def line_direction(line):
        if not ARABIC_UI_TEXT.search(line) or line.startswith('\u202b'):
            return line
        # Use embedding marks understood by Windows text controls. Layout
        # direction is independent from the placement of the widgets.
        return '\u202b' + LTR_UI_TEXT.sub(lambda match: '\u202a' + match[0] + '\u202c', line) + '\u202c'

    return '\n'.join(line_direction(line) for line in text.split('\n'))


def set_windows_caption(window, dark, background, foreground, border):
    if sys.platform != 'win32' or window.overrideredirect():
        return
    try:
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        user32.GetAncestor.argtypes = (wintypes.HWND, wintypes.UINT)
        user32.GetAncestor.restype = wintypes.HWND
        hwnd = user32.GetAncestor(window.winfo_id(), 2)
        if not hwnd:
            return
        set_attribute = ctypes.windll.dwmapi.DwmSetWindowAttribute
        set_attribute.argtypes = (wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD)
        set_attribute.restype = ctypes.c_long
        enabled = wintypes.BOOL(dark)
        result = set_attribute(hwnd, 20, ctypes.byref(enabled), ctypes.sizeof(enabled))
        if result:
            set_attribute(hwnd, 19, ctypes.byref(enabled), ctypes.sizeof(enabled))
        for attribute, hex_color in ((35, background), (36, foreground), (34, border)):
            red, green, blue = bytes.fromhex(hex_color[1:])
            value = wintypes.DWORD(red | (green << 8) | (blue << 16))
            set_attribute(hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))
    except (OSError, AttributeError, tk.TclError):
        # Older systems can keep their native caption.
        pass


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


def page_selection_expression(pages, total):
    # Blank is the existing representation for all pages.
    if len(pages) == total:
        return ''
    ranges = []
    for _, group in groupby(enumerate(sorted(pages)), lambda pair: pair[1] - pair[0]):
        first = last = next(group)[1] + 1
        for _, page in group:
            last = page + 1
        ranges.append(str(first) if first == last else f'{first}-{last}')
    return ' '.join(ranges)


def merge_files(paths: list, destination: Path, progress, *, image_pages=None):
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
                    first_page = len(writer.pages)
                    writer.append(PdfReader(stream), import_outline=False)
                    if image_pages is not None:
                        image_pages.update(range(first_page, len(writer.pages)))
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



PACKAGE_NS = 'http://schemas.openxmlformats.org/package/2006'
OFFICE_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PRESENTATION_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'


def write_relationships(package, name, links):
    root = ET.Element('Relationships', xmlns=f'{PACKAGE_NS}/relationships')
    for identifier, kind, target in links:
        ET.SubElement(root, 'Relationship', Id=identifier, Type=f'{OFFICE_NS}/{kind}', Target=target)
    package.writestr(name, ET.tostring(root, encoding='utf-8'))


def write_content_types(package, parts):
    root = ET.Element('Types', xmlns=f'{PACKAGE_NS}/content-types')
    for extension, content_type in (
        ('rels', 'application/vnd.openxmlformats-package.relationships+xml'),
        ('xml', 'application/xml'), ('png', 'image/png'),
    ):
        ET.SubElement(root, 'Default', Extension=extension, ContentType=content_type)
    for name, content_type in parts:
        ET.SubElement(root, 'Override', PartName=f'/{name}',
                      ContentType=f'application/vnd.openxmlformats-officedocument.{content_type}+xml')
    package.writestr('[Content_Types].xml', ET.tostring(root, encoding='utf-8'))


def page_jpeg(page):
    # Render at up to 160 dpi, with a pixel bound even for poster-sized PDFs.
    scale = min(160 / 72, 2600 / max(page.rect.width, page.rect.height))
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
    return pixmap.tobytes('jpeg', jpg_quality=92)


def fitted_size(rect, width, height):
    scale = min(width / rect.width, height / rect.height)
    return max(1, int(rect.width * scale)), max(1, int(rect.height * scale))


def drawing_picture(prefix, identifier, relationship, width, height, x=0, y=0, rotation=0):
    return f'''<{prefix}:pic>
        <{prefix}:nvPicPr><{prefix}:cNvPr id="{identifier}" name="Image {identifier}"/>
          <{prefix}:cNvPicPr><a:picLocks noChangeAspect="1"/></{prefix}:cNvPicPr>
          {'<p:nvPr/>' if prefix == 'p' else ''}
        </{prefix}:nvPicPr>
        <{prefix}:blipFill><a:blip r:embed="{relationship}"/>
          <a:stretch><a:fillRect/></a:stretch>
        </{prefix}:blipFill>
        <{prefix}:spPr><a:xfrm rot="{rotation}"><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </{prefix}:spPr>
    </{prefix}:pic>'''



PDF_FONT_ALIASES = {
    'Helvetica': 'Arial', 'HelveticaNeue': 'Arial', 'ArialMT': 'Arial',
    'Times': 'Times New Roman', 'Times-Roman': 'Times New Roman',
    'TimesRoman': 'Times New Roman', 'TimesNewRomanPS': 'Times New Roman',
    'TimesNewRomanPSMT': 'Times New Roman', 'Courier': 'Courier New',
    'CourierNewPS': 'Courier New', 'CourierNewPSMT': 'Courier New',
}
OCR_PLACEHOLDER = 'OCR required to recover text from this page.'
OFFICE_WARNINGS = {
    'ocr': 'No usable text on output page(s) {pages}. OCR is required to recover text from scans.',
    'page_image': 'Page-sized raster images omitted on output page(s) {pages}.',
    'image': 'Some images could not be extracted on output page(s) {pages}.',
    'table': 'Table analysis failed on output page(s) {pages}. Extracted text was retained.',
    'drawing': 'Some vector drawings were omitted on output page(s) {pages}.',
}


def office_text(text):
    return escape(re.sub(r'[\x00-\x08\x0b-\x1f\ud800-\udfff\ufffe\uffff]', '', text))


def office_font(name):
    name = re.sub(r'^[A-Z]{6}\+', '', name)
    name = re.sub(r'[-, ]?(?:BoldItalic|BoldOblique|Bold|Italic|Oblique|Regular)(?:MT)?$',
                  '', name, flags=re.IGNORECASE).rstrip(' ,-')
    return quoteattr(PDF_FONT_ALIASES.get(name, name) or 'Arial')


def is_rtl(text):
    return next((direction in ('R', 'AL') for ch in text
                 if (direction := unicodedata.bidirectional(ch)) in ('L', 'R', 'AL')), False)


def positioned_pdf_spans(span, direction):
    # Split synthetic padding and large PDF text-position jumps into separate runs.
    chars = span['chars']
    dx, dy = direction
    groups, current, pending = [], [], []
    end = None
    for char in chars:
        if char['c'].isspace():
            if current:
                pending.append(char)
            continue
        x, y = char['origin']
        along = x * dx + y * dy
        corners = fitz.Rect(char['bbox'])
        trailing = max(point.x * dx + point.y * dy
                       for point in (corners.tl, corners.tr, corners.bl, corners.br))
        if current and (len(pending) >= 3 or (end is not None and along - end > span['size'] * .8)
                        or char['c'] == '\ufffd' or current[-1]['c'] == '\ufffd'):
            groups.append(current)
            current, pending = [], []
        current.extend(pending)
        pending = []
        current.append(char)
        end = trailing
    if current:
        groups.append(current)
    formatted = {key: value for key, value in span.items() if key != 'chars'}
    for group in groups:
        bbox = fitz.Rect(group[0]['bbox'])
        for char in group[1:]:
            bbox |= fitz.Rect(char['bbox'])
        yield dict(formatted, text=''.join(char['c'] for char in group),
                   bbox=tuple(bbox), origin=group[0]['origin'])


def spans_on_baseline(group, candidate, direction):
    first, last = group[0], group[-1]
    dx, dy = direction
    x, y = candidate['origin']
    bx, by = first['origin']
    fraction = 'fraction' in candidate or 'fraction' in last
    tolerance = .45 if fraction else .15
    if abs((y - by) * dx - (x - bx) * dy) > max(1, first['size'] * tolerance):
        return False
    bounds = fitz.Rect(last['bbox'])
    end = max(point.x * dx + point.y * dy
              for point in (bounds.tl, bounds.tr, bounds.bl, bounds.br))
    delimiter = 'fraction_parenthesis' in candidate or 'fraction_parenthesis' in last
    gap = 1.5 if fraction or delimiter else .7
    return abs(x * dx + y * dy - end) <= max(1, last['size'] * gap)


def joined_pdf_spans(spans, direction, *, separate_parentheses=False):
    # Retain explicit word spaces through extraction and reconstruct ordinary
    # spaces between font runs from their positions.
    group = []
    dx, dy = direction
    for span in spans:
        standalone = separate_parentheses and ('fraction_parenthesis' in span or
                                                (group and 'fraction_parenthesis' in group[-1]))
        if group and (standalone or not spans_on_baseline(group, span, direction)):
            yield group
            group = []
        if group and dx > .9 and not is_rtl(span['text']):
            previous = group[-1]
            gap = span['bbox'][0] - previous['bbox'][2]
            if gap > max(1, previous['size'] * .13):
                group[-1] = dict(previous, text=previous['text'] + ' ')
        group.append(span)
    if group:
        yield group


def extract_office_page(page, imported_image=False):
    flags = fitz.TEXTFLAGS_RAWDICT & ~fitz.TEXT_CID_FOR_UNKNOWN_UNICODE
    blocks = page.get_text('rawdict', sort=True, flags=flags)['blocks']
    page_rect = page.rect * page.derotation_matrix
    has_text = False
    content = []
    warnings = set()
    for block in blocks:
        if block['type'] == 0:
            lines = []
            for line in block['lines']:
                spans = [run for span in line['spans']
                         for run in positioned_pdf_spans(span, line['dir'])]
                if spans:
                    lines.append(dict(line, spans=spans))
            block['lines'] = lines
            if not block['lines']:
                continue
            has_text |= any(span['text'].replace('\ufffd', '').strip()
                            for line in block['lines'] for span in line['spans'])
        elif block['type'] == 1:
            if not imported_image and (
                fitz.Rect(block['bbox']) & page_rect
            ).get_area() >= 0.90 * page_rect.get_area():
                warnings.add('page_image')
                continue
        else:
            continue
        content.append(block)
    if not has_text and not imported_image:
        try:
            has_content = bool(blocks or page.get_bboxlog())
        except (RuntimeError, ValueError):
            has_content = True
        if has_content:
            warnings.add('ocr')
    return content, warnings


def write_office_image(package, folder, image_id, block):
    path = f'media/image-{image_id}.png'
    with Image.open(io.BytesIO(block['image'])) as source:
        bitmap = source.convert('RGBA' if 'A' in source.getbands() else 'RGB')
        if block.get('mask'):
            with Image.open(io.BytesIO(block['mask'])) as mask:
                bitmap.putalpha(mask.convert('L').resize(bitmap.size))
        output = io.BytesIO()
        bitmap.save(output, 'PNG')
    package.writestr(f'{folder}/{path}', output.getvalue())
    return path


def record_office_warnings(report, number, warnings):
    for reason in warnings:
        report.setdefault(reason, []).append(number)


def pdf_fraction_runs(lines, drawings):
    # A short horizontal rule with aligned small glyphs above and below is a
    # fraction; require both parts to avoid treating ordinary underlines as math.
    candidates = [span for line in lines if line['dir'][0] > .95 for span in line['spans']
                  if re.fullmatch(r'[\w+−-]{1,5}', span['text'])]
    fractions, consumed = [], set()
    for drawing in drawings:
        for item in drawing['items']:
            if item[0] == 'l':
                x0, y = item[1]
                x1, y1 = item[2]
                if abs(y - y1) > .35:
                    continue
                x0, x1 = sorted((x0, x1))
            elif item[0] == 're' and item[1].height <= 1.5:
                x0, y, x1, _ = item[1]
            else:
                continue
            if not 2 <= x1 - x0 <= 38:
                continue
            nearby = []
            for span in candidates:
                if id(span) in consumed:
                    continue
                rect = fitz.Rect(span['bbox'])
                center = (rect.x0 + rect.x1) / 2
                delta = span['origin'][1] - y
                if x0 - 1 <= center <= x1 + 1 and rect.width <= x1 - x0 + 3:
                    if -1.2 * span['size'] <= delta <= 1.7 * span['size']:
                        nearby.append((delta, span))
            above = [span for delta, span in nearby if delta <= .35]
            below = [span for delta, span in nearby if delta > 2]
            if len(above) != 1 or len(below) != 1:
                continue
            numerator, denominator = above[0], below[0]
            if not .65 <= numerator['size'] / denominator['size'] <= 1.5:
                continue
            consumed.update((id(numerator), id(denominator)))
            bbox = fitz.Rect(numerator['bbox']) | fitz.Rect(denominator['bbox'])
            fraction = dict(numerator, bbox=tuple(bbox), origin=(x0, y + numerator['size'] * .35),
                            fraction=(numerator['text'], denominator['text']), rule_right=x1,
                            text=f"{numerator['text']}/{denominator['text']}")
            fractions.append(fraction)
    return fractions, consumed


def recover_fraction_parentheses(lines, fractions):
    spans = [span for line in lines if line['dir'][0] > .95 for span in line['spans']]
    unmapped = [span for span in spans if span['text'] == '\ufffd']
    for fraction in fractions:
        x, y = fraction['origin']
        right_edge = fraction['rule_right']
        size = fraction['size']
        nearby = [span for span in unmapped
                  if abs((span['bbox'][1] + span['bbox'][3]) / 2 - y) <= size * 1.6
                  and span['bbox'][2] - span['bbox'][0] <= size * 1.8]
        left = [span for span in nearby if -size * .4 <= x - span['bbox'][2] <= size * 5]
        right = [span for span in nearby
                 if -size * .4 <= span['bbox'][0] - right_edge <= size * 2]
        if len(left) != 1 or len(right) != 1 or left[0] is right[0]:
            continue
        opening, closing = left[0], right[0]
        # A tight pair can enclose just the fraction. A wider pair needs a
        # visible arithmetic operator between the missing glyphs.
        if x - opening['bbox'][2] > size * 1.5 and not any(
            opening['bbox'][0] <= span['bbox'][2] and span['bbox'][0] <= x
            and abs(span['origin'][1] - y) <= size
            and any(operator in span['text'] for operator in ('+', '-', '\u2212'))
            for span in spans
        ):
            continue
        for span, character in ((opening, '('), (closing, ')')):
            span['text'] = character
            span['font'] = 'Cambria Math'
            span['fraction_parenthesis'] = character
            unmapped.remove(span)
        fraction['parentheses'] = (opening, closing)


def apply_word_fractions(blocks, fractions, consumed):
    lines = [line for block in blocks for line in block['lines']]
    for line in lines:
        line['spans'] = [span for span in line['spans'] if id(span) not in consumed]
    for fraction in fractions:
        x, y = fraction['origin']
        neighbors = []
        parentheses = fraction.get('parentheses', ())
        for line in lines:
            if line['dir'][0] <= .95:
                continue
            for span in line['spans']:
                if span in parentheses:
                    continue
                if abs(span['origin'][1] - y) > max(3, span['size'] * .55):
                    continue
                start = parentheses[0]['bbox'][0] if parentheses else x
                stop = parentheses[1]['bbox'][2] if parentheses else fraction['bbox'][2]
                gap = max(start - span['bbox'][2], span['bbox'][0] - stop, 0) if parentheses else min(
                    abs(span['bbox'][2] - x), abs(span['bbox'][0] - fraction['bbox'][2]))
                if gap <= span['size'] * 1.5:
                    neighbors.append((gap, line, span))
        if not neighbors:
            blocks.append({'lines': [{'spans': [fraction], 'dir': (1, 0)}]})
            continue
        _, target, base = max(neighbors, key=lambda item: (item[2]['size'], -item[0]))
        baseline = base['origin'][1]
        for delimiter in parentheses:
            for line in lines:
                if delimiter in line['spans']:
                    line['spans'].remove(delimiter)
                    break
            target['spans'].append(dict(delimiter, origin=(delimiter['origin'][0], baseline),
                                        size=base['size'], flags=delimiter['flags'] & ~1))
        x_end = max(span['bbox'][2] for _, _, span in neighbors)
        for _, line, span in neighbors:
            if line is not target and abs(span['origin'][1] - baseline) <= max(2, base['size'] * .25):
                line['spans'].remove(span)
                target['spans'].append(span)
        for line in lines:
            for span in line['spans'][:]:
                rise = baseline - span['origin'][1]
                if (span['size'] <= base['size'] * .85 and 2 <= rise <= base['size']
                        and fraction['bbox'][2] <= span['bbox'][0] <= x_end + base['size'] * .5
                        and re.fullmatch(r'[\w+−-]{1,5}', span['text'])):
                    line['spans'].remove(span)
                    target['spans'].append(dict(span, origin=(span['origin'][0], baseline),
                                                flags=span['flags'] | 1))
        target['spans'].append(dict(fraction, origin=(x, baseline)))
        target['spans'].sort(key=lambda span: span['bbox'][0])


def word_math_run(text, size):
    return f'<m:r><w:rPr><w:sz w:val="{size}"/></w:rPr><m:t>{office_text(text)}</m:t></m:r>'


def word_fraction(span, size_scale):
    numerator, denominator = span['fraction']
    size = max(2, min(1638, round(span['size'] * size_scale * 2.7)))
    return (f'<m:f><m:num>{word_math_run(numerator, size)}</m:num>'
            f'<m:den>{word_math_run(denominator, size)}</m:den></m:f>')


def word_line_runs(line, size_scale):
    result = []
    pending = []
    for span in line:
        delimiter = span.get('fraction_parenthesis')
        if delimiter == '(' and not pending:
            pending.append(span)
            continue
        if not pending:
            result.append(word_run(span, size_scale))
            continue
        pending.append(span)
        if delimiter == ')':
            inner = pending[1:-1]
            if any('fraction' in item for item in inner):
                size = max(2, min(1638, round(pending[0]['size'] * size_scale * 2)))
                equation = ''.join(word_fraction(item, size_scale) if 'fraction' in item
                                   else word_math_run(item['text'], size) for item in inner)
                result.append('<m:oMath><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/>'
                              f'</m:dPr><m:e>{equation}</m:e></m:d></m:oMath>')
            else:
                result.extend(word_run(item, size_scale) for item in pending)
            pending = []
    result.extend(word_run(item, size_scale) for item in pending)
    return ''.join(result)


def word_run(span, size_scale=1):
    if 'fraction' in span:
        return f'<m:oMath>{word_fraction(span, size_scale)}</m:oMath>'
    flags = span.get('flags', 0)
    size = max(2, min(1638, round(span.get('size', 11) * size_scale * 2)))
    font = office_font(span.get('font', 'Arial'))
    color = f"{span.get('color', 0) & 0xFFFFFF:06X}"
    text = span['text']
    return f'''<w:r><w:rPr><w:rFonts w:ascii={font} w:hAnsi={font} w:eastAsia={font} w:cs={font}/>
        {'<w:b/><w:bCs/>' if flags & 16 else ''}{'<w:i/><w:iCs/>' if flags & 2 else ''}
        <w:color w:val="{color}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>
        {'<w:vertAlign w:val="superscript"/>' if flags & 1 else ''}
        {'<w:rtl/>' if is_rtl(text) else ''}</w:rPr>
        <w:t xml:space="preserve">{office_text(text)}</w:t></w:r>'''


def word_paragraph(lines, heading=False, indent=0, size_scale=1, cell=False, center=False):
    spans = [span for line in lines for span in line]
    rtl = is_rtl(' '.join(span['text'] for span in spans))
    runs = '<w:r><w:br/></w:r>'.join(word_line_runs(line, size_scale) for line in lines)
    return f'''<w:p><w:pPr>{'<w:keepNext/>' if heading else ''}
        {'<w:bidi/>' if rtl else ''}<w:spacing w:before="0" w:after="{0 if cell else 60}"/>
        <w:ind {'w:right' if rtl else 'w:left'}="{indent}"/>
        <w:jc w:val="{'center' if center else 'right' if rtl else 'left'}"/>
        {'<w:outlineLvl w:val="0"/>' if heading else ''}</w:pPr>{runs}</w:p>'''


def table_cell_contains(cell, span):
    rect = fitz.Rect(span['bbox'])
    return (rect.tl + (rect.br - rect.tl) * 0.5) in cell


def pdf_rgb(rgb):
    return ''.join(f'{min(255, max(0, round(value * 255))):02X}' for value in rgb)


def table_style(table, drawings):
    bounds = fitz.Rect(table.bbox)
    header = fitz.Rect(table.rows[0].bbox)
    border_color, fill = 'A0A0A0', None
    largest_fill = 0
    for drawing in drawings:
        rect = drawing.get('rect')
        if rect is None or not bounds.intersects(fitz.Rect(rect) + (-1, -1, 1, 1)):
            continue
        if (drawing.get('color') is not None and abs(rect.y0 - bounds.y0) <= 1
                and rect.width >= bounds.width * .4):
            border_color = pdf_rgb(drawing['color'])
        overlap = (rect & header).get_area()
        if (drawing.get('fill') is not None and overlap > largest_fill
                and rect.height <= header.height + 2 and overlap >= header.get_area() * .4):
            fill = pdf_rgb(drawing['fill'])
            largest_fill = overlap
    return border_color, fill


def word_table(table, raw_rows, lines, available_width, size_scale, margin, drawings):
    col_count = table.col_count
    width = min(available_width, max(col_count * 20, round(fitz.Rect(table.bbox).width * 20 * size_scale)))
    boundaries = sorted({round(x, 2) for cell in table.cells if cell
                         for x in (cell[0], cell[2])})
    if len(boundaries) == col_count + 1 and boundaries[-1] > boundaries[0]:
        widths = [max(20, round(width * (right - left) / (boundaries[-1] - boundaries[0])))
                  for left, right in zip(boundaries, boundaries[1:])]
    else:
        widths = [max(20, width // col_count)] * col_count
    grid = ''.join(f'<w:gridCol w:w="{value}"/>' for value in widths)
    border_color, header_fill = table_style(table, drawings)
    border = ''.join(f'<w:{edge} w:val="single" w:sz="4" w:color="{border_color}"/>'
                     for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    rows = []
    for row_index, row in enumerate(table.rows):
        cells = []
        for col_index, cell in enumerate(row.cells):
            contents = []
            if cell is not None:
                rect = fitz.Rect(cell)
                for line in lines:
                    for group in joined_pdf_spans(
                        [span for span in line['spans'] if table_cell_contains(rect, span)], line['dir']
                    ):
                        contents.append(group)
                value = raw_rows[row_index][col_index]
                if not contents and value:
                    contents = [[{'text': text}] for text in value.splitlines()]
            is_header = row_index == 0 and header_fill is not None
            shading = f'<w:shd w:val="clear" w:fill="{header_fill}"/>' if is_header else ''
            paragraph = word_paragraph(contents, size_scale=size_scale, cell=True, center=is_header)
            cells.append(f'''<w:tc><w:tcPr><w:tcW w:w="{widths[col_index]}" w:type="dxa"/>{shading}
                             <w:vAlign w:val="center"/></w:tcPr>
                             {paragraph}</w:tc>''')
        rows.append(f"<w:tr>{''.join(cells)}</w:tr>")
    indent = max(0, min(available_width - sum(widths), round(table.bbox[0] * 20 * size_scale) - margin))
    return f'''<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>
        <w:tblInd w:w="{indent}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/>
        <w:tblCellMar><w:left w:w="75" w:type="dxa"/><w:right w:w="75" w:type="dxa"/></w:tblCellMar></w:tblPr>
        <w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl>'''


def word_column_split(lines, page_rect):
    # Require substantial text on both sides of a common gutter, with overlapping
    # vertical ranges. Short labels and right-aligned page numbers are not columns.
    substantial = [rect for rect, spans in lines if rect.width >= page_rect.width * .075
                   and sum(len(span['text'].strip()) for span in spans) >= 12]
    best = None
    for fraction in range(39, 63):
        split = page_rect.x0 + page_rect.width * fraction / 100
        padding = page_rect.width * .012
        left = [rect for rect in substantial if rect.x1 <= split - padding]
        right = [rect for rect in substantial if rect.x0 >= split + padding]
        if min(len(left), len(right)) < 3:
            continue
        top = max(min(rect.y0 for rect in left), min(rect.y0 for rect in right))
        bottom = min(max(rect.y1 for rect in left), max(rect.y1 for rect in right))
        left_links = [rect for rect in left if top <= (rect.y0 + rect.y1) / 2 <= bottom]
        right_links = [rect for rect in right if top <= (rect.y0 + rect.y1) / 2 <= bottom]
        if min(len(left_links), len(right_links)) < 2:
            continue
        crossing = [rect for rect in substantial if rect.x0 < split < rect.x1 and top <= rect.y0 <= bottom]
        if len(crossing) >= min(len(left_links), len(right_links)):
            continue
        score = (min(len(left_links), len(right_links)) - len(crossing),
                 len(left) + len(right), -abs(fraction - 50))
        if best is None or score > best[0]:
            best = score, split, min(rect.x0 for rect in right)
    return best[1:] if best else (None, 0)


def word_column_side(rect, split):
    if split is None or rect.x0 < split < rect.x1:
        return None
    return 0 if rect.x1 <= split else 1


def word_reading_order(items, split, rtl):
    ordered = sorted(items, key=lambda item: (round(item[0].y0, 1), item[0].x0))
    if split is None:
        return [xml for _, xml in ordered]
    result = []
    columns = ([], [])
    for rect, xml in ordered:
        side = word_column_side(rect, split)
        if side is not None:
            columns[side].append(xml)
            continue
        for column in reversed(columns) if rtl else columns:
            result.extend(column)
            column.clear()
        result.append(xml)
    for column in reversed(columns) if rtl else columns:
        result.extend(column)
    return result


def write_word_export(package, document, progress, image_pages=()):
    main = 'word/document.xml'
    links, parts, report = [], [], {}
    image_id = 0
    for number, page in enumerate(document, 1):
        progress(number - 1, len(document), f'Page {number}')
        blocks, warnings = extract_office_page(page, number - 1 in image_pages)
        text_blocks = [block for block in blocks if block['type'] == 0]
        lines = [line for block in text_blocks for line in block['lines']]
        if lines:
            try:
                drawings = page.get_drawings()
            except (RuntimeError, ValueError):
                drawings = []
            fractions, consumed = pdf_fraction_runs(lines, drawings)
            if fractions:
                recover_fraction_parentheses(lines, fractions)
                apply_word_fractions(text_blocks, fractions, consumed)
                lines = [line for block in text_blocks for line in block['lines']]
        size_scale = min(1, 28800 / (20 * max(page.rect.width, page.rect.height)))
        page_width = max(1440, round(page.rect.width * 20 * size_scale))
        page_height = max(1440, round(page.rect.height * 20 * size_scale))
        margin = min(540, page_width // 8, page_height // 8)
        orientation = ' w:orient="landscape"' if page_width > page_height else ''
        section = f'''<w:sectPr><w:pgSz w:w="{page_width}" w:h="{page_height}"{orientation}/>
          <w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" w:left="{margin}"
                   w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'''
        available_width = page_width - 2 * margin
        body_size = median([span['size'] for line in lines for span in line['spans']
                            for _ in range(min(len(span['text']), 100))] or [11])
        items = []
        table_regions = []
        if lines:
            try:
                tables = page.find_tables(strategy='lines_strict', use_layout=False).tables
                if not tables:
                    tables = page.find_tables().tables
            except (RuntimeError, ValueError, TypeError, IndexError):
                tables = []
                warnings.add('table')
            for table in tables:
                if table.row_count < 2 or table.col_count < 2:
                    continue
                try:
                    raw_rows = table.extract()
                    if not any(value and value.strip() for row in raw_rows for value in row):
                        continue
                    table_xml = word_table(table, raw_rows, lines, available_width, size_scale, margin, drawings)
                except (RuntimeError, ValueError, TypeError, IndexError):
                    warnings.add('table')
                    continue
                table_rect = fitz.Rect(table.bbox)
                items.append((table_rect, table_xml))
                table_regions.append(table_rect)
        paragraph_blocks = []
        for block in text_blocks:
            content = []
            for line in block['lines']:
                spans = [span for span in line['spans'] if not any(
                    table_cell_contains(region, span) for region in table_regions)]
                for group in joined_pdf_spans(spans, line['dir']):
                    rect = fitz.Rect(group[0]['bbox'])
                    for span in group[1:]:
                        rect |= fitz.Rect(span['bbox'])
                    content.append((rect, group))
            if content:
                paragraph_blocks.append(content)
        content_lines = [line for block in paragraph_blocks for line in block]
        split, right_origin = word_column_split(content_lines, page.rect * page.derotation_matrix)
        rtl = is_rtl(' '.join(span['text'] for _, spans in content_lines for span in spans))

        def word_indent(rect):
            x = rect.x0
            if word_column_side(rect, split) == 1:
                x -= right_origin - margin / (20 * size_scale)
            return min(available_width // 3, max(0, round(x * 20 * size_scale) - margin))

        for block in paragraph_blocks:
            # A PDF text block can contain lines from both columns.
            for _, group in groupby(block, key=lambda line: word_column_side(line[0], split)):
                group_lines = list(group)
                bbox = fitz.Rect(group_lines[0][0])
                for rect, _ in group_lines[1:]:
                    bbox |= rect
                paragraph_lines = []
                baseline = None
                for _, spans in group_lines:
                    origin_y = spans[0]['origin'][1]
                    if baseline is not None and abs(origin_y - baseline) <= max(1, spans[0]['size'] * .28):
                        previous = paragraph_lines[-1][-1]
                        if 'fraction' not in previous and spans[0]['bbox'][0] > previous['bbox'][2] + 1:
                            paragraph_lines[-1][-1] = dict(previous, text=previous['text'] + ' ')
                        paragraph_lines[-1].extend(spans)
                    else:
                        paragraph_lines.append(list(spans))
                        baseline = origin_y
                heading = len(paragraph_lines) <= 2 and max(
                    span['size'] for line in paragraph_lines for span in line) >= body_size * 1.3
                items.append((bbox, word_paragraph(paragraph_lines, heading, word_indent(bbox), size_scale)))

        for block in (item for item in blocks if item['type'] == 1):
            bbox = fitz.Rect(block['bbox'])
            if bbox.is_empty or bbox.is_infinite:
                continue
            indent = word_indent(bbox)
            image_id += 1
            try:
                image = write_office_image(package, 'word', image_id, block)
            except (ValueError, OSError, RuntimeError):
                warnings.add('image')
                continue
            relation = f'rId{image_id}'
            links.append((relation, 'image', image))
            width, height = fitted_size(bbox, min((available_width - indent) * 635, bbox.width * 12700 * size_scale),
                                        min((page_height - 2 * margin - 120) * 635,
                                            bbox.height * 12700 * size_scale))
            items.append((bbox, f'''<w:p><w:pPr><w:spacing w:after="60"/><w:ind w:left="{indent}"/></w:pPr>
                <w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="{width}" cy="{height}"/><wp:docPr id="{image_id}" name="Image {image_id}"/>
                <a:graphic><a:graphicData uri="{DRAWING_NS}/picture">
                {drawing_picture('pic', image_id, relation, width, height)}
                </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'''))
        ordered = word_reading_order(items, split, rtl)
        if 'ocr' in warnings:
            ordered.insert(0, word_paragraph([[{'text': OCR_PLACEHOLDER, 'color': 0x755B49, 'flags': 2}]]))
        if number < len(document):
            ordered.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                           f'{section}</w:pPr></w:p>')
        else:
            ordered.append(section)
        parts.extend(ordered)
        record_office_warnings(report, number, warnings)
    package.writestr(main, f'''<w:document xmlns:w="{WORD_NS}" xmlns:a="{DRAWING_NS}/main"
      xmlns:wp="{DRAWING_NS}/wordprocessingDrawing" xmlns:pic="{DRAWING_NS}/picture" xmlns:r="{OFFICE_NS}"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:body>{''.join(parts)}</w:body></w:document>''')
    write_relationships(package, '_rels/.rels', [('rId1', 'officeDocument', main)])
    write_relationships(package, 'word/_rels/document.xml.rels', links)
    write_content_types(package, [(main, 'wordprocessingml.document.main')])
    return report


def write_presentation_theme(package):
    colors = {
        'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '26313D', 'lt2': 'F8F5EF',
        'accent1': '287D74', 'accent2': '446785', 'accent3': '755B49',
        'accent4': '78857A', 'accent5': '6B596C', 'accent6': '8D684C',
        'hlink': '0563C1', 'folHlink': '954F72',
    }
    scheme = ''.join(f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>' for name, value in colors.items())
    fonts = '<a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/>'
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    lines = ''.join(f'<a:ln w="{width}">{fill}<a:prstDash val="solid"/></a:ln>'
                    for width in (6350, 12700, 19050))
    effects = '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3
    package.writestr('ppt/theme/theme1.xml', f'''<a:theme xmlns:a="{DRAWING_NS}/main" name="Wadaa">
      <a:themeElements><a:clrScheme name="Wadaa">{scheme}</a:clrScheme>
      <a:fontScheme name="Arial"><a:majorFont>{fonts}</a:majorFont><a:minorFont>{fonts}</a:minorFont></a:fontScheme>
      <a:fmtScheme name="Wadaa"><a:fillStyleLst>{fill * 3}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst>
        <a:effectStyleLst>{effects}</a:effectStyleLst><a:bgFillStyleLst>{fill * 3}</a:bgFillStyleLst>
      </a:fmtScheme></a:themeElements></a:theme>''')



def slide_transform(box, rotation=0):
    x, y, width, height = box
    return f'''<a:xfrm rot="{rotation}"><a:off x="{x}" y="{y}"/>
               <a:ext cx="{max(1, width)}" cy="{max(1, height)}"/></a:xfrm>'''


def slide_text_box(identifier, spans, box, scale, rotation=0):
    rtl = is_rtl(''.join(span['text'] for span in spans))
    runs = []
    for span in spans:
        flags = span.get('flags', 0)
        size = max(100, min(400000, round(span.get('size', 11) * scale / 127)))
        color = f"{span.get('color', 0) & 0xFFFFFF:06X}"
        font = office_font(span.get('font', 'Arial'))
        runs.append(f'''<a:r><a:rPr sz="{size}" b="{int(bool(flags & 16))}"
          i="{int(bool(flags & 2))}" baseline="{30000 if flags & 1 else 0}">
          <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
          <a:latin typeface={font}/><a:ea typeface={font}/><a:cs typeface={font}/></a:rPr>
          <a:t>{office_text(span['text'])}</a:t></a:r>''')
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="Text {identifier}"/>
      <p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr>{slide_transform(box, rotation)}
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr">
      <a:noAutofit/></a:bodyPr><a:lstStyle/><a:p>
      <a:pPr algn="{'r' if rtl else 'l'}" rtl="{int(rtl)}"><a:buNone/></a:pPr>
      {''.join(runs)}</a:p></p:txBody></p:sp>'''


def slide_text_geometry(line, position, scale):
    try:
        quad = fitz.recover_line_quad(line)
    except (ValueError, RuntimeError, ZeroDivisionError):
        quad = fitz.Rect(line['bbox']).quad
    ul, ur, ll, lr = map(position, (quad.ul, quad.ur, quad.ll, quad.lr))
    dx, dy = ur[0] - ul[0], ur[1] - ul[1]
    reserve = max(1, round(scale * 1.5))
    width = round(math.hypot(dx, dy)) + reserve
    height = max(1, round(math.dist(ul, ll)))
    angle = math.atan2(dy, dx)
    cx = (ul[0] + lr[0]) / 2 + reserve * math.cos(angle) / 2
    cy = (ul[1] + lr[1]) / 2 + reserve * math.sin(angle) / 2
    return (round(cx - width / 2), round(cy - height / 2), width, height), round(
        math.degrees(angle) * 60000) % 21600000


def slide_path(identifier, drawing, position, scale):
    paths = []
    for item in drawing['items']:
        if item[0] == 'l':
            paths.append((False, [position(item[1]), position(item[2])]))
        elif item[0] == 'c':
            paths.append((False, [position(point) for point in item[1:]]))
        elif item[0] == 're':
            rect = item[1]
            corners = (rect.tl, rect.tr, rect.br, rect.bl)
            if item[2] < 0:
                corners = corners[::-1]
            paths.append((True, [position(point) for point in corners]))
        else:
            return ''
    all_points = [point for _, points in paths for point in points]
    if not all_points:
        return ''
    x = min(point[0] for point in all_points)
    y = min(point[1] for point in all_points)
    width = max(1, max(point[0] for point in all_points) - x)
    height = max(1, max(point[1] for point in all_points) - y)
    commands = []
    previous = None
    for closed, points in paths:
        if previous != points[0] or closed:
            commands.append(f'<a:moveTo><a:pt x="{points[0][0]-x}" y="{points[0][1]-y}"/></a:moveTo>')
        if len(points) == 4 and not closed:
            controls = ''.join(f'<a:pt x="{px-x}" y="{py-y}"/>' for px, py in points[1:])
            commands.append(f'<a:cubicBezTo>{controls}</a:cubicBezTo>')
        else:
            for px, py in points[1:]:
                commands.append(f'<a:lnTo><a:pt x="{px-x}" y="{py-y}"/></a:lnTo>')
        if closed:
            commands.append('<a:close/>')
        previous = None if closed else points[-1]
    if drawing.get('closePath') and previous is not None:
        commands.append('<a:close/>')

    def paint(rgb, opacity):
        if rgb is None:
            return '<a:noFill/>'
        color = pdf_rgb(rgb)
        return (f'<a:solidFill><a:srgbClr val="{color}">'
                f'<a:alpha val="{round(100000 * opacity)}"/></a:srgbClr></a:solidFill>')

    fill = paint(drawing.get('fill'), drawing.get('fill_opacity') or 0)
    stroke = paint(drawing.get('color'), drawing.get('stroke_opacity') or 0)
    line_width = min(20116800, max(1, round((drawing.get('width') or 0.5) * scale)))
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="Drawing {identifier}"/>
      <p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{slide_transform((x, y, width, height))}
      <a:custGeom><a:pathLst><a:path w="{width}" h="{height}">{''.join(commands)}</a:path></a:pathLst>
      </a:custGeom>{fill}<a:ln w="{line_width}">{stroke}</a:ln></p:spPr></p:sp>'''


def slide_objects(package, page, number, slide_width, slide_height, imported_image=False):
    blocks, warnings = extract_office_page(page, imported_image)
    scale = min(slide_width / page.rect.width, slide_height / page.rect.height)
    x = (slide_width - page.rect.width * scale) / 2
    y = (slide_height - page.rect.height * scale) / 2
    matrix = page.rotation_matrix

    def position(point):
        point = fitz.Point(point) * matrix
        return round(x + point.x * scale), round(y + point.y * scale)

    shapes = []
    links = [('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')]
    try:
        drawings = page.get_drawings()
    except (ValueError, RuntimeError):
        drawings = []
        warnings.add('drawing')
    lines = [line for block in blocks if block['type'] == 0 for line in block['lines']]
    if drawings and any(span['text'] == '\ufffd' for line in lines for span in line['spans']):
        fractions, _ = pdf_fraction_runs(lines, drawings)
        recover_fraction_parentheses(lines, fractions)
    for drawing in drawings:
        try:
            shape = slide_path(len(shapes) + 2, drawing, position, scale)
        except (ValueError, TypeError, IndexError):
            shape = ''
        if shape:
            shapes.append(shape)
        else:
            warnings.add('drawing')
    for block in (item for item in blocks if item['type'] == 1):
        rect = fitz.Rect(block['bbox'])
        if rect.is_empty or rect.is_infinite:
            continue
        transform = fitz.Matrix(block['transform']) * matrix
        width = max(1, round(scale * math.hypot(transform.a, transform.b)))
        height = max(1, round(scale * math.hypot(transform.c, transform.d)))
        cx, cy = position((rect.tl + rect.br) * 0.5)
        angle = round(math.degrees(math.atan2(transform.b, transform.a)) * 60000) % 21600000
        identifier = len(shapes) + 2
        try:
            image = write_office_image(package, 'ppt', f'{number}-{identifier}', block)
        except (ValueError, OSError, RuntimeError):
            warnings.add('image')
            continue
        relation = f'rId{len(links) + 1}'
        links.append((relation, 'image', f'../{image}'))
        shapes.append(drawing_picture('p', identifier, relation, width, height,
                                      cx - width // 2, cy - height // 2, rotation=angle))
    for block in (item for item in blocks if item['type'] == 0):
        for line in block['lines']:
            for spans in joined_pdf_spans(line['spans'], line['dir'], separate_parentheses=True):
                box, angle = slide_text_geometry(dict(line, spans=spans), position, scale)
                # Each run group is already positioned on its PDF baseline.
                positioned = [dict(span, flags=span['flags'] & ~1) for span in spans]
                shapes.append(slide_text_box(len(shapes) + 2, positioned, box, scale, angle))
    if 'ocr' in warnings:
        note = {'text': OCR_PLACEHOLDER, 'size': 14, 'flags': 2, 'color': 0x755B49}
        box = (round(x + 18 * scale), round(y + 18 * scale),
               round((page.rect.width - 36) * scale), round(24 * scale))
        shapes.append(slide_text_box(len(shapes) + 2, [note], box, scale))
    return ''.join(shapes), links, warnings


def write_powerpoint_export(package, document, progress, image_pages=()):
    main = 'ppt/presentation.xml'
    parts = [(main, 'presentationml.presentation.main'),
             ('ppt/slideMasters/slideMaster1.xml', 'presentationml.slideMaster'),
             ('ppt/slideLayouts/slideLayout1.xml', 'presentationml.slideLayout'),
             ('ppt/theme/theme1.xml', 'theme')]
    namespaces = f'xmlns:p="{PRESENTATION_NS}" xmlns:a="{DRAWING_NS}/main" xmlns:r="{OFFICE_NS}"'
    shape_tree = '''<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'''
    color_map = ('accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
                 'accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" '
                 'folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"')
    package.writestr('ppt/slideMasters/slideMaster1.xml', f'''<p:sldMaster {namespaces}>
      <p:cSld name="Blank"><p:spTree>{shape_tree}</p:spTree></p:cSld><p:clrMap {color_map}/>
      <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
      </p:sldMaster>''')
    package.writestr('ppt/slideLayouts/slideLayout1.xml', f'''<p:sldLayout {namespaces} type="blank" preserve="1">
      <p:cSld name="Blank"><p:spTree>{shape_tree}</p:spTree></p:cSld>
      <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>''')
    write_presentation_theme(package)
    write_relationships(package, 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
                        [('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'),
                         ('rId2', 'theme', '../theme/theme1.xml')])
    write_relationships(package, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
                        [('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')])
    slide_width, slide_height = fitted_size(document[0].rect, 9144000, 9144000)
    slide_width, slide_height = max(914400, slide_width), max(914400, slide_height)
    links = [('rId1', 'slideMaster', 'slideMasters/slideMaster1.xml')]
    slide_ids, report = [], {}
    for number, page in enumerate(document, 1):
        progress(number - 1, len(document), f'Page {number}')
        slide = f'slide{number}.xml'
        part = f'ppt/slides/{slide}'
        shapes, slide_links, warnings = slide_objects(
            package, page, number, slide_width, slide_height, number - 1 in image_pages)
        package.writestr(part, f'''<p:sld {namespaces}><p:cSld>
          <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
          <p:spTree>{shape_tree}{shapes}</p:spTree></p:cSld>
          <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>''')
        write_relationships(package, f'ppt/slides/_rels/{slide}.rels', slide_links)
        record_office_warnings(report, number, warnings)
        relation = f'rId{number + 1}'
        links.append((relation, 'slide', f'slides/{slide}'))
        slide_ids.append(f'<p:sldId id="{255 + number}" r:id="{relation}"/>')
        parts.append((part, 'presentationml.slide'))
    package.writestr(main, f'''<p:presentation {namespaces}>
      <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
      <p:sldIdLst>{''.join(slide_ids)}</p:sldIdLst>
      <p:sldSz cx="{slide_width}" cy="{slide_height}"/>
      <p:notesSz cx="6858000" cy="9144000"/></p:presentation>''')
    write_relationships(package, '_rels/.rels', [('rId1', 'officeDocument', main)])
    write_relationships(package, 'ppt/_rels/presentation.xml.rels', links)
    write_content_types(package, parts)
    return report


def export_merged_pages(jobs, destination, kind, progress):
    with tempfile.TemporaryDirectory(prefix='.wadaa-export-', dir=destination.parent) as temp:
        workdir = Path(temp)
        merged = workdir / 'merged.pdf'
        image_pages = set() if kind in ('docx', 'pptx') else None
        merge_files(jobs, merged, lambda done, total, name: progress(done * 50 // total, 100, name),
                    image_pages=image_pages)
        def render_progress(done, total, name):
            progress(50 + done * 50 // total, 100, name)
        staged = workdir / ('images' if kind == 'jpg' else f'output.{kind}')
        report = {}
        with fitz.open(str(merged)) as document:
            total = len(document)
            if not total:
                raise ValueError('The merged PDF has no pages.')
            if kind == 'jpg':
                staged.mkdir()
                digits = max(3, len(str(total)))
                for number, page in enumerate(document, 1):
                    render_progress(number - 1, total, f'Page {number}')
                    (staged / f'page-{number:0{digits}d}.jpg').write_bytes(page_jpeg(page))
            else:
                writer = write_word_export if kind == 'docx' else write_powerpoint_export
                with ZipFile(staged, 'w', compression=ZIP_DEFLATED) as package:
                    report = writer(package, document, render_progress, image_pages)
        if kind == 'jpg':
            if destination.exists():
                raise FileExistsError(f'The output folder already exists: {destination}')
            staged.rename(destination)
        else:
            staged.replace(destination)
        return total, report


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
        self.create_rectangle(0, 0, width, height, fill=self._root().color(SURFACE), outline='')
        if self.last - self.first >= 0.99999:
            return
        top, bottom, _, _ = self._bounds()
        color = self._root().color(ACCENT if self._grab_offset is not None else MUTED)
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
        'normal': (SURFACE, HOVER, SELECT, BORDER, TEXT),
        'navigation': (SURFACE, HOVER, SELECT, BORDER, '#287D74'),
        'primary': (ACCENT, ACCENT_HOVER, ACCENT_PRESSED, ACCENT, ACCENT_TEXT),
        'donate': (SURFACE, HOVER, SELECT, BORDER, ACCENT),
    }

    def __init__(self, parent, text, command, kind='normal', **kwargs):
        self.font_obj = tkfont.Font(family='Segoe UI', size=10,
                                    weight='bold' if kind in ('primary', 'donate', 'navigation') else 'normal')
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
            self.configure(width=max(104, self.font_obj.measure(self._root().tr(self._label)) + 30))
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
            fill, edge, text_color = SURFACE, BORDER, MUTED
        elif self._pressed:
            fill = pressed
        elif self._hover:
            fill = hover
        # Premultiply the rounded shape onto the exact parent background.
        # Oversampling gives genuinely round anti-aliased corner pixels.
        k = 3
        color = self._root().color
        art = Image.new('RGB', (width*k, height*k), color(BG))
        draw = ImageDraw.Draw(art)
        draw.rounded_rectangle((k, k, (width-2)*k, (height-2)*k),
                               radius=10*k, fill=color(fill), outline=color(edge), width=k)
        self._photo = ImageTk.PhotoImage(art.resize((width,height), Image.Resampling.LANCZOS), master=self)
        self.delete('all')
        self.create_image(0, 0, image=self._photo, anchor='nw')
        self.create_text(width//2, height//2, text=self._root().tr(self._label),
                         font=self.font_obj, fill=color(text_color), anchor='center')

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


class DarkMenuPopup:
    """Draw the existing menu entries without the Windows popup's bright frame."""

    def __init__(self, app):
        self.app = app
        self.panels = []
        self.key = None
        self.previous_focus = None

    def close(self):
        if not self.panels:
            return
        first, _, _, _ = self.panels[0]
        if first.grab_current() == first:
            first.grab_release()
        self.remove_from(0)
        self.key = None
        if self.previous_focus is not None and self.previous_focus.winfo_exists():
            self.previous_focus.focus_set()
        self.previous_focus = None

    def remove_from(self, level):
        for popup, _, _, _ in reversed(self.panels[level:]):
            popup.destroy()
        del self.panels[level:]

    def post(self, key, *, toggle=False):
        if self.key == key:
            if toggle:
                self.close()
            return 'break'
        self.close()
        self.key = key
        self.previous_focus = self.app.focus_get()
        button = self.app.menu_buttons[key]
        menu = button.nametowidget(button.cget('menu'))
        popup = self.show(menu, button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        popup.grab_set()
        popup.focus_set()
        return 'break'

    def show(self, menu, x, y, *, anchor=None):
        level = len(self.panels)
        popup = tk.Toplevel(self.panels[0][0] if level else self.app.menu_bar, bg=SURFACE, bd=0,
                            highlightbackground=BORDER, highlightcolor=BORDER, highlightthickness=1)
        popup.withdraw()
        popup.overrideredirect(True)
        rows = {}
        active = None
        self.panels.append([popup, menu, rows, active])
        for index in range(menu.index('end') + 1):
            kind = menu.type(index)
            if kind == 'separator':
                tk.Frame(popup, bg=BORDER, height=1).pack(fill='x', pady=3)
                continue
            checked = (kind == 'radiobutton'
                       and str(menu.getvar(menu.entrycget(index, 'variable'))) ==
                       str(menu.entrycget(index, 'value')))
            row = tk.Frame(popup, bg=SURFACE, padx=7)
            row.pack(fill='x')
            mark = tk.Label(row, text='●' if checked else '', width=2, bg=SURFACE, fg=ACCENT,
                            font=('Segoe UI', 8))
            mark.pack(side='left')
            label = tk.Label(row, text=menu.entrycget(index, 'label'), bg=SURFACE, fg=TEXT,
                             font=('Segoe UI', 10), pady=2)
            label.pack(side='left')
            arrow = tk.Label(row, text='›' if kind == 'cascade' else '', width=2, bg=SURFACE,
                             fg=MUTED, font=('Segoe UI', 10))
            arrow.pack(side='right')
            rows[index] = row
            for widget in (row, mark, label, arrow):
                widget.bind('<Enter>', lambda event, i=index: self.highlight(level, i, expand=True))
                widget.bind('<ButtonRelease-1>', lambda event, i=index: self.invoke(level, i))
        popup.bind('<ButtonPress-1>', self.outside_click)
        popup.bind('<KeyPress>', self.keypress)
        popup.bind('<FocusOut>', lambda event: self.app.after_idle(self.dismiss_if_inactive))
        popup.update_idletasks()
        width, height = popup.winfo_reqwidth(), popup.winfo_reqheight()
        right, bottom = popup.winfo_screenwidth(), popup.winfo_screenheight()
        if anchor is not None and x + width > right:
            x = anchor.winfo_rootx() - width
        popup.geometry(f'+{max(0, min(x, right - width))}+{max(0, min(y, bottom - height))}')
        popup.deiconify()
        popup.lift()
        return popup

    def highlight(self, level, index, *, expand=False):
        if level >= len(self.panels):
            return
        popup, menu, rows, active = self.panels[level]
        if active != index:
            self.remove_from(level + 1)
            for i in (active, index):
                if isinstance(i, int):
                    row = rows[i]
                    color = HOVER if i == index else SURFACE
                    for widget in (row, *row.winfo_children()):
                        widget.configure(bg=color)
            self.panels[level][3] = index
        if expand and menu.type(index) == 'cascade' and level + 1 == len(self.panels):
            submenu = menu.nametowidget(menu.entrycget(index, 'menu'))
            row = rows[index]
            self.show(submenu, popup.winfo_rootx() + popup.winfo_width(), row.winfo_rooty(), anchor=row)

    def invoke(self, level, index):
        menu = self.panels[level][1]
        if menu.type(index) == 'cascade':
            self.highlight(level, index, expand=True)
        else:
            self.close()
            menu.invoke(index)
        return 'break'

    def outside_click(self, event):
        widget = self.app.winfo_containing(event.x_root, event.y_root)
        for key, button in self.app.menu_buttons.items():
            if widget == button:
                return self.post(key, toggle=True)
        window = widget.winfo_toplevel() if widget is not None else None
        if not any(window == popup for popup, _, _, _ in self.panels):
            self.close()
            return 'break'

    def dismiss_if_inactive(self):
        if self.panels and self.app.focus_displayof() is None:
            self.close()

    def keypress(self, event):
        if not self.panels:
            return
        level = len(self.panels) - 1
        _, _, rows, active = self.panels[level]
        indices = list(rows)
        key = event.keysym
        if key in ('Up', 'Down'):
            if active is None:
                index = indices[0 if key == 'Down' else -1]
            else:
                step = 1 if key == 'Down' else -1
                index = indices[(indices.index(active) + step) % len(indices)]
            self.highlight(level, index)
        elif key in ('Return', 'space') and active is not None:
            self.invoke(level, active)
        elif key == 'Right' and active is not None and self.panels[level][1].type(active) == 'cascade':
            self.highlight(level, active, expand=True)
            self.highlight(level + 1, next(iter(self.panels[level + 1][2])))
        elif key == 'Escape':
            self.close()
        elif key == 'Left' and level:
            self.remove_from(level)
        elif key in ('Left', 'Right'):
            keys = list(self.app.menu_buttons)
            self.post(keys[(keys.index(self.key) + (1 if key == 'Right' else -1)) % len(keys)])
        else:
            return
        return 'break'


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title('Wadaa PDF Manager')
        self.version = resource_path('VERSION.txt').read_text(encoding='ascii').strip()
        self.about_window = None
        self.options_window = None
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
        self.preferences = self.load_preferences()
        self.theme = self.preferences.get('theme', 'Dark')
        self.language = self.preferences.get('language', 'English')
        self.last_session_enabled = self.preferences.get('enabled') is not False
        self.theme_choice = tk.StringVar(value=self.theme)
        self.language_choice = tk.StringVar(value=self.language)
        self.session_choice = tk.BooleanVar(value=self.last_session_enabled)
        self.total_text = tk.StringVar(value='0 files  •  0 pages  •  0 B')
        self.open_after_merge = tk.BooleanVar(value=False)
        self.status.set(self.tr(self.status.get()))
        self.menu_bar = None
        self.menu_buttons = {}
        self.dark_menu = DarkMenuPopup(self)
        for key in ('f', 'e', 'w'):
            self.bind(f'<Alt-{key}>', lambda event, key=key: self.post_menu(key))
        self.bind('<F10>', lambda event: self.menu_buttons['f'].focus_set())
        self.bind_all('<Map>', self.on_window_map, add='+')
        self.build_ui()
        self.update_appearance()
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
                self.iconbitmap(default=str(ico))
            elif png.exists():
                self._icon_image = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._icon_image)
        except (tk.TclError, OSError):
            pass

    def color(self, color):
        return LIGHT_COLORS.get(color, color) if self.theme == 'Light' else color

    def tr(self, label, **fields):
        if self.language == 'Arabic':
            label = UI_AR.get(label, label)
        text = label.format(**fields) if fields else label
        return rtl_display(text) if self.language == 'Arabic' else text

    def apply_caption(self, window):
        set_windows_caption(window, self.theme == 'Dark', self.color(BG),
                            self.color(TEXT), self.color(BORDER))

    def on_window_map(self, event):
        if isinstance(event.widget, (tk.Tk, tk.Toplevel)):
            self.apply_caption(event.widget)

    def themed_dialog(self, title, message, *, confirm=False, parent=None):
        parent = parent or self
        dialog = tk.Toplevel(parent, bg=BG)
        dialog.withdraw()
        dialog.title(self.tr(title))
        dialog.transient(parent)
        dialog.resizable(False, False)
        frame = tk.Frame(dialog, bg=BG, padx=22, pady=18)
        frame.pack(fill='both', expand=True)
        tk.Label(frame, text=self.tr(message), bg=BG, fg=TEXT, wraplength=480,
                 justify='right' if self.language == 'Arabic' else 'left',
                 font=('Segoe UI', 10)).pack(fill='x')
        actions = tk.Frame(frame, bg=BG)
        actions.pack(fill='x', pady=(16, 0))
        result = False

        def close(choice=False):
            nonlocal result
            result = choice
            dialog.destroy()

        if confirm:
            self.button(actions, 'Yes', lambda: close(True), primary=True).pack(side='right')
            default = self.button(actions, 'No', close)
            default.pack(side='right', padx=(0, 8))
        else:
            default = self.button(actions, 'OK', close, primary=True)
            default.pack(side='right')
        dialog.protocol('WM_DELETE_WINDOW', close)
        dialog.bind('<Escape>', lambda event: close())
        self.style_widgets(dialog)
        dialog.update_idletasks()
        x = max(0, parent.winfo_rootx() + (parent.winfo_width() - dialog.winfo_reqwidth()) // 2)
        y = max(0, parent.winfo_rooty() + (parent.winfo_height() - dialog.winfo_reqheight()) // 2)
        dialog.geometry(f'+{x}+{y}')
        dialog.deiconify()
        dialog.wait_visibility()
        dialog.grab_set()
        default.focus_set()
        parent.wait_window(dialog)
        return result

    def info(self, title, message, *, parent=None):
        self.themed_dialog(title, message, parent=parent)

    def configure_styles(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        c = self.color
        style.configure('.', background=c(BG), foreground=c(TEXT), bordercolor=c(BORDER))
        style.configure('Dark.TFrame', background=c(BG))
        style.configure('Surface.TFrame', background=c(SURFACE))
        style.configure('Dark.TLabel', background=c(BG), foreground=c(TEXT), font=('Segoe UI', 10))
        style.configure('Muted.TLabel', background=c(BG), foreground=c(MUTED), font=('Segoe UI', 9))
        style.configure('Title.TLabel', background=c(BG), foreground=c(TEXT), font=('Segoe UI Semibold', 23))
        normal, hover, pressed, edge, fg = map(c, RoundedButton.PALETTES['normal'])
        style.configure('Dark.TButton', background=normal, foreground=fg,
                        bordercolor=edge, lightcolor=normal, darkcolor=normal,
                        relief='flat', borderwidth=1, padding=(10, 0), font=('Segoe UI', 10),
                        focuscolor=normal)
        style.map('Dark.TButton', background=[('pressed', pressed), ('active', hover),
                                             ('disabled', c(BG))],
                  foreground=[('disabled', c(MUTED))])
        style.configure('Dark.Treeview', background=c(SURFACE), fieldbackground=c(SURFACE),
                        foreground=c(TEXT), rowheight=31, font=('Segoe UI', 10),
                        borderwidth=0, relief='flat')
        style.configure('Dark.Treeview.Heading', background=c(BORDER), foreground=c(TEXT),
                        borderwidth=0, font=('Segoe UI Semibold', 10), padding=(8, 8))
        style.map('Dark.Treeview', background=[('selected', c(SELECT))],
                  foreground=[('selected', c(TEXT))])
        style.map('Dark.Treeview.Heading', background=[('active', c(HOVER))])
        style.configure('Dark.Horizontal.TProgressbar', background=c(ACCENT), troughcolor=c(SURFACE),
                        lightcolor=c(ACCENT), darkcolor=c(ACCENT), borderwidth=0, thickness=10)
        style.configure('Preview.Vertical.TScrollbar', background=c(BORDER), troughcolor=c(SURFACE),
                        bordercolor=c(BORDER), lightcolor=c(BORDER), darkcolor=c(BORDER),
                        arrowcolor=c(TEXT))
        style.map('Preview.Vertical.TScrollbar', background=[('active', c(HOVER))])

    def style_widgets(self, widget):
        # Cache each classic widget's original color roles and untranslated text.
        if not hasattr(widget, '_appearance_options'):
            options = widget.keys()
            color_keys = ('background', 'foreground', 'activebackground', 'activeforeground',
                          'selectbackground', 'selectforeground', 'selectcolor', 'insertbackground',
                          'highlightbackground', 'highlightcolor', 'disabledforeground')
            widget._appearance_options = {
                key: str(widget.cget(key)) for key in color_keys
                if key in options and str(widget.cget(key)) in LIGHT_COLORS
            }
            if 'text' in options and str(widget.cget('text')) in UI_AR:
                widget._appearance_options['text'] = str(widget.cget('text'))
        if widget._appearance_options:
            widget.configure(**{
                key: self.tr(value) if key == 'text' else self.color(value)
                for key, value in widget._appearance_options.items()
            })
        if isinstance(widget, RoundedButton):
            widget.configure(text=widget._label)
        elif isinstance(widget, DarkScrollbar):
            widget.redraw()
        for child in widget.winfo_children():
            if child is not self.menu_bar and not isinstance(child, tk.Menu):
                self.style_widgets(child)
        if hasattr(widget, '_refresh_appearance'):
            widget._refresh_appearance()
        if isinstance(widget, (tk.Tk, tk.Toplevel)) and widget.winfo_ismapped():
            self.apply_caption(widget)

    def update_appearance(self):
        self.configure_styles()
        self.option_add('*Toplevel.background', self.color(BG))
        self.build_menu_bar()
        self.style_widgets(self)
        for col, label in (('name', 'File name (current merge order)'), ('type', 'Type'), ('pages', 'Pages')):
            self.tree.heading(col, text=self.tr(label))
        self._update_totals()
        self.update_idletasks()
        self.minsize(max(800, self._min_content_width()), max(565, self.winfo_reqheight() + 5))

    def build_menu_bar(self):
        self.dark_menu.close()
        if self.menu_bar is not None:
            self.menu_bar.destroy()
        c = self.color
        self.menu_bar = tk.Frame(self, bg=c(SURFACE), highlightthickness=1,
                                 highlightbackground=c(BORDER))
        self.menu_bar.pack(fill='x', before=self.wrapper)
        menu_style = dict(tearoff=False, bg=c(SURFACE), fg=c(TEXT), activebackground=c(SELECT),
                          activeforeground=c(TEXT), selectcolor=c(ACCENT),
                          disabledforeground=c(MUTED), font=('Segoe UI', 10))

        def top(label, key):
            button = tk.Menubutton(self.menu_bar, text=self.tr(label), bg=c(SURFACE), fg=c(TEXT),
                                   activebackground=c(SELECT), activeforeground=c(TEXT),
                                   font=('Segoe UI', 10), padx=11, pady=3, bd=0, takefocus=True,
                                   highlightthickness=1, highlightbackground=c(SURFACE),
                                   highlightcolor=c(ACCENT), underline=0 if self.language == 'English' else -1)
            menu = tk.Menu(button, **menu_style)
            button.configure(menu=menu)
            button.pack(side='left')
            self.menu_buttons[key] = button
            if self.theme == 'Dark':
                button.bind('<ButtonPress-1>', lambda event: self.dark_menu.post(key, toggle=True))
                button.bind('<ButtonRelease-1>', lambda event: 'break')
                for shortcut in ('<Return>', '<space>', '<Down>'):
                    button.bind(shortcut, lambda event: self.dark_menu.post(key))
            return menu

        def command(menu, label, action):
            menu.add_command(label=self.tr(label), command=action)

        def cascade(menu, label):
            submenu = tk.Menu(menu, **menu_style)
            menu.add_cascade(label=self.tr(label), menu=submenu)
            return submenu

        file_menu = top('File', 'f')
        command(file_menu, 'Add Files', self.add_files)
        command(file_menu, 'Open Folder', self.open_folder)
        file_menu.add_separator()
        for label, kind in (('Export File to PDF', 'pdf'), ('Export File to Word', 'docx'),
                            ('Export File to PowerPoint', 'pptx'), ('Export File to Images', 'jpg')):
            command(file_menu, label, lambda kind=kind: self.start_merge(kind))
        command(file_menu, 'Open Output Folder', self.open_output_folder)
        file_menu.add_separator()
        command(file_menu, 'Exit', self.close_app)

        edits = top('Edits', 'e')
        for label, variable, choices, update in self.preference_controls():
            submenu = cascade(edits, label)
            for caption, value in choices:
                submenu.add_radiobutton(label=self.tr(caption), variable=variable,
                                        value=value, command=update)
        edits.add_separator()
        command(edits, 'Options...', self.show_options)

        wadaa = top('WadaaDev', 'w')
        command(wadaa, 'Check for Updates', lambda: self.updater.check(manual=True))
        command(wadaa, 'Donate and Support', lambda: webbrowser.open(DONATE))
        command(wadaa, 'About', self.show_about)
        contact = cascade(wadaa, 'Contact and Report')
        command(contact, 'Contact Support',
                lambda: self.open_support_link(self, support_mailto(self.version), mail=True))
        command(contact, 'Report a Bug', lambda: self.open_support_link(self, BUG_REPORT_URL))
        command(contact, 'GitHub Repository', lambda: self.open_support_link(self, REPOSITORY_URL))

    def post_menu(self, key):
        if self.theme == 'Dark':
            return self.dark_menu.post(key)
        button = self.menu_buttons[key]
        self.tk.call('tk::MbPost', str(button))
        return 'break'

    def preference_controls(self):
        return (
            ('Theme', self.theme_choice, (('Dark', 'Dark'), ('Light', 'Light')),
             lambda: self.set_preference('theme', self.theme_choice.get())),
            ('Language', self.language_choice, (('English', 'English'), ('Arabic', 'Arabic')),
             lambda: self.set_preference('language', self.language_choice.get())),
            ('Save Last Session', self.session_choice, (('ON', True), ('OFF', False)),
             lambda: self.set_preference('enabled', self.session_choice.get())),
        )

    def show_options(self):
        if self.options_window is not None and self.options_window.winfo_exists():
            self.options_window.lift()
            self.options_window.focus_set()
            return
        dialog = tk.Toplevel(self, bg=BG)
        self.options_window = dialog
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog._refresh_appearance = lambda: dialog.title(self.tr('Options'))
        content = ttk.Frame(dialog, style='Dark.TFrame', padding=(22, 18))
        content.pack(fill='both', expand=True)
        for row, (label, variable, choices, update) in enumerate(self.preference_controls()):
            ttk.Label(content, text=label, style='Dark.TLabel').grid(row=row, column=0,
                                                                  sticky='w', padx=(0, 20), pady=6)
            for column, (caption, value) in enumerate(choices, 1):
                tk.Radiobutton(content, text=caption, variable=variable, value=value, command=update,
                               bg=BG, fg=TEXT, activebackground=HOVER, activeforeground=TEXT,
                               selectcolor=SURFACE, highlightbackground=BG, highlightcolor=ACCENT,
                               font=('Segoe UI', 10)).grid(row=row, column=column, sticky='w', pady=6)
        self.button(content, 'Close', dialog.destroy).grid(row=3, column=0, columnspan=3,
                                                         sticky='e', pady=(12, 0))
        dialog.bind('<Escape>', lambda event: dialog.destroy())
        self.style_widgets(dialog)
        dialog.focus_set()

    def button(self, parent, label, command, primary=False, donate=False, navigation=False, **kwargs):
        kind = 'navigation' if navigation else 'donate' if donate else 'primary' if primary else 'normal'
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

        location_label = ttk.Label(self.wrapper, style='Muted.TLabel')
        location_label._refresh_appearance = lambda: location_label.configure(text=self.tr(self.location.get()))
        self.location.trace_add('write', lambda *_: location_label._refresh_appearance())
        location_label.pack(anchor='w', pady=(0, 5))
        ttk.Label(self.wrapper, text='Drop files into the list  •  Drag a row to move it  •  Ctrl/Shift to select multiple',
                  style='Muted.TLabel').pack(anchor='w', pady=(0, 8))

        searchbar = ttk.Frame(self.wrapper, style='Dark.TFrame')
        searchbar.pack(fill='x', pady=(0, 6))
        ttk.Label(searchbar, text='Search files:', style='Muted.TLabel').pack(side='left', padx=(0, 8))
        self.search_entry = tk.Entry(searchbar, textvariable=self.search_text, bg=SURFACE, fg=TEXT,
            insertbackground=TEXT, highlightthickness=1, highlightcolor=HOVER,
            highlightbackground=BORDER, relief='flat', font=('Segoe UI', 10))
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
                              highlightbackground=BORDER, relief='flat', font=('Consolas', 12))
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
        dialog._refresh_appearance = lambda: dialog.title(self.tr('About & Support'))
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
        version_label = ttk.Label(content, style='Muted.TLabel')
        version_label._refresh_appearance = lambda: version_label.configure(text=self.tr(
            'Version: {version}', version=f'v{self.version.lstrip("vV")}'))
        version_label.pack(anchor='w', pady=(1, 12))
        ttk.Label(content, text='Developed by WadaaDev', style='Dark.TLabel').pack(anchor='w')
        ttk.Label(content, text='Powered by WadaaDev', style='Muted.TLabel').pack(anchor='w')
        email_label = ttk.Label(content, style='Muted.TLabel')
        email_label._refresh_appearance = lambda: email_label.configure(text=self.tr('Email: {email}',
                                                                                     email=SUPPORT_EMAIL))
        email_label.pack(anchor='w', pady=(2, 14))

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

        self.style_widgets(dialog)
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
            messagebox.showerror(self.tr('Unable to open link'), hint, parent=parent)

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
        self.status.set(self.tr('{count} new file(s) added; {duplicates} duplicate(s) skipped. {total} total.',
                                count=len(unique), duplicates=max(duplicates, 0), total=len(self.paths)))

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
        label = self.tr('{files} files  •  {pages} pages', files=len(self.paths), pages=known_pages)
        if unknown:
            label += self.tr(' (+ {count} uncounted files)', count=unknown)
        label += f'  •  {size:.1f} {unit}' if unit != 'B' else f'  •  {total_bytes} B'
        self.total_text.set(label)

    def _snapshot(self):
        return (self.paths.copy(), self.location.get(), self.page_choices.copy())

    def _session_path(self):
        home = Path(os.environ.get('APPDATA') or (Path.home() / '.config'))
        return home / 'WadaaPDFManager' / 'session.json'

    def _session_preference_path(self):
        return self._session_path().with_name('session-preference.json')

    def load_preferences(self):
        try:
            data = json.loads(self._session_preference_path().read_text(encoding='utf-8'))
            if isinstance(data, dict):
                if data.get('theme') not in ('Dark', 'Light'):
                    data['theme'] = 'Dark'
                if data.get('language') not in ('English', 'Arabic'):
                    data['language'] = 'English'
                return data
        except (OSError, ValueError):
            pass
        return {}

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
        self.set_preference('enabled', not self.last_session_enabled)

    def set_preference(self, key, value):
        attribute, variable = {
            'theme': ('theme', self.theme_choice),
            'language': ('language', self.language_choice),
            'enabled': ('last_session_enabled', self.session_choice),
        }[key]
        previous = getattr(self, attribute)
        if value == previous:
            return
        preferences = {**self.preferences, key: value}
        try:
            self.write_session_json(self._session_preference_path(), preferences)
        except OSError as exc:
            variable.set(previous)
            self.status.set(self.tr('Unable to save settings: {error}', error=exc))
            messagebox.showerror(self.tr('Options'), str(exc), parent=self.options_window
                                 if self.options_window and self.options_window.winfo_exists() else self)
            return
        self.preferences = preferences
        setattr(self, attribute, value)
        variable.set(value)
        if key in ('theme', 'language'):
            self.update_appearance()
            if key == 'language':
                current = self.status.get()
                if value == 'English':
                    current = next((en for en, ar in UI_AR.items() if rtl_display(ar) == current), current)
                self.status.set(self.tr(current))
        else:
            self.last_session_button.configure(text=self.session_button_label())
            if value and not self.paths and self._session_path().is_file():
                if messagebox.askyesno(self.tr('Last Session'), self.tr('Restore the previously saved list now?'),
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
            self.status.set(self.tr('Unable to save session: {error}', error=exc))
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
            self.status.set(self.tr('Restored {count} file(s) from the previous session.', count=len(paths)))
        except (ValueError, OSError, TypeError, KeyError) as exc:
            self.status.set(self.tr('Previous session could not be restored: {error}', error=exc))

    def close_app(self):
        if self.busy:
            messagebox.showinfo(self.tr('Export'), self.tr('Wait for the export to finish before closing.'),
                                parent=self)
            return
        if not self.save_session():
            messagebox.showwarning(self.tr('Save Last Session'), self.status.get(), parent=self)
            return
        self.updater.close()
        self.updater.apply_pending()
        self.destroy()

    def choose_pdf_pages(self):
        if self.busy:
            return
        selection = self.selected_indices()
        if len(selection) != 1:
            self.info('Select PDF Pages', 'Select one PDF file in the list first.')
            return
        path = self.paths[selection[0]]
        if path.suffix.lower() != '.pdf':
            self.info('Select PDF Pages', 'Page selection is available for PDFs only.')
            return
        try:
            reader = PdfReader(str(path), strict=False)
            if reader.is_encrypted and reader.decrypt('') == 0:
                raise ValueError('This PDF is password protected.')
            total = len(reader.pages)
        except Exception as exc:
            messagebox.showerror(self.tr('Select PDF Pages'),
                                 self.tr('Cannot read {name}: {error}', name=path.name, error=exc), parent=self)
            return
        dialog = tk.Toplevel(self, bg=BG)
        dialog.title(f'Select pages — {path.name}')
        dialog.geometry('530x220')
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        frame = tk.Frame(dialog, bg=BG, padx=18, pady=16)
        frame.pack(fill='both', expand=True)
        heading = tk.Label(frame, bg=BG, fg=TEXT, font=('Segoe UI', 10))
        heading._refresh_appearance = lambda: heading.configure(text=self.tr(
            '{name}  •  {total} pages', name=path.name, total=total))
        heading.pack(anchor='w')
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
                    messagebox.showerror(self.tr('Invalid pages'), str(exc), parent=dialog)
                    return
            self.apply_page_choice(path, expr)
            dialog.destroy()
        self.button(buttons, 'Apply', apply, primary=True).pack(side='right')
        self.button(buttons, 'Cancel', dialog.destroy).pack(side='right', padx=(0, 8))
        dialog.bind('<Return>', lambda e: apply())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        self.style_widgets(dialog)

    def apply_page_choice(self, path, expr):
        if self.busy or path not in self.paths or expr == self.page_choices.get(path, ''):
            return
        self._remember()
        if expr:
            self.page_choices[path] = expr
        else:
            self.page_choices.pop(path, None)
        self.refresh()
        self.status.set(self.tr('Page selection saved for {name}: {pages}',
                                name=path.name, pages=expr or self.tr('all pages')))

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
        self.status.set(self.tr('Undo completed.'))

    def redo(self):
        if self.busy or not self.redo_stack:
            return
        self.undo_stack.append(self._snapshot())
        self.paths, where, choices = self.redo_stack.pop()
        self.page_choices = choices
        self.location.set(where)
        self.refresh()
        self.status.set(self.tr('Redo completed.'))

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
            self.info('Preview Pages', 'Select a PDF or image in the file list first.')
            return
        path = self.paths[selected[0]]
        if path.suffix.lower() in OFFICE_EXTS:
            self.info('Preview Pages', 'Office page preview is not available before conversion. Select a PDF or image.')
            return
        document = None
        is_pdf = path.suffix.lower() == '.pdf'
        try:
            if is_pdf:
                document = fitz.open(str(path))
                if document.needs_pass:
                    raise ValueError(self.tr('This PDF is password protected.'))
                length = len(document)
            else:
                with Image.open(path) as img:
                    length = getattr(img, 'n_frames', 1)
            if length < 1:
                raise ValueError(self.tr('File has no pages.'))
        except Exception as exc:
            if document is not None:
                document.close()
            messagebox.showerror(self.tr('Preview error'), str(exc), parent=self)
            return

        pop = tk.Toplevel(self, bg=BG)
        pop.transient(self)
        pop.geometry('900x700' if is_pdf else '680x670')
        pop.minsize(650 if is_pdf else 440, 390)
        current = 0
        zoom = 1.0
        offset = (0.0, 0.0)
        page_size = None
        chosen = set()
        anchor = None
        existing = self.page_choices.get(path, '')
        if is_pdf and existing:
            try:
                chosen.update(selected_page_indices(existing, length))
            except ValueError:
                # Keep preview available for PDFs that changed on disk.
                pass
        outer = tk.Frame(pop, bg=BG)
        outer.pack(fill='both', expand=True, padx=14, pady=12)
        if is_pdf:
            tk.Label(outer, text='Click a page to include or remove it.', bg=BG, fg=MUTED,
                     font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 7))
        body = tk.Frame(outer, bg=BG)
        body.pack(fill='both', expand=True)
        display = tk.Frame(body, bg=SURFACE, highlightthickness=2, highlightbackground=BORDER)
        display.pack(side='left', fill='both', expand=True)
        label = tk.Label(display, bg=SURFACE, fg=TEXT, text='Rendering...', relief='flat')
        label.place(relwidth=1, relheight=1)
        mark = tk.Label(display, bg=ACCENT, fg=ACCENT_TEXT, text='✓', font=('Segoe UI', 13, 'bold'),
                        padx=6)
        footer = tk.Frame(outer, bg=BG)
        footer.pack(fill='x', pady=(9, 0))
        indicator = tk.StringVar(master=pop)
        picture = None
        render_job = None
        thumb_job = None
        thumb_cache = OrderedDict()
        page_status = tk.StringVar(master=pop)

        def paint_state():
            pop.title(self.tr('Page preview - {name}', name=path.name))
            indicator.set(self.tr('Page {page} of {total}  •  {name}',
                                  page=current + 1, total=length, name=path.name))
            page_status.set(self.tr('{count} selected', count=len(chosen)))
            display.configure(highlightbackground=self.color(ACCENT if current in chosen else BORDER))
            if current in chosen:
                mark.place(x=7, y=7)
                mark.lift()
            else:
                mark.place_forget()
            if is_pdf:
                schedule_thumbs()

        def viewport(size, magnification, shift):
            width = max(display.winfo_width() - 20, 30)
            height = max(display.winfo_height() - 20, 30)
            source_width, source_height = size
            scale_limit = 1.6 if is_pdf else 1
            scale = min(scale_limit, width / source_width, height / source_height) * magnification

            def position(view, content, displacement):
                center = (view - content) / 2
                return center if content <= view else min(0, max(view - content, center + displacement))

            left = position(width, source_width * scale, shift[0])
            top = position(height, source_height * scale, shift[1])
            return width, height, scale, left, top

        def render():
            nonlocal picture, render_job, page_size
            render_job = None
            try:
                if is_pdf:
                    page = document.load_page(current)
                    box = page.rect
                    size = (max(1, box.width), max(1, box.height))
                else:
                    with Image.open(path) as img:
                        img.seek(current)
                        source = ImageOps.exif_transpose(img.convert('RGB'))
                    size = source.size

                width, height, scale, left, top = viewport(size, zoom, offset)
                x0, y0 = max(0, -left / scale), max(0, -top / scale)
                x1 = min(size[0], (width - left) / scale)
                y1 = min(size[1], (height - top) / scale)
                if is_pdf:
                    # Rasterize only the visible part, so high zoom doesn't
                    # allocate a full-page bitmap larger than the window.
                    clip = fitz.Rect(box.x0 + x0, box.y0 + y0, box.x0 + x1, box.y0 + y1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
                    image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                else:
                    target = (max(1, round((x1 - x0) * scale)), max(1, round((y1 - y0) * scale)))
                    image = source.resize(target, Image.Resampling.LANCZOS, box=(x0, y0, x1, y1))

                picture = ImageTk.PhotoImage(image, master=pop)
                label.configure(image=picture, text='')
                page_size = size
            except Exception as exc:
                picture = None
                page_size = None
                label.configure(image='', text=f"{self.tr('Cannot preview page:')} {exc}")

        def schedule_render(event=None):
            nonlocal render_job
            if render_job is not None:
                pop.after_cancel(render_job)
            render_job = pop.after(60, render)

        def navigate(index):
            nonlocal current, zoom, offset, page_size
            target = max(0, min(length - 1, index))
            if target != current:
                current = target
                zoom, offset, page_size = 1.0, (0.0, 0.0), None
            if is_pdf:
                top = thumbs.canvasy(0)
                y = (current // columns) * row_height
                if y < top:
                    thumbs.yview_moveto(y / canvas_height)
                elif y + row_height > top + thumbs.winfo_height():
                    thumbs.yview_moveto((y + row_height - thumbs.winfo_height()) / canvas_height)
            paint_state()
            schedule_render()

        def preview_wheel(event):
            nonlocal zoom, offset
            delta = event.delta
            if event.num == 4:
                delta = 120
            elif event.num == 5:
                delta = -120
            x = event.x_root - display.winfo_rootx()
            y = event.y_root - display.winfo_rooty()
            if not (0 <= x < display.winfo_width() and 0 <= y < display.winfo_height()):
                if is_pdf and delta and (thumbs.winfo_rootx() <= event.x_root
                        < thumbs.winfo_rootx() + thumbs.winfo_width()
                        and thumbs.winfo_rooty() <= event.y_root
                        < thumbs.winfo_rooty() + thumbs.winfo_height()):
                    return wheel(delta)
                return None
            if not delta:
                return 'break'
            steps = max(1, abs(delta) // 120)
            direction = 1 if delta > 0 else -1
            if not event.state & 0x0004:  # Tk ControlMask
                navigate(current - direction * steps)
            elif page_size is not None:
                new_zoom = min(8.0, max(1.0, zoom * 1.25 ** (direction * steps)))
                if new_zoom != zoom:
                    width, height, scale, left, top = viewport(page_size, zoom, offset)
                    x = min(width, max(0, x - (display.winfo_width() - width) / 2))
                    y = min(height, max(0, y - (display.winfo_height() - height) / 2))
                    source_x = min(page_size[0], max(0, (x - left) / scale))
                    source_y = min(page_size[1], max(0, (y - top) / scale))
                    ratio = new_zoom / zoom
                    new_width, new_height = page_size[0] * scale * ratio, page_size[1] * scale * ratio
                    origin_x = min(0, max(width - new_width, x - source_x * scale * ratio))
                    origin_y = min(0, max(height - new_height, y - source_y * scale * ratio))
                    offset = (origin_x - (width - new_width) / 2,
                              origin_y - (height - new_height) / 2)
                    zoom = new_zoom
                    schedule_render()
            return 'break'

        def toggle(index, shift=False):
            nonlocal anchor
            if shift and anchor is not None:
                # Shift+click selects every page between the last clicked page and this one.
                low, high = sorted((anchor, index))
                chosen.update(range(low, high + 1))
            elif index in chosen:
                chosen.remove(index)
            else:
                chosen.add(index)
            if not (shift and anchor is not None):
                anchor = index
            navigate(index)

        if is_pdf:
            columns, card_width, row_height = 2, 116, 145
            canvas_height = ((length + columns - 1) // columns) * row_height
            sidebar = tk.Frame(body, bg=BG)
            sidebar.pack(side='right', fill='y', padx=(10, 0))
            thumbs = tk.Canvas(sidebar, bg=BG, width=columns * card_width,
                               bd=0, highlightthickness=0, yscrollincrement=18,
                               scrollregion=(0, 0, columns * card_width, canvas_height))
            scrollbar = ttk.Scrollbar(sidebar, command=thumbs.yview, style='Preview.Vertical.TScrollbar')
            scrollbar.pack(side='right', fill='y')
            thumbs.pack(side='left', fill='y', expand=True)

            def thumbnail(index):
                if index not in thumb_cache:
                    try:
                        page = document.load_page(index)
                        box = page.rect
                        scale = min(100 / box.width, 108 / box.height)
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                        thumb_cache[index] = tk.PhotoImage(data=pix.tobytes('ppm'), master=pop)
                    except Exception:
                        thumb_cache[index] = None
                thumb_cache.move_to_end(index)
                return thumb_cache[index]

            def draw_thumbnails():
                nonlocal thumb_job
                thumb_job = None
                thumbs.delete('all')
                first = max(0, int(thumbs.canvasy(0)) // row_height) * columns
                last_row = int(thumbs.canvasy(thumbs.winfo_height())) // row_height
                for index in range(first, min(length, (last_row + 1) * columns)):
                    x = (index % columns) * card_width
                    y = (index // columns) * row_height
                    selected = index in chosen
                    thumbs.create_rectangle(x+3, y+3, x+card_width-4, y+row_height-4,
                                            fill=self.color(SELECT if index == current else SURFACE),
                                            outline=self.color(ACCENT if selected else BORDER),
                                            width=3 if selected else 1)
                    photo = thumbnail(index)
                    if photo is not None:
                        thumbs.create_image(x+card_width//2, y+60, image=photo)
                    thumbs.create_text(x+card_width//2, y+124, text=f'✓ {index+1}' if selected else str(index+1),
                                       fill=self.color(TEXT), font=('Segoe UI', 10, 'bold' if selected else 'normal'))
                # Keep visible images and at most 48 thumbnails for revisiting nearby pages.
                limit = max(48, (last_row + 1) * columns - first)
                while len(thumb_cache) > limit:
                    thumb_cache.popitem(last=False)

            def schedule_thumbs(event=None):
                nonlocal thumb_job
                if thumb_job is None:
                    thumb_job = pop.after_idle(draw_thumbnails)

            def scroll_changed(first, last):
                scrollbar.set(first, last)
                schedule_thumbs()

            def click_thumbnail(event):
                row = int(thumbs.canvasy(event.y)) // row_height
                column = event.x // card_width
                index = row * columns + column
                if 0 <= column < columns and 0 <= index < length:
                    toggle(index, bool(event.state & 0x0001))

            def wheel(delta):
                direction = -1 if delta > 0 else 1
                thumbs.yview_scroll(direction * 5, 'units')
                return 'break'

            thumbs.configure(yscrollcommand=scroll_changed)
            thumbs.bind('<Configure>', schedule_thumbs)
            thumbs.bind('<Button-1>', click_thumbnail)
            for widget in (display, label, mark):
                widget.configure(cursor='hand2')
                widget.bind('<Button-1>', lambda event: toggle(current))
            actions = tk.Frame(outer, bg=BG)
            actions.pack(fill='x', pady=(8, 0))

            def set_all(checked):
                chosen.clear()
                if checked:
                    chosen.update(range(length))
                paint_state()

            def apply():
                if self.busy:
                    self.info('Export', 'Wait for the export to finish before changing pages.', parent=pop)
                    return
                if path in self.paths:
                    self.apply_page_choice(path, page_selection_expression(chosen, length))
                pop.destroy()

            self.button(actions, 'Select All', lambda: set_all(True)).pack(side='left')
            self.button(actions, 'Clear Selection', lambda: set_all(False)).pack(side='left', padx=8)
            self.button(actions, 'Apply Selection', apply, primary=True).pack(side='right')
            tk.Label(actions, textvariable=page_status, bg=BG, fg=MUTED,
                     font=('Segoe UI', 9)).pack(side='right', padx=8)
            pop.bind('<Key-space>', lambda event: toggle(current) if event.widget is pop else None)
        self.button(footer, 'Previous', lambda: navigate(current - 1), navigation=True).pack(side='left')
        self.button(footer, 'Next', lambda: navigate(current + 1), navigation=True).pack(side='right')
        tk.Label(footer, textvariable=indicator, bg=BG, fg=MUTED, font=('Segoe UI', 9),
                 wraplength=400).pack(side='left', expand=True)
        display.bind('<Configure>', schedule_render)
        for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
            pop.bind(sequence, preview_wheel)
        pop.bind('<Left>', lambda event: navigate(current - 1))
        pop.bind('<Right>', lambda event: navigate(current + 1))
        pop.bind('<Escape>', lambda event: pop.destroy())

        def cleanup(event):
            if event.widget is not pop:
                return
            for pending in (render_job, thumb_job):
                if pending is not None:
                    pop.after_cancel(pending)
            if document is not None:
                document.close()
            thumb_cache.clear()

        pop.bind('<Destroy>', cleanup)
        pop._refresh_appearance = paint_state
        self.style_widgets(pop)
        pop.focus_set()
        schedule_render()

    def open_folder(self):
        if self.busy:
            return
        name = filedialog.askdirectory(parent=self, title=self.tr('Choose folder with PDFs / images / Office files'))
        if not name:
            return
        self._remember()
        self.paths = sorted((p for p in Path(name).iterdir() if p.is_file() and p.suffix.lower() in EXTS),
                            key=sort_key)
        self.page_choices = {p: expr for p, expr in self.page_choices.items() if p in self.paths}
        self.location.set(name)
        self.refresh()
        self.status.set(self.tr('{count} file(s) loaded in natural order.', count=len(self.paths)))

    def add_files(self):
        if self.busy:
            return
        names = filedialog.askopenfilenames(parent=self, title=self.tr('Add PDFs / images / Office files'),
            filetypes=[(self.tr('PDF, images and Office'), '*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.doc *.docx *.docm *.xls *.xlsx *.xlsm *.ppt *.pptx *.pptm'), (self.tr('All files'), '*.*')])
        if names:
            self.location.set('Files from one or more locations')
            self.add_paths([Path(n) for n in names], sort_new=True)

    def sort_naturally(self):
        if self.busy:
            return
        self._remember()
        self.paths.sort(key=sort_key)
        self.refresh()
        self.status.set(self.tr('List sorted naturally.'))

    def clear(self):
        if self.busy:
            return
        if self.paths:
            self._remember()
        self.paths.clear()
        self.page_choices.clear()
        self.location.set('No folder selected')
        self.refresh()
        self.status.set(self.tr('List cleared.'))

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
        self.status.set(self.tr('Removed {count} row(s).', count=len(indices)))

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
            self.status.set(self.tr('Moved {name} to position {position}.', name=moved.name, position=dest + 1))

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
            messagebox.showerror(self.tr('Selection error'), str(exc), parent=self)
            return
        pop = tk.Toplevel(self, bg=BG)
        pop.title(f'Merge order — {len(files)} files')
        pop.geometry('690x470')
        preview = tk.Text(pop, bg=SURFACE, fg=TEXT, insertbackground=TEXT, borderwidth=0,
                          font=('Consolas', 11), wrap='none', padx=14, pady=14)
        preview.pack(fill='both', expand=True, padx=12, pady=12)
        preview.insert('1.0', '\n'.join(f'{i}. {p.name}' for i, p in enumerate(files, 1)))
        preview.config(state='disabled')
        self.style_widgets(pop)

    def start_merge(self, kind='pdf'):
        if self.busy:
            return
        try:
            files = self.selected_jobs()
        except Exception as exc:
            messagebox.showerror(self.tr('Selection error'), str(exc), parent=self)
            return
        stem = f'Wadaa_PDF_{datetime.now():%Y-%m-%d_%H-%M-%S}'
        if kind == 'jpg':
            parent = filedialog.askdirectory(parent=self, mustexist=True,
                                             title=self.tr('Choose a parent folder for the JPG export'))
            if not parent:
                return
            directory = Path(parent)
            destination = directory / f'{stem}_JPG'
            suffix = 2
            while destination.exists():
                destination = directory / f'{stem}_JPG_{suffix}'
                suffix += 1
        else:
            title, label = {
                'pdf': ('Save merged PDF', 'PDF files'),
                'docx': ('Save Word document', 'Word documents'),
                'pptx': ('Save PowerPoint presentation', 'PowerPoint presentations'),
            }[kind]
            target = filedialog.asksaveasfilename(parent=self, title=self.tr(title),
                defaultextension=f'.{kind}', initialfile=f'{stem}.{kind}',
                filetypes=[(self.tr(label), f'*.{kind}')])
            if not target:
                return
            destination = Path(target)
            if destination.suffix.casefold() != f'.{kind}':
                messagebox.showerror(self.tr('Unsafe output'), self.tr('Use the .{kind} extension.', kind=kind),
                                     parent=self)
                return
            if any(destination.resolve() == path.resolve() for path, _ in files):
                messagebox.showerror(self.tr('Unsafe output'),
                                     self.tr('Choose a different file name; cannot overwrite a source file.'),
                                     parent=self)
                return
            if destination.exists() and not messagebox.askyesno(
                self.tr('Overwrite?'), self.tr('Replace {name}?', name=destination.name), parent=self,
            ):
                return
        self.busy = True
        self.merge_button.configure(state='disabled')
        self.progress['value'] = 0
        self.status.set(self.tr('Merging...'))
        self._update_history_buttons()
        threading.Thread(target=self.worker, args=(files, destination, kind), daemon=True).start()

    def worker(self, files, destination, kind='pdf'):
        def progress(done, total, name):
            self.after(0, self.update_progress, done, total, name)

        try:
            if kind == 'pdf':
                merge_files(files, destination, progress)
                text = self.tr('Merged {count} file(s).', count=len(files))
            else:
                count, report = export_merged_pages(files, destination, kind, progress)
                text = self.tr('Exported {count} page(s).', count=count)
                for reason, template in OFFICE_WARNINGS.items():
                    if reason in report:
                        pages = ', '.join(str(page) for page in report[reason][:12])
                        if len(report[reason]) > 12:
                            pages += ', …'
                        text += '\n' + self.tr(template, pages=pages)
        except Exception as exc:
            self.after(0, self.finish, False, str(exc), destination, kind)
        else:
            self.after(0, self.finish, True, text, destination, kind)

    def update_progress(self, done, total, name):
        self.progress['value'] = done * 100 / total if total else 0
        self.status.set(f'{done}/{total} — {name}')

    def finish(self, ok, text, destination, kind='pdf'):
        self.busy = False
        self.merge_button.configure(state='normal')
        self._update_history_buttons()
        if ok:
            self.output_file = destination
            self.progress['value'] = 100
            self.status.set(self.tr('Completed: {path}', path=destination))
            if kind == 'pdf' and self.open_after_merge.get():
                try:
                    if sys.platform == 'win32':
                        os.startfile(str(destination))
                    else:
                        webbrowser.open(destination.as_uri())
                except OSError as exc:
                    messagebox.showwarning(self.tr('Open PDF'), str(exc), parent=self)
            elif self.themed_dialog('Merge completed' if kind == 'pdf' else 'Export completed',
                                     f"{text}\n\n{destination}\n\n{self.tr('Open output folder?')}",
                                     confirm=True):
                self.open_output_folder()
        else:
            self.status.set(self.tr('Error: {error}', error=text))
            messagebox.showerror(self.tr('Merge failed' if kind == 'pdf' else 'Export failed'),
                                 text, parent=self)

    def open_output_folder(self):
        if not self.output_file:
            self.info('Output folder', 'Export a file first.')
            return
        try:
            if sys.platform == 'win32':
                if self.output_file.is_dir():
                    os.startfile(str(self.output_file))
                else:
                    subprocess.Popen(['explorer', '/select,', str(self.output_file)])
            else:
                folder = self.output_file if self.output_file.is_dir() else self.output_file.parent
                webbrowser.open(folder.resolve().as_uri())
        except OSError as exc:
            self.themed_dialog('Open folder', str(exc))


if __name__ == '__main__':
    App().mainloop()
