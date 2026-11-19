from PyInstaller.utils.hooks import collect_data_files

# Preserve package data used by the stock hook. PDF preview uses core rendering;
# the optional document layout add-on and C/C++ SDK are unused by this app.
datas = collect_data_files('pymupdf', excludes=['layout/**', 'mupdf-devel/**'])
