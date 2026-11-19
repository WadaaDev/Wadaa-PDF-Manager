@echo off
setlocal
cd /d "%~dp0"
set "PY=python"
set "PYINSTALLER_CONFIG_DIR=%CD%\build\cache"
if defined WADAA_BUILD_PYTHON set "PY=%WADAA_BUILD_PYTHON%"
"%PY%" -c "import PyInstaller, tkinter, tkinterdnd2, pymupdf, PIL, pypdf, pythoncom, win32com.client" || goto MISSING
if not exist "assets\icon.ico" goto MISSING
if not exist "VERSION.txt" goto MISSING
"%PY%" -m PyInstaller --noconfirm --clean --onedir --windowed --name "Wadaa PDF Manager" --workpath "build\pyinstaller" --specpath "build" --distpath "dist" --icon "%CD%\assets\icon.ico" --add-data "%CD%\assets;assets" --add-data "%CD%\update_settings.json;." --add-data "%CD%\VERSION.txt;." --additional-hooks-dir "%CD%\pyinstaller_hooks" --collect-data tkinterdnd2 --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes "Wadaa_PDF_Manager.py"
if errorlevel 1 exit /b 1
if not exist "dist\Wadaa PDF Manager\_internal\assets\icon.ico" exit /b 1
echo Built dist\Wadaa PDF Manager\Wadaa PDF Manager.exe
exit /b 0
:MISSING
echo A local build prerequisite or the original icon is missing. See README_BUILD.txt.
exit /b 1



