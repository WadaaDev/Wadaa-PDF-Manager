Wadaa PDF Manager 1.0 - Windows x64

COMPILE SETUP FROM THE EXTRACTED ZIP
1. Extract the complete ZIP to a local folder, such as C:\Build.
2. Open Wadaa_PDF_Manager.iss in Inno Setup 7.1.1 and click Compile.
3. The result is installer\Wadaa PDF Manager Setup v1.0.exe.

The ZIP already includes the complete dist\Wadaa PDF Manager folder.
Keep the application executable and its _internal folder together.
Compiling Setup from the supplied bundle needs neither Python nor internet.
The Setup and app are unsigned. The Start Menu shortcut uses the app's
embedded original icon. This installer follows the previous Start Menu-only
shortcut configuration.

REBUILD THE APPLICATION FROM SOURCE (OPTIONAL)
The build was verified with Windows x64 Python 3.14.10 and the package
versions listed in requirements.txt. Use Python with Tcl/Tk and those
packages already installed. If supplying your own offline wheel directory,
install the build prerequisites with:
    python -m pip install --no-index --find-links C:\Wheelhouse -r requirements.txt
Point WADAA_BUILD_PYTHON at the full path of python.exe if it is not on PATH.
Run BUILD WADAA PDF.bat. It uses only the configured local Python and
recreates dist\Wadaa PDF Manager. It does not install packages. Compile
the ISS again if you want Setup to contain your new build.

FILES
  Wadaa_PDF_Manager.py          Existing Tkinter application
  BUILD WADAA PDF.bat           Offline One-Folder build command
  Wadaa_PDF_Manager.iss         Final installer script
  Wadaa_PDF_Manager_original.iss  Original installer script backup
  backup\                       Pre-release source, BAT, ISS and original README
  assets\                       Unchanged original assets
  updater.py, update_settings.json  Optional updater and configuration
  VERSION.txt                  Shared version for About, updater and Setup
  pyinstaller_hooks\            Collects runtime PyMuPDF assets without optional
                                document-layout data and C/C++ SDK files
  dist\Wadaa PDF Manager\       Built executable and bundled runtime
  installer\                    Output directory for Inno Setup

ABOUT & SUPPORT
Use About & Support in the main window footer. Contact Support opens
your default mail app with a draft. Report a Bug and GitHub Repository
open their addresses in your default browser only when clicked.

LAST SESSION
The Last Session toggle sits beside Clear All. It starts enabled for
existing users. When enabled, the app restores the saved files at startup
and saves on close. Turning it off leaves the current list and earlier
save intact. The next launch starts empty while disabled. Re-enabling
leaves an open list intact. If the list is empty and an earlier save exists,
the app offers to restore it. Closing with the toggle enabled writes the
current list.

Session data is stored in %APPDATA%\WadaaPDFManager\session.json.
The toggle setting is stored beside it in session-preference.json.
Installation and uninstall leave these files alone.

FUTURE UPDATES
Online updates are disabled in this delivery. When you are ready to test
with an official public release, set enabled to true in update_settings.json
and rebuild. For a new release, change VERSION.txt before rebuilding. Setup
checks the bundled version before compiling. The updater expects a stable version tag (for example v1.1 or v1.1.0),
a matching Wadaa PDF Manager Setup v<version>.exe asset, and the official
release metadata's SHA-256 digest. Missing digests block download. Validate
a published test release before enabling updates for users, including
startup retry, cancellation, interrupted resume, the elevation prompt,
installation, and automatic reopening. No live release or online service
was used to build or test this ZIP.

TEST COVERAGE
Locally checked Python syntax and imports, the PyInstaller build, frozen
application launch and clean exit, PDF merge and rendering, Tk loading the original icon,
Last Session across launches, and offline updater fixtures for version
comparison, host restrictions, range resume, asset changes, hash rejection,
and concurrent access. Compiled Setup with Inno Setup 7.1.1 and compared
embedded icon image bytes in both executables against the original ICO.
Compiled Setup directly from an extracted copy of this project. The path
checks covered drive roots, duplicate folder names, Next correction, and
simulated preservation of a previous installation path. A local automated
wizard test opened Browse and confirmed the final path appeared on the
destination page immediately after confirmation, before Next.
An isolated, non-registering test installer also installed and reinstalled
the runtime files at the corrected destination. An isolated uninstall
removed the application and left a test user PDF in place.

Still to verify on a test Windows installation: a registered upgrade
using the real AppId, the Start Menu shortcut as
rendered by Windows, and Microsoft Office or LibreOffice conversion.
Office document conversion still requires one of those applications.
Test the enabled updater only after you have an official test release.
