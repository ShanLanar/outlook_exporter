# OutlookExporter.spec
# PyInstaller Spec-Datei – reproduzierbarer Build einer Standalone-.exe.
# Verwendung:  pyinstaller OutlookExporter.spec   (oder einfach build.bat)
# Ausgabe:     dist\OutlookExporter.exe  (kein Python auf dem Zielrechner nötig)

from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

a = Analysis(
    [str(ROOT / "outlook_advanced_exporter.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=[
        # pywin32 / COM-Zugriff auf Outlook
        "win32com", "win32com.client", "win32timezone",
        "pythoncom", "pywintypes",
        # tkinter-GUI
        "tkinter", "tkinter.ttk", "tkinter.scrolledtext",
        "tkinter.filedialog", "tkinter.messagebox",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # Nicht benötigte, große Pakete ausschließen → kleinere .exe
        "matplotlib", "scipy", "numpy", "pandas", "PIL",
        "PyQt5", "PyQt6", "wx", "IPython", "jupyter",
        "test", "unittest",
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="OutlookExporter",
    debug=False,
    strip=False,
    upx=True,
    console=False,      # kein schwarzes Konsolenfenster zur GUI
    # icon="icon.ico",  # optional: eigenes Icon einbinden
)
