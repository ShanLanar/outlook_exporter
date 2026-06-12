@echo off
setlocal enabledelayedexpansion
:: build.bat - Erzeugt dist\OutlookExporter.exe (kein Python auf Zielrechner noetig)
:: Voraussetzung: Python 3.10+ auf dem BUILD-Rechner.

echo ============================================================
echo  Outlook Advanced Exporter - EXE-Build
echo ============================================================
echo.

:: --- Python finden ---------------------------------------------------------
python --version >nul 2>&1
if %errorlevel%==0 ( set "PYTHON=python" & goto found )
py -3 --version >nul 2>&1
if %errorlevel%==0 ( set "PYTHON=py -3" & goto found )
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
) do ( if exist %%P ( set "PYTHON=%%P" & goto found ) )
echo FEHLER: Python nicht gefunden. Bitte Python 3.10+ installieren.
pause & exit /b 1

:found
echo Verwende Python: %PYTHON%
echo.

:: --- Build-Abhaengigkeiten installieren ------------------------------------
echo Installiere pywin32 + pyinstaller ...
%PYTHON% -m pip install --upgrade pip >nul
%PYTHON% -m pip install pywin32 pyinstaller || ( echo FEHLER bei pip install & pause & exit /b 1 )
echo.

:: --- Bauen -----------------------------------------------------------------
echo Baue OutlookExporter.exe ...
%PYTHON% -m PyInstaller --noconfirm --clean OutlookExporter.spec || ( echo FEHLER beim Build & pause & exit /b 1 )

echo.
echo ============================================================
echo  Fertig:  dist\OutlookExporter.exe
echo ============================================================
pause
