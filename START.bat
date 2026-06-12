@echo off
REM Outlook Advanced Exporter - Start Script
REM Prüft ob Python verfügbar ist und startet das Programm

echo.
echo ============================================
echo Outlook Advanced Exporter
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python nicht gefunden!
    echo Bitte installieren Sie Python 3.6+ von python.org
    pause
    exit /b 1
)

echo Python gefunden. Starte Exporter...
echo.

python outlook_advanced_exporter.py

if errorlevel 1 (
    echo.
    echo FEHLER beim Start des Exporters.
    echo Bitte stellen Sie sicher, dass:
    echo - Outlook geöffnet ist
    echo - pywin32 installiert ist (pip install pywin32)
    pause
)

pause
