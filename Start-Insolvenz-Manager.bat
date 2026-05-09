@echo off
chcp 65001 >nul
title Insolvenz-Manager v2.0

echo ============================================================
echo   Insolvenz-Manager v2.0
echo   Startet lokalen Server und öffnet Browser automatisch...
echo ============================================================
echo.

REM Prüfen ob Python vorhanden
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python nicht gefunden!
    echo Bitte Python 3.x von https://www.python.org installieren.
    pause
    exit /b 1
)

REM Wechsle in Verzeichnis der Batch-Datei
cd /d "%~dp0"

REM Abhängigkeiten prüfen / installieren
echo Prüfe Abhängigkeiten...
python -c "import pypdf" >nul 2>&1
if errorlevel 1 (
    echo Installiere pypdf...
    pip install pypdf --quiet
)

python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installiere python-docx...
    pip install python-docx --quiet
)

echo.
echo Server wird gestartet...
echo Browser öffnet sich automatisch.
echo.
echo Zum Beenden: Dieses Fenster schließen oder Strg+C drücken.
echo.

python insolvenz_server.py

pause
