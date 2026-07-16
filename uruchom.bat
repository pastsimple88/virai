@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
    set PY=python
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set PY=py
    ) else (
        echo Nie znaleziono Pythona na tym komputerze.
        echo Pobierz go z https://www.python.org/downloads/
        echo WAZNE: podczas instalacji zaznacz "Add python.exe to PATH".
        echo.
        pause
        exit /b 1
    )
)

echo Instaluje potrzebna biblioteke "requests" (jesli jeszcze jej nie masz)...
%PY% -m pip install --quiet requests

echo.
echo Uruchamiam skrypt, to moze potrwac chwile...
echo.
%PY% find_app_leads.py

echo.
echo ============================================
echo Gotowe. Wyniki sa w pliku app_leads.csv
echo w tym samym folderze co ten plik .bat
echo ============================================
echo.
pause
