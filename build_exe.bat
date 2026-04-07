@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1 && set "PY=py" || set "PY=python"
%PY% -m pip install --upgrade pip
%PY% -m pip install "pyinstaller>=6.0.0"

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name SerialChecker app.py

echo.
echo Done. Executable: dist\SerialChecker.exe
echo You can copy only that .exe anywhere; first run may be slower (Windows one-file unpack).
pause
