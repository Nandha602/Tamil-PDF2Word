@echo off
rem Builds dist\PDF2Word.exe (a single portable file). Needs Python 3.10+ on PATH.
cd /d "%~dp0"
py -3 -m pip install -r requirements-build.txt || (echo Install failed & pause & exit /b 1)
rem the optional pymupdf-layout add-on is not used and breaks packaged builds
py -3 -m pip uninstall -y pymupdf-layout pymupdf4llm >nul 2>&1
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --exclude-module pymupdf.layout --collect-all uharfbuzz --name PDF2Word pdf2word.py || (echo Build failed & pause & exit /b 1)
echo.
echo Done: dist\PDF2Word.exe
pause
