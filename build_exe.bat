@echo off
echo ========================================
echo   PostNord TPL - Bygger exe-fil
echo ========================================
echo.

echo Installerar beroenden...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Bygger exe-fil...
pyinstaller --onefile --windowed --name "PostNord_TPL_Michelle" --clean excel_merger_app.py

echo.
echo ========================================
echo   Klart!
echo   Exe-filen finns i: dist\PostNord_TPL_Michelle.exe
echo ========================================
echo.
pause

