@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Cleaning previous builds...
rmdir /s /q build dist
del *.spec

echo Building application...
pyinstaller --onefile --windowed --name "TaxChecker" --clean --noconfirm gui_app_qt.py

echo Build complete. App is located in dist\TaxChecker.exe
pause
