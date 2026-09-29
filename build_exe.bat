@echo off
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building Drive Sync.exe...
pyinstaller --onefile --windowed --name "Drive Sync" gdrive_sync_gui.py

echo.
echo Done. Your .exe is in the dist\ folder.
echo Copy dist\Drive Sync.exe anywhere you like, including your Desktop.
echo.
pause
