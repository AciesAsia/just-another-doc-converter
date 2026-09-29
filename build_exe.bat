@echo off
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building DriveSync.exe...
pyinstaller --onefile --windowed --name DriveSync gdrive_sync_gui.py

echo.
echo Done. Your .exe is in the dist\ folder.
echo Copy dist\DriveSync.exe anywhere you like, including your Desktop.
echo.
pause
