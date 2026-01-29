@echo off
REM Build script for Windows

echo Building 错题集生成器 for Windows...

cd /d "%~dp0\.."

REM Create venv if needed
if not exist ".venv" (
    python -m venv .venv
    .venv\Scripts\pip install -r requirements.txt
)

REM Install pyinstaller
.venv\Scripts\pip install pyinstaller

REM Build
.venv\Scripts\pyinstaller ^
    --onefile ^
    --windowed ^
    --name="错题集生成器" ^
    --add-data="ui/icons;ui/icons" ^
    main.py

echo Build complete! Executable: dist\错题集生成器.exe
pause
