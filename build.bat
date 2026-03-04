@echo off
REM Build servicio_tool.exe on Windows
REM Requires: Python + pip install -r requirements.txt -r requirements-build.txt

echo Installing dependencies...
pip install -r requirements.txt -r requirements-build.txt --quiet

echo Building executable...
pyinstaller servicio_tool.spec --noconfirm

echo Done. Single exe: dist\servicio_tool.exe
pause
