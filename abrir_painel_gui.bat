@echo off
cd /d "%~dp0"
set "TCL_LIBRARY=C:\Users\LFVieira\AppData\Local\Programs\Python\Python313\tcl\tcl8.6"
set "TK_LIBRARY=C:\Users\LFVieira\AppData\Local\Programs\Python\Python313\tcl\tk8.6"
".\.venv\Scripts\pythonw.exe" ".\painel_admin_gui.py"
if errorlevel 1 (
  ".\.venv\Scripts\python.exe" ".\painel_admin_gui.py"
  pause
)
