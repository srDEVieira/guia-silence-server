@echo off
cd /d "%~dp0"
powershell -NoExit -ExecutionPolicy Bypass -File ".\painel_admin.ps1"
