@echo off
REM Drag&Drop: DXL-Datei(en) oder Ordner auf diese Datei ziehen
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0dxl2html.ps1" -Open %*
if errorlevel 1 pause
