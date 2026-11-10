@echo off
rem Abre o SetupPC com interface grafica (sem janela preta)
cd /d "%~dp0"
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SetupPC.ps1"
