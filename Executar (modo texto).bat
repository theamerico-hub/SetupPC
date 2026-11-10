@echo off
rem Abre o SetupPC no modo texto (menu no console)
title SetupPC - modo texto
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0SetupPC.ps1" -Console
