@echo off
SETLOCAL

REM Script pour nettoyer les fichiers .DEPRECATED
REM Créé automatiquement pour le framework APEX

echo ===================================================
echo  APEX Framework - Nettoyage des fichiers .DEPRECATED
echo ===================================================
echo.

set SCRIPT_PATH=%~dp0clean_deprecated.ps1

echo Lancement du script de nettoyage...
powershell -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

REM Le script PowerShell gère sa propre pause 