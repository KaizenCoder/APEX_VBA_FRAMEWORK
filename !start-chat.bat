@echo off
setlocal

echo ===================================
echo    APEX Framework - Chat Hub
echo ===================================
echo.

:: Vérifier si PowerShell est disponible
where pwsh >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Lancement avec PowerShell Core...
    pwsh -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0coordination\hub\Start-ApexChatGUI.ps1'"
) else (
    echo Lancement avec Windows PowerShell...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0coordination\hub\Start-ApexChatGUI.ps1'"
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo Erreur lors du lancement de l'application.
    echo Vérifiez que PowerShell est installé et que les scripts sont autorisés.
    pause
    exit /b 1
)

exit /b 0 