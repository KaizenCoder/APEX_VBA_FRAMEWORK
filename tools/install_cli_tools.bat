@echo off
SETLOCAL

REM Script pour installer les outils CLI APEX Framework
REM Créé automatiquement

echo ===================================================
echo  APEX Framework - Installation des outils CLI
echo ===================================================
echo.

set CLI_PATH=%~dp0python\apex_rename_logs

echo Installation de l'outil 'apex-rename-logs'...
cd "%CLI_PATH%"
pip install -e .

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo  Installation réussie!
    echo  Vous pouvez maintenant utiliser:
    echo.
    echo  apex-rename-logs --help
    echo ===================================================
) else (
    echo.
    echo ===================================================
    echo  ERREUR: L'installation a échoué.
    echo  Vérifiez les messages d'erreur ci-dessus.
    echo ===================================================
)

echo.
echo Appuyez sur une touche pour fermer cette fenêtre...
pause > nul 