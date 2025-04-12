@echo off
SETLOCAL

REM Script pour renommer les fichiers journaux obsolètes
REM Créé automatiquement pour le framework APEX

echo ===================================================
echo  APEX Framework - Renommage des logs obsolètes
echo ===================================================
echo.

set SCRIPT_PATH=%~dp0python\rename_deprecated_logs.py
set PROJECT_ROOT=%~dp0..

REM Définir les arguments par défaut
set DRY_RUN=--dry-run
set VERBOSE=-v
set TARGET_DIR=%PROJECT_ROOT%

REM Analyser les arguments
:parse_args
if "%~1"=="" goto :done_args
if /i "%~1"=="--no-dry-run" (
    set DRY_RUN=
) else if /i "%~1"=="--dir" (
    set TARGET_DIR=%~2
    shift
)
shift
goto :parse_args
:done_args

echo Configuration:
echo - Répertoire cible: %TARGET_DIR%
echo - Mode simulation: %DRY_RUN%
echo - Logs détaillés: %VERBOSE%
echo.

REM Exécuter le script Python
echo Démarrage du renommage...
python "%SCRIPT_PATH%" --dir "%TARGET_DIR%" %DRY_RUN% %VERBOSE% --export-csv "%PROJECT_ROOT%\rename_logs_report.csv"

REM Vérifier le résultat
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo  Opération terminée avec succès!
    echo  Les résultats détaillés sont disponibles dans le
    echo  fichier de log créé dans le répertoire courant.
    echo ===================================================
) else (
    echo.
    echo ===================================================
    echo  ERREUR: Échec de l'opération.
    echo  Vérifiez les messages d'erreur ci-dessus.
    echo ===================================================
)

echo.
echo Appuyez sur une touche pour fermer cette fenêtre...
pause > nul 