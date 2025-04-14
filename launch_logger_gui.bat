@echo off
echo ===================================================
echo =   Lancement de l'interface APEX Logger GUI      =
echo ===================================================
echo.

:: Vérifier si Python est installé
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERREUR] Python n'est pas installé ou n'est pas dans le PATH.
    echo Veuillez installer Python et réessayer.
    echo.
    pause
    exit /b 1
)

:: Vérifier si le fichier de l'interface graphique existe
if not exist "%~dp0apex_cursor_logger\logger_gui.py" (
    echo [ERREUR] Le fichier logger_gui.py n'a pas été trouvé dans le répertoire apex_cursor_logger.
    echo Assurez-vous que le fichier existe et que vous exécutez ce script à partir du bon répertoire.
    echo.
    pause
    exit /b 1
)

echo [INFO] Lancement de l'interface graphique APEX Logger...
echo.

:: Lancer l'interface graphique
python "%~dp0apex_cursor_logger\logger_gui.py"

:: Vérifier si le programme s'est terminé avec une erreur
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERREUR] L'interface graphique s'est terminée avec une erreur (code %ERRORLEVEL%).
    echo Consultez les logs pour plus de détails.
) else (
    echo.
    echo [SUCCÈS] L'interface graphique s'est terminée normalement.
)

echo.
pause