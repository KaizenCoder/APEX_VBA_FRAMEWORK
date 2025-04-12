@echo off
REM ==========================================================================
REM Script : BuildRelease.bat
REM Version : 1.2
REM Purpose : Construction d'une version de release du framework APEX
REM Date : 10/04/2025
REM ==========================================================================

echo ===== CONSTRUCTION DE LA RELEASE APEX FRAMEWORK =====
echo.

REM --- Variables ---
set VERSION=1.1.0
set RELEASE_DIR=release
set OUTPUT_DIR=%RELEASE_DIR%\ApexFramework_v%VERSION%
set LOG_FILE=build_log.txt
set TEMP_DIR=temp_build

REM --- Initialisation ---
echo [INFO] Initialisation de la construction...
if exist %LOG_FILE% del %LOG_FILE%
echo Début de la construction: %date% %time% > %LOG_FILE%

REM --- Vérification des prérequis ---
echo [INFO] Vérification des prérequis...
if not exist apex-core (
    echo [ERREUR] Répertoire apex-core manquant
    echo [ERREUR] Répertoire apex-core manquant >> %LOG_FILE%
    goto :error
)
if not exist apex-metier (
    echo [ERREUR] Répertoire apex-metier manquant
    echo [ERREUR] Répertoire apex-metier manquant >> %LOG_FILE%
    goto :error
)
if not exist apex-ui (
    echo [ERREUR] Répertoire apex-ui manquant
    echo [ERREUR] Répertoire apex-ui manquant >> %LOG_FILE%
    goto :error
)
if not exist config (
    echo [ERREUR] Répertoire config manquant
    echo [ERREUR] Répertoire config manquant >> %LOG_FILE%
    goto :error
)

REM --- Création des répertoires ---
echo [INFO] Création des répertoires de sortie...
if not exist %RELEASE_DIR% mkdir %RELEASE_DIR%
if exist %OUTPUT_DIR% rmdir /S /Q %OUTPUT_DIR%
mkdir %OUTPUT_DIR%
mkdir %OUTPUT_DIR%\apex-core
mkdir %OUTPUT_DIR%\apex-metier
mkdir %OUTPUT_DIR%\apex-ui
mkdir %OUTPUT_DIR%\config
mkdir %OUTPUT_DIR%\docs
mkdir %OUTPUT_DIR%\examples

REM --- Copie des fichiers sources ---
echo [INFO] Copie des fichiers sources...
xcopy /E /Y apex-core\*.* %OUTPUT_DIR%\apex-core\ >> %LOG_FILE% 2>&1
xcopy /E /Y apex-metier\*.* %OUTPUT_DIR%\apex-metier\ >> %LOG_FILE% 2>&1
xcopy /E /Y apex-ui\*.* %OUTPUT_DIR%\apex-ui\ >> %LOG_FILE% 2>&1

REM --- Copie des fichiers de configuration ---
echo [INFO] Copie des fichiers de configuration...
xcopy /E /Y config\*.ini %OUTPUT_DIR%\config\ >> %LOG_FILE% 2>&1
xcopy /E /Y config\*.json %OUTPUT_DIR%\config\ >> %LOG_FILE% 2>&1
xcopy /E /Y config\*.xml %OUTPUT_DIR%\config\ >> %LOG_FILE% 2>&1

REM --- Copie de la documentation ---
echo [INFO] Copie de la documentation...
xcopy /E /Y docs\*.md %OUTPUT_DIR%\docs\ >> %LOG_FILE% 2>&1
xcopy /E /Y CHANGELOG.md %OUTPUT_DIR%\docs\ >> %LOG_FILE% 2>&1
xcopy /E /Y ROADMAP.md %OUTPUT_DIR%\docs\ >> %LOG_FILE% 2>&1

REM --- Copie des exemples ---
echo [INFO] Copie des exemples...
if exist examples xcopy /E /Y examples %OUTPUT_DIR%\examples\ >> %LOG_FILE% 2>&1
if exist samples xcopy /E /Y samples %OUTPUT_DIR%\examples\ >> %LOG_FILE% 2>&1

REM --- Validation de la release ---
echo [INFO] Validation de la structure de la release...
set ERROR_COUNT=0

REM Vérification des fichiers essentiels Core
if not exist "%OUTPUT_DIR%\apex-core\clsLogger.cls" (
    echo [ERREUR] Fichier manquant: apex-core\clsLogger.cls >> %LOG_FILE%
    set /A ERROR_COUNT+=1
)
if not exist "%OUTPUT_DIR%\apex-core\modConfigManager.bas" (
    echo [ERREUR] Fichier manquant: apex-core\modConfigManager.bas >> %LOG_FILE%
    set /A ERROR_COUNT+=1
)

REM Vérification des fichiers essentiels Métier
if not exist "%OUTPUT_DIR%\apex-metier\recette\modRecipeComparer.bas" (
    echo [ERREUR] Fichier manquant: apex-metier\recette\modRecipeComparer.bas >> %LOG_FILE%
    set /A ERROR_COUNT+=1
)

REM Vérification des fichiers essentiels UI
if not exist "%OUTPUT_DIR%\apex-ui\ribbon\customUI.xml" (
    echo [ERREUR] Fichier manquant: apex-ui\ribbon\customUI.xml >> %LOG_FILE%
    set /A ERROR_COUNT+=1
)

REM --- Génération du fichier de version ---
echo [INFO] Génération du fichier de version...
echo Version: %VERSION% > %OUTPUT_DIR%\VERSION.txt
echo Date de création: %date% >> %OUTPUT_DIR%\VERSION.txt
echo Heure de création: %time% >> %OUTPUT_DIR%\VERSION.txt
echo. >> %OUTPUT_DIR%\VERSION.txt
echo Architecture Apex Framework v1.1 à trois couches: >> %OUTPUT_DIR%\VERSION.txt
echo - Core: Modules techniques et transversaux >> %OUTPUT_DIR%\VERSION.txt
echo - Métier: Modules applicatifs et fonctionnels >> %OUTPUT_DIR%\VERSION.txt
echo - UI: Interface utilisateur et interactions >> %OUTPUT_DIR%\VERSION.txt
echo. >> %OUTPUT_DIR%\VERSION.txt
echo Fichiers inclus: >> %OUTPUT_DIR%\VERSION.txt
dir /B /S %OUTPUT_DIR%\apex-core\*.cls %OUTPUT_DIR%\apex-core\*.bas %OUTPUT_DIR%\apex-core\*.frm | find /c "." >> %OUTPUT_DIR%\VERSION.txt
dir /B /S %OUTPUT_DIR%\apex-metier\*.cls %OUTPUT_DIR%\apex-metier\*.bas %OUTPUT_DIR%\apex-metier\*.frm | find /c "." >> %OUTPUT_DIR%\VERSION.txt
dir /B /S %OUTPUT_DIR%\apex-ui\*.cls %OUTPUT_DIR%\apex-ui\*.bas %OUTPUT_DIR%\apex-ui\*.frm | find /c "." >> %OUTPUT_DIR%\VERSION.txt

REM --- Génération du fichier de licence ---
echo [INFO] Génération du fichier de licence...
echo APEX FRAMEWORK >> %OUTPUT_DIR%\LICENSE.txt
echo Copyright (c) 2025 APEX Team >> %OUTPUT_DIR%\LICENSE.txt
echo. >> %OUTPUT_DIR%\LICENSE.txt
echo Tous droits réservés. >> %OUTPUT_DIR%\LICENSE.txt

REM --- Génération du fichier README ---
echo [INFO] Génération du fichier README...
echo # APEX Framework v%VERSION% > %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo ## Architecture >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo Le framework est organisé en trois couches distinctes : >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo 1. **Apex.Core** - Modules techniques et transversaux >> %OUTPUT_DIR%\README.md
echo 2. **Apex.Métier** - Modules applicatifs et fonctionnels >> %OUTPUT_DIR%\README.md
echo 3. **Apex.UI** - Interface utilisateur et interactions >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo ## Installation >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo 1. Décompressez l'archive dans un répertoire de votre choix >> %OUTPUT_DIR%\README.md
echo 2. Importez les modules VBA nécessaires dans votre projet >> %OUTPUT_DIR%\README.md
echo 3. Référencez les dépendances requises >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo ## Documentation >> %OUTPUT_DIR%\README.md
echo. >> %OUTPUT_DIR%\README.md
echo Consultez le répertoire 'docs' pour la documentation complète. >> %OUTPUT_DIR%\README.md

REM --- Création de l'archive ZIP ---
echo [INFO] Création de l'archive ZIP...
powershell -Command "& {if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) { Compress-Archive -Path '%OUTPUT_DIR%\*' -DestinationPath '%RELEASE_DIR%\ApexFramework_v%VERSION%.zip' -Force } else { Add-Type -Assembly 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::CreateFromDirectory('%OUTPUT_DIR%', '%RELEASE_DIR%\ApexFramework_v%VERSION%.zip') }}"

if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] Échec de la création de l'archive ZIP
    echo [ERREUR] Échec de la création de l'archive ZIP >> %LOG_FILE%
    set /A ERROR_COUNT+=1
)

REM --- Rapport final ---
echo [INFO] Finalisation...
echo Fin de la construction: %date% %time% >> %LOG_FILE%

if %ERROR_COUNT% GTR 0 (
    echo [ATTENTION] La construction s'est terminée avec %ERROR_COUNT% erreurs.
    echo Consultez le fichier %LOG_FILE% pour plus de détails.
) else (
    echo [SUCCÈS] La construction s'est terminée avec succès.
    echo Release disponible dans %RELEASE_DIR%\ApexFramework_v%VERSION%.zip
)

echo.
echo ===== FIN DE LA CONSTRUCTION =====
goto :end

:error
echo [ERREUR] La construction a échoué. Consultez le fichier %LOG_FILE% pour plus de détails.
exit /B 1

:end
exit /B 0
