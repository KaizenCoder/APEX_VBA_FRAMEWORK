@echo off
REM ==========================================================================
REM Script : GenerateDocs.bat
REM Version : 1.0
REM Purpose : Génération de la documentation du framework APEX
REM Date : 10/04/2025
REM ==========================================================================

echo ===== GÉNÉRATION DE LA DOCUMENTATION APEX FRAMEWORK =====
echo.

REM --- Variables ---
set DOCS_DIR=docs
set OUTPUT_DIR=docs_html
set TEMP_DIR=docs_temp
set LOG_FILE=docs_gen_log.txt

REM --- Vérification de l'installation de pandoc ---
echo [INFO] Vérification des dépendances...
where pandoc >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] Pandoc n'est pas installé ou n'est pas dans le PATH.
    echo Veuillez installer Pandoc depuis https://pandoc.org/installing.html
    goto :EOF
)

REM --- Initialisation ---
echo [INFO] Initialisation de la génération de documentation...
if exist %LOG_FILE% del %LOG_FILE%
echo Début de la génération: %date% %time% > %LOG_FILE%

REM --- Création des répertoires ---
echo [INFO] Création des répertoires de sortie...
if not exist %OUTPUT_DIR% mkdir %OUTPUT_DIR%
if exist %TEMP_DIR% rmdir /S /Q %TEMP_DIR%
mkdir %TEMP_DIR%

REM --- Copie des fichiers Markdown ---
echo [INFO] Préparation des fichiers Markdown...
xcopy /E /Y %DOCS_DIR%\*.md %TEMP_DIR%\ >> %LOG_FILE%

REM --- Génération du fichier d'index ---
echo [INFO] Génération du fichier d'index...
echo ^<!DOCTYPE html^> > %TEMP_DIR%\index.html
echo ^<html^> >> %TEMP_DIR%\index.html
echo ^<head^> >> %TEMP_DIR%\index.html
echo ^<meta charset="UTF-8"^> >> %TEMP_DIR%\index.html
echo ^<title^>Documentation APEX Framework^</title^> >> %TEMP_DIR%\index.html
echo ^<style^> >> %TEMP_DIR%\index.html
echo body { font-family: Arial, sans-serif; max-width: 1000px; margin: 0 auto; padding: 20px; } >> %TEMP_DIR%\index.html
echo h1 { color: #2c3e50; } >> %TEMP_DIR%\index.html
echo ul { list-style-type: none; padding: 0; } >> %TEMP_DIR%\index.html
echo li { margin: 10px 0; } >> %TEMP_DIR%\index.html
echo a { text-decoration: none; color: #3498db; } >> %TEMP_DIR%\index.html
echo a:hover { text-decoration: underline; } >> %TEMP_DIR%\index.html
echo .category { margin-top: 30px; } >> %TEMP_DIR%\index.html
echo ^</style^> >> %TEMP_DIR%\index.html
echo ^</head^> >> %TEMP_DIR%\index.html
echo ^<body^> >> %TEMP_DIR%\index.html
echo ^<h1^>Documentation APEX Framework^</h1^> >> %TEMP_DIR%\index.html
echo ^<p^>Bienvenue dans la documentation du Framework APEX. Sélectionnez un document ci-dessous pour commencer.^</p^> >> %TEMP_DIR%\index.html

echo ^<div class="category"^> >> %TEMP_DIR%\index.html
echo ^<h2^>Guides^</h2^> >> %TEMP_DIR%\index.html
echo ^<ul^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="README.html"^>README - Présentation générale^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="QuickStartGuide.html"^>Guide de démarrage rapide^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="CHANGELOG.html"^>Journal des modifications^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^</ul^> >> %TEMP_DIR%\index.html
echo ^</div^> >> %TEMP_DIR%\index.html

echo ^<div class="category"^> >> %TEMP_DIR%\index.html
echo ^<h2^>Tutoriels^</h2^> >> %TEMP_DIR%\index.html
echo ^<ul^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="RecipeComparer_Tutorial.html"^>Tutoriel: Module de recette^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="Logger_Structure.html"^>Structure du système de journalisation^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^<li^>^<a href="TestFramework_Overview.html"^>Aperçu du framework de test^</a^>^</li^> >> %TEMP_DIR%\index.html
echo ^</ul^> >> %TEMP_DIR%\index.html
echo ^</div^> >> %TEMP_DIR%\index.html

echo ^</body^> >> %TEMP_DIR%\index.html
echo ^</html^> >> %TEMP_DIR%\index.html

REM --- Conversion des fichiers Markdown en HTML ---
echo [INFO] Conversion des fichiers Markdown en HTML...
for %%f in (%TEMP_DIR%\*.md) do (
    echo Traitement de %%~nxf...
    pandoc -f markdown -t html "%%f" -o "%OUTPUT_DIR%\%%~nf.html" --standalone --metadata title="APEX Framework - %%~nf" 2>> %LOG_FILE%
)

REM --- Copie du fichier d'index ---
echo [INFO] Copie du fichier d'index...
copy %TEMP_DIR%\index.html %OUTPUT_DIR%\index.html >> %LOG_FILE%

REM --- Nettoyage ---
echo [INFO] Nettoyage des fichiers temporaires...
rmdir /S /Q %TEMP_DIR%

REM --- Rapport final ---
echo [INFO] Finalisation...
echo Fin de la génération: %date% %time% >> %LOG_FILE%
echo [SUCCÈS] La documentation a été générée dans le répertoire %OUTPUT_DIR%.
echo Ouvrez %OUTPUT_DIR%\index.html pour consulter la documentation.

echo.
echo ===== FIN DE LA GÉNÉRATION DE DOCUMENTATION =====
