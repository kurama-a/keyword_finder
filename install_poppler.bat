@echo off
:: Définir le chemin temporaire et le dossier d'extraction
set "TEMP_DIR=%~dp0"
set "POPPLER_DIR=C:\poppler-24.07.0"

:: Vérifier si le fichier ZIP de Poppler est présent dans le dossier du script
if not exist "%TEMP_DIR%\Release-24.07.0-0.zip" (
    echo Fichier Poppler ZIP non trouvé dans le dossier du script.
    exit /b 1
)

:: Extraire Poppler
echo Extraction de Poppler...
powershell -Command "Expand-Archive -Path '%TEMP_DIR%\Release-24.07.0-0.zip' -DestinationPath 'C:\'"

:: Vérifier si l'extraction a réussi
if not exist "%POPPLER_DIR%" (
    echo Échec de l'extraction de Poppler.
    exit /b 1
)

:: Ajouter Poppler au PATH si ce n'est pas déjà présent
echo Ajout de Poppler au PATH...
set "NEW_PATH=%POPPLER_DIR%\Library\bin"

:: Vérifier si le chemin de Poppler est déjà dans le PATH
echo %PATH% | find /i "%NEW_PATH%" >nul
if errorlevel 1 (
    :: Poppler n'est pas dans le PATH, donc on l'ajoute
    setx PATH "%PATH%;%NEW_PATH%"
    echo Poppler ajouté au PATH.
) else (
    echo Poppler est déjà dans le PATH.
)

echo Poppler installé avec succès !

:: Nettoyage des fichiers temporaires
echo Nettoyage des fichiers temporaires...
del "%TEMP_DIR%\Release-24.07.0-0.zip"

echo Installation terminée.
pause
