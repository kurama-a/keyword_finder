@echo off
:: Définir le chemin temporaire et le dossier d'extraction
set "TEMP_DIR=C:\Temp"
set "POPPLER_DIR=C:\poppler-24.07.0"

:: Créer le répertoire temporaire s'il n'existe pas
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

:: Télécharger Poppler
echo Téléchargement de Poppler...
curl -L -o "%TEMP_DIR%\poppler.zip" https://github.com/oschwartz10612/poppler-windows/releases/download/v24.07.0-0/Release-24.07.0-0.zip

:: Vérifier si le téléchargement a réussi
if not exist "%TEMP_DIR%\poppler.zip" (
    echo Échec du téléchargement de Poppler.
    exit /b 1
)

:: Extraire Poppler
echo Extraction de Poppler...
powershell -Command "Expand-Archive -Path '%TEMP_DIR%\poppler.zip' -DestinationPath 'C:\'"

:: Vérifier si l'extraction a réussi
if not exist "%POPPLER_DIR%" (
    echo Échec de l'extraction de Poppler.
    exit /b 1
)

:: Ajouter Poppler au PATH
echo Ajout de Poppler au PATH...
set "NEW_PATH=%POPPLER_DIR%\Library\bin"
setx PATH "%PATH%;%NEW_PATH%"

:: Vérifier si l'ajout au PATH a réussi
if errorlevel 1 (
    echo Erreur lors de l'ajout de Poppler au PATH.
    exit /b 1
)

echo Poppler installé avec succès !

:: Nettoyage des fichiers temporaires
echo Nettoyage des fichiers temporaires...
del "%TEMP_DIR%\poppler.zip"

echo Installation terminée.
pause
