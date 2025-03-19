@echo off
:: Définir le dossier où se trouve le script
set "TEMP_DIR=%~dp0"
set "POPPLER_DIR=C:\poppler-24.07.0"
set "ZIP_FILE=Release-24.07.0-0.zip"

:: Vérifier si le fichier ZIP de Poppler est présent dans le dossier du script
if not exist "%TEMP_DIR%\%ZIP_FILE%" (
    echo Fichier ZIP de Poppler non trouvé dans le dossier du script.
    exit /b 1
)

:: Vérifier si Poppler est déjà installé
if exist "%POPPLER_DIR%" (
    echo Poppler semble déjà être installé dans %POPPLER_DIR%.
) else (
    :: Extraire Poppler
    echo Extraction de Poppler...
    powershell -Command "Expand-Archive -Path '%TEMP_DIR%\%ZIP_FILE%' -DestinationPath 'C:\'"
    
    :: Vérifier si l'extraction a réussi
    if not exist "%POPPLER_DIR%" (
        echo Échec de l'extraction de Poppler.
        exit /b 1
    ) else (
        echo Poppler extrait avec succès !
    )
)

:: Définir le chemin de la bibliothèque Poppler
set "NEW_PATH=%POPPLER_DIR%\Library\bin"

:: Vérifier si Poppler est déjà dans le PATH de l'utilisateur
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USER_PATH=%%B"

echo %USER_PATH% | find /i "%NEW_PATH%" >nul
if errorlevel 1 (
    echo Ajout de Poppler au PATH de l'utilisateur...
    setx PATH "%USER_PATH%;%NEW_PATH%"
    echo Poppler a été ajouté au PATH de l'utilisateur.
) else (
    echo Poppler est déjà dans le PATH de l'utilisateur.
)

echo Installation terminée avec succès !
pause
