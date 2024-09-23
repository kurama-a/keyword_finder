@echo off
:: Télécharger et installer Poppler
echo Téléchargement de Poppler...
curl -L -o C:\Temp\poppler.zip https://github.com/oschwartz10612/poppler-windows/releases/download/v24.07.0-0/Release-24.07.0-0.zip

echo Extraction de Poppler...
powershell -Command "Expand-Archive -Path 'C:\Temp\poppler.zip' -DestinationPath 'C:\Program Files\poppler'"

echo Ajout de Poppler au PATH...
setx PATH "%PATH%;C:\Program Files\poppler\Library\bin"

echo Poppler installé avec succès !
