@echo off
echo Installing Picture Pros Folder Script as Windows Service...

REM Download NSSM if not present
if not exist "nssm.exe" (
    echo Downloading NSSM...
    powershell -Command "Invoke-WebRequest -Uri 'https://nssm.cc/release/nssm-2.24.zip' -OutFile 'nssm.zip'"
    powershell -Command "Expand-Archive -Path 'nssm.zip' -DestinationPath '.' -Force"
    copy "nssm-2.24\win64\nssm.exe" "nssm.exe"
    rmdir /s "nssm-2.24"
    del "nssm.zip"
)

REM Install the service
nssm.exe install "PictureProsFolderScript" "python.exe" "%~dp0picture_pros_folder_script.py"
nssm.exe set "PictureProsFolderScript" AppDirectory "%~dp0"
nssm.exe set "PictureProsFolderScript" Description "Picture Pros Folder Script - Watches for photo/label files and moves them to printer folders"
nssm.exe set "PictureProsFolderScript" Start SERVICE_AUTO_START

echo Service installed successfully!
echo To start the service: net start PictureProsFolderScript
echo To stop the service: net stop PictureProsFolderScript
pause
