@echo off
echo Creating scheduled task for Picture Pros Folder Script...

REM Create a scheduled task that runs at startup and restarts if it fails
schtasks /create /tn "PictureProsFolderScript" /tr "python.exe \"%~dp0picture_pros_folder_script.py\"" /sc onstart /ru "SYSTEM" /f

echo Scheduled task created successfully!
echo The script will now run automatically when Windows starts.
echo To manually start: schtasks /run /tn "PictureProsFolderScript"
echo To stop: schtasks /end /tn "PictureProsFolderScript"
pause
