@echo off
setlocal
title Add Python profile to Windows Terminal

if /I "%~1"=="/remove" goto remove

where wt.exe >nul 2>nul
if errorlevel 1 (
    echo Windows Terminal was not found.
    echo Install or update Windows Terminal, then run this file again.
    pause
    exit /b 1
)

where powershell.exe >nul 2>nul
if errorlevel 1 (
    echo Windows PowerShell was not found.
    pause
    exit /b 1
)

echo Adding a permanent Python profile to the Windows Terminal new-tab menu...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $launcher=Get-Command py.exe -ErrorAction SilentlyContinue; if($launcher){$command=([char]34)+$launcher.Source+([char]34)+' -3 -i'}else{$launcher=Get-Command python.exe -ErrorAction SilentlyContinue; if(-not $launcher){throw 'Python 3 was not found'}; $command=([char]34)+$launcher.Source+([char]34)+' -i'}; $folder=Join-Path $env:LOCALAPPDATA 'Microsoft\Windows Terminal\Fragments\AddressMend'; New-Item -ItemType Directory -Path $folder -Force | Out-Null; $file=Join-Path $folder 'python.json'; $profile=[ordered]@{guid='{b5a51a23-8d75-55bc-bda8-9d2c7e92e8c5}'; name='Python'; commandline=$command; startingDirectory=$env:USERPROFILE; hidden=$false}; $json=([ordered]@{profiles=@($profile)} | ConvertTo-Json -Depth 5); [System.IO.File]::WriteAllText($file,$json,(New-Object System.Text.UTF8Encoding($false))); Write-Host ('Profile file: '+$file)"
if errorlevel 1 (
    echo.
    echo The Python profile could not be added.
    pause
    exit /b 1
)

echo.
echo Python has been added to the Windows Terminal new-tab menu.
echo Close every Windows Terminal window and reopen it to see the profile.
echo To remove it later, run: %~nx0 /remove
pause
exit /b 0

:remove
echo Removing the AddressMend Python profile from Windows Terminal...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$file=Join-Path $env:LOCALAPPDATA 'Microsoft\Windows Terminal\Fragments\AddressMend\python.json'; if(Test-Path -LiteralPath $file){Remove-Item -LiteralPath $file -Force; Write-Host 'Python profile removed.'}else{Write-Host 'The AddressMend Python profile was not installed.'}"
if errorlevel 1 (
    echo The Python profile could not be removed.
    pause
    exit /b 1
)
echo Close every Windows Terminal window and reopen it to refresh the menu.
pause
exit /b 0
