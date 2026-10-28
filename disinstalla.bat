@echo off
title BatchPrinter - disinstallazione
echo Rimozione collegamenti e voci menu contestuale...
del "%APPDATA%\Microsoft\Windows\SendTo\BatchPrinter.lnk" >nul 2>&1
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\BatchPrinter.lnk" >nul 2>&1
set "PSREG=%TEMP%\bp_regdel_%RANDOM%.ps1"
> "%PSREG%" echo [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree('Software\Classes\*\shell\BatchPrinter', $false)
>> "%PSREG%" echo [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree('Software\Classes\Directory\shell\BatchPrinter', $false)
powershell -NoProfile -ExecutionPolicy Bypass -File "%PSREG%" >nul 2>&1
del "%PSREG%" >nul 2>&1
echo Rimozione programma...
rd /s /q "%LOCALAPPDATA%\BatchPrinter" >nul 2>&1
echo Fatto. Impostazioni e log restano in %APPDATA%\BatchPrinter
echo (cancellala a mano se non ti serve piu').
pause
