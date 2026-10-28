@echo off
setlocal EnableDelayedExpansion
title BatchPrinter - installazione

echo ============================================
echo   BATCHPRINTER - INSTALLAZIONE
echo ============================================
echo.

rem --- 1. Verifica Windows 10/11 (motore PDF Windows.Data.Pdf)
echo [1/4] Verifica versione Windows...
ver | findstr /i "10\.0\." >nul
if errorlevel 1 (
    echo   ATTENZIONE: serve Windows 10 o 11 per la stampa PDF integrata.
    echo   Il programma funziona comunque per immagini e documenti Office.
) else (
    echo   OK: Windows 10/11 rilevato.
)

rem --- 2. Verifica .NET Framework 4.8 (compilatore integrato, nessun download)
echo [2/4] Verifica .NET Framework 4.8...
set "CSC="
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not defined CSC (
    echo   ERRORE: .NET Framework 4.8 non trovato.
    echo   Scaricalo da: https://dotnet.microsoft.com/download/dotnet-framework/net48
    pause
    exit /b 1
)
echo   OK: compilatore trovato.

rem --- 3. Compilazione e installazione in %LOCALAPPDATA%\BatchPrinter
echo [3/4] Compilazione...
set "DEST=%LOCALAPPDATA%\BatchPrinter"
if not exist "%DEST%" mkdir "%DEST%"

set "SR=%WINDIR%\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll"
set "SRW=%WINDIR%\Microsoft.NET\assembly\GAC_MSIL\System.Runtime.WindowsRuntime\v4.0_4.0.0.0__b77a5c561934e089\System.Runtime.WindowsRuntime.dll"
set "WM=%WINDIR%\System32\WinMetadata"

"%CSC%" /nologo /target:winexe /optimize+ /win32icon:"%~dp0BatchPrinter.ico" /out:"%DEST%\BatchPrinter.exe" ^
  /r:"%SR%" /r:"%SRW%" ^
  /r:"%WM%\Windows.Data.winmd" /r:"%WM%\Windows.Foundation.winmd" ^
  /r:"%WM%\Windows.Storage.winmd" /r:"%WM%\Windows.UI.winmd" ^
  /r:System.Windows.Forms.dll /r:System.Drawing.dll "%~dp0BatchPrinter.cs"
if errorlevel 1 (
    echo   ERRORE di compilazione.
    pause
    exit /b 1
)
copy /y "%~dp0README.txt" "%DEST%\README.txt" >nul
echo   OK: installato in %DEST%

rem --- 4. Collegamenti: menu Start + "Invia a" + menu contestuale (tasto dx)
echo [4/4] Creazione collegamenti...
set "VBS=%TEMP%\bp_shortcut_%RANDOM%.vbs"
> "%VBS%" echo Set sh = CreateObject("WScript.Shell")
>> "%VBS%" echo p = sh.SpecialFolders("StartMenu") ^& "\Programs\BatchPrinter.lnk"
>> "%VBS%" echo Set lnk = sh.CreateShortcut(p)
>> "%VBS%" echo lnk.TargetPath = "%DEST%\BatchPrinter.exe"
>> "%VBS%" echo lnk.WorkingDirectory = "%DEST%"
>> "%VBS%" echo lnk.Description = "Stampa in blocco"
>> "%VBS%" echo lnk.Save
>> "%VBS%" echo s = sh.SpecialFolders("SendTo") ^& "\BatchPrinter.lnk"
>> "%VBS%" echo Set l2 = sh.CreateShortcut(s)
>> "%VBS%" echo l2.TargetPath = "%DEST%\BatchPrinter.exe"
>> "%VBS%" echo l2.WorkingDirectory = "%DEST%"
>> "%VBS%" echo l2.Description = "Aggiungi alla coda di stampa BatchPrinter"
>> "%VBS%" echo l2.Save
cscript //nologo "%VBS%" >nul 2>&1
del "%VBS%" >nul 2>&1
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\BatchPrinter.lnk" (
    echo   OK: menu Start -^> BatchPrinter
) else (
    echo   Nota: collegamento menu Start non creato ^(usa l'exe direttamente^).
)
if exist "%APPDATA%\Microsoft\Windows\SendTo\BatchPrinter.lnk" (
    echo   OK: tasto destro -^> Invia a -^> BatchPrinter
) else (
    echo   Nota: collegamento Invia a non creato.
)
echo.
choice /c SN /m "Aggiungere anche 'Stampa con BatchPrinter' nel menu contestuale (tasto destro su file e cartelle)"
if errorlevel 2 goto fine
set "PSREG=%TEMP%\bp_reg_%RANDOM%.ps1"
> "%PSREG%" echo $exe = '%DEST%\BatchPrinter.exe'
>> "%PSREG%" echo [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\Classes\*\shell\BatchPrinter\command').Close()
>> "%PSREG%" echo $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\*\shell\BatchPrinter', $true)
>> "%PSREG%" echo $k.SetValue('', 'Stampa con BatchPrinter'); $k.SetValue('Icon', $exe); $k.Close()
>> "%PSREG%" echo $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\*\shell\BatchPrinter\command', $true)
>> "%PSREG%" echo $k.SetValue('', ('"' + $exe + '" "%1"')); $k.Close()
>> "%PSREG%" echo [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\Classes\Directory\shell\BatchPrinter\command').Close()
>> "%PSREG%" echo $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\Directory\shell\BatchPrinter', $true)
>> "%PSREG%" echo $k.SetValue('', 'Stampa contenuto con BatchPrinter'); $k.SetValue('Icon', $exe); $k.Close()
>> "%PSREG%" echo $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\Directory\shell\BatchPrinter\command', $true)
>> "%PSREG%" echo $k.SetValue('', ('"' + $exe + '" "%1"')); $k.Close()
powershell -NoProfile -ExecutionPolicy Bypass -File "%PSREG%" >nul 2>&1
set "RC=%ERRORLEVEL%"
del "%PSREG%" >nul 2>&1
if "%RC%"=="0" (
    echo   OK: menu contestuale aggiunto (per rimuovere: disinstalla.bat)
) else (
    echo   Nota: menu contestuale non aggiunto ^(%RC%^) - resta comunque "Invia a".
)

echo.
echo ============================================
echo   INSTALLAZIONE COMPLETATA
echo   Avvio: menu Start -^> BatchPrinter
echo   oppure: %DEST%\BatchPrinter.exe
echo ============================================
echo.
choice /c SN /m "Avviare BatchPrinter ora"
if errorlevel 2 goto fine
start "" "%DEST%\BatchPrinter.exe"
:fine
endlocal
