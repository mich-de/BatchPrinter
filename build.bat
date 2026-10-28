@echo off
set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist %CSC% set CSC=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
set SR=C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll
set SRW=C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime.WindowsRuntime\v4.0_4.0.0.0__b77a5c561934e089\System.Runtime.WindowsRuntime.dll
set WM=C:\Windows\System32\WinMetadata
%CSC% /nologo /target:winexe /optimize+ /win32icon:"%~dp0BatchPrinter.ico" /out:BatchPrinter.exe /r:"%SR%" /r:"%SRW%" /r:"%WM%\Windows.Data.winmd" /r:"%WM%\Windows.Foundation.winmd" /r:"%WM%\Windows.Storage.winmd" /r:"%WM%\Windows.UI.winmd" /r:System.Windows.Forms.dll /r:System.Drawing.dll BatchPrinter.cs
if exist BatchPrinter.exe (echo Build OK: BatchPrinter.exe) else (echo Build FAILED)
pause
