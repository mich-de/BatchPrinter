# BatchPrinter

A free batch-printing tool for Windows, in the spirit of Print Conductor. Build a queue of PDFs, images and Office documents, pick a printer, and print everything in one go.

Single C# source file, no external dependencies: it compiles with the C# compiler that ships with Windows.

> **Note:** the application interface is currently in Italian.

## Features

- Queue with drag & drop, reordering and removal
- Add whole folders (recursive scan)
- Any installed printer, copies (1-999), collate, duplex, page range (e.g. `1-3,5`)
- **Print preview** before each document, with "OK, print all" to skip the remaining previews on large queues
- Native PDF rendering via `Windows.Data.Pdf`: no PDF reader needed, no window opens
- Full-page printing (scaled to the printable area of the sheet)
- Single instance: files sent from "Send to" or the right-click menu join the running queue
- Saveable profiles, print log, 8 colour themes (including a dark title bar that follows the theme)

## Requirements

- Windows 10 or 11 (PDF printing uses the built-in `Windows.Data.Pdf`; on older systems only images and Office documents work)
- .NET Framework 4.8 (already included in Windows 10 1903+ and Windows 11)
- Office documents (docx, xlsx, pptx, ...) need Microsoft Office or LibreOffice installed, because they are printed through the associated application
- An installed printer with its driver

## Install

**From a release:** download `BatchPrinter.exe` from the [Releases](../../releases) page and run it. Windows SmartScreen may warn because the file is not code-signed.

**From source:** run `install.bat`. It compiles the source with the Windows built-in compiler, installs to `%LOCALAPPDATA%\BatchPrinter`, creates the Start menu and "Send to" shortcuts and, if you accept, adds a right-click entry "Stampa con BatchPrinter" for files and folders. Run `disinstalla.bat` to remove everything.

## Build

Double-click `build.bat`, or run:

```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo /target:winexe /optimize+ /win32icon:BatchPrinter.ico /out:BatchPrinter.exe ^
  /r:"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll" ^
  /r:"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime.WindowsRuntime\v4.0_4.0.0.0__b77a5c561934e089\System.Runtime.WindowsRuntime.dll" ^
  /r:C:\Windows\System32\WinMetadata\Windows.Data.winmd ^
  /r:C:\Windows\System32\WinMetadata\Windows.Foundation.winmd ^
  /r:C:\Windows\System32\WinMetadata\Windows.Storage.winmd ^
  /r:C:\Windows\System32\WinMetadata\Windows.UI.winmd ^
  /r:System.Windows.Forms.dll /r:System.Drawing.dll BatchPrinter.cs
```

The references are not optional: a plain `csc` call fails. On 32-bit systems use `Framework` instead of `Framework64`.

## How each file type is printed

| Type | Method |
| --- | --- |
| PDF | In-process rendering with `Windows.Data.Pdf`, then silent GDI printing. Copies, collate, duplex and page range all work. |
| Images (png, jpg, bmp, gif, multi-page tiff, emf, wmf) | Native GDI printing with all options. |
| Word, Excel, PowerPoint, text, others | Shell `printto` / `print` verb through the associated application. No preview; copies and advanced settings are handled by that application. |

## Data and logs

User data lives in `%APPDATA%\BatchPrinter\`: `impostazioni.txt` (settings), `profilo.bpr` (profile), `printlog.txt` (print log) and `errorlog.txt` (diagnostics).

## License

[MIT](LICENSE)
