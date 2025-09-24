# IgorKrup

A Windows automation and utility library for .NET, providing COM-accessible controls for UI automation, PDF manipulation, video recording, and keyboard input injection. 

## How to install
- Add RegKrupReg.vbs (https://github.com/igorkrupitsky/IgorKrup/blob/main/vbs/RegKrupReg.vbs) source in your VBS file.
- Run the following command.  This command will download and register IgorKrup.dll and the depended DLLs.

```vbscript
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
InstallIgorKrup
```

## Features

- **UI Automation**: Control windows, send keystrokes, mouse events, and automate UI tasks (see `Control.vb`).
- **PDF Tools**: Extract pages and count pages in PDF files using iTextSharp (see `PDF.vb`).
- **Video Recording**: Record desktop or window video using FFmpeg (see `VideoRecorder.vb`).
- **Keyboard Injection**: Low-level keyboard input via a native DLL (`Keyboard.dll`).
- **Edge Browser Automation**: Automate Microsoft Edge browser using a built-in WebDriver client (see `EdgeDriver.vb`).
- **COM Registration**: Register .NET classes for COM automation via `RegKrupReg.vbs`.

## EdgeDriver (Web Automation)

`EdgeDriver.vb` provides a COM-accessible automation interface for Microsoft Edge, similar to SeleniumBasic but with local-user registration support. It can:

- Launch and control Edge browser sessions (headless or visible)
- Navigate to URLs, fill forms, click elements, and execute JavaScript
- Find elements by CSS, XPath, ID, name, tag, class, and link text
- Take screenshots (full page or element)
- Download and update the correct version of `msedgedriver.exe` automatically
- Use Chrome DevTools Protocol (CDP) for advanced features (network, performance, device emulation, etc.)
- Upload files to file inputs
- Handle alerts, cookies, and browser windows/tabs

### Usage Example (VBScript)

```vbscript
Set edge = CreateObject("IgorKrup.EdgeDriver")
edge.UpdateDriver
edge.GetUrl "https://github.com/igorkrupitsky/IgorKrup"

elementId = edge.FindElementByCss("input[aria-label='Go to file']")

Do While edge.IsElementDisplayed(elementId) = False
    WScript.Sleep 100
Loop

Do While edge.IsElementEnabled(elementId) = False
    WScript.Sleep 100
Loop

edge.ClickElement elementId 

WScript.Sleep 1000
edge.SendKeysToElement elementId, "Hello world!"

MsgBox "Done!"
```

### Notes
- Place the correct `msedgedriver.exe` in your output directory, or use `UpdateDriver()` to download it automatically.
- EdgeDriver supports both visible and headless modes (set `useHeadless` property).
- For advanced automation, use methods like `ExecuteScript`, `SendCdpCommand`, and element search helpers.

## Project Structure

- `IgorKrup.sln` / `IgorKrup.vbproj`: Main solution and project files.
- `Control.vb`: UI automation and window control functions.
- `PDF.vb`: PDF page extraction and page count.
- `VideoRecorder.vb`: Desktop/window video recording using FFmpeg.
- `Module1.vb`: Keyboard injection interop.
- `Keyboard/Keyboard.cpp`: Native DLL for keyboard scan code injection.
- `RegKrupReg.vbs`: Script for registering COM classes in the registry.
- `bin/` and `obj/`: Build output directories.
- `IgorKrupTest/`: Test project with a sample form.

## Requirements

- .NET Framework 3.5  
	(This is the only framework version that allows COM registration of the library for the local user without requiring admin rights.)
- [iTextSharp](https://github.com/itext/itextsharp) (for PDF features)
- [ICSharpCode.SharpZipLib](https://github.com/icsharpcode/SharpZipLib) (dependency)
- [FFmpeg](https://ffmpeg.org/) (for video recording; place `ffmpeg.exe` in the output directory)
- Windows OS

## Building

1. Open `IgorKrup.sln` in Visual Studio.
2. Build the solution (ensure all dependencies are present in `bin/Debug` or `bin/Release`).
3. Build the native `Keyboard.dll` using Visual Studio's C++ tools if needed.

## Registering for COM

To use the library via COM (e.g., from VBScript):

1. Run `RegKrupReg.vbs` (edit the script to set the correct remote and local paths if needed).
2. This will copy required DLLs and register the COM classes in the registry under `HKCU`.

## Example Usage (VBScript)

```vbscript
Set pdf = CreateObject("IgorKrup.PDF")
WScript.Echo "Pages: " & pdf.PageCount("C:\temp\test.pdf")
pdf.ExtractPage "C:\temp\test.pdf", "C:\temp\page1.pdf", 1
```

## Notes

- For video recording, ensure `ffmpeg.exe` is present in the application directory.
- For keyboard injection, `Keyboard.dll` must be built and present in the output directory.
- All COM classes are registered per-user (HKCU) and do not require admin rights.

## License

MIT License (see source code headers for details)
