# IgorKrup

A Windows automation and utility library for .NET, providing COM-accessible controls for UI automation, PDF manipulation, video recording, and keyboard input injection. 

## Features

- **UI Automation**: Control windows, send keystrokes, mouse events, and automate UI tasks (see `Control.vb`).
- **PDF Tools**: Extract pages and count pages in PDF files using iTextSharp (see `PDF.vb`).
- **Video Recording**: Record desktop or window video using FFmpeg (see `VideoRecorder.vb`).
- **Keyboard Injection**: Low-level keyboard input via a native DLL (`Keyboard.dll`).
- **COM Registration**: Register .NET classes for COM automation via `RegKrupReg.vbs`.

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
WScript.Echo "Pages: " & pdf.PageCount("C:\\test.pdf")
pdf.ExtractPage "C:\\test.pdf", "C:\\page1.pdf", 1

Set ctrl = CreateObject("IgorKrup.Control")
ctrl.Run "notepad.exe"
```

## Notes

- For video recording, ensure `ffmpeg.exe` is present in the application directory.
- For keyboard injection, `Keyboard.dll` must be built and present in the output directory.
- All COM classes are registered per-user (HKCU) and do not require admin rights.

## License

MIT License (see source code headers for details)
