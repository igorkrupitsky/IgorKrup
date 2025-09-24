Set fso = CreateObject("Scripting.FileSystemObject")

fso.CopyFile "C:\Users\80014379\source\repos\IgorKrup\bin\Debug\IgorKrup.dll",    "\\pwas0038\download\IgorKrup\IgorKrup.dll"
fso.CopyFile "C:\Users\80014379\source\repos\IgorKrup\bin\Debug\IgorKrup.pdb",    "\\pwas0038\download\IgorKrup\IgorKrup.pdb"

MsgBox "Done"

'Keyboard.dll
'itextsharp.dll
'ICSharpCode.SharpZipLib.dll
