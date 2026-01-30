Dim fso: Set fso = CreateObject("Scripting.FileSystemObject") 'PDF_pages=>pdf_img
Dim shell : Set shell = CreateObject("WScript.Shell")

Dim sGhostscriptPath: sGhostscriptPath = "C:\Igor\GitHub\PdfOcr\Ghostscript\bin\gswin64.exe" 
If fso.FileExists(sGhostscriptPath) = False Then
    sGhostscriptPath = "\\pwdb3030\download\Software\Ghostscript\bin\gswin64.exe"
End If

Dim sInFolderPath: sInFolderPath = ""

If WScript.Arguments.Count = 1 Then    
    If fso.FolderExists(WScript.Arguments(0)) Then
        sInFolderPath = WScript.Arguments(0)
    End If    
End If

If sInFolderPath = "" Then
    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.BrowseForFolder(0, "Select Folder", 0, "")
    If Not oFolder Is Nothing Then
        sInFolderPath = oFolder.Self.Path
    End If
End If

If sInFolderPath = "" Then
    'sInFolderPath = "C:\Users\80014379\Desktop\Protocol_Docs\PDF_pages"
End If

If fso.FolderExists(sInFolderPath) = False Then
    WScript.Echo "Folder does not exist: " & sInFolderPath
    WScript.Quit
End If

Dim sOutFolderPath: sOutFolderPath = fso.GetFolder(sInFolderPath).ParentFolder.Path & "\pdf_img"
If fso.FolderExists(sOutFolderPath) = False Then
    fso.CreateFolder sOutFolderPath   
End If

MsgBox "Output Folder: " & sOutFolderPath

Dim oFolder: Set oFolder = fso.GetFolder(sInFolderPath)
For Each oFile In oFolder.Files
    sImgFilePath = sOutFolderPath & "\" & fso.GetBaseName(oFile.Path) & ".png"

    If fso.GetExtensionName(oFile.Name) = "pdf" And fso.FileExists(sImgFilePath) = False Then
        '-r300 - Print quality
        '-r600- High-quality print/scanning
        shell.run """" & sGhostscriptPath & """ -dNOPAUSE -q -r600 -sDEVICE=png16m -dBATCH -sOutputFile=""" & sImgFilePath & """ """ & oFile.Path & """ -c quit", 0 , True
    End If
Next

MsgBox "Done"