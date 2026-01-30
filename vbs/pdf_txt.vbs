Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim oPdf: Set oPdf = CreateObject("IgorKrup.PDF")

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
    'sInFolderPath = "C:\Users\80014379\Desktop\Protocol_Docs\PDF_txt"
End If

If fso.FolderExists(sInFolderPath) = False Then
    WScript.Echo "Folder does not exist: " & sInFolderPath
    WScript.Quit
End If

Dim sOutFolderPath: sOutFolderPath = fso.GetFolder(sInFolderPath).ParentFolder.Path & "\PDF_txt"
If fso.FolderExists(sOutFolderPath) = False Then
    fso.CreateFolder sOutFolderPath   
End If

MsgBox "Output Folder: " & sOutFolderPath

Dim oFolder: Set oFolder = fso.GetFolder(sInFolderPath)
For Each oFile In oFolder.Files
    sOutFilePath = sOutFolderPath & "\" & fso.GetBaseName(oFile.Name) & ".md"
    If fso.FileExists(sOutFilePath) = False Then
        
        on error resume next
        sText = oPdf.GetFileText(oFile.Path)
        If Err.number <> 0 Then
            MsgBox oFile.Name & ", " & Err.Description
            sText = ""
        End If
        on error goto 0

        If Len(trim(sText)) > 5 Then
            Set oOutFile = fso.CreateTextFile(sOutFilePath,True,True)
            oOutFile.Write sText
            oOutFile.Close
            Set oOutFile = Nothing
        End If
    End If
Next

MsgBox "Done"

