Set fso = CreateObject("Scripting.FileSystemObject") 'markdown_pages=>markdowns

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
    'sInFolderPath = "C:\Users\80014379\Desktop\Financials_Contracts\markdown_pages"
End If

If fso.FolderExists(sInFolderPath) = False Then
    WScript.Echo "Folder does not exist: " & sInFolderPath
    WScript.Quit
End If

Dim sOutFolderPath: sOutFolderPath = fso.GetFolder(sInFolderPath).ParentFolder.Path & "\markdowns"
If fso.FolderExists(sOutFolderPath) = False Then
    fso.CreateFolder sOutFolderPath   
End If

Dim sOldMcc: sOldMcc = ""
Dim oOutFile

Set oFolder = fso.GetFolder(sInFolderPath)
For Each oFile In oFolder.Files
    If fso.GetExtensionName(oFile.Name) = "md" Then    
        sBaseName = fso.GetBaseName(oFile.Name)
        oBaseName = Split(sBaseName, "_")
        sMcc = oBaseName(0)
        sFileNumber = oBaseName(1)

        sOutFilePath  = sOutFolderPath & "\" & sMcc & ".md"

        If sOldMcc <> sMcc Then
            If sOldMcc <> "" Then oOutFile.Close
            Set oOutFile = fso.CreateTextFile(sOutFilePath, True, True)
        Else
            oOutFile.WriteLine "___"
        End If
        
        oOutFile.WriteLine "## Page " & Cint(sFileNumber)
        oOutFile.WriteLine ""
        oOutFile.Write ReadTextFile(oFile.Path)
        
        sOldMcc = sMcc
    End If 
Next

oOutFile.Close
MsgBox "Done"

Function ReadTextFile(sPath)
    Dim sRet: sRet = ""
    Dim file: Set file = fso.OpenTextFile(sPath, 1, False, -2)
    Do Until file.AtEndOfStream
        sRet = sRet & file.ReadLine & vbCrLf
    Loop
    file.Close
    Set file = Nothing
    ReadTextFile = sRet
End Function


