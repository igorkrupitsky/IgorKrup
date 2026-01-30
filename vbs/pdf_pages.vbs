Dim fso: Set fso = CreateObject("Scripting.FileSystemObject") 'PDFs=>pdf_pages
Dim oPdf: Set oPdf = CreateObject("IgorKrup.PDF")
 
Dim sInFolderPath: sInFolderPath = ""

If WScript.Arguments.Count = 1 Then   
    sParam = WScript.Arguments(0)
    If fso.FolderExists(sParam) Then
        sInFolderPath = sParam
    ElseIf fso.FileExists(sParam) Then
        BreakePdfToPages sParam, fso.GetFile(sParam).ParentFolder.Path
        WScript.Quit
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
    'sInFolderPath = "C:\Users\80014379\Desktop\Protocol_Docs\PDFs"
End If

If fso.FolderExists(sInFolderPath) = False Then
    WScript.Echo "Folder does not exist: " & sInFolderPath
    WScript.Quit
End If

Dim sOutFolderPath: sOutFolderPath = fso.GetFolder(sInFolderPath).ParentFolder.Path & "\pdf_pages"
If fso.FolderExists(sOutFolderPath) = False Then
    fso.CreateFolder sOutFolderPath   
End If

MsgBox "Output Folder: " & sOutFolderPath

Dim oFolder: Set oFolder = fso.GetFolder(sInFolderPath)
For Each oFile In oFolder.Files
    If fso.GetExtensionName(oFile.Path) = "pdf" Then    
        BreakePdfToPages oFile.Path, sOutFolderPath
    End If
Next

MsgBox "Done"

'===============================
Sub BreakePdfToPages(sFilePath, sOutFolderPath)
    Dim iPageCount: iPageCount = 0
    on error resume next
    iPageCount = oPdf.PageCount(sFilePath)
    If Err.number <> 0 Then
        MsgBox Err.Description & ", PageCount: " & sFilePath
    End If
    on error goto 0

    If iPageCount = 0 Then
        Exit Sub
    End If

    If iPageCount = 1 Then
        Dim sDestFile: sDestFile = sOutFolderPath & "\" & fso.GetBaseName(sFilePath) & "_001.pdf"
        If fso.FileExists(sDestFile) = False Then
            fso.CopyFile sFilePath, sDestFile
        End If
        Exit Sub
    End If

    Dim iPage, sOutputFile
    For iPage = 1 to iPageCount
        sOutputFile = sOutFolderPath & "\" & fso.GetBaseName(sFilePath) & "_" & Right("000" & iPage, 3) & ".pdf"

        on error resume next
        oPdf.ExtractPage sFilePath, sOutputFile, iPage
        If Err.number <> 0 Then
            MsgBox Err.Description & ", ExtractPage(" & iPage & "): " & sFilePath
        End If
        on error goto 0
    Next
End Sub
