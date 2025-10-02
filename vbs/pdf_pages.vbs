Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim sRunUser: sRunUser = CreateObject("WScript.Network").UserName
Dim oPdf: Set oPdf = CreateObject("IgorKrup.PDF")
Dim sFilePath: sFilePath = ""
Dim ie

If WScript.Arguments.Count = 1 Then
    sFilePath = WScript.Arguments(0)
Else
    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.BrowseForFolder(0, "Select Folder", 0, "")
    If oFolder Is Nothing Then
        WScript.Quit
    Else
        sFolderPath = oFolder.Self.Path
        If sFolderPath = "" Then
            WScript.Quit 
        End If

        sHtml = "<p><select id='file_name' style='width: 100%'>" & GetFileSelect(sFolderPath)  & "</select></p>" & _
                "<p align=center><input type='button' value='Convert PDF to MD' onclick='Send()'> " & _
                                "<input type='button' value='Close' onclick='self.Close()'></p>"

        Set oRet = OpenDialog(sHtml, "file_name", 300, 200, "Convert PDF to MD")
        If oRet.Count = 0 Then
            'User clicked 'Close' button
            WScript.Quit
        End If

        file_name = oRet("file_name")
        If file_name = "" Then
            WScript.Quit
        End If

        If file_name = "0" Then
            sFilePath = sFolderPath
        Else
            sFilePath = sFolderPath & "\" & file_name
        End If
    End If
End If

If sFilePath = "" Then
    MsgBox "Please drag and drop a PDF file or folder on top this file to convert it to Markdown."
    WScript.Quit 
End If

If fso.FileExists(sFilePath) Then
    ProcessFile sFilePath
ElseIf fso.FolderExists(sFilePath) Then
    ProcessFolder sFilePath
End If

MsgBox "Done"

'===============================
Sub ProcessFolder (sFilePath)
    Dim oFolder: Set oFolder = fso.GetFolder(sFilePath)
    For Each oFile In oFolder.Files
        ProcessFile oFile.Path
    Next
End Sub

Sub ProcessFile (sFilePath)
    If fso.GetExtensionName(sFilePath) <> "pdf" Then    
        Exit Sub
    End If

    BreakePdfToPages sFilePath
End Sub

Function GetFileSelect(sFolderPath)
    Dim sRet: sRet = "<option value='0'>Convert All</options>" & vbCrLf

    Dim oFolder: Set oFolder = fso.GetFolder(sFolderPath)
    For Each oFile In oFolder.Files
        If fso.GetExtensionName(oFile.Name) = "pdf" Then  
            sRet = sRet & "<option value=""" & oFile.Name & """>" & oFile.Name & "</options>" & vbCrLf
        End If
    Next
    GetFileSelect = sRet
End Function

Sub BreakePdfToPages(sFilePath)
    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path
    Dim sTempFolder: sTempFolder = sFolderPath & "\PDF_pages"

    sDestFile = sTempFolder & "\" & fso.GetBaseName(sFilePath) & "_001.pdf"
    If fso.FileExists(sDestFile)  Then
        Exit Sub 
    End If

    If fso.FolderExists(sTempFolder) = False Then
        fso.CreateFolder sTempFolder   
    End If

    Dim iPageCount

    on error resume next
    iPageCount = oPdf.PageCount(sFilePath)
    If Err.number <> 0 Then
        MsgBox Err.Description & ", PageCount: " & sFilePath
    End If
    on error goto 0

    If iPageCount = 1 Then
        Dim sDestFile: sDestFile = sTempFolder & "\" & fso.GetBaseName(sFilePath) & "_001.pdf"
        If fso.FileExists(sDestFile) = False Then
            fso.CopyFile sFilePath, sDestFile
        End If
        Exit Sub
    End If

    Dim iPage, sOutputFile
    For iPage = 1 to iPageCount
        sOutputFile = sTempFolder & "\" & fso.GetBaseName(sFilePath) & "_" & Right("000" & iPage, 3) & ".pdf"

        on error resume next
        oPdf.ExtractPage sFilePath, sOutputFile, iPage
        If Err.number <> 0 Then
            MsgBox Err.Description & ", ExtractPage(" & iPage & "): " & sFilePath
        End If
        on error goto 0
    Next
End Sub

Function OpenDialog(sHtml, sFields,iWidth,iHeight, sTitle)
  sHtaFilePath = Wscript.ScriptFullName & ".hta"

  CreateHtaFile sHtaFilePath, sHtml, sFields,iWidth,iHeight,sTitle

  Set f = fso.GetFile(sHtaFilePath)
  f.attributes = f.attributes + 2 'Hidden

  Dim oShell: Set oShell = CreateObject("WScript.Shell")
  
  oShell.Run """" & sHtaFilePath & """", 1, True

  If fso.FileExists(sHtaFilePath) Then
    fso.DeleteFile sHtaFilePath
  End If

  Set OpenDialog = ReadXmlFile(sHtaFilePath & ".xml", sFields, True)
End Function

Function ReadXmlFile(sFilePath, sFields, bDeleteAfterRead)
  Set oRet = CreateObject("Scripting.Dictionary")

  'Load return data from XML File
  If fso.FileExists(sFilePath) Then
      Set oXml = CreateObject("Microsoft.XMLDOM")
      oXML.async = False
      oXML.load sFilePath

      For each sField In Split(sFields,",")
        Set oNode = oXML.SelectSingleNode("/root/" & trim(sField))
        If Not oNode is Nothing Then
            oRet.Add trim(sField), oNode.text
        End If
      Next

      if bDeleteAfterRead Then
        fso.DeleteFile sFilePath
      End If
  End If

  Set ReadXmlFile = oRet
End Function

Sub CreateHtaFile(sHtaFilePath, sHtml, sFields, iWidth, iHeight, sTitle)

    If fso.FileExists(sHtaFilePath) Then
        MsgBox "You double-clicked the script twice. Dialog is already opened and may be hidded behind some other window. " & sHtaFilePath
        WScript.Quit
    End If

    Set f = fso.CreateTextFile(sHtaFilePath, True)
    f.WriteLine "<html><title>Convert PDF to MD</title><head><HTA:APPLICATION ID=oHTA SINGLEINSTANCE=""yes"" SCROLL=""no""/></head>"
    f.WriteLine "<script language=""vbscript"">"
    f.WriteLine "Window.ResizeTo " & iWidth & ", " & iHeight
    f.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine ""
    f.WriteLine "Sub Send()"
    f.WriteLine " Dim sFilePath: sFilePath = Replace(location.href,""file:///"","""")"
    f.WriteLine " sFilePath = Replace(sFilePath,""/"",""\"")"
    f.WriteLine " sFilePath = Replace(sFilePath,""%20"","" "")"
    f.WriteLine " Set oXml = CreateObject(""Microsoft.XMLDOM"")"
    f.WriteLine " Set oRoot = oXml.createElement(""root"")"
    f.WriteLine " oXml.appendChild oRoot"

    For each sField In Split(sFields,",")
        f.WriteLine " AddXmlVal oXml, oRoot, """ & sField & """, GetVal(" & sField & ")"
    Next

    f.WriteLine " oXml.Save sFilePath & "".xml"""
    f.WriteLine " self.Close()"
    f.WriteLine "End Sub"
    f.WriteLine ""
    f.WriteLine "Sub AddXmlVal(oXml, oRoot, sName, sVal)"
    f.WriteLine " Set oNode = oXml.createElement(sName)"
    f.WriteLine " oNode.Text = sVal"
    f.WriteLine " oRoot.appendChild oNode"
    f.WriteLine "End Sub"
    f.WriteLine ""
    f.WriteLine "Function GetVal(o)"
    f.WriteLine " GetVal = o.value"
    f.WriteLine " If o.Type = ""checkbox"" Then"
    f.WriteLine "   If o.checked = False Then"
    f.WriteLine "     GetVal = """""
    f.WriteLine "   End If"
    f.WriteLine " End If"
    f.WriteLine "End Function"  
    f.WriteLine "</script>"
    f.WriteLine "<body>"
    f.WriteLine sHtml
    f.WriteLine "</body></html>"
    f.Close
End Sub
