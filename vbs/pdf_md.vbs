Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim sRunUser: sRunUser = CreateObject("WScript.Network").UserName
Dim sFilePath: sFilePath = ""
Dim ie
Dim sGhostscriptPath: sGhostscriptPath = "C:\Igor\GitHub\PdfOcr\Ghostscript\bin\gswin64.exe"
'sGhostscriptPath = GetFolderPath() & "\Ghostscript\bin\gswin64.exe"

on error resume next
Set ie = CreateObject("IgorKrup.EdgeDriver")
If Err.number <> 0 Then
    WScript.Echo "Please download and run: IgorKrup.vbs"
    WScript.Quit 
End If
on error goto 0

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

If fso.FileExists(sFilePath) Or fso.FolderExists(sFilePath) Then
    Set ie = CreateObject("IgorKrup.EdgeDriver")
    ie.UpdateDriver
    ie.Get "https://m365.cloud.microsoft/chat" 
End If

If fso.FileExists(sFilePath) Then
    ProcessFile sFilePath
ElseIf fso.FolderExists(sFilePath) Then
    ProcessFolder sFilePath
End If

ie.Quit
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
    PdfToMarkDown sFilePath
    MergeMdFiles sFilePath 
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

Sub MergeMdFiles (sFilePath)
    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path & "\" & fso.GetBaseName(oFile.Name)   
    Dim sOutPath: sOutPath = sFolderPath & ".md"
        
    If fso.FileExists(sOutPath) Then
        Exit Sub
    End If

    Dim bFirstPage: bFirstPage = True
    Dim oOutFile: Set oOutFile = fso.CreateTextFile(sOutPath, True, True)
    Dim oFolder: Set oFolder = fso.GetFolder(sFolderPath)

    For Each oFile In oFolder.Files
        If fso.GetExtensionName(oFile.Name) = "md" Then
            
            If bFirstPage = False Then
                oOutFile.WriteLine ""
                oOutFile.WriteLine "___"
                oOutFile.WriteLine ""
            End If

            oOutFile.WriteLine "## Page " & fso.GetBaseName(oFile.Name)

            sText = ReadTextFile(oFile.Path)
            oOutFile.Write sText
            bFirstPage = False
        End If
    Next

    oOutFile.Close
    Set oOutFile = Nothing
End Sub

Sub PdfToMarkDown(sFilePath)

    Dim oPdf: Set oPdf = CreateObject("IgorKrup.PDF")
    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path & "\" & fso.GetBaseName(oFile.Name)    
    If fso.FolderExists(sFolderPath) = False Then
        MsgBox "Folder does not exist: " & sFolderPath
        WScript.Quit
    End If

    WaitForIE
    'Wait for chat to laod
    Do While ie.ExecuteScript("return document.querySelectorAll(""span[contenteditable]"").length") = 0
        WScript.Sleep 100
    Loop

    AddConvertElementToMarkdown

    bFirstQuestion = True
    sPrompt = "Convert attached image file to text.  If there is a table transcribe the entire table in full detail row by row. Do not give any files to download. Do not provide any comments, just give output."

    Dim oFolder: Set oFolder = fso.GetFolder(sFolderPath)
    For Each oFile In oFolder.Files
        If fso.GetExtensionName(oFile.Name) = "pdf" Then
            
            sTextFilePath = sFolderPath & "\" & fso.GetBaseName(oFile.Path) & ".md"
            If fso.FileExists(sTextFilePath) = False Then

                If bFirstQuestion = False Then                
                    NewQuestion
                End If

                sImgFilePath = sFolderPath & "\" & fso.GetBaseName(oFile.Path) & ".png"

                If fso.FileExists(sGhostscriptPath) And fso.FileExists(sImgFilePath) = False Then
                    'https://ghostscript.com/download/gsdnld.html
                    '-r300 - Print quality
                    '-r600- High-quality print/scanning
                    shell.run """" & sGhostscriptPath & """ -dNOPAUSE -q -r600 -sDEVICE=png16m -dBATCH -sOutputFile=""" & sImgFilePath & """ """ & oFile.Path & """ -c quit", 0 , True
                End If

                If fso.FileExists(sImgFilePath) Then
                    sAnswer = SendFile(sPrompt, sImgFilePath)
                Else
                    sAnswer = SendFile(sPrompt, oFile.Path)
                End If

                bFirstQuestion = False

                If Len(sAnswer) < 100 Then
                    sPdfText = oPdf.GetFileText(oFile.Path)
                    If Len(sPdfText) / 2 > Len(sAnswer) Then
                        'Use OCR Text
                        sAnswer = sPdfText
                    End If
                End If

                If sAnswer <> "" Then
                    WriteTextFile sTextFilePath, sAnswer
                End If

            End If
        End If
    Next

End Sub

Sub AddConvertElementToMarkdown
    'Load turndown.min.js
    ie.ExecuteScript "window.iTurnDownStatus = 0;" & _
        "const script = document.createElement('script');" & _
        "script.src = 'https://cdn.jsdelivr.net/npm/turndown/dist/turndown.min.js';" & _
        "script.onload  = () => {window.iTurnDownStatus =  1};" & _
        "script.onerror = () => {window.iTurnDownStatus = -1};" & _
        "document.head.appendChild(script)"

    Do While ie.ExecuteScript("return window.iTurnDownStatus") = 0
        WScript.Sleep 100
    Loop

    ie.ExecuteScript "window.ConvertElementToMarkdown = function (element) {" & _
    "  const td = new TurndownService({ codeBlockStyle: 'fenced' });" & _
    "  td.addRule('table', {" & _
    "    filter: 'table'," & _
    "    replacement: (_, node) => {" & _
    "      if (node.querySelector('[colspan], [rowspan]')) {" & _
    "        return node.outerHTML;" & _
    "      }" & _
    "      const rows = node.querySelectorAll('tr');" & _
    "      return Array.from(rows).map((row, i) => {" & _
    "        const cells = row.querySelectorAll('th, td');" & _
    "        const line = '| ' + Array.from(cells).map(c => c.textContent.trim()).join(' | ') + ' |\n';" & _
    "        const header = i === 0 ? '| ' + '--- |'.repeat(cells.length) + '\n' : '';" & _
    "        return line + header;" & _
    "      }).join('');" & _
    "    }" & _
    "  });" & _
    "  return td.turndown(element.innerHTML);" & _
    "}" 
End Sub

Sub NewQuestion
    ie.ExecuteScript "document.getElementById('new-chat-button').click()" 'New Question
    WScript.Sleep 1000
End Sub

Function SendFile(sPrompt, sFilePath)
    ie.ExecuteScript "const oSpan = document.querySelector('span[contenteditable]');"& _
    "oSpan.focus();" & _
    "const event = new InputEvent('beforeinput', {inputType: 'insertText', data: `" & Replace(sPrompt,"`","\`") & "`, bubbles: true, cancelable: true});" & _
    "oSpan.dispatchEvent(event)"

    WScript.Sleep 1000
    Do While ie.ExecuteScript("return document.querySelectorAll(""button[data-testid='PlusMenuButton']"").length") = 0
        WScript.Sleep 100
    Loop

    ie.ExecuteScript "document.querySelector(""button[data-testid='PlusMenuButton']"").click()"
    ie.ExecuteScript "document.querySelector(""div[role='menuitem']"").click()"
    ie.ExecuteScript "document.querySelector(""button[data-testid='upload-local-file']"").click()"

    sFileInputElementId = ie.FindElementByCss("input[type='file']")
    ie.UploadFileToElement sFilePath, sFileInputElementId

    WScript.Sleep 2000 'Wait for uplaod finish

    'Wait for reply
    Do While ie.ExecuteScript("return document.querySelectorAll(""div[data-testid='markdown-reply']"").length") = 0

        ie.ExecuteScript "var o = document.querySelector(""button[type='submit']""); if (o) o.click()"
        WScript.Sleep 100
    Loop

    Do While ie.ExecuteScript("var o = document.querySelector(""div[data-testid='markdown-reply']""); return o ? o.innerText : ''") = ""
        WScript.Sleep 100
    Loop

    Do While ie.ExecuteScript("var o = document.querySelector(""button[type='submit']""); return o ? o.getAttribute('aria-label') : ''") = "Stop generating"
        WScript.Sleep 100
    Loop

    Do While GetDataMessageTypeAttr() <> "Chat" and GetDataMessageTypeAttr() <> "GeneratedCode"
        WScript.Sleep 100
    Loop

    SendFile = ie.ExecuteScript("return ConvertElementToMarkdown(document.querySelector(""div[data-testid='markdown-reply']""))")
End Function

Function GetDataMessageTypeAttr()
    GetDataMessageTypeAttr = ie.ExecuteScript("return document.querySelector(""div[data-testid='markdown-reply']"").getAttribute('data-message-type')")
End Function

Sub WaitForIE
    Do While ie.ExecuteScript("return document.readyState") <> "complete"
        WScript.Sleep 100
    Loop
End Sub

Sub WriteTextFile(sPath, sText)
    Dim file: Set file = fso.CreateTextFile(sPath, True, True)
    file.Write sText
    file.Close
    Set file = Nothing
End Sub

Function ReadTextFile(sPath)
    Dim sRet: sRet = ""
    Dim file: Set file = fso.OpenTextFile(sPath, 1, False, -1)
    Do Until file.AtEndOfStream
        sRet = sRet & file.ReadLine & vbCrLf
    Loop
    file.Close
    Set file = Nothing
    ReadTextFile = sRet
End Function

Sub BreakePdfToPages(sFilePath)
    Dim oPdf: Set oPdf = CreateObject("IgorKrup.PDF")
    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path
    Dim sTempFolder: sTempFolder = sFolderPath & "\" & fso.GetBaseName(sFilePath)

    If fso.FolderExists(sTempFolder) = False Then
        fso.CreateFolder sTempFolder   
    End If

    Dim iPageCount: iPageCount = oPdf.PageCount(sFilePath)
    If iPageCount = 1 Then
        Dim sDestFile: sDestFile = sTempFolder & "\" & fso.GetFileName(sFilePath)
        If fso.FileExists(sDestFile) = False Then
            fso.CopyFile sFilePath, sDestFile
        End If
        Exit Sub
    End If

    Dim iPage, sOutputFile
    For iPage = 1 to iPageCount
        sOutputFile = sTempFolder & "\" & Right("000" & iPage, 3) & ".pdf"
        oPdf.ExtractPage sFilePath, sOutputFile, iPage
    Next

    Set oPdf = Nothing
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

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder
End Function

