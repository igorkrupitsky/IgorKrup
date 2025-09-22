Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim sRunUser: sRunUser = CreateObject("WScript.Network").UserName
Dim sFilePath: sFilePath = ""
Dim ie

InstallIgorKrup

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
    CopyPdfToOneDrive sFilePath
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

    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path & "\" & fso.GetBaseName(oFile.Name)    
    If fso.FolderExists(sFolderPath) = False Then
        MsgBox "Folder does not exist: " & sFolderPath
        WScript.Quit
    End If
    
    Set ie = CreateObject("Selenium.EdgeDriver")
    ie.Get "https://m365.cloud.microsoft/chat" 

    WaitForIE
    'Wait for chat to laod
    Do While ie.ExecuteScript("return document.querySelectorAll(""span[contenteditable]"").length") = 0
        WScript.Sleep 100
    Loop

    AddConvertElementToMarkdown

    bFirstQuestion = True
    sPrompt = "OCR attached PDF.  Do not provide any comments, just output."

    Dim sCloudFileFolder: sCloudFileFolder = GetCloudFolder(sFilePath)

    Dim oFolder: Set oFolder = fso.GetFolder(sFolderPath)
    For Each oFile In oFolder.Files
        If fso.GetExtensionName(oFile.Name) = "pdf" Then
            
            sTextFilePath = sFolderPath & "\" & fso.GetBaseName(oFile.Path) & ".md"
            If fso.FileExists(sTextFilePath) = False Then

                sCloudPath = "Microsoft Copilot Chat Files\Temp\" & fso.GetBaseName(sFilePath) & "\" & oFile.Name 'Microsoft Copilot Chat Files\Temp\FileName\001.pdf"
                        
                If bFirstQuestion = False Then                
                    NewQuestion
                End If

                sAnswer = SendFile(sPrompt, sCloudPath)
                bFirstQuestion = False

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

Function SendFile(sPrompt, sCloudPath)
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
    ie.ExecuteScript "document.querySelector(""button[data-testid='upload-cloud-file']"").click()"

    Set iframe = ie.FindElementByXPath("//iframe[@title='File Picker']")
    ie.SwitchToFrame iframe

    'Pick Files Dialog
    Do While ie.ExecuteScript("return document.querySelectorAll(""button[data-automationid='nav-collapse-button']"").length") = 0
        WScript.Sleep 100
    Loop

    ie.ExecuteScript "document.querySelector(""div[name='My files']"").click()"
    WScript.Sleep 3000

    Do While ie.ExecuteScript("return document.querySelectorAll(""input[role='searchbox']"").length") = 0
        WScript.Sleep 100
    Loop

    'Search Function
    ie.ExecuteScript "window.SearchItem = function(searchText){" & _
            "var o = document.querySelector(""input[role='searchbox']"");" & _
            "o.focus();" & _
            "o.value = searchText;" & _
            "o.dispatchEvent(new Event('input', { bubbles: true }));" & _
            "o.dispatchEvent(new Event('change', { bubbles: true }));" & _
            "o.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));" & _
            "o.dispatchEvent(new KeyboardEvent('keyup',   { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));" & _
        "}"

    oCloudPath = Split(sCloudPath, "\")

    For Each sPart in oCloudPath

        Do While ie.ExecuteScript("return document.querySelectorAll('div[data-is-scrollable]').length") = 0
            WScript.Sleep 100
        Loop

        WScript.Sleep 200

        iWait = 1
        Do While ie.ExecuteScript("window.items = document.querySelectorAll(""[data-automationid='FieldRenderer-name'][title='" & sPart & "']""); return items.length") = 0
            WScript.Sleep 10
            iWait = iWait + 1
                
            If iWait = 100 * 1 and InStr(sPart, ".pdf") <> 0 Then
                '1 sec PDF
                ie.ExecuteScript "SearchItem('" & sPart & "')"

            ElseIf iWait = 100 * 2 and InStr(sPart, ".pdf") = 0 Then
                '2 sec Folder
                ie.ExecuteScript "var o = document.querySelector('div[data-is-scrollable]'); o.scrollTo({top: o.scrollHeight,  behavior: 'smooth'})"

            ElseIf iWait = 100 * 5 Then
                '5 sec
                ie.ExecuteScript "SearchItem('')" 'Reset search
                ie.ExecuteScript "var o = document.querySelector('div[data-is-scrollable]'); " & _
                                    "var dir = 1; var step = 5;" & _
                                    "setInterval(function(){ " & _
                                    "o.scrollTop += step * dir;" & _
                                    "if (o.scrollTop + o.clientHeight >= o.scrollHeight - 5) {dir = -1; step = Math.max(step - 1, 2);} " & _
                                    "if (o.scrollTop <= 5) {dir  = 1; step = Math.max(step - 1, 2);} " & _
                                    "}, 10)"    

            ElseIf iWait > 100 * 60 * 5 Then
                '5 mins
                ie.ExecuteScript "SearchItem('" & sPart & "')"
                MsgBox "Could not select " & sPart & ". Please select " & sPart & " manually."
                ie.ExecuteScript "SearchItem('')" 'Reset search
                ie.ExecuteScript "window.items = []"
            End If
        Loop

        ie.ExecuteScript "if (window.items.length > 0) window.items[0].click()"
        WScript.Sleep 1000
    Next

    ie.ExecuteScript "document.querySelector(""button[data-automationid='picker-complete']"").click()"
    WScript.Sleep 200

    ie.SwitchToDefaultContent

    'Wait for uplaod finish
    Do While ie.ExecuteScript("return document.querySelectorAll(""button[type='submit']"").length") = 0
        WScript.Sleep 100
    Loop

    ie.ExecuteScript "document.querySelector(""button[type='submit']"").click()"

    'Wait for reply
    Do While ie.ExecuteScript("return document.querySelectorAll(""div[data-testid='markdown-reply']"").length") = 0
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

Function GetCloudFolder(sFilePath)
    Const sCloudBase = "Microsoft Copilot Chat Files"
    Dim sCloudLocal: sCloudLocal = "C:\Users\" & sRunUser & "\OneDrive - Moffitt Cancer Center\" & sCloudBase
    If fso.FolderExists(sCloudLocal) = False Then
        MsgBox "OneDrive folder does not exist: " & sCloudLocal
        WScript.Quit
    End If

    'Create Cloud: Temp
    Dim sCloudLocalSubFolder: sCloudLocalSubFolder = sCloudLocal & "\Temp"
    If fso.FolderExists(sCloudLocalSubFolder) = False Then
        fso.CreateFolder sCloudLocalSubFolder
    End If
    
    'Create Cloud: Temp\FileName
    Dim sCloudFileFolder: sCloudFileFolder = sCloudLocalSubFolder & "\" & fso.GetBaseName(sFilePath)
    If fso.FolderExists(sCloudFileFolder) = False Then
        fso.CreateFolder sCloudFileFolder
    End If

    GetCloudFolder = sCloudFileFolder
End Function

Sub CopyPdfToOneDrive(sFilePath)

    Dim oFile: Set oFile = fso.GetFile(sFilePath)
    Dim sFolderPath: sFolderPath = oFile.ParentFolder.Path & "\" & fso.GetBaseName(oFile.Name)    
    If fso.FolderExists(sFolderPath) = False Then
        MsgBox "CopyPdfToOneDrive. Folder does not exist: " & sFolderPath
        WScript.Quit
    End If

    Dim sCloudFileFolder: sCloudFileFolder = GetCloudFolder(sFilePath)

    Dim sCloudFilePath
    Dim oPageFile
    Dim oPageFolder: Set oPageFolder = fso.GetFolder(sFolderPath)

    For Each oPageFile In oPageFolder.Files
        If fso.GetExtensionName(oPageFile.Name) = "pdf" Then
            sCloudFilePath = sCloudFileFolder & "\" & oPageFile.Name
            If fso.FileExists(sCloudFilePath) = False Then
                fso.CopyFile oPageFile.Path, sCloudFilePath
            End If
        End If
    Next
End Sub

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
        Dim sDestFile: sDestFile = sTempFolder & "\" & fso.GetBaseName(sFilePath)
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


'======================================
Sub InstallIgorKrup

    Dim oDllList: oDllList = Array("IgorKrup.dll", "itextsharp.dll", "Keyboard.dll", "ICSharpCode.SharpZipLib.dll")
    Dim sRemoteAppFolder: sRemoteAppFolder = "" '"\\pwas0038\download\IgorKrup" 'Copy above files to your own file share (if you don't want to download file from github)        
    Dim sFolderPath: sFolderPath = GetFolderPath()
    Dim sLibPath: sLibPath = sFolderPath & "\bin\Debug\IgorKrup.dll"

    If fso.FileExists(sLibPath) = False Then
        sAppFolder = GetAppFolder()

        If sRemoteAppFolder <> "" Then
            DownloadDlls oDllList, sRemoteAppFolder, sAppFolder        
        Else
            DownloadGitHubFiles oDllList, "https://github.com/igorkrupitsky/IgorKrup/blob/main/bin/Debug", sAppFolder
        End If

        sLibPath = sAppFolder   & "\IgorKrup.dll"
    End If

    If fso.FileExists(sLibPath) Then
        RegisterComClass "IgorKrup", "{6332A55E-EEFF-433B-9976-41A4A0420877}", sLibPath, "IgorKrup.PDF"
        RegisterComClass "IgorKrup", "{179F44FC-862E-472E-AD91-2BFAFD7763ED}", sLibPath, "IgorKrup.EdgeDriver"
        RegisterComClass "IgorKrup", "{7142584B-8680-4FD6-9F60-8649F6BF6966}", sLibPath, "IgorKrup.Control"
    
        on error resume next
        Set o = CreateObject("IgorKrup.Control")
    
        If Err.number = 0 Then
            'WScript.Echo "Registed " & sLibPath
        Else
            WScript.Echo "Could not register " & sLibPath & ", " & Err.number & ", " & Err.Description
        End If

        on error goto 0
    Else
        WScript.Echo "Cound not find " & sLibPath
    End if

End Sub

Sub DownloadGitHubFiles(oDllList, sFromUrlFolder, sToFolder)
    Dim sRemoteFilePath, sLocalFilePath

    For Each sDllName in oDllList
        sFromUrl = sFromUrlFolder & "/" & sDllName
        sLocalFilePath  = sToFolder  & "\" & sDllName        

        If fso.FileExists(sLocalFilePath) Then
            sTempFilePath  = sToFolder  & "\temp_" & sDllName
            
            If fso.FileExists(sTempFilePath) Then
                fso.DeleteFile sTempFilePath, True
            End If
                        
            DownloadGitHubFile sFromUrl, sTempFilePath

            If fso.GetFileVersion(sLocalFilePath) <> fso.GetFileVersion(sTempFilePath) Then
                'Versions are different
                fso.CopyFile sTempFilePath, sLocalFilePath, True
            End If

            If fso.FileExists(sTempFilePath) Then
                fso.DeleteFile sTempFilePath, True
            End If

        Else
            DownloadGitHubFile sFromUrl, sLocalFilePath
        End If        
    Next
End Sub

Function GetAppFolder()
    Dim sUserAppFolder: sUserAppFolder = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    If fso.FolderExists(sUserAppFolder) = False Then
        WScript.Echo "User App Folder does not exist: " & sUserAppFolder
    End If

    Dim sAppFolder: sAppFolder = sUserAppFolder & "\IgorKrup"
    If fso.FolderExists(sAppFolder) = False Then
        fso.CreateFolder sAppFolder
    End If

    GetAppFolder = sAppFolder
End Function

Sub DownloadDlls(oDllList, sFromFolder, sToFolder)
    Dim sRemoteFilePath, sLocalFilePath

    For Each sDllName in oDllList
        sRemoteFilePath = sFromFolder & "\" & sDllName
        sLocalFilePath  = sToFolder  & "\" & sDllName
        If fso.FileExists(sRemoteFilePath) Then
            If fso.FileExists(sLocalFilePath) Then
                If fso.GetFileVersion(sLocalFilePath) <> fso.GetFileVersion(sRemoteFilePath) Then
                    'Versions are different
                    fso.CopyFile sRemoteFilePath, sLocalFilePath, True
                End If
            Else
                fso.CopyFile sRemoteFilePath, sLocalFilePath
            End If
        End If
    Next
End Sub

Sub RegisterComClass(assemblyName, clsid, dllFullPath, progId)
    Dim baseKey, codebase
    baseKey = "HKCU\Software\Classes\"
    codebase = "file:///" & Replace(dllFullPath, "\", "/")

    On Error Resume Next
    If shell.RegRead(baseKey & "CLSID\" & clsid & "\InprocServer32\CodeBase") = codebase Then
        'WScript.Echo "Class is already registered: " & clsid
        'Exit Sub
    End If
    On Error GoTo 0

    ' ProgID key
    shell.RegWrite baseKey & progId & "\", progId, "REG_SZ"
    shell.RegWrite baseKey & progId & "\CLSID\", clsid, "REG_SZ"

    ' CLSID root
    shell.RegWrite baseKey & "CLSID\" & clsid & "\", progId, "REG_SZ"

    ' InprocServer32 base
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\", "mscoree.dll", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\ThreadingModel", "Both", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\Class", progId, "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\Assembly", assemblyName & ", Version=1.0.0.0, Culture=neutral, PublicKeyToken=null", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\RuntimeVersion", "v4.0.30319", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\CodeBase", codebase, "REG_SZ"

    ' Version-specific entry
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\1.0.0.0\Class", progId, "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\1.0.0.0\Assembly", assemblyName & ", Version=1.0.0.0, Culture=neutral, PublicKeyToken=null", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\1.0.0.0\RuntimeVersion", "v4.0.30319", "REG_SZ"
    shell.RegWrite baseKey & "CLSID\" & clsid & "\InprocServer32\1.0.0.0\CodeBase", codebase, "REG_SZ"

    ' ProgID again under CLSID
    shell.RegWrite baseKey & "CLSID\" & clsid & "\ProgId\", progId, "REG_SZ"

    ' Optional: Mark as safe for scripting & initialization
    shell.RegWrite baseKey & "CLSID\" & clsid & "\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}", "", "REG_SZ" ' Safe for scripting
    shell.RegWrite baseKey & "CLSID\" & clsid & "\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}", "", "REG_SZ" ' Safe for initializing
End Sub

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder.ParentFolder
End Function

Sub DownloadGitHubFile(sUrl, sFilePath)
    Dim sUrl2: sUrl2 = NormalizeGitHubUrl(sUrl)

    Dim oHTTP: Set oHTTP = CreateObject("Microsoft.XMLHTTP")
    oHTTP.Open "GET", sUrl2, False
    oHTTP.Send

    If oHTTP.Status = 200 Then 
        Set oStream = CreateObject("ADODB.Stream") 
        oStream.Open 
        oStream.Type = 1 
        oStream.Write oHTTP.ResponseBody 
        oStream.SaveToFile sFilePath, 2 
        oStream.Close 
    Else
        WScript.Echo "Error Status: " & oHTTP.Status & ", URL:" & sUrl2
    End If

    Set oFile = fso.GetFile(sFilePath)
    If oFile.Size > 500000 Then
        'File should be OK (not html) if over 0.5 MB
    Else
        sFileText = GetFileText(sFilePath)
        If InStr(sFileText, "<!doctype html") > 0 Or InStr(sFileText, "<html") > 0 Then
            WScript.Echo "Could not download: " & sUrl & cvCrLf & sFileText
        End If
    End If
End Sub

Function GetFileText(filePath)
    Dim ts: Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' 1 = ForReading
    GetFileText = ts.ReadAll
    ts.Close
End Function

Function NormalizeGitHubUrl(ByVal url)
    Dim re, m
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^https?://github\.com/([^/]+)/([^/]+)/blob/([^/]+)/(.+)$"
    re.IgnoreCase = True

    If re.Test(url) Then
        Set m = re.Execute(url)(0)
        ' owner / repo / branch / path
        NormalizeGitHubUrl = "https://raw.githubusercontent.com/" & _
                             m.SubMatches(0) & "/" & _
                             m.SubMatches(1) & "/" & _
                             m.SubMatches(2) & "/" & _
                             m.SubMatches(3)
    Else
        NormalizeGitHubUrl = url
    End If
End Function

