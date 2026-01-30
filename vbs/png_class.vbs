Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim dic: Set dic = CreateObject("Scripting.Dictionary")

sInFolderPath = "C:\Users\80014379\Desktop\Financials_Contracts\PDF_page_img"
sOutFilePath = sInFolderPath & "\HasTable.txt"

If fso.FileExists(sOutFilePath) Then
    Set oTextFile = fso.OpenTextFile(sOutFilePath)
    Do Until oTextFile.AtEndOfStream
        sLine = oTextFile.ReadLine
        oLine = Split(sLine, vbTab)
        If Trim(oLine(1)) <> "" Then
            dic(oLine(0)) = oLine(1)
        End If
    Loop
    oTextFile.Close
End If


on error resume next
Set ie = CreateObject("IgorKrup.EdgeDriver")
If Err.number <> 0 Then
    WScript.Echo "Please download and run: IgorKrup.vbs"
    WScript.Quit 
End If
on error goto 0

ie.UpdateDriver
ie.Get "https://m365.cloud.microsoft/chat" 

WaitForIE
'Wait for chat to laod
Do While ie.ExecuteScript("return document.querySelectorAll(""span[contenteditable]"").length") = 0
    WScript.Sleep 100
Loop

AddConvertElementToMarkdown

bFirstQuestion = True
sPrompt = "Does the attached image contain table?  Do not provide any comments, just say: Yes or No."

Set oLogFile = fso.CreateTextFile(sOutFilePath)

Dim oFolder: Set oFolder = fso.GetFolder(sInFolderPath)
For Each oFile In oFolder.Files
    If dic.Exists(oFile.Path) = False Then
        If bFirstQuestion = False Then                
            NewQuestion
        End If

        sAnswer = SendFile(sPrompt, oFile.Path)
        oLogFile.WriteLine oFile.Path & vbTab & sAnswer
        bFirstQuestion = False
    End If
Next

oLogFile.Close
ie.Quit
MsgBox "Done"

'========================================
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

    iTries = 0
    Do While ie.ExecuteScript("var o = document.querySelector(""div[data-testid='markdown-reply']""); return o ? o.innerText : ''") = ""
        iTries = iTries + 1
        If iTries > 10 * 60 * 5 Then 
            '5 min timemout
            
            Do While ie.ExecuteScript("var o = document.querySelector("".fai-CopilotMessage__content""); return o ? o.innerText : ''") = ""
                iTries = iTries + 1
                If iTries > 10 * 60 * 20 Then 
                    '20 min timemout
                    Exit Function
                End If
                WScript.Sleep 100
            Loop

            SendFile = ie.ExecuteScript("var o = document.querySelector("".fai-CopilotMessage__content""); return o ? o.innerText : ''")
            Exit Function
        End If
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


