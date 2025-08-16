oDllList = Array("IgorKrup.dll", "itextsharp.dll", "Keyboard.dll", "ICSharpCode.SharpZipLib.dll")
sRemoteAppFolder = "" '"\\pwas0038\download\IgorKrup" 'Copy DLLs
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
sFolderPath = GetFolderPath()
sLibPath = "" 'sFolderPath & "\bin\Debug\IgorKrup.dll"

If fso.FileExists(sLibPath) = False Then
    sAppFolder = GetAppFolder()

    If sRemoteAppFolder <> "" Then
        DownloadDlls sRemoteAppFolder, sAppFolder        
    Else
        DownloadFiles "https://github.com/igorkrupitsky/IgorKrup/blob/main/bin/Debug", sAppFolder
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
        MsgBox "Registed " & sLibPath
    Else
        MsgBox "Could not register " & sLibPath & ", " & Err.number & ", " & Err.Description
    End If

    on error goto 0
Else
    MsgBox "Cound not find " & sLibPath
End if

'======================================
Sub DownloadFiles(sFromUrlFolder, sToFolder)
    Dim sRemoteFilePath, sLocalFilePath

    For Each sDllName in oDllList
        sFromUrl = sFromUrlFolder & "/" & sDllName
        sLocalFilePath  = sToFolder  & "\" & sDllName        

        If fso.FileExists(sLocalFilePath) Then
            sTempFilePath  = sToFolder  & "\temp_" & sDllName
            
            If fso.FileExists(sTempFilePath) Then
                fso.DeleteFile sTempFilePath, True
            End If

            DownloadFile sFromUrl, sTempFilePath

            If fso.GetFileVersion(sLocalFilePath) <> fso.GetFileVersion(sTempFilePath) Then
                'Versions are different
                fso.CopyFile sTempFilePath, sLocalFilePath, True
            End If

            If fso.FileExists(sTempFilePath) Then
                fso.DeleteFile sTempFilePath, True
            End If

        Else
            DownloadFile sFromUrl, sLocalFilePath
        End If        
    Next
End Sub

Function GetAppFolder()
    Dim sUserAppFolder: sUserAppFolder = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    If fso.FolderExists(sUserAppFolder) = False Then
        MsgBox "User App Folder does not exist: " & sUserAppFolder
    End If

    Dim sAppFolder: sAppFolder = sUserAppFolder & "\IgorKrup"
    If fso.FolderExists(sAppFolder) = False Then
        fso.CreateFolder sAppFolder
    End If

    GetAppFolder = sAppFolder
End Function

Sub DownloadDlls(sFromFolder, sToFolder)
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
        'MsgBox "Class is already registered: " & clsid
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
	GetFolderPath = oFile.ParentFolder
End Function

Sub DownloadFile(srcUrl, destPath)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folderPath: folderPath = fso.GetParentFolderName(destPath)
  If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

  Dim url: url = ToRawGithubUrl(srcUrl)

  Dim http: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
  Const WinHttpOption_SecureProtocols = 9
  Const TLS1_0 = &H80, TLS1_1 = &H200, TLS1_2 = &H800
  http.Option(WinHttpOption_SecureProtocols) = TLS1_0 Or TLS1_1 Or TLS1_2

  ' Follow redirects manually
  Dim i
  For i = 1 To 5
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "VBScript-Downloader"
    http.Send

    If http.Status >= 300 And http.Status < 400 Then
      Dim loc: loc = http.GetResponseHeader("Location")
      If Len(loc) = 0 Then Exit For
      url = loc
    Else
      Exit For
    End If
  Next

  If http.Status = 200 Then
    ' Optional sanity check: if first byte is "<" it’s probably HTML.
    Dim rb: rb = http.ResponseBody
    Dim looksHtml: looksHtml = False
    If IsArray(rb) Then
      On Error Resume Next
      If UBound(rb) >= 0 Then looksHtml = (rb(0) = 60) ' 60 = Asc("<")
      On Error GoTo 0
    End If

    Dim ct: ct = LCase(http.GetResponseHeader("Content-Type"))

    If looksHtml Or InStr(ct, "text/html") > 0 Then
      WScript.Echo "Server returned HTML instead of a binary file." & vbCrLf & _
                   "Final URL: " & url & vbCrLf & "Content-Type: " & ct & vbCrLf & _
                   "Tip: make sure it’s a RAW GitHub URL and the repo is public."
      Exit Sub
    End If

    Dim stm: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write http.ResponseBody
    stm.SaveToFile destPath, 2 ' adSaveCreateOverWrite
    stm.Close
    'WScript.Echo "Saved to: " & destPath
  Else
    WScript.Echo "HTTP " & http.Status & " " & http.StatusText & vbCrLf & "Final URL: " & url
  End If
End Sub

Function ToRawGithubUrl(u)
  If InStr(u, "://github.com/") > 0 And InStr(u, "/blob/") > 0 Then
    Dim s: s = Split(Replace(u, "://github.com/", "://raw.githubusercontent.com/"), "/blob/")
    ToRawGithubUrl = s(0) & "/" & s(1)
  Else
    ToRawGithubUrl = u
  End If
End Function

