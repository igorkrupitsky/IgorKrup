Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
InstallIgorKrup

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
            DownloadFiles oDllList, "https://github.com/igorkrupitsky/IgorKrup/blob/main/bin/Debug", sAppFolder
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
            WScript.Echo "Registed " & sLibPath
        Else
            WScript.Echo "Could not register " & sLibPath & ", " & Err.number & ", " & Err.Description
        End If

        on error goto 0
    Else
        WScript.Echo "Cound not find " & sLibPath
    End if

End Sub

Sub DownloadFiles(oDllList, sFromUrlFolder, sToFolder)
    Dim sRemoteFilePath, sLocalFilePath

    For Each sDllName in oDllList
        sFromUrl = sFromUrlFolder & "/" & sDllName
        sLocalFilePath  = sToFolder  & "\" & sDllName        

        If fso.FileExists(sLocalFilePath) Then
            sTempFilePath  = sToFolder  & "\temp_" & sDllName
            
            If fso.FileExists(sTempFilePath) Then
                fso.DeleteFile sTempFilePath, True
            End If
                        
            DownloadFile ToRawGithubUrl(sFromUrl), sTempFilePath

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

Function ToRawGithubUrl(u)
  If InStr(u, "://github.com/") > 0 And InStr(u, "/blob/") > 0 Then
    Dim s: s = Split(Replace(u, "://github.com/", "://raw.githubusercontent.com/"), "/blob/")
    ToRawGithubUrl = s(0) & "/" & s(1)
  Else
    ToRawGithubUrl = u
  End If
End Function

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

'    libid = "{7142584B-8680-4FD6-9F60-8649F6BF1234}"
'    tlbVersion = "1.0"
'    tlbFullPath = "C:\Path\To\Your.tlb"
'    shell.RegWrite baseKey & "CLSID\" & clsid & "\TypeLib\", libid, "REG_SZ"
'    shell.RegWrite baseKey & "CLSID\" & clsid & "\Version\", tlbVersion, "REG_SZ"
'    shell.RegWrite baseKey & "TypeLib\" & libid & "\" & tlbVersion & "\0\win32\", tlbFullPath, "REG_SZ"
'    shell.RegWrite baseKey & "TypeLib\" & libid & "\" & tlbVersion & "\0\win64\", tlbFullPath, "REG_SZ"
'    shell.RegWrite baseKey & "TypeLib\" & libid & "\" & tlbVersion & "\FLAGS", 0, "REG_DWORD"

'    shell.RegWrite baseKey & progId & "\CurVer\", progId, "REG_SZ"
'    shell.RegWrite baseKey & "CLSID\" & clsid & "\VersionIndependentProgID\", progId, "REG_SZ"
End Sub

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder.ParentFolder
End Function

Sub DownloadFile(sUrl, sFilePath)
  Dim oHTTP: Set oHTTP = CreateObject("Microsoft.XMLHTTP")
  oHTTP.Open "GET", sUrl, False
  oHTTP.Send

  If oHTTP.Status = 200 Then 
    Set oStream = CreateObject("ADODB.Stream") 
    oStream.Open 
    oStream.Type = 1 
    oStream.Write oHTTP.ResponseBody 
    oStream.SaveToFile sFilePath, 2 
    oStream.Close 
  Else
    WScript.Echo "Error Status: " & oHTTP.Status & ", URL:" & sUrl
  End If
End Sub
