Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
sDllName = "itextsharp.dll"
sFromUrlFolder = "https://github.com/igorkrupitsky/IgorKrup/blob/main/bin/Debug"
sFromUrl = sFromUrlFolder & "/" & sDllName
DownloadFile ToRawGithubUrl(sFromUrl), GetFolderPath() & "\" & sDllName
MsgBox "Done"

Function ToRawGithubUrl(u)
  If InStr(u, "://github.com/") > 0 And InStr(u, "/blob/") > 0 Then
    Dim s: s = Split(Replace(u, "://github.com/", "://raw.githubusercontent.com/"), "/blob/")
    ToRawGithubUrl = s(0) & "/" & s(1)
  Else
    ToRawGithubUrl = u
  End If
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

Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder 
End Function
