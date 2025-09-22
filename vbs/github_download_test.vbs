Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
sDllName = "itextsharp.dll"
sFromUrlFolder = "https://github.com/igorkrupitsky/IgorKrup/blob/main/bin/Debug"
sFromUrl = sFromUrlFolder & "/" & sDllName
sToFilePath = GetFolderPath() & "\" & sDllName
DownloadFile NormalizeGitHubUrl(sFromUrl), sToFilePath

Set oFile = fso.GetFile(sToFilePath)
If oFile.Size > 500000 Then
    'File should be HTML if over 0.5 MB
Else
    sFileText = GetFileText(sToFilePath)
    If InStr(sFileText, "<!doctype html") > 0 Or InStr(sFileText, "<html") > 0 Then
        MsgBox sFileText
    End If
End If

MsgBox "Done"

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

Function GetFileText(filePath)
    Dim ts: Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' 1 = ForReading
    GetFileText = ts.ReadAll
    ts.Close
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
