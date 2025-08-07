' Full implementation of SeleniumBasic.EdgeDriver replacement 

Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Runtime.InteropServices
Imports System.IO.Compression
Imports System.Diagnostics
Imports System.Text.RegularExpressions

<ProgId("IgorKrup.EdgeDriver")>
<Guid("179F44FC-862E-472E-AD91-2BFAFD7763ED")>
<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class EdgeDriver

    Public sSharedDownloadFolder As String = "\\pwdb3030\download\Macros\Selenium\Selenium"
    Public sEdgeDriverPath As String = ""
    Public iPort As Integer = 9515

    Dim proc As Process = Nothing
    Dim sessionId As String = ""

    Public Sub New()
        sEdgeDriverPath = GetEdgeDriverPath()
    End Sub

    Public Sub GetUrl(url As String)

        If proc Is Nothing Then
            If sEdgeDriverPath = "" Then
                MsgBox($"msedgedriver.exe is missing. Run UpdateDriver() or manually download msedgedriver.exe to {AppDomain.CurrentDomain.BaseDirectory} from https://developer.microsoft.com/en-us/microsoft-edge/tools/webdrive")
                Exit Sub
            End If

            Init()
        End If

        Dim serializer As New JavaScriptSerializer()
        Dim payload = New Dictionary(Of String, Object) From {{"url", url}}
        Dim sJson = serializer.Serialize(payload)

        SendRequest($"http://localhost:{iPort}/session/{sessionId}/url", "POST", sJson)
    End Sub

    Public Function ExecuteScript(sJS As String) As Object

        If sessionId = "" Then
            Throw New Exception("First use GetUrl to initialize")
        End If

        Dim serializer As New JavaScriptSerializer()
        Dim payload = New Dictionary(Of String, Object) From {{"script", sJS}, {"args", New Object() {}}}
        Dim sJson = serializer.Serialize(payload)

        Dim sRetJson = SendRequest($"http://localhost:{iPort}/session/{sessionId}/execute/sync", "POST", sJson)
        Dim oRet As Object = serializer.DeserializeObject(sRetJson)
        Return oRet("value")
    End Function

    Public Sub SwitchToFrame(identifier As Object)
        Dim serializer As New JavaScriptSerializer()
        Dim idPayload As New Dictionary(Of String, Object)()

        If TypeOf identifier Is Integer OrElse TypeOf identifier Is String Then
            idPayload("id") = identifier
        ElseIf TypeOf identifier Is Dictionary(Of String, Object) Then
            idPayload("id") = identifier ' WebElement JSON object
        Else
            Throw New ArgumentException("Invalid frame identifier")
        End If

        Dim json = serializer.Serialize(idPayload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/frame", "POST", json)
    End Sub

    Public Sub Quit()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}", "DELETE", "")
        'proc.Kill()
    End Sub

    'Private support functions

    Private Sub Init()
        proc = New Process()
        proc.StartInfo.FileName = sEdgeDriverPath
        proc.StartInfo.Arguments = "--port=" & iPort
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        proc.StartInfo.RedirectStandardError = True
        proc.StartInfo.CreateNoWindow = True
        proc.Start()

        Dim success = WaitForDriver("http://localhost:" & iPort & "/status", 5000)
        If Not success Then
            MsgBox("Error: msedgedriver did not respond.")
            Return
        End If

        Dim sessionJson As String = "{""capabilities"": {""alwaysMatch"": {""browserName"": ""MicrosoftEdge""}}}"
        Dim sessionResponse = SendRequest($"http://localhost:{iPort}/session", "POST", sessionJson)
        sessionId = New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(sessionResponse)("value")("sessionId").ToString()
    End Sub

    Private Function WaitForDriver(url As String, timeoutMs As Integer) As Boolean
        Dim sw = Stopwatch.StartNew()
        Do
            Try
                Dim req = WebRequest.Create(url)
                Using resp = req.GetResponse()
                    Return True
                End Using
            Catch ex As WebException
                Threading.Thread.Sleep(200)
            End Try
        Loop While sw.ElapsedMilliseconds < timeoutMs
        Return False
    End Function

    Private Function SendRequest(url As String, method As String, body As String) As String
        Dim req = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = method
        req.ContentType = "application/json"

        If Not String.IsNullOrEmpty(body) Then
            Dim bytes = Encoding.UTF8.GetBytes(body)
            req.ContentLength = bytes.Length
            Using stream = req.GetRequestStream()
                stream.Write(bytes, 0, bytes.Length)
            End Using
        End If

        Using resp = CType(req.GetResponse(), HttpWebResponse)
            Using reader = New StreamReader(resp.GetResponseStream())
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function


    Public Sub UpdateDriver()
        If sEdgeDriverPath = "" Then
            GetSelenium()
            Exit Sub
        End If

        Dim sEdgeVersion As String = GetEdgeVersion()
        Dim sDriverVersion As String = GetMajorVersion(FileVersionInfo.GetVersionInfo(sEdgeDriverPath).FileVersion)

        If sEdgeVersion <> sDriverVersion Then
            GetSelenium()
            Exit Sub
        End If
    End Sub

    Private Function GetEdgeDriverPath()
        Dim iEdgeVersion As Integer = GetEdgeVersion()

        For Each sFileName As String In {
            "edgedriver_" & iEdgeVersion & ".exe",
            "msedgedriver.exe",
            "edgedriver.exe",
            "edgedriver_" & (iEdgeVersion - 1) & ".exe",
            "edgedriver_" & (iEdgeVersion - 2) & ".exe"
        }

            Dim sPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, sFileName)
            If IO.File.Exists(sPath) Then
                Return sPath
            End If
        Next
        Return ""
    End Function

    Private Function GetEdgeVersion() As Integer
        Dim fullVersion As String = CStr(Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon", "version", String.Empty))
        If Not String.IsNullOrEmpty(fullVersion) Then
            Return GetMajorVersion(fullVersion)
        Else
            Return 10
        End If
    End Function

    Private Sub GetSelenium()
        If sSharedDownloadFolder <> "" Then
            'Copy from shareled location
            Dim sEdgeVersion As String = GetEdgeVersion()
            Dim sFileName = $"edgedriver_{sEdgeVersion}.exe"
            Dim sSharedPath = Path.Combine(sSharedDownloadFolder, sFileName)
            If IO.File.Exists(sSharedPath) Then
                Dim sLocalPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, sFileName)
                If IO.File.Exists(sLocalPath) = False Then
                    File.Copy(sSharedPath, sLocalPath)
                End If

                sEdgeDriverPath = sLocalPath
                Exit Sub
            End If
        End If

        DownloadLatestSelenium()
        sEdgeDriverPath = GetEdgeDriverPath()
    End Sub

    Private Sub DownloadLatestSelenium()
        ' Force use of TLS 1.2
        System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)


        Dim client As New WebClient()
        Dim pageHtml As String = client.DownloadString("https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver")

        Dim match = Regex.Match(pageHtml, "(https:\/\/[^\s""]*edgedriver_win64\.zip)")
        If Not match.Success Then
            Console.WriteLine("Failed to find download URL.")
            Return
        End If

        Dim downloadUrl = match.Groups(1).Value
        Dim baseFolder = AppDomain.CurrentDomain.BaseDirectory
        Dim zipPath = Path.Combine(baseFolder, "edgedriver_win64.zip")

        If File.Exists(zipPath) Then File.Delete(zipPath)

        Console.WriteLine("Downloading: " & downloadUrl)
        client.DownloadFile(downloadUrl, zipPath)

        Dim edgeDriverPath = Path.Combine(baseFolder, "msedgedriver.exe")
        If File.Exists(edgeDriverPath) Then File.Delete(edgeDriverPath)

        UnzipFile(zipPath, baseFolder)

        If Not File.Exists(edgeDriverPath) Then
            Console.WriteLine("Unzipped file not found.")
            Return
        End If

        If File.Exists(zipPath) Then File.Delete(zipPath)

        Dim driver_Notes = Path.Combine(baseFolder, "Driver_Notes")
        If IO.Directory.Exists(driver_Notes) Then IO.Directory.Delete(driver_Notes, True)

        Dim newVersion = GetMajorVersion(FileVersionInfo.GetVersionInfo(edgeDriverPath).FileVersion)
        Dim newFileName = $"edgedriver_{newVersion}.exe"
        Dim localPath = Path.Combine(baseFolder, newFileName)

        'Save to local
        If Not File.Exists(localPath) Then
            File.Copy(edgeDriverPath, localPath)
        End If

        'Save to sSharedDownloadFolder
        If sSharedDownloadFolder <> "" Then
            Dim sSharedDownloadPath = Path.Combine(sSharedDownloadFolder, newFileName)
            If Not File.Exists(sSharedDownloadPath) Then
                File.Copy(edgeDriverPath, sSharedDownloadPath)
            End If
        End If

    End Sub

    Private Sub UnzipFile(zipPath As String, destFolder As String)
        Dim winrarPath As String = "C:\Program Files\WinRAR\WinRAR.exe"
        Dim fso As Object = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(winrarPath) Then
            ' Use WinRAR to extract the zip
            Dim cmd As String = Chr(34) & winrarPath & Chr(34) & " x -o+ -ibck " &
                            Chr(34) & zipPath & Chr(34) & " " & Chr(34) & destFolder & Chr(34)

            Dim shell = CreateObject("WScript.Shell")
            shell.Run(cmd, 0, True)
        Else
            Try
                ' Use ICSharpCode.SharpZipLib.Zip
                UnzipFile2(zipPath, destFolder)
            Catch ex As Exception
                Console.WriteLine("Shell extraction error: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub UnzipFile2(zipPath As String, destFolder As String)
        ' Ensure destination directory exists
        If Not Directory.Exists(destFolder) Then
            Directory.CreateDirectory(destFolder)
        End If

        ' Open the ZIP file for reading
        Using zipStream As FileStream = File.OpenRead(zipPath)
            Using zipInputStream As New ICSharpCode.SharpZipLib.Zip.ZipInputStream(zipStream)
                Dim entry As ICSharpCode.SharpZipLib.Zip.ZipEntry = zipInputStream.GetNextEntry()

                While entry IsNot Nothing
                    Dim entryFileName As String = entry.Name

                    ' Skip directories
                    If Not String.IsNullOrEmpty(entryFileName) AndAlso Not entry.IsDirectory Then
                        Dim fullPath As String = Path.Combine(destFolder, entryFileName)

                        ' Create subdirectories if necessary
                        Dim directoryName As String = Path.GetDirectoryName(fullPath)
                        If Not String.IsNullOrEmpty(directoryName) AndAlso Not Directory.Exists(directoryName) Then
                            Directory.CreateDirectory(directoryName)
                        End If

                        ' Extract the file
                        Using outputStream As FileStream = File.Create(fullPath)
                            Dim buffer(4096) As Byte
                            Dim bytesRead As Integer = zipInputStream.Read(buffer, 0, buffer.Length)

                            While bytesRead > 0
                                outputStream.Write(buffer, 0, bytesRead)
                                bytesRead = zipInputStream.Read(buffer, 0, buffer.Length)
                            End While
                        End Using
                    End If

                    entry = zipInputStream.GetNextEntry()
                End While
            End Using
        End Using
    End Sub

    Private Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

    Private Function GetMajorVersion(fullVersion As String) As String
        Dim i = fullVersion.IndexOf("."c)
        If i > 0 Then
            Return fullVersion.Substring(0, i)
        End If
        Return fullVersion
    End Function

End Class
