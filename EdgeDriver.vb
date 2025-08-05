' Full implementation of SeleniumBasic.EdgeDriver replacement 

Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Runtime.InteropServices

<ProgId("IgorKrup.EdgeDriver")>
<Guid("179F44FC-862E-472E-AD91-2BFAFD7763ED")>
<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class EdgeDriver

    Public sEdgeDriverPath As String = "C:\Users\80014379\Desktop\Update Selimium\msedgedriver.exe"
    Public iPort As Integer = 9515

    Dim proc As Process = Nothing
    Dim sessionId As String = ""

    Public Sub New()
    End Sub

    Public Sub GetUrl(url As String)
        If proc Is Nothing Then
            Init()
        End If

        Dim sJson = "{""url"":""" & PadJson(url) & """}"
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/url", "POST", sJson)
    End Sub

    Public Function ExecuteScript(sJS As String) As Object

        If sessionId = "" Then
            Throw New Exception("First use GetUrl to initialize")
        End If

        Dim sJson = "{""script"":""" & PadJson(sJS) & """,""args"":[]}"
        Dim sRetJson = SendRequest($"http://localhost:{iPort}/session/{sessionId}/execute/sync", "POST", sJson)
        Dim serializer As New JavaScriptSerializer()
        Dim oRet As Object = serializer.DeserializeObject(sRetJson)
        Return oRet("value")
    End Function

    Public Sub SwitchToFrame(identifier As Object)
        Dim idJson As String

        If TypeOf identifier Is Integer OrElse TypeOf identifier Is String Then
            ' Identifier is an index or frame name/id
            idJson = $"""id"": {If(TypeOf identifier Is Integer, identifier.ToString(), $"""{identifier}""")}
"
        ElseIf TypeOf identifier Is String AndAlso identifier.ToString().Trim().StartsWith("{") Then
            ' Assume raw JSON object (e.g., WebElement reference)
            idJson = $"""id"": {identifier}"
        Else
            Throw New ArgumentException("Invalid frame identifier")
        End If

        Dim json = "{" & idJson & "}"
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/frame", "POST", json)
    End Sub

    Public Sub Quit()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}", "DELETE", "")
        'proc.Kill()
    End Sub

    'Private support functions
    Private Function PadJson(ByVal s As String) As String
        s = Replace(s, "\", "\\")
        s = Replace(s, vbCrLf, "\r\n")
        s = Replace(s, vbCr, "\r")
        s = Replace(s, vbLf, "\n")
        s = Replace(s, vbTab, "\t")
        Return Replace(s, """", "\""")
    End Function

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

End Class
