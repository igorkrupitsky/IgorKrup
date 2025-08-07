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

    Public sSharedDownloadFolder As String = ""
    Public sEdgeDriverPath As String = ""
    Public iPort As Integer = 9515

    Dim proc As Process = Nothing
    Dim sessionId As String = ""

    Public Sub New()
        sEdgeDriverPath = GetEdgeDriverPath()
    End Sub

    Public Sub GetUrl(url As String, Optional username As String = "", Optional password As String = "")
        If proc Is Nothing Then
            If sEdgeDriverPath = "" Then
                MsgBox($"msedgedriver.exe is missing. Run UpdateDriver() or manually download msedgedriver.exe to {AppDomain.CurrentDomain.BaseDirectory} from https://developer.microsoft.com/en-us/microsoft-edge/tools/webdrive")
                Exit Sub
            End If
            Init()
        End If

        If username <> "" And password <> "" Then
            ' Insert credentials into the URL
            Dim uri As New Uri(url)
            url = uri.Scheme & "://" & username & ":" & password & "@" & uri.Host & uri.PathAndQuery
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

    ' Switch to parent frame
    Public Sub SwitchToParentFrame()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/frame/parent", "POST", "{}")
    End Sub

    Public Sub Quit()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}", "DELETE", "")
        'proc.Kill()
    End Sub

    Public Sub CloseAllWindows()
        SendRequest($"http://localhost:{iPort}/session", "DELETE", "")
    End Sub


    ' Navigate back in browser history
    Public Sub NavigateBack()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/back", "POST", "{}")
    End Sub

    ' Navigate forward in browser history
    Public Sub NavigateForward()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/forward", "POST", "{}")
    End Sub

    ' Refresh the current page
    Public Sub Refresh()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/refresh", "POST", "{}")
    End Sub

    ' Get the current URL
    Public Function GetCurrentUrl() As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/url", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    ' Get the title of the current page
    Public Function GetTitle() As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/title", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    ' Add a cookie
    Public Sub AddCookie(name As String, value As String)
        Dim payload = New Dictionary(Of String, Object) From {{"cookie", New Dictionary(Of String, Object) From {{"name", name}, {"value", value}}}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/cookie", "POST", sJson)
    End Sub

    ' Delete all cookies
    Public Sub DeleteAllCookies()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/cookie", "DELETE", "")
    End Sub


    ' Accept JavaScript alert
    Public Sub AcceptAlert()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/alert/accept", "POST", "{}")
    End Sub

    ' Dismiss JavaScript alert
    Public Sub DismissAlert()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/alert/dismiss", "POST", "{}")
    End Sub

    ' Get alert text
    Public Function GetAlertText() As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/alert/text", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    ' Set implicit wait timeout (milliseconds)
    Public Sub SetImplicitWait(milliseconds As Integer)
        Dim payload = New Dictionary(Of String, Object) From {{"implicit", milliseconds}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/timeouts", "POST", sJson)
    End Sub


    ' Maximize browser window
    Public Sub MaximizeWindow()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/window/maximize", "POST", "{}")
    End Sub

    ' Minimize browser window
    Public Sub MinimizeWindow()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/window/minimize", "POST", "{}")
    End Sub

    ' Set browser window size
    Public Sub SetWindowSize(width As Integer, height As Integer)
        Dim payload = New Dictionary(Of String, Object) From {
        {"width", width},
        {"height", height}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/window/rect", "POST", sJson)
    End Sub

    ' Get browser window size
    Public Function GetWindowSize() As Dictionary(Of String, Object)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/window/rect", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
    End Function

    ' Close the current window
    Public Sub CloseWindow()
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/window", "DELETE", "")
    End Sub

    Public Function GetPageSource() As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/source", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    Public Function GetWindowHandles() As String()
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/window/handles", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
    End Function
    Public Sub SwitchToWindow(windowHandle As String)
        Dim payload = New Dictionary(Of String, Object) From {{"handle", windowHandle}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/window", "POST", sJson)
    End Sub

    ' Take screenshot and save to path
    Public Sub TakeScreenshot(savePath As String)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/screenshot", "GET", "")
        Dim base64String = New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
        Dim bytes = Convert.FromBase64String(base64String)
        File.WriteAllBytes(savePath, bytes)
    End Sub

    Public Sub CaptureFullPageScreenshot(savePath As String)
        Dim metrics = GetWindowSize()
        Dim params = New Dictionary(Of String, Object) From {
        {"format", "png"},
        {"fromSurface", True},
        {"clip", New Dictionary(Of String, Object) From {
            {"x", 0},
            {"y", 0},
            {"width", metrics("width")},
            {"height", metrics("height")},
            {"scale", 1.0}
        }}
    }
        Dim payload = New Dictionary(Of String, Object) From {
        {"cmd", "Page.captureScreenshot"},
        {"params", params}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", sJson)
        Dim base64 = New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")("data").ToString()
        Dim bytes = Convert.FromBase64String(base64)
        File.WriteAllBytes(savePath, bytes)
    End Sub


    Public Function FindElementsByCss(selector As String) As String()
        Return FindElementsBy("css selector", selector)
    End Function

    Public Function FindElementsByXpath(selector As String) As String()
        Return FindElementsBy("xpath", selector)
    End Function

    Public Function FindElementsById(selector As String) As String()
        Return FindElementsBy("id", selector)
    End Function

    Public Function FindElementsByName(selector As String) As String()
        Return FindElementsBy("name", selector)
    End Function

    Public Function FindElementsByTagName(selector As String) As String()
        Return FindElementsBy("tag name", selector)
    End Function

    Public Function FindElementsByClassName(selector As String) As String()
        Return FindElementsBy("class name", selector)
    End Function

    Public Function FindElementsByLinkText(selector As String) As String()
        Return FindElementsBy("link text", selector)
    End Function

    Public Function FindElementsByPartialLinkText(selector As String) As String()
        Return FindElementsBy("partial link text", selector)
    End Function

    Public Function FindElementsBy(sUsing As String, selector As String) As String()
        Dim payload = New Dictionary(Of String, Object) From {{"using", sUsing}, {"value", selector}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/elements", "POST", sJson)

        Dim rawArray = New JavaScriptSerializer().Deserialize(Of Object())(resp)
        Dim result As New List(Of String)

        For Each item As Object In rawArray
            Dim elementDict = CType(item, Dictionary(Of String, Object))
            result.Add(elementDict("element-6066-11e4-a52e-4f735466cecf").ToString())
        Next

        Return result.ToArray()
    End Function

    Public Function FindElementByCss(selector As String) As String
        Return FindElementBy("css selector", selector)
    End Function

    Public Function FindElementByXpath(selector As String) As String
        Return FindElementBy("xpath", selector)
    End Function

    Public Function FindElementById(selector As String) As String
        Return FindElementBy("id", selector)
    End Function

    Public Function FindElementByName(selector As String) As String
        Return FindElementBy("name", selector) 'name="email"
    End Function

    Public Function FindElementByTagName(selector As String) As String
        Return FindElementBy("tag name", selector) 'div, input
    End Function

    Public Function FindElementByLinkText(selector As String) As String
        Return FindElementBy("link text", selector) 'Exact match of anchor (<a>) text
    End Function

    Public Function FindElementByPartialLinkText(selector As String) As String
        Return FindElementBy("partial link text", selector) 'single class name
    End Function

    Public Function FindElementBy(sUsing As String, selector As String) As String
        Dim payload = New Dictionary(Of String, Object) From {{"using", sUsing}, {"value", selector}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element", "POST", sJson)
        Dim result = New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
        Return result("element-6066-11e4-a52e-4f735466cecf").ToString()
    End Function

    Public Function GetElementText(elementId As String) As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/text", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    Public Function IsElementDisplayed(elementId As String) As Boolean
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/displayed", "GET", "")
        Return CBool(New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value"))
    End Function

    Public Function IsElementEnabled(elementId As String) As Boolean
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/enabled", "GET", "")
        Return CBool(New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value"))
    End Function

    Public Function IsElementSelected(elementId As String) As Boolean
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/selected", "GET", "")
        Return CBool(New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value"))
    End Function

    Public Sub ClearElement(elementId As String)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/clear", "POST", "{}")
    End Sub

    Public Sub SubmitElement(elementId As String)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/submit", "POST", "{}")
    End Sub

    Public Function GetCssValue(elementId As String, propertyName As String) As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/css/{propertyName}", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    Public Sub SendKeysToElement(elementId As String, keys As String)
        Dim payload = New Dictionary(Of String, Object) From {
        {"text", keys},
        {"value", keys.ToCharArray()}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/value", "POST", sJson)
    End Sub
    Public Sub ClickElement(elementId As String)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/click", "POST", "{}")
    End Sub
    Public Function GetElementAttribute(elementId As String, attributeName As String) As String
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/element/{elementId}/attribute/{attributeName}", "GET", "")
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    Public Sub PerformActions(rawJson As String)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/actions", "POST", rawJson)
    End Sub

    Public Sub MoveToElement(elementId As String)
        Dim json As String = "{" &
        """actions"":[{" &
            """type"":""pointer""," &
            """id"":""mouse""," &
            """parameters"":{""pointerType"":""mouse""}," &
            """actions"":[{" &
                """type"":""pointerMove""," &
                """origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & elementId & """}," &
                """x"":0,""y"":0,""duration"":100" &
            "}]" &
        "}]" &
    "}"
        PerformActions(json)
    End Sub

    Public Sub DragAndDrop(sourceId As String, targetId As String)
        Dim json As String = "{" &
        """actions"":[{" &
            """type"":""pointer""," &
            """id"":""mouse""," &
            """parameters"":{""pointerType"":""mouse""}," &
            """actions"":[{" &
                """type"":""pointerMove""," &
                """origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & sourceId & """}," &
                """x"":0,""y"":0,""duration"":100" &
            "},{" &
                """type"":""pointerDown"",""button"":0" &
            "},{" &
                """type"":""pointerMove""," &
                """origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & targetId & """}," &
                """x"":0,""y"":0,""duration"":100" &
            "},{" &
                """type"":""pointerUp"",""button"":0" &
            "}]" &
        "}]" &
    "}"
        PerformActions(json)
    End Sub


    ' Uploads file to WebDriver and returns the remote path for input[type="file"]
    Public Function UploadFile(localPath As String) As String
        Dim fileBytes = File.ReadAllBytes(localPath)
        Dim base64 = Convert.ToBase64String(fileBytes)
        Dim payload = New Dictionary(Of String, Object) From {
        {"file", base64}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/file", "POST", sJson)
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value").ToString()
    End Function

    ' Uploads file and sends its remote path to a file input element (by element ID)
    Public Sub UploadFileToElement(localPath As String, elementId As String)
        Dim remotePath = UploadFile(localPath)
        SendKeysToElement(elementId, remotePath)
    End Sub

    ' Shortcut: Uploads file to a file input using HTML id
    Public Sub UploadFileById(inputId As String, localPath As String)
        Dim elementId As String = FindElementById(inputId)
        UploadFileToElement(localPath, elementId)
    End Sub

    'Chrome DevTools Protocol (CDP) =======================
    Public Sub SendCdpCommand(command As String, params As Dictionary(Of String, Object))
        Dim payload = New Dictionary(Of String, Object) From {
        {"cmd", command},
        {"params", params}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", sJson)
    End Sub

    Public Function GetBrowserLogs() As Object()
        Dim payload = New Dictionary(Of String, Object) From {{"type", "browser"}}
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/log", "POST", sJson)
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
    End Function

    Public Function GetPerformanceMetrics() As Dictionary(Of String, Object)
        SendCdpCommand("Performance.enable", New Dictionary(Of String, Object))
        Dim payload = New Dictionary(Of String, Object) From {
        {"cmd", "Performance.getMetrics"},
        {"params", New Dictionary(Of String, Object)}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", sJson)
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
    End Function

    Public Sub EnableNetworkLogging()
        SendCdpCommand("Network.enable", New Dictionary(Of String, Object))
    End Sub

    Public Sub EnableConsoleLogging()
        SendCdpCommand("Log.enable", New Dictionary(Of String, Object))
    End Sub

    Public Sub EmulateNetworkConditions(offline As Boolean, latency As Integer, downloadThroughput As Integer, uploadThroughput As Integer)
        Dim params = New Dictionary(Of String, Object) From {
        {"offline", offline},
        {"latency", latency},
        {"downloadThroughput", downloadThroughput},
        {"uploadThroughput", uploadThroughput},
        {"connectionType", "cellular3g"}
    }
        SendCdpCommand("Network.emulateNetworkConditions", params)
    End Sub

    Public Sub EmulateDeviceMetrics(width As Integer, height As Integer, deviceScaleFactor As Double, mobile As Boolean)
        Dim params = New Dictionary(Of String, Object) From {
        {"width", width},
        {"height", height},
        {"deviceScaleFactor", deviceScaleFactor},
        {"mobile", mobile}
    }
        SendCdpCommand("Emulation.setDeviceMetricsOverride", params)
    End Sub

    ' Enable Fetch request interception
    Public Sub EnableRequestInterception()
        Dim fetchEnableParams = New Dictionary(Of String, Object) From {
        {"patterns", New Object() {}}
    }
        SendCdpCommand("Fetch.enable", fetchEnableParams)
    End Sub

    ' Capture DOM snapshot for auditing or static analysis
    Public Function CaptureDomSnapshot() As Object
        Dim params = New Dictionary(Of String, Object) From {
        {"computedStyles", New String() {"color", "font-size", "display"}}
    }
        SendCdpCommand("DOMSnapshot.enable", New Dictionary(Of String, Object))
        Dim payload = New Dictionary(Of String, Object) From {
        {"cmd", "DOMSnapshot.captureSnapshot"},
        {"params", params}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", sJson)
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(resp)("value")
    End Function

    ' Start performance trace
    Public Sub StartPerformanceTrace()
        Dim params = New Dictionary(Of String, Object) From {
        {"categories", "devtools.timeline"},
        {"transferMode", "ReturnAsStream"}
    }
        SendCdpCommand("Tracing.start", params)
    End Sub

    ' Stop performance trace and return base64 stream handle
    Public Function StopPerformanceTrace() As String
        Dim payload = New Dictionary(Of String, Object) From {
        {"cmd", "Tracing.end"},
        {"params", New Dictionary(Of String, Object)}
    }
        Dim sJson = New JavaScriptSerializer().Serialize(payload)
        Dim resp = SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", sJson)
        Return resp ' You may parse the response to extract stream handle for trace data
    End Function

    ' Start precise JS coverage
    Public Sub EnablePreciseCoverage()
        SendCdpCommand("Profiler.enable", New Dictionary(Of String, Object))
        SendCdpCommand("Profiler.startPreciseCoverage", New Dictionary(Of String, Object) From {{"callCount", True}, {"detailed", True}})
    End Sub

    ' Stop and return JS coverage report
    Public Function GetPreciseCoverage() As Object
        Dim result = SendRequest($"http://localhost:{iPort}/session/{sessionId}/chromium/send_command", "POST", New JavaScriptSerializer().Serialize(New Dictionary(Of String, Object) From {
        {"cmd", "Profiler.takePreciseCoverage"},
        {"params", New Dictionary(Of String, Object)}
    }))
        Return New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(result)("value")
    End Function


    '===========================================
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

    ' UpdateDriver =====================================

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

    Private Function GetMajorVersion(fullVersion As String) As String
        Dim i = fullVersion.IndexOf("."c)
        If i > 0 Then
            Return fullVersion.Substring(0, i)
        End If
        Return fullVersion
    End Function

End Class
