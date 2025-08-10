Imports System.Runtime.InteropServices
Imports System.Windows.Automation
Imports System.Diagnostics
Imports System.Text
Imports System.Web.Script.Serialization

<ProgId("IgorKrup.TaskRecorder")>
<Guid("7E8D6E65-7B8B-40C3-9F7D-7D0F0C6C4A11")>
<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class TaskRecorder

    ' ===== Public options =====
    Public MaskKeywords As String = "password;pwd;secret;ssn;token"
    Public CaptureScreenshots As Boolean = False
    Public ScreenshotDirectory As String = ""

    ' ===== Recorded action model =====
    <ComVisible(True)>
    Public Class RecordedAction
        Public TimestampUtc As String
        Public ActionType As String ' click|text|key|window
        Public Button As String     ' left|right|middle (for click)
        Public X As Integer
        Public Y As Integer
        Public WindowTitle As String
        Public ProcessName As String
        Public ElementName As String
        Public ElementClass As String
        Public AutomationId As String
        Public ControlType As String
        Public Text As String
        Public IsSensitive As Boolean
    End Class

    Private _log As New List(Of RecordedAction)()
    Private _currentText As New StringBuilder()
    Private _currentTextSensitive As Boolean = False
    Private _currentFocusKey As String = ""

    ' ===== Low-level hooks =====
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_MBUTTONDOWN As Integer = &H207

    <StructLayout(LayoutKind.Sequential)>
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInteger
        Public scanCode As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure MSLLHOOKSTRUCT
        Public pt As POINT
        Public mouseData As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    Private Delegate Function LowLevelKeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    Private Delegate Function LowLevelMouseProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelKeyboardProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelMouseProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    <DllImport("user32.dll")> Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll")> Private Shared Function GetKeyboardState(lpKeyState() As Byte) As Boolean
    End Function
    <DllImport("user32.dll")> Private Shared Function MapVirtualKey(uCode As UInteger, uMapType As UInteger) As UInteger
    End Function
    <DllImport("user32.dll")> Private Shared Function ToUnicode(wVirtKey As UInteger, wScanCode As UInteger, lpKeyState() As Byte, lpwTransKey As StringBuilder, cchBuff As Integer, wFlags As UInteger) As Integer
    End Function

    Private _kbHook As IntPtr = IntPtr.Zero
    Private _mouseHook As IntPtr = IntPtr.Zero
    Private _kbProc As LowLevelKeyboardProc = AddressOf KeyboardProc
    Private _mouseProc As LowLevelMouseProc = AddressOf MouseProc
    Private ReadOnly _gate As New Object()
    Public ReadOnly Property IsRunning As Boolean
        Get
            Return _kbHook <> IntPtr.Zero OrElse _mouseHook <> IntPtr.Zero
        End Get
    End Property


    ' ===== Public control =====
    Public Sub Start()
        If IsRunning Then Exit Sub
        _kbHook = SetWindowsHookEx(WH_KEYBOARD_LL, _kbProc, IntPtr.Zero, 0UI)
        _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, IntPtr.Zero, 0UI)
    End Sub

    Public Sub [Stop]()
        SyncLock _gate
            FlushTextBuffer()
            If _kbHook <> IntPtr.Zero Then UnhookWindowsHookEx(_kbHook) : _kbHook = IntPtr.Zero
            If _mouseHook <> IntPtr.Zero Then UnhookWindowsHookEx(_mouseHook) : _mouseHook = IntPtr.Zero
        End SyncLock
    End Sub

    Public Sub Clear()
        _log.Clear()
        _currentText.Length = 0
        _currentFocusKey = ""
        _currentTextSensitive = False
    End Sub

    Public Function GetLogJson() As String
        Dim ser = New JavaScriptSerializer()
        ser.MaxJsonLength = Integer.MaxValue
        Return ser.Serialize(_log)
    End Function

    Public Sub SaveLog(path As String)
        System.IO.File.WriteAllText(path, GetLogJson(), Encoding.UTF8)
    End Sub

    ' ===== Exporters =====
    Public Sub ExportVbs(outPath As String)
        System.IO.File.WriteAllText(outPath, GenerateVbs(), Encoding.UTF8)
    End Sub

    Public Function GenerateVbs() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("' Generated by IgorKrup.TaskRecorder")
        sb.AppendLine("Option Explicit")
        sb.AppendLine("Dim ctrl : On Error Resume Next : Set ctrl = CreateObject(""IgorKrup.Control"")")
        sb.AppendLine("If Err.Number <> 0 Then WScript.Echo ""Install/COM-register IgorKrup.Control first."" : WScript.Quit 1")
        sb.AppendLine("On Error Goto 0")
        Dim lastWin As String = ""
        Const q As String = """"

        For Each a In CoalesceText(_log)
            If a.ActionType = "click" OrElse a.ActionType = "text" Then
                If a.WindowTitle <> lastWin Then
                    sb.AppendLine($"ctrl.WinActivate ""{EscapeVbs(a.WindowTitle)}""")
                    sb.AppendLine("WScript.Sleep 200")
                    lastWin = a.WindowTitle
                End If
            End If

            Select Case a.ActionType
                Case "click"
                    Dim btn = If(String.IsNullOrEmpty(a.Button), "left", a.Button)
                    sb.AppendLine($"ctrl.MouseClick {q}{btn}{q}, {a.X}, {a.Y}")

                Case "text"
                    Dim t = If(a.IsSensitive, "***", a.Text)
                    sb.AppendLine($"ctrl.ControlSend {q}{EscapeVbs(a.WindowTitle)}{q}, {q}{q}, {q}{q}, {q}{VbsLiteral(t)}{q}")

                Case "key"
                    sb.AppendLine($"ctrl.ControlSend {q}{EscapeVbs(a.WindowTitle)}{q}, {q}{q}, {q}{q}, {q}[{a.Text}]{q}")
            End Select
        Next

        Return sb.ToString()
    End Function

    Public Sub ExportVbaModule(outPath As String)
        System.IO.File.WriteAllText(outPath, GenerateVbaModule(), Encoding.UTF8)
    End Sub

    Public Function GenerateVbaModule() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("' Generated by IgorKrup.TaskRecorder")
        sb.AppendLine("Option Explicit")
        sb.AppendLine("Public Sub RunRecorded()")
        sb.AppendLine("    Dim ctrl As Object: Set ctrl = CreateObject(""IgorKrup.Control"")")
        Dim lastWin As String = ""
        Const q As String = """"

        For Each a In CoalesceText(_log)
            If a.ActionType = "click" OrElse a.ActionType = "text" Then
                If a.WindowTitle <> lastWin Then
                    sb.AppendLine($"    ctrl.WinActivate ""{EscapeVba(a.WindowTitle)}""")
                    sb.AppendLine("    Application.Wait Now + TimeSerial(0,0,1)")
                    lastWin = a.WindowTitle
                End If
            End If

            Select Case a.ActionType
                Case "click"
                    Dim btn = If(String.IsNullOrEmpty(a.Button), "left", a.Button)
                    sb.AppendLine($"    ctrl.MouseClick {q}{btn}{q}, {a.X}, {a.Y}")

                Case "text"
                    Dim t = If(a.IsSensitive, "***", a.Text)
                    sb.AppendLine($"    ctrl.ControlSend {q}{EscapeVba(a.WindowTitle)}{q}, {q}{q}, {q}{q}, {q}{VbaLiteral(t)}{q}")

                Case "key"
                    sb.AppendLine($"    ctrl.ControlSend {q}{EscapeVba(a.WindowTitle)}{q}, {q}{q}, {q}{q}, {q}[{a.Text}]{q}")
            End Select
        Next
        sb.AppendLine("End Sub")
        Return sb.ToString()
    End Function

    ' ===== Hook callbacks =====
    Private Function KeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 Then
            Dim k As KBDLLHOOKSTRUCT = DirectCast(
                Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT)),
                KBDLLHOOKSTRUCT)

            If wParam = CType(WM_KEYDOWN, IntPtr) OrElse wParam = CType(WM_SYSKEYDOWN, IntPtr) Then
                Dim ch = VkToChar(k.vkCode, k.scanCode)
                Dim isPrintable = ch.Length = 1 AndAlso Not Char.IsControl(ch(0))
                Dim keyName As String = If(isPrintable, ch, VkToName(k.vkCode))

                Dim fw = GetForegroundWindow()
                Dim t = GetWindowTitle(fw)
                Dim procName = GetProcessName(fw)
                Dim focusKey = t & "|" & procName

                Dim sensitiveHere = _currentTextSensitive OrElse LooksSensitiveUnderCursor()

                If isPrintable Then
                    SyncLock _gate
                        If _currentFocusKey <> focusKey Then FlushTextBuffer()
                        _currentFocusKey = focusKey
                        _currentText.Append(ch)
                        _currentTextSensitive = sensitiveHere
                    End SyncLock
                Else
                    FlushTextBuffer()
                    SyncLock _gate
                        _log.Add(New RecordedAction With {
                            .TimestampUtc = DateTime.UtcNow.ToString("o"),
                            .ActionType = "key",
                            .Text = keyName,
                            .WindowTitle = t,
                            .ProcessName = procName
                        })
                    End SyncLock
                End If
            End If
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    Private Function MouseProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 Then
            Dim m As MSLLHOOKSTRUCT = DirectCast(Marshal.PtrToStructure(lParam, GetType(MSLLHOOKSTRUCT)), MSLLHOOKSTRUCT)

            Dim msg = CInt(wParam)
            If msg = WM_LBUTTONDOWN OrElse msg = WM_RBUTTONDOWN OrElse msg = WM_MBUTTONDOWN Then
                FlushTextBuffer()

                Dim btn As String = If(msg = WM_LBUTTONDOWN, "left", If(msg = WM_RBUTTONDOWN, "right", "middle"))
                Dim fw = GetForegroundWindow()
                Dim t = GetWindowTitle(fw)
                Dim procName = GetProcessName(fw)

                Dim elName As String = ""
                Dim elClass As String = ""
                Dim elAutoId As String = ""
                Dim elType As String = ""
                Try
                    Dim ae = AutomationElement.FromPoint(New System.Windows.Point(m.pt.X, m.pt.Y))
                    If ae IsNot Nothing Then
                        elName = SafeStr(ae.Current.Name)
                        elClass = SafeStr(ae.Current.ClassName)
                        elType = SafeStr(ae.Current.ControlType.ProgrammaticName)
                        elAutoId = SafeStr(ae.Current.AutomationId)
                    End If
                Catch
                End Try

                SyncLock _gate
                    _log.Add(New RecordedAction With {
                        .TimestampUtc = DateTime.UtcNow.ToString("o"),
                        .ActionType = "click",
                        .Button = btn,
                        .X = m.pt.X, .Y = m.pt.Y,
                        .WindowTitle = t, .ProcessName = procName,
                        .ElementName = elName, .ElementClass = elClass,
                        .AutomationId = elAutoId, .ControlType = elType
                    })
                End SyncLock
            End If
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    ' ===== Helpers =====
    Private Sub FlushTextBuffer()
        SyncLock _gate
            If _currentText.Length = 0 Then Exit Sub
            Dim fw = GetForegroundWindow()
            Dim t = GetWindowTitle(fw)
            Dim procName = GetProcessName(fw)
            _log.Add(New RecordedAction With {
            .TimestampUtc = DateTime.UtcNow.ToString("o"),
            .ActionType = "text",
            .Text = _currentText.ToString(),
            .WindowTitle = t,
            .ProcessName = procName,
            .IsSensitive = _currentTextSensitive
        })
            _currentText.Length = 0
            _currentTextSensitive = False
        End SyncLock
    End Sub

    Private Function LooksSensitiveUnderCursor() As Boolean
        Try
            Dim p As System.Drawing.Point = System.Windows.Forms.Cursor.Position
            Dim pt As New System.Windows.Point(CDbl(p.X), CDbl(p.Y))

            Dim ae As System.Windows.Automation.AutomationElement = System.Windows.Automation.AutomationElement.FromPoint(pt)
            If ae Is Nothing Then Return False

            Dim name As String = (SafeStr(ae.Current.Name) & "" & SafeStr(ae.Current.AutomationId)).ToLowerInvariant()
            Dim classes As String = SafeStr(ae.Current.ClassName).ToLowerInvariant()

            For Each kw As String In MaskKeywords.ToLowerInvariant().Split(";"c)
                If kw.Length = 0 Then Continue For
                If name.Contains(kw) OrElse classes.Contains(kw) Then Return True
            Next
        Catch
        End Try
        Return False
    End Function

    Private Function GetWindowTitle(hWnd As IntPtr) As String
        If hWnd = IntPtr.Zero Then Return ""
        Dim sb As New StringBuilder(512)
        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    Private Function GetProcessName(hWnd As IntPtr) As String
        Try
            Dim pid As Integer
            GetWindowThreadProcessId(hWnd, pid)
            Dim p = Process.GetProcessById(pid)
            Return p.ProcessName
        Catch
            Return ""
        End Try
    End Function

    Private Function SafeStr(s As String) As String
        If s Is Nothing Then Return ""
        Return s
    End Function

    Private Function VkToName(vk As UInteger) As String
        Try
            Dim k = CType(vk, Windows.Forms.Keys)
            Return k.ToString().ToUpperInvariant()
        Catch
            Return "VK_" & vk
        End Try
    End Function

    Private Function VkToChar(vk As UInteger, scan As UInteger) As String
        Dim ks(255) As Byte
        If Not GetKeyboardState(ks) Then Return ""
        Dim sb As New StringBuilder(4)
        Dim rc = ToUnicode(vk, scan, ks, sb, sb.Capacity, 0UI)
        If rc = 1 Then Return sb.ToString()
        If rc < 0 Then Return "" ' dead key
        Return ""
    End Function

    Private Function CoalesceText(src As List(Of RecordedAction)) As IEnumerable(Of RecordedAction)
        Dim out As New List(Of RecordedAction)()
        Dim buf As New StringBuilder()
        Dim bufSensitive As Boolean = False
        Dim lastWin As String = ""
        For Each a In src
            If a.ActionType = "text" Then
                If lastWin <> a.WindowTitle Then
                    If buf.Length > 0 Then
                        out.Add(New RecordedAction With {.ActionType = "text", .Text = buf.ToString(), .WindowTitle = lastWin, .IsSensitive = bufSensitive})
                        buf.Length = 0
                        bufSensitive = False
                    End If
                    lastWin = a.WindowTitle
                End If
                buf.Append(a.Text)
                bufSensitive = bufSensitive OrElse a.IsSensitive
            Else
                If buf.Length > 0 Then
                    out.Add(New RecordedAction With {.ActionType = "text", .Text = buf.ToString(), .WindowTitle = lastWin, .IsSensitive = bufSensitive})
                    buf.Length = 0
                    bufSensitive = False
                End If
                out.Add(a)
            End If
        Next
        If buf.Length > 0 Then
            out.Add(New RecordedAction With {.ActionType = "text", .Text = buf.ToString(), .WindowTitle = lastWin, .IsSensitive = bufSensitive})
        End If
        Return out
    End Function

    Private Function EscapeVbs(s As String) As String
        If s Is Nothing Then Return ""
        ' In VBScript, only quotes need doubling inside string literals
        Return s.Replace("""", """""")
    End Function

    Private Function EscapeVba(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("""", """""")
    End Function

    Private Function VbsLiteral(s As String) As String
        Dim x = EscapeVbs(s)
        x = x.Replace(vbCrLf, """ & vbCrLf & """).Replace(vbCr, """ & vbCr & """).Replace(vbLf, """ & vbLf & """)
        Return x
    End Function

    Private Function VbaLiteral(s As String) As String
        Dim x = EscapeVba(s)
        x = x.Replace(vbCrLf, """ & vbCrLf & """).Replace(vbCr, """ & vbCr & """).Replace(vbLf, """ & vbLf & """)
        Return x
    End Function

End Class