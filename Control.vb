' Full implementation of an AutoItX3.Control replacement in VB.NET

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Drawing

Public Class Control

    ' === Enums and Constants ===
    Public winTitleMatchMode As Integer = 1 ' 1: starts with, 2: contains, 3: exact, 4: regex
    Public openAiApiKey As String = "" ' Environment.GetEnvironmentVariable("OPENAI_API_KEY")

    Private sendKeyDelay As Integer = 0
    Private Shared ReadOnly INPUT_SIZE As Integer = Marshal.SizeOf(GetType(INPUT))

    ' Mouse event constants
    Private Const MOUSEEVENTF_LEFTDOWN As UInteger = &H2
    Private Const MOUSEEVENTF_LEFTUP As UInteger = &H4
    Private Const MOUSEEVENTF_RIGHTDOWN As UInteger = &H8
    Private Const MOUSEEVENTF_RIGHTUP As UInteger = &H10
    Private Const MOUSEEVENTF_MIDDLEDOWN As UInteger = &H20
    Private Const MOUSEEVENTF_MIDDLEUP As UInteger = &H40

    ' Input event constants
    Private Const INPUT_KEYBOARD = 1
    Private Const KEYEVENTF_KEYUP = &H2
    Private Const KEYEVENTF_UNICODE = &H4
    Private Const VK_SHIFT = &HA0
    Private Const VK_CONTROL = &HA2
    Private Const VK_MENU = &HA4

    ' Windows API structures
    <StructLayout(LayoutKind.Sequential)> Private Structure RECT
        Public Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> Private Structure INPUT
        Public type As Integer
        Public u As INPUTUNION
    End Structure

    <StructLayout(LayoutKind.Explicit)> Private Structure INPUTUNION
        <FieldOffset(0)> Public ki As KEYBDINPUT
    End Structure

    <StructLayout(LayoutKind.Sequential)> Private Structure KEYBDINPUT
        Public wVk As UShort, wScan As UShort, dwFlags As UInteger, time As Integer, dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    ' === Win32 API Declarations ===
    <DllImport("user32.dll")> Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")> Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function
    <DllImport("user32.dll")> Private Shared Sub SetCursorPos(x As Integer, y As Integer)
    End Sub
    <DllImport("user32.dll")> Private Shared Sub mouse_event(dwFlags As UInteger, dx As UInteger, dy As UInteger, cButtons As UInteger, dwExtraInfo As IntPtr)
    End Sub
    <DllImport("user32.dll")> Private Shared Function SendInput(nInputs As Integer, ByRef pInputs As INPUT, cbSize As Integer) As UInteger
    End Function
    <DllImport("user32.dll")> Private Shared Function VkKeyScan(ch As Char) As Short
    End Function
    <DllImport("user32.dll")> Private Shared Function FindWindow(lpClass As String, lpTitle As String) As IntPtr
    End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll")> Private Shared Function EnumWindows(callback As EnumWindowsProc, lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")> Private Shared Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")> Private Shared Function MapVirtualKey(uCode As UInteger, uMapType As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetCursorPos(ByRef lpPoint As POINT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowEx(hwndParent As IntPtr, hwndChildAfter As IntPtr, lpszClass As String, lpszWindow As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetFocus(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SetWindowText(hWnd As IntPtr, lpString As String) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetDlgItem(hWnd As IntPtr, nIDDlgItem As Integer) As IntPtr
    End Function


    Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

    ' === Automation Functions ===

    Public Sub Opt(optionName As String, value As Integer)
        Select Case optionName.ToUpper()
            Case "WINTITLEMATCHMODE"
                winTitleMatchMode = value
            Case "SENDKEYDELAY"
                sendKeyDelay = value
        End Select
    End Sub

    Private lastWindowHandle As IntPtr = IntPtr.Zero

    Public Sub New()
        openAiApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") & ""
        'setx OPENAI_API_KEY "sk-yourkeyhere"
        'System.Configuration.ConfigurationManager.AppSettings("OPENAI_API_KEY")
    End Sub

    Public Function WinExists(titleHint As String) As Boolean
        lastWindowHandle = FindWindowByTitleHint(titleHint)
        Return lastWindowHandle <> IntPtr.Zero
    End Function

    Public Sub WinActivate(titleHint As String)
        If lastWindowHandle = IntPtr.Zero Then
            lastWindowHandle = FindWindowByTitleHint(titleHint)
        End If
        If lastWindowHandle <> IntPtr.Zero Then
            SetForegroundWindow(lastWindowHandle)
        End If
    End Sub

    Public Function WinWait(titleHint As String, Optional textHint As String = "", Optional timeoutSec As Integer = 0) As Boolean
        ' Wait for window to exist, then activate and wait for it to become active
        Dim sw As New Stopwatch()
        sw.Start()
        Do
            If WinExists(titleHint) Then
                WinActivate(titleHint)
                Return True
            End If
            Thread.Sleep(250)
        Loop While sw.ElapsedMilliseconds < timeoutSec * 1000 Or timeoutSec = 0
        Return False
    End Function

    Public Function WinWaitClose(titleHint As String, Optional timeoutSec As Integer = 0) As Boolean
        ' Waits until the specified window is closed or timeout elapses.
        ' Returns True if the window was closed, False if timed out.

        Dim sw As New Stopwatch()
        sw.Start()

        Do
            If Not WinExists(titleHint) Then
                Return True ' Window is gone
            End If

            Thread.Sleep(250)

            ' If timeoutSec is 0, wait forever
            If timeoutSec > 0 AndAlso sw.Elapsed.TotalSeconds >= timeoutSec Then
                Exit Do
            End If
        Loop

        Return False ' Timed out
    End Function


    Public Function WinWaitActive(titleHint As String, Optional timeoutSec As Integer = 5) As Boolean
        WinActivate(titleHint)
        Dim sw As New Stopwatch()
        sw.Start()
        Do While sw.ElapsedMilliseconds < timeoutSec * 1000
            If GetForegroundWindow() = lastWindowHandle Then Return True
            Thread.Sleep(100)
        Loop
        Return False
    End Function

    Public Sub WinClose(titleHint As String)
        Dim hWnd = FindWindowByTitleHint(titleHint)
        If hWnd <> IntPtr.Zero Then PostMessage(hWnd, &H10, IntPtr.Zero, IntPtr.Zero) ' WM_CLOSE
    End Sub

    Public Sub Run(exePath As String)
        Process.Start(exePath)
    End Sub

    Public Function ProcessExists(name As String) As Boolean
        Dim n = name.ToLower().Replace(".exe", "")
        Return Process.GetProcessesByName(n).Length > 0
    End Function

    Public Function WinGetTitle(titleHint As String) As String
        Dim hWnd = FindWindowByTitleHint(titleHint)
        If hWnd = IntPtr.Zero Then Return String.Empty
        Dim length = GetWindowTextLength(hWnd)
        If length = 0 Then Return ""
        Dim sb As New StringBuilder(length + 1)
        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    Public Function WinGetHandle(titleHint As String) As String
        Dim hWnd = FindWindowByTitleHint(titleHint)
        Return hWnd.ToInt64().ToString("X")
    End Function

    Public Function WinGetPosX(titleHint As String) As Integer
        Return GetWindowRectFor(titleHint).Left
    End Function

    Public Function WinGetPosY(titleHint As String) As Integer
        Return GetWindowRectFor(titleHint).Top
    End Function

    Public Function WinGetPosWidth(titleHint As String) As Integer
        Dim r = GetWindowRectFor(titleHint)
        Return r.Right - r.Left
    End Function

    Public Function WinGetPosHeight(titleHint As String) As Integer
        Dim r = GetWindowRectFor(titleHint)
        Return r.Bottom - r.Top
    End Function

    Public Function WinGetPos(titleHint As String) As Integer()
        Dim r = GetWindowRectFor(titleHint)
        Return {r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top}
    End Function

    'Return New Object() {x + sx, y + sy, bmpSearch.Width, bmpSearch.Height}

    Public Sub MouseClick(button As String, x As Integer, y As Integer)
        SetCursorPos(x, y)
        Thread.Sleep(50)
        MouseDown(button)
        MouseUp(button)
    End Sub

    Public Sub MouseDown(button As String)
        Select Case button.ToLower()
            Case "left" : mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)
            Case "right" : mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)
            Case "middle" : mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, IntPtr.Zero)
        End Select
    End Sub

    Public Sub MouseUp(button As String)
        Select Case button.ToLower()
            Case "left" : mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)
            Case "right" : mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)
            Case "middle" : mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, IntPtr.Zero)
        End Select
    End Sub

    Public Sub MouseMove(x As Integer, y As Integer)
        SetCursorPos(x, y)
    End Sub

    Public Sub Sleep(ms As Integer)
        Thread.Sleep(ms)
    End Sub

    Private Function ExtractStackedModifiers(ByRef text As String, ByRef i As Integer) As List(Of UShort)
        Dim mods As New List(Of UShort)
        While i < text.Length AndAlso (text(i) = "^"c Or text(i) = "+"c Or text(i) = "!"c)
            Select Case text(i)
                Case "^"c : mods.Add(VK_CONTROL)
                Case "+"c : mods.Add(VK_SHIFT)
                Case "!"c : mods.Add(VK_MENU)
            End Select
            i += 1
        End While
        Return mods
    End Function

    Private Sub SendAsciiChar(ch As Char)
        Dim vkFull = VkKeyScan(ch)
        Dim vk = CByte(vkFull And &HFF)
        Dim shiftState = (vkFull >> 8) And &HFF
        Dim scan = VirtualKeyToScanCode(vk)
        Dim ext = IsExtendedKey(scan)

        If (shiftState And 1) <> 0 Then SendKeyScan(VirtualKeyToScanCode(VK_SHIFT), True, False)
        SendKeyScan(scan, True, ext)
        Thread.Sleep(sendKeyDelay)
        SendKeyScan(scan, False, ext)
        If (shiftState And 1) <> 0 Then SendKeyScan(VirtualKeyToScanCode(VK_SHIFT), False, False)
    End Sub

    Private Sub SendUnicodeChar(ch As Char)
        Dim backup As String = SafeGetClipboardText()

        If Not SafeSetClipboardText(ch.ToString()) Then
            Throw New Exception("Failed to set clipboard text.")
            Return
        End If

        ' Simulate Ctrl+V
        SendKeyScan(VirtualKeyToScanCode(VK_CONTROL), True, False)
        SendKeyScan(VirtualKeyToScanCode(&H56), True, False) ' V
        Thread.Sleep(sendKeyDelay)
        SendKeyScan(VirtualKeyToScanCode(&H56), False, False)
        SendKeyScan(VirtualKeyToScanCode(VK_CONTROL), False, False)

        Thread.Sleep(sendKeyDelay)

        SafeSetClipboardText(backup)
    End Sub

    Private Function SafeGetClipboardText(Optional retries As Integer = 10) As String
        For i = 1 To retries
            Try
                Return Clipboard.GetText()
            Catch ex As ExternalException
                Thread.Sleep(50)
            End Try
        Next
        Return String.Empty
    End Function

    Private Function SafeSetClipboardText(text As String, Optional retries As Integer = 10) As Boolean
        For i = 1 To retries
            Try
                Clipboard.SetText(text)
                Return True
            Catch ex As ExternalException
                Thread.Sleep(50)
            End Try
        Next
        Return False
    End Function

    Public Sub Send(text As String)
        Dim rawMode As Boolean = False
        Dim unicodeMode As Boolean = False

        Dim vk As UShort
        Dim scan As UShort
        Dim ext As Boolean

        Dim i As Integer = 0
        Dim modStack As New Stack(Of UShort)

        While i < text.Length
            ' === Escaped Braces ===
            If i + 3 <= text.Length AndAlso text.Substring(i, 3) = "{{}" Then
                SendAsciiChar("{"c)
                i += 3
                Continue While
            ElseIf i + 2 <= text.Length AndAlso text.Substring(i, 2) = "}}" Then
                SendAsciiChar("}"c)
                i += 2
                Continue While
            End If

            ' === Block Parser ===
            If text(i) = "{"c Then
                Dim endIdx = text.IndexOf("}"c, i)
                If endIdx < 0 Then Exit While

                Dim raw = text.Substring(i + 1, endIdx - i - 1).Trim()

                Select Case raw.ToUpper()
                    Case "RAW" : rawMode = True
                    Case "/RAW" : rawMode = False
                    Case "UNICODE" : unicodeMode = True
                    Case "/UNICODE" : unicodeMode = False
                    Case Else
                        If raw.ToUpper().StartsWith("SLEEP ") Then
                            Dim sleepMs As Integer
                            If Integer.TryParse(raw.Substring(6).Trim(), sleepMs) Then Thread.Sleep(sleepMs)
                        ElseIf raw.ToUpper().StartsWith("ASC ") Then
                            Dim code As Integer
                            If Integer.TryParse(raw.Substring(4).Trim(), code) Then
                                SendUnicodeChar(ChrW(code))
                            End If
                        Else
                            Dim parts = raw.Split(" "c)
                            Dim key = parts(0).ToUpper()
                            Dim count As Integer = If(parts.Length > 1, Integer.Parse(parts(1)), 1)

                            vk = GetVirtualKey(key.Replace("DOWN", "").Replace("UP", "").Trim())
                            scan = VirtualKeyToScanCode(vk)
                            ext = IsExtendedKey(scan)

                            If key.EndsWith("DOWN") Then
                                SendKeyScan(scan, True, ext)
                                modStack.Push(scan)
                            ElseIf key.EndsWith("UP") Then
                                SendKeyScan(scan, False, ext)
                            Else
                                For j = 1 To count
                                    SendKeyScan(scan, True, ext)
                                    Thread.Sleep(sendKeyDelay)
                                    SendKeyScan(scan, False, ext)
                                Next
                            End If
                        End If
                End Select

                i = endIdx + 1
                Continue While
            End If

            ' === RAW Mode ===
            If rawMode Then
                SendAsciiChar(text(i))
                i += 1
                Continue While
            End If

            ' === UNICODE Mode ===
            If unicodeMode Then
                SendUnicodeChar(text(i))
                i += 1
                Continue While
            End If

            ' === Normal Mode ===
            Dim mods = ExtractStackedModifiers(text, i)
            If i >= text.Length Then Exit While
            Dim ch = text(i)

            Dim vkFull = VkKeyScan(ch)
            vk = CByte(vkFull And &HFF)
            Dim shiftState = (vkFull >> 8) And &HFF
            scan = VirtualKeyToScanCode(vk)
            ext = IsExtendedKey(scan)

            For Each m In mods
                SendKeyScan(VirtualKeyToScanCode(m), True, IsExtendedKey(VirtualKeyToScanCode(m)))
            Next

            If (shiftState And 1) <> 0 Then
                SendKeyScan(VirtualKeyToScanCode(VK_SHIFT), True, False)
            End If

            SendKeyScan(scan, True, ext)
            Thread.Sleep(sendKeyDelay)
            SendKeyScan(scan, False, ext)

            If (shiftState And 1) <> 0 Then
                SendKeyScan(VirtualKeyToScanCode(VK_SHIFT), False, False)
            End If

            For Each m In mods.AsEnumerable().Reverse()
                SendKeyScan(VirtualKeyToScanCode(m), False, IsExtendedKey(VirtualKeyToScanCode(m)))
            Next

            Thread.Sleep(sendKeyDelay)
            i += 1
        End While

        ' === Cleanup remaining held keys from KEYDOWNs ===
        While modStack.Count > 0
            Dim s = modStack.Pop()
            SendKeyScan(s, False, IsExtendedKey(s))
        End While
    End Sub

    Private Function VirtualKeyToScanCode(vk As UShort) As UShort
        Return CShort(MapVirtualKey(vk, 0))
    End Function

    Private Function IsExtendedKey(scanCode As UShort) As Boolean
        ' Minimal reliable list of extended keys
        Return scanCode = &H1D OrElse  ' Right Ctrl
           scanCode = &H38 OrElse  ' Right Alt
           scanCode = &H35 OrElse  ' NumPad /
           scanCode = &H1C OrElse  ' NumPad Enter
           scanCode = &H4B OrElse  ' Left
           scanCode = &H4D OrElse  ' Right
           scanCode = &H48 OrElse  ' Up
           scanCode = &H50 OrElse  ' Down
           scanCode = &H47 OrElse  ' Home
           scanCode = &H4F OrElse  ' End
           scanCode = &H49 OrElse  ' Page Up
           scanCode = &H51 OrElse  ' Page Down
           scanCode = &H52 OrElse  ' Insert
           scanCode = &H53         ' Delete
    End Function

    Private Function GetVirtualKeyFromChar(ch As Char) As UShort
        Return CByte(VkKeyScan(ch) And &HFF)
    End Function
    Private Function GetVirtualKey(name As String) As UShort
        Select Case name
            Case "ENTER" : Return &HD
            Case "TAB" : Return &H9
            Case "ESC", "ESCAPE" : Return &H1B
            Case "SPACE" : Return &H20
            Case "LEFT" : Return &H25
            Case "UP" : Return &H26
            Case "RIGHT" : Return &H27
            Case "DOWN" : Return &H28
            Case "CTRL", "CONTROL" : Return VK_CONTROL
            Case "ALT" : Return VK_MENU
            Case "SHIFT" : Return VK_SHIFT
            Case "DEL", "DELETE" : Return &H2E
            Case "BACKSPACE" : Return &H8
            Case "HOME" : Return &H24
            Case "END" : Return &H23
            Case "PGUP" : Return &H21
            Case "PGDN" : Return &H22
            Case Else
                If name.StartsWith("F") AndAlso Integer.TryParse(name.Substring(1), Nothing) Then
                    Return CByte(&H70 + Integer.Parse(name.Substring(1)) - 1)
                End If
                Return GetVirtualKeyFromChar(name(0))
        End Select
    End Function

    Private Function FindWindowByTitleHint(hint As String) As IntPtr
        Dim result As IntPtr = IntPtr.Zero
        EnumWindows(Function(hWnd, lParam)
                        Dim len = GetWindowTextLength(hWnd)
                        If len = 0 Then Return True
                        Dim sb As New StringBuilder(len + 1)
                        GetWindowText(hWnd, sb, sb.Capacity)
                        Dim title = sb.ToString()
                        Select Case winTitleMatchMode
                            Case 1 : If title.StartsWith(hint, StringComparison.OrdinalIgnoreCase) Then result = hWnd : Return False
                            Case 2 : If title.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0 Then result = hWnd : Return False
                            Case 3 : If title.Equals(hint, StringComparison.OrdinalIgnoreCase) Then result = hWnd : Return False
                            Case 4 : If System.Text.RegularExpressions.Regex.IsMatch(title, hint, System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then result = hWnd : Return False
                        End Select
                        Return True
                    End Function, IntPtr.Zero)
        Return result
    End Function

    Private Function GetWindowRectFor(titleHint As String) As RECT
        Dim hWnd = FindWindowByTitleHint(titleHint)
        Dim r As RECT
        GetWindowRect(hWnd, r)
        Return r
    End Function

    Public Function MouseGetPosX() As Integer
        Dim pt As POINT
        If GetCursorPos(pt) Then
            Return pt.X
        Else
            Return -1 ' Or throw an exception/log failure
        End If
    End Function

    Public Function MouseGetPosY() As Integer
        Dim pt As POINT
        If GetCursorPos(pt) Then
            Return pt.Y
        Else
            Return -1
        End If
    End Function

    Private Function FindControl(parentTitle As String, controlClass As String, controlText As String) As IntPtr
        Dim parentHwnd = FindWindowByTitleHint(parentTitle)
        If parentHwnd = IntPtr.Zero Then Return IntPtr.Zero

        Dim child As IntPtr = IntPtr.Zero
        Do
            child = FindWindowEx(parentHwnd, child, If(controlClass = "", Nothing, controlClass), If(controlText = "", Nothing, controlText))
            If child <> IntPtr.Zero Then Return child
        Loop While child <> IntPtr.Zero

        Return IntPtr.Zero
    End Function

    Public Sub ControlSend(windowTitle As String, controlClass As String, controlText As String, sendText As String)
        'Simulates key presses to a control (just like a user typing).

        Dim hCtrl = FindControl(windowTitle, controlClass, controlText)
        If hCtrl = IntPtr.Zero Then Return

        SetFocus(hCtrl) ' Optional: give it focus

        For Each ch As Char In sendText
            SendMessage(hCtrl, &H102, CType(AscW(ch), IntPtr), IntPtr.Zero) ' WM_CHAR
        Next
    End Sub

    Public Sub ControlFocus(windowTitle As String, controlClass As String, controlText As String)
        Dim hCtrl = FindControl(windowTitle, controlClass, controlText)
        If hCtrl <> IntPtr.Zero Then SetFocus(hCtrl)
    End Sub

    Public Sub ControlClick(windowTitle As String, controlClass As String, controlText As String)
        Dim hCtrl = FindControl(windowTitle, controlClass, controlText)
        If hCtrl = IntPtr.Zero Then Return

        Const BM_CLICK As UInteger = &HF5
        SendMessage(hCtrl, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
    End Sub

    Public Sub ControlSetText(windowTitle As String, controlClass As String, controlText As String, newText As String)
        'Doesn't trigger keypress events or hotkeys
        Dim hCtrl = FindControl(windowTitle, controlClass, controlText)
        If hCtrl <> IntPtr.Zero Then SetWindowText(hCtrl, newText)
    End Sub

    Public Function ControlGetText(windowTitle As String, controlClass As String, controlText As String) As String
        Dim hCtrl = FindControl(windowTitle, controlClass, controlText)
        If hCtrl = IntPtr.Zero Then Return String.Empty

        Dim length = GetWindowTextLength(hCtrl)
        If length <= 0 Then Return String.Empty

        Dim sb As New StringBuilder(length + 1)
        GetWindowText(hCtrl, sb, sb.Capacity)
        Return sb.ToString()
    End Function


    Private Function ColorFromHex(hex As String) As System.Drawing.Color
        If hex.StartsWith("#") Then hex = hex.Substring(1)
        If hex.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then hex = hex.Substring(2)

        If hex.Length <> 6 Then Throw New ArgumentException("Hex color must be 6 characters (RRGGBB)")

        Dim r = Convert.ToInt32(hex.Substring(0, 2), 16)
        Dim g = Convert.ToInt32(hex.Substring(2, 2), 16)
        Dim b = Convert.ToInt32(hex.Substring(4, 2), 16)

        Return System.Drawing.Color.FromArgb(r, g, b)
    End Function

    Public Function PixelSearch(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, hexColor As String) As System.Drawing.Point?
        Dim targetColor As System.Drawing.Color = ColorFromHex(hexColor)
        Dim width As Integer = x2 - x1
        Dim height As Integer = y2 - y1

        Dim bmp As New System.Drawing.Bitmap(width, height)
        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
            g.CopyFromScreen(x1, y1, 0, 0, New System.Drawing.Size(width, height))
        End Using

        For y As Integer = 0 To height - 1
            For x As Integer = 0 To width - 1
                If bmp.GetPixel(x, y).ToArgb() = targetColor.ToArgb() Then
                    Return New System.Drawing.Point(x1 + x, y1 + y)
                End If
            Next
        Next

        Return Nothing
    End Function


    Public Function PixelGetColor(x As Integer, y As Integer) As String
        Dim bmp As New System.Drawing.Bitmap(1, 1)
        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
            g.CopyFromScreen(x, y, 0, 0, New System.Drawing.Size(1, 1))
        End Using
        Dim c As System.Drawing.Color = bmp.GetPixel(0, 0)
        Return $"{c.R:X2}{c.G:X2}{c.B:X2}" ' Return as "RRGGBB"
    End Function

    Public Sub ScreenCapture(filePath As String, Optional x As Integer = 0, Optional y As Integer = 0, Optional width As Integer = 0, Optional height As Integer = 0)
        Dim bounds As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.Bounds
        If width <= 0 Then width = bounds.Width
        If height <= 0 Then height = bounds.Height

        Dim bmp As New System.Drawing.Bitmap(width, height)

        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
            g.CopyFromScreen(x, y, 0, 0, New System.Drawing.Size(width, height))
        End Using

        bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png)
    End Sub

    Public Function SearchImage2(imagePath As String, winPos As Object, tolerance As Integer, Optional returnCenter As Boolean = False) As Object
        If winPos Is Nothing Then Return Nothing

        Dim arr As System.Array = TryCast(winPos, System.Array)
        If arr Is Nothing OrElse arr.Rank <> 1 OrElse arr.Length < 4 Then Return Nothing

        Dim lb As Integer = arr.GetLowerBound(0)

        Dim x As Integer = CInt(arr.GetValue(lb + 0))
        Dim y As Integer = CInt(arr.GetValue(lb + 1))
        Dim w As Integer = CInt(arr.GetValue(lb + 2))
        Dim h As Integer = CInt(arr.GetValue(lb + 3))

        If w <= 0 OrElse h <= 0 Then Return Nothing ' avoid GDI+ ArgumentException

        Return SearchImage(imagePath, x, y, w, h, tolerance, returnCenter)
    End Function

    Public Function SearchImage(imagePath As String, x As Integer, y As Integer, width As Integer, height As Integer,
                                tolerance As Integer, Optional returnCenter As Boolean = False) As Object

        If Not IO.File.Exists(imagePath) Then Return Nothing

        Using bmpSearch As New Bitmap(imagePath)
            Using bmpScreen As New Bitmap(width, height)
                Using g As Graphics = Graphics.FromImage(bmpScreen)
                    g.CopyFromScreen(New System.Drawing.Point(x, y), System.Drawing.Point.Empty, bmpScreen.Size)
                End Using

                For sy As Integer = 0 To bmpScreen.Height - bmpSearch.Height
                    For sx As Integer = 0 To bmpScreen.Width - bmpSearch.Width
                        Dim match As Boolean = True

                        For j As Integer = 0 To bmpSearch.Height - 1
                            For i As Integer = 0 To bmpSearch.Width - 1
                                Dim c1 = bmpSearch.GetPixel(i, j)
                                Dim c2 = bmpScreen.GetPixel(sx + i, sy + j)

                                If Not CompareColors(c1, c2, tolerance) Then
                                    match = False
                                    Exit For
                                End If
                            Next
                            If Not match Then Exit For
                        Next

                        If match Then
                            If returnCenter Then
                                ' Return center point
                                Dim centerX As Integer = x + sx + bmpSearch.Width \ 2
                                Dim centerY As Integer = y + sy + bmpSearch.Height \ 2
                                Return New Object() {centerX, centerY}
                            Else
                                ' Return top-left and size
                                Return New Object() {x + sx, y + sy, bmpSearch.Width, bmpSearch.Height}
                            End If
                        End If
                    Next
                Next
            End Using
        End Using

        Return Nothing
    End Function

    Private Function CompareColors(c1 As Color, c2 As Color, tolerance As Integer) As Boolean
        Return Math.Abs(Int(c1.R) - Int(c2.R)) <= tolerance AndAlso
               Math.Abs(Int(c1.G) - Int(c2.G)) <= tolerance AndAlso
               Math.Abs(Int(c1.B) - Int(c2.B)) <= tolerance
    End Function

    ' Overload for FindTextOnScreen that accepts WinGetPos array
    Public Function FindTextOnScreen2(winPos As Integer(), searchText As String,
                                     Optional tessdataPath As String = "", Optional lang As String = "eng",
                                     Optional returnCenter As Boolean = False) As Object

        If winPos Is Nothing Then Return Nothing

        Dim arr As System.Array = TryCast(winPos, System.Array)
        If arr Is Nothing OrElse arr.Rank <> 1 OrElse arr.Length < 4 Then Return Nothing

        Dim lb As Integer = arr.GetLowerBound(0)

        Dim x As Integer = CInt(arr.GetValue(lb + 0))
        Dim y As Integer = CInt(arr.GetValue(lb + 1))
        Dim w As Integer = CInt(arr.GetValue(lb + 2))
        Dim h As Integer = CInt(arr.GetValue(lb + 3))

        If w <= 0 OrElse h <= 0 Then Return Nothing ' avoid GDI+ ArgumentException

        Return FindTextOnScreen(x, y, w, h, searchText, tessdataPath, lang, returnCenter)
    End Function

    Public Function FindTextOnScreen(screenX As Integer, screenY As Integer, width As Integer, height As Integer, searchText As String,
                                 Optional tessdataPath As String = "",
                                 Optional lang As String = "eng",
                                 Optional returnCenter As Boolean = False) As Object

        If tessdataPath = "" OrElse Not IO.Directory.Exists(tessdataPath) Then
            tessdataPath = GetTessdataPath()
        End If

        If Not IO.Directory.Exists(tessdataPath) Then
            Throw New IO.DirectoryNotFoundException("Tessdata folder not found: " & tessdataPath)
        End If

        Dim tempImagePath As String = IO.Path.GetTempFileName() & ".png"
        Dim tempOutputPath As String = IO.Path.GetTempFileName()
        Dim tesseractExe As String = "tesseract.exe" ' Full path if not in PATH

        Try
            ' Capture screen area
            Using bmp As New Drawing.Bitmap(width, height, Drawing.Imaging.PixelFormat.Format32bppArgb)
                Using g As Drawing.Graphics = Drawing.Graphics.FromImage(bmp)
                    g.CopyFromScreen(screenX, screenY, 0, 0, bmp.Size, Drawing.CopyPixelOperation.SourceCopy)
                End Using
                bmp.Save(tempImagePath, Drawing.Imaging.ImageFormat.Png)
            End Using

            ' Run tesseract with TSV output
            Dim psi As New ProcessStartInfo()
            psi.FileName = tesseractExe
            psi.Arguments = $"""{tempImagePath}"" ""{tempOutputPath}"" -l {lang} --tessdata-dir ""{tessdataPath}"" tsv"
            psi.UseShellExecute = False
            psi.CreateNoWindow = True
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True

            Using proc As Process = Process.Start(psi)
                proc.WaitForExit()
            End Using

            ' Parse TSV result
            Dim tsvPath As String = tempOutputPath & ".tsv"
            If IO.File.Exists(tsvPath) Then
                Dim lines = IO.File.ReadAllLines(tsvPath)
                For i As Integer = 1 To lines.Length - 1 ' Skip header line
                    Dim parts = lines(i).Split(ControlChars.Tab)
                    If parts.Length >= 12 Then
                        Dim word = parts(11).Trim().ToLower()
                        If word = searchText.Trim().ToLower() Then
                            Dim left = Integer.Parse(parts(6))
                            Dim top = Integer.Parse(parts(7))
                            Dim w = Integer.Parse(parts(8))
                            Dim h = Integer.Parse(parts(9))

                            If returnCenter Then
                                Dim cx = screenX + left + (w \ 2)
                                Dim cy = screenY + top + (h \ 2)
                                Return New Object() {cx, cy}
                            Else
                                Return New Integer() {screenX + left, screenY + top, w, h}
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            Return "[ERROR running tesseract TSV] " & ex.Message

        Finally
            ' Cleanup
            If IO.File.Exists(tempImagePath) Then IO.File.Delete(tempImagePath)
            If IO.File.Exists(tempOutputPath & ".tsv") Then IO.File.Delete(tempOutputPath & ".tsv")
        End Try

        Return Nothing
    End Function

    Public Shared Function GetTessdataPath() As String

        Dim sPath As String = "C:\Program Files\Tesseract-OCR\tessdata"
        If IO.Directory.Exists(sPath) Then
            Return sPath
        End If

        Dim uninstallKeyPaths As String() = {
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Tesseract-OCR",
        "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Tesseract-OCR"
        }

        For Each keyPath As String In uninstallKeyPaths
            Try
                Dim uninstallStr As Object = Microsoft.Win32.Registry.GetValue(keyPath, "UninstallString", Nothing)
                If uninstallStr IsNot Nothing Then
                    ' Strip the executable name to get install folder
                    Dim installDir As String = System.IO.Path.GetDirectoryName(uninstallStr.ToString())
                    If Not String.IsNullOrEmpty(installDir) Then
                        Dim tessdataPath As String = System.IO.Path.Combine(installDir, "tessdata")
                        If System.IO.Directory.Exists(tessdataPath) Then
                            Return tessdataPath
                        End If
                    End If
                End If
            Catch
                ' Continue to next registry path if there's an error
            End Try
        Next

        Return ""
    End Function







    Public Function Vision_TextboxFocused(imagePath As String, ByRef bbox As Integer()) As Boolean
        Using bmp As New Bitmap(imagePath)
            Return Vision_TextboxFocused(bmp, bbox)
        End Using
    End Function

    Public Function Vision_TextboxFocused(bmp As Bitmap, ByRef bbox As Integer()) As Boolean
        bbox = Nothing
        Dim dataUrl = BitmapToDataUrlPng(bmp)

        Dim prompt As String =
            "Analyze this UI screenshot and return ONLY valid JSON:
        { ""textbox_focused"": boolean, ""textbox_bbox"": [x,y,w,h] or null }
        Rules:
        - ""textbox_focused"" is true only if the text caret is inside a textbox OR the textbox clearly shows focused/active styling.
        - ""textbox_bbox"" are absolute screen coordinates [x,y,w,h] of the active textbox, else null."

        Dim obj = Vision_RequestJson(dataUrl, "gpt-4o", prompt)
        If obj Is Nothing Then Return False

        Dim focused As Boolean = False
        If obj.ContainsKey("textbox_focused") Then
            Try : focused = Convert.ToBoolean(obj("textbox_focused")) : Catch : focused = False : End Try
        End If

        If obj.ContainsKey("textbox_bbox") AndAlso Not obj("textbox_bbox") Is Nothing Then
            Dim arr = TryCast(obj("textbox_bbox"), Object())
            If arr Is Nothing Then
                Dim al = TryCast(obj("textbox_bbox"), System.Collections.ArrayList)
                If Not al Is Nothing Then arr = al.ToArray()
            End If
            If Not arr Is Nothing AndAlso arr.Length = 4 Then
                Try
                    bbox = New Integer() {CInt(arr(0)), CInt(arr(1)), CInt(arr(2)), CInt(arr(3))}
                Catch
                    bbox = Nothing
                End Try
            End If
        End If

        Return focused
    End Function

    Private Function Vision_RequestJson(imageDataUrl As String, model As String, prompt As String) As Dictionary(Of String, Object)
        Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()

        ' Payload for Chat Completions with image input
        Dim payloadObj As Object = New Dictionary(Of String, Object) From {
            {"model", model},
            {"response_format", New Dictionary(Of String, Object) From {{"type", "json_object"}}},
            {"messages", New Object() {
                New Dictionary(Of String, Object) From {
                    {"role", "user"},
                    {"content", New Object() {
                        New Dictionary(Of String, Object) From {{"type", "text"}, {"text", prompt}},
                        New Dictionary(Of String, Object) From {
                            {"type", "image_url"},
                            {"image_url", New Dictionary(Of String, Object) From {{"url", imageDataUrl}}}
                        }
                    }}
                }
            }}
        }
        Dim payload As String = jss.Serialize(payloadObj)

        ' TLS 1.2 shim (if available on OS)
        Try : System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType) : Catch : End Try

        Dim req As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create("https://api.openai.com/v1/chat/completions"), System.Net.HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/json"
        req.Headers(System.Net.HttpRequestHeader.Authorization) = "Bearer " & openAiApiKey

        Dim bytes = Encoding.UTF8.GetBytes(payload)
        Using rs = req.GetRequestStream()
            rs.Write(bytes, 0, bytes.Length)
        End Using

        Dim respText As String
        Using resp = CType(req.GetResponse(), System.Net.HttpWebResponse)
            Using sr As New IO.StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                respText = sr.ReadToEnd()
            End Using
        End Using

        ' Parse OpenAI envelope -> choices[0].message.content (our JSON string)
        Dim env = CType(jss.DeserializeObject(respText), Dictionary(Of String, Object))
        Dim choices = CType(env("choices"), Object())
        If choices Is Nothing OrElse choices.Length = 0 Then Return Nothing

        Dim msg = CType(CType(choices(0), Dictionary(Of String, Object))("message"), Dictionary(Of String, Object))
        Dim content As String = CStr(msg("content"))
        If String.IsNullOrEmpty(content) Then Return Nothing

        ' Parse the JSON string the model returned
        Return CType(jss.DeserializeObject(content), Dictionary(Of String, Object))
    End Function

    Private Function BitmapToDataUrlPng(bmp As Bitmap) As String
        Using ms As New IO.MemoryStream()
            bmp.Save(ms, Imaging.ImageFormat.Png)
            Return "data:image/png;base64," & Convert.ToBase64String(ms.ToArray())
        End Using
    End Function


End Class
