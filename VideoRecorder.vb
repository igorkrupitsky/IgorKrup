Imports System.Diagnostics
Imports System.IO

Public Class VideoRecorder

    Private recorderProc As Process = Nothing
    Private recordingFilePath As String = ""
    Public sFfmpegFilePath As String = ""

    Public Sub New()
        'Download https://github.com/FFmpeg/FFmpeg
        Dim sPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe")
        If IO.File.Exists(sPath) Then
            sFfmpegFilePath = sPath
        End If
    End Sub

    Public Sub StartRecording(outputPath As String,
                              Optional windowTitle As String = "",
                              Optional audioDeviceName As String = "",
                              Optional monitorRegion As String = "")

        'audioDeviceName - ListFfmpegDevices()
        'monitorRegion - GetMonitorCaptureInfo()

        If String.IsNullOrEmpty(sFfmpegFilePath) Then
            Throw New ArgumentException("sFfmpegFilePath is not set to ffmpeg.exe location. Or you can copy the file to the dll folder")
        End If

        If recorderProc IsNot Nothing AndAlso Not recorderProc.HasExited Then
            Throw New InvalidOperationException("Recording is already in progress.")
        End If

        If Not outputPath.ToLower().EndsWith(".mp4") Then
            Throw New ArgumentException("Output file must end with .mp4")
        End If

        Dim dir = Path.GetDirectoryName(outputPath)
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

        recordingFilePath = outputPath

        Dim videoInput As String

        If windowTitle <> "" Then
            ' Record a specific window
            videoInput = $"-f gdigrab -framerate 25 -i title=""{windowTitle}"""
        ElseIf monitorRegion <> "" Then
            ' Record a specific monitor region like "1920,0,1920x1080"
            Dim parts = monitorRegion.Split(","c)
            If parts.Length = 3 Then
                Dim offsetX = parts(0).Trim()
                Dim offsetY = parts(1).Trim()
                Dim size = parts(2).Trim() ' e.g. 1920x1080
                videoInput = $"-f gdigrab -framerate 25 -offset_x {offsetX} -offset_y {offsetY} -video_size {size} -i desktop"
            Else
                Throw New ArgumentException("Invalid monitor region format. Use 'X,Y,WidthxHeight'.")
            End If
        Else
            ' Default to primary monitor
            videoInput = "-f gdigrab -framerate 25 -i desktop"
        End If

        Dim audioInput As String = ""
        If audioDeviceName <> "" Then
            audioInput = $" -f dshow -i audio=""{audioDeviceName}"""
        End If

        Dim outputSettings = $"-pix_fmt yuv420p ""{outputPath}"""

        Dim ffmpegArgs As String = $"-y {videoInput}{audioInput} {outputSettings}"

        recorderProc = New Process()
        recorderProc.StartInfo.FileName = sFfmpegFilePath
        recorderProc.StartInfo.Arguments = ffmpegArgs
        recorderProc.StartInfo.UseShellExecute = False
        recorderProc.StartInfo.CreateNoWindow = True
        recorderProc.StartInfo.RedirectStandardInput = True
        recorderProc.StartInfo.RedirectStandardOutput = True
        recorderProc.StartInfo.RedirectStandardError = True

        recorderProc.Start()

        AddHandler recorderProc.ErrorDataReceived, Sub(s, e)
                                                       If e.Data IsNot Nothing Then
                                                           Debug.WriteLine("FFmpeg STDERR: " & e.Data)
                                                       End If
                                                   End Sub
        recorderProc.BeginErrorReadLine()
    End Sub

    Public Sub StopRecording()
        If recorderProc Is Nothing Then Exit Sub

        Try
            If Not recorderProc.HasExited Then
                recorderProc.StandardInput.WriteLine("q")
                If Not recorderProc.WaitForExit(5000) Then
                    recorderProc.Kill()
                End If
            End If
        Catch ex As Exception
            recorderProc.Kill()
        Finally
            recorderProc = Nothing
        End Try
    End Sub

    Public Function GetLastRecordingPath() As String
        Return recordingFilePath
    End Function

    Public Function IsRecording() As Boolean
        Return recorderProc IsNot Nothing AndAlso Not recorderProc.HasExited
    End Function

    Public Function ListFfmpegDevices() As List(Of String)
        Dim result As New List(Of String)

        Dim proc As New Process()
        proc.StartInfo.FileName = "ffmpeg"
        proc.StartInfo.Arguments = "-list_devices true -f dshow -i dummy"
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardError = True ' ffmpeg outputs device list to stderr
        proc.StartInfo.CreateNoWindow = True

        AddHandler proc.ErrorDataReceived, Sub(sender, e)
                                               If e.Data IsNot Nothing Then
                                                   ' Filter lines showing devices
                                                   If e.Data.Contains("DirectShow audio devices") OrElse e.Data.Contains("DirectShow video devices") OrElse e.Data.Trim().StartsWith("""") Then
                                                       result.Add(e.Data.Trim())
                                                   End If
                                               End If
                                           End Sub

        proc.Start()
        proc.BeginErrorReadLine()
        proc.WaitForExit()

        Return result
    End Function

    Public Function GetMonitorCaptureInfo() As List(Of String)
        Dim monitorList As New List(Of String)

        For Each scr As System.Windows.Forms.Screen In System.Windows.Forms.Screen.AllScreens
            Dim bounds = scr.Bounds
            Dim info = $"{scr.DeviceName}, {bounds.X},{bounds.Y},{bounds.Width}x{bounds.Height}"
            monitorList.Add(info)
        Next

        Return monitorList
    End Function

End Class
