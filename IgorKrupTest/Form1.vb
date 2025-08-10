Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        'VideoRecorderTest()
        'EdgeTest()
        'ImageSearchTest()
        'OcrTest()
        'PdfTest()
    End Sub

    Sub VideoRecorderTest()
        Dim vr As New IgorKrup.VideoRecorder
        vr.sFfmpegFilePath = "C:\Igor\GitHub\ElevenLabs\ffmpeg\bin\ffmpeg.exe"
        Dim oDevices = vr.ListFfmpegDevices()
        Dim oMonitor = vr.GetMonitorCaptureInfo()
        vr.StartRecording("C:\Temp\Test\test.mp4", , oDevices(1), oMonitor(2))
        'vr.StartRecording("C:\Temp\Test\test.mp4", "Zoom Workplace")
        MsgBox("recording")
        vr.StopRecording()
    End Sub

    Sub EdgeTest()
        Dim ie As New IgorKrup.EdgeDriver
        ie.sSharedDownloadFolder = ""
        ie.UpdateDriver()
        ie.GetUrl("https://google.com")
        Dim r = ie.ExecuteScript("return document.title")
        MsgBox(r)
        ie.Quit()
    End Sub

    Sub PdfTest()
        Dim oPDF As New IgorKrup.PDF
        Dim sPath = "C:\Users\80014379\Desktop\How to Generate a Certificate Signing Request.pdf"
        MsgBox(oPDF.PageCount(sPath))
    End Sub

    Sub ImageSearchTest()
        Dim oAutoIt As New IgorKrup.Control()
        Dim titleHint = "Outside "
        If oAutoIt.WinExists(titleHint) Then
            Dim fullTitle = oAutoIt.WinGetTitle(titleHint)
            Dim winX = oAutoIt.WinGetPosX(fullTitle)
            Dim winY = oAutoIt.WinGetPosY(fullTitle)
            Dim winW = oAutoIt.WinGetPosWidth(fullTitle)
            Dim winH = oAutoIt.WinGetPosHeight(fullTitle)

            Dim sSearchImage = "C:\Users\80014379\Desktop\print_button.png"
            Dim result2 = oAutoIt.SearchImage(sSearchImage, winX, winY, winW, winH, 10)
            If IsArray(result2) Then
                Dim x = result2(0)
                Dim y = result2(1)
                Dim w = result2(2)
                Dim h = result2(3)

                MsgBox("Found at X=" & x & ", Y=" & y & " (" & w & "," & h & ")")

                Dim sColor = oAutoIt.PixelGetColor(x, y)
                MsgBox(sColor)
            End If
        End If

    End Sub

    Sub OcrTest()
        Dim oAutoIt As New IgorKrup.Control()
        Dim titleHint = "Outside "

        If oAutoIt.WinExists(titleHint) Then
            Dim fullTitle = oAutoIt.WinGetTitle(titleHint)
            Dim winX = oAutoIt.WinGetPosX(fullTitle)
            Dim winY = oAutoIt.WinGetPosY(fullTitle)
            Dim winW = oAutoIt.WinGetPosWidth(fullTitle)
            Dim winH = oAutoIt.WinGetPosHeight(fullTitle)

            Dim result = oAutoIt.FindTextOnScreen(winX, winY, winW, winH, "Summary")
            If IsArray(result) Then
                Dim x = result(0)
                Dim y = result(1)
                Dim w = result(2)
                Dim h = result(3)

                MsgBox("Found at X=" & x & ", Y=" & y & " (" & w & "," & h & ")")

                Dim sColor = oAutoIt.PixelGetColor(x, y)
                MsgBox(sColor)
            End If
        End If
    End Sub

End Class
