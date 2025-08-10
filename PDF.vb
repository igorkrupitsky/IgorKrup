'Control for iTextSharp

Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Runtime.InteropServices

<ProgId("IgorKrup.PDF")>
<Guid("6332A55E-EEFF-433B-9976-41A4A0420877")>
<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class PDF

    Function PageCount(ByVal sInPdf As String) As Integer
        Dim doc As PdfReader = Nothing
        Try
            doc = New PdfReader(sInPdf)
            Return doc.NumberOfPages
        Finally
            If doc IsNot Nothing Then doc.Close()
        End Try
    End Function

    Sub ExtractPage(ByVal sInFilePath As String, ByVal sOutFilePath As String, iPage As Integer)
        Dim oPdfReader As PdfReader = Nothing
        Dim oPdfDoc As Document = Nothing
        Dim oPdfWriter As PdfWriter = Nothing

        Try
            oPdfReader = New PdfReader(sInFilePath)
            oPdfDoc = New Document(oPdfReader.GetPageSizeWithRotation(iPage))
            oPdfWriter = PdfWriter.GetInstance(oPdfDoc, New IO.FileStream(sOutFilePath, IO.FileMode.Create))

            oPdfDoc.Open()
            Dim oDirectContent As PdfContentByte = oPdfWriter.DirectContent
            Dim iRotation As Integer = oPdfReader.GetPageRotation(iPage)
            Dim oPdfImportedPage As PdfImportedPage = oPdfWriter.GetImportedPage(oPdfReader, iPage)

            oPdfDoc.NewPage()

            If iRotation = 90 Or iRotation = 270 Then
                oDirectContent.AddTemplate(oPdfImportedPage, 0, -1.0F, 1.0F, 0, 0, oPdfReader.GetPageSizeWithRotation(iPage).Height)
            Else
                oDirectContent.AddTemplate(oPdfImportedPage, 1.0F, 0, 0, 1.0F, 0, 0)
            End If
        Finally
            If oPdfDoc IsNot Nothing Then oPdfDoc.Close()
            If oPdfWriter IsNot Nothing Then oPdfWriter.Close()
            If oPdfReader IsNot Nothing Then oPdfReader.Close()
        End Try
    End Sub

    Function GetFileText(ByVal sInPdf As String) As String
        Dim doc As New iTextSharp.text.pdf.PdfReader(sInPdf)
        Dim sb As New System.Text.StringBuilder()

        For iPage As Integer = 1 To doc.NumberOfPages
            Dim pg As iTextSharp.text.pdf.PdfDictionary = doc.GetPageN(iPage)
            Dim ir As Object = pg.Get(iTextSharp.text.pdf.PdfName.CONTENTS)
            Dim value As iTextSharp.text.pdf.PdfObject = doc.GetPdfObject(ir.Number)
            If value.IsStream() Then
                Dim stream As iTextSharp.text.pdf.PRStream = value
                Dim streamBytes As Byte() = iTextSharp.text.pdf.PdfReader.GetStreamBytes(stream)
                Dim tokenizer As New iTextSharp.text.pdf.PRTokeniser(New iTextSharp.text.pdf.RandomAccessFileOrArray(streamBytes))

                Try
                    While tokenizer.NextToken()
                        If tokenizer.TokenType = iTextSharp.text.pdf.PRTokeniser.TK_STRING Then
                            Dim str As String = tokenizer.StringValue
                            sb.Append(str)
                        End If
                    End While
                Catch ex As Exception
                    'Ignore
                Finally
                    tokenizer.Close()
                End Try

            End If
        Next
        doc.Close()

        Return sb.ToString()
    End Function

    Function GetPageText(ByVal sInPdf As String, iPage As Integer) As String
        Dim doc As New iTextSharp.text.pdf.PdfReader(sInPdf)
        Dim sb As New System.Text.StringBuilder()
        Dim pg As iTextSharp.text.pdf.PdfDictionary = doc.GetPageN(iPage)
        Dim ir As Object = pg.Get(iTextSharp.text.pdf.PdfName.CONTENTS)
        Dim value As iTextSharp.text.pdf.PdfObject = doc.GetPdfObject(ir.Number)

        If value.IsStream() Then
            Dim stream As iTextSharp.text.pdf.PRStream = value
            Dim streamBytes As Byte() = iTextSharp.text.pdf.PdfReader.GetStreamBytes(stream)
            Dim tokenizer As New iTextSharp.text.pdf.PRTokeniser(New iTextSharp.text.pdf.RandomAccessFileOrArray(streamBytes))

            Try
                While tokenizer.NextToken()
                    If tokenizer.TokenType = iTextSharp.text.pdf.PRTokeniser.TK_STRING Then
                        Dim str As String = tokenizer.StringValue
                        sb.Append(str)
                    End If
                End While
            Catch ex As Exception
                'Ignore
            Finally
                tokenizer.Close()
            End Try

        End If
        doc.Close()

        Return sb.ToString()
    End Function


    Sub MergeFileInFolder(ByVal sFolderPath As String,
                          ByVal sOutFilePath As String,
                          ByVal bResize As Boolean,
                          Optional sFileType As String = "All")

        Dim oOcrTempFiles As New ArrayList()
        Dim oFiles As String() = IO.Directory.GetFiles(sFolderPath)

        Dim oPdfDoc As New iTextSharp.text.Document()
        Dim oPdfWriter As PdfWriter = PdfWriter.GetInstance(oPdfDoc, New IO.FileStream(sOutFilePath, IO.FileMode.Create))
        oPdfDoc.Open()

        System.Array.Sort(Of String)(oFiles)

        For i As Integer = 0 To oFiles.Length - 1
            Dim sFromFilePath As String = oFiles(i)
            Dim oFileInfo As New IO.FileInfo(sFromFilePath)
            Dim sExt As String = PadExt(oFileInfo.Extension)

            Try
                Dim bAddPdf As Boolean = False
                Dim bAddImage As Boolean = False
                Select Case sFileType
                    Case "All"
                        If sExt = "PDF" Then
                            bAddPdf = True
                        ElseIf sExt = "JPG" Or sExt = "TIF" Then
                            bAddImage = True
                        End If
                    Case "PDF"
                        If sExt = "PDF" Then
                            bAddPdf = True
                        End If
                    Case "JPG", "TIF"
                        If sExt = "JPG" Or sExt = "TIF" Then
                            bAddImage = True
                        End If
                End Select

                If bAddPdf Or bAddImage Then
                    Dim sBookmarkTitle As String = oFileInfo.Name
                    Dim sOcrTiffFile As String = ""
                    Dim sOcrPdfFile As String = ""
                    Dim sError As String = ""

                    If bAddPdf Or bAddImage Then
                        AddBookmark(oPdfDoc, sBookmarkTitle)
                    End If

                    If bAddPdf Then
                        AddPdf(sFromFilePath, oPdfDoc, oPdfWriter, bResize)
                    ElseIf bAddImage Then
                        AddImage(sFromFilePath, oPdfDoc, oPdfWriter, sExt, bResize)
                    End If

                End If

            Catch ex As Exception

            End Try
        Next

        Try
            oPdfDoc.Close()
            oPdfWriter.Close()
        Catch ex As Exception
            Try
                IO.File.Delete(sOutFilePath)
            Catch ex2 As Exception
            End Try
        End Try

    End Sub


    Sub AddPdf(ByVal sInFilePath As String, ByRef oPdfDoc As iTextSharp.text.Document,
               ByRef oPdfWriter As PdfWriter, bResize As Boolean)

        Dim oDirectContent As iTextSharp.text.pdf.PdfContentByte = oPdfWriter.DirectContent
        Dim oPdfReader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(sInFilePath)
        Dim iNumberOfPages As Integer = oPdfReader.NumberOfPages
        Dim iPage As Integer = 0

        Do While (iPage < iNumberOfPages)
            iPage += 1

            Dim iRotation As Integer = oPdfReader.GetPageRotation(iPage)
            Dim oPdfImportedPage As iTextSharp.text.pdf.PdfImportedPage = oPdfWriter.GetImportedPage(oPdfReader, iPage)

            If bResize Then
                If (oPdfImportedPage.Width <= oPdfImportedPage.Height) Then
                    oPdfDoc.SetPageSize(iTextSharp.text.PageSize.LETTER)
                Else
                    oPdfDoc.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate())
                End If

                oPdfDoc.NewPage()

                Dim iWidthFactor As Single = oPdfDoc.PageSize.Width / oPdfReader.GetPageSize(iPage).Width
                Dim iHeightFactor As Single = oPdfDoc.PageSize.Height / oPdfReader.GetPageSize(iPage).Height
                Dim iFactor As Single = Math.Min(iWidthFactor, iHeightFactor)

                Dim iOffsetX As Single = (oPdfDoc.PageSize.Width - (oPdfImportedPage.Width * iFactor)) / 2
                Dim iOffsetY As Single = (oPdfDoc.PageSize.Height - (oPdfImportedPage.Height * iFactor)) / 2

                oDirectContent.AddTemplate(oPdfImportedPage, iFactor, 0, 0, iFactor, iOffsetX, iOffsetY)

            Else
                oPdfDoc.SetPageSize(oPdfReader.GetPageSizeWithRotation(iPage))
                oPdfDoc.NewPage()

                If iRotation = 90 Then
                    oDirectContent.AddTemplate(oPdfImportedPage, 0, -1.0F, 1.0F, 0, 0, oPdfReader.GetPageSizeWithRotation(iPage).Height)

                ElseIf iRotation = 270 Then
                    oDirectContent.AddTemplate(oPdfImportedPage, 0, 1.0F, -1.0F, 0, oPdfReader.GetPageSizeWithRotation(iPage).Width, 0)

                ElseIf iRotation = 180 Then
                    oDirectContent.AddTemplate(oPdfImportedPage, -1.0F, 0, 0, -1.0F, oPdfReader.GetPageSizeWithRotation(iPage).Width, oPdfReader.GetPageSizeWithRotation(iPage).Height)

                Else
                    oDirectContent.AddTemplate(oPdfImportedPage, 1.0F, 0, 0, 1.0F, 0, 0)
                End If
            End If
        Loop

    End Sub

    Sub AddImage(ByVal sInFilePath As String, ByRef oPdfDoc As iTextSharp.text.Document,
                 ByRef oPdfWriter As PdfWriter, ByVal sExt As String, bResize As Boolean)


        If bResize = False Then
            Dim oDirectContent As iTextSharp.text.pdf.PdfContentByte = oPdfWriter.DirectContent
            Dim oPdfImage As iTextSharp.text.Image
            oPdfImage = iTextSharp.text.Image.GetInstance(sInFilePath)
            oPdfImage.SetAbsolutePosition(1, 1)
            oPdfDoc.SetPageSize(New iTextSharp.text.Rectangle(oPdfImage.Width, oPdfImage.Height))
            oPdfDoc.NewPage()
            oDirectContent.AddImage(oPdfImage)
            Exit Sub
        End If

        Dim oImage As System.Drawing.Image = System.Drawing.Image.FromFile(sInFilePath)

        'Multi-Page Tiff
        If sExt = "TIF" Then
            Dim iPageCount As Integer = oImage.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page)
            If iPageCount > 1 Then
                For iPage As Integer = 0 To iPageCount - 1
                    oImage.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, iPage)
                    Dim oMemoryStream As New IO.MemoryStream()
                    oImage.Save(oMemoryStream, System.Drawing.Imaging.ImageFormat.Png)
                    Dim oImage2 As System.Drawing.Image = System.Drawing.Image.FromStream(oMemoryStream)
                    AddImage2(oImage2, oPdfDoc, oPdfWriter)
                    oMemoryStream.Close()
                Next
                Exit Sub
            End If
        End If

        AddImage2(oImage, oPdfDoc, oPdfWriter)
        oImage.Dispose()

    End Sub


    Sub AddImage2(ByRef oImage As System.Drawing.Image, ByRef oPdfDoc As iTextSharp.text.Document, ByRef oPdfWriter As PdfWriter)

        Dim oDirectContent As iTextSharp.text.pdf.PdfContentByte = oPdfWriter.DirectContent
        Dim oPdfImage As iTextSharp.text.Image
        Dim iWidth As Single = oImage.Width
        Dim iHeight As Single = oImage.Height
        Dim iAspectRatio As Double = iWidth / iHeight

        Dim iWidthPage As Single = 0
        Dim iHeightPage As Single = 0

        If iAspectRatio < 1 Then
            'Landscape image
            iWidthPage = iTextSharp.text.PageSize.LETTER.Width
            iHeightPage = iTextSharp.text.PageSize.LETTER.Height
        Else
            iHeightPage = iTextSharp.text.PageSize.LETTER.Width
            iWidthPage = iTextSharp.text.PageSize.LETTER.Height
        End If

        Dim iPageAspectRatio As Double = iWidthPage / iHeightPage

        Dim iWidthGoal As Single = 0
        Dim iHeightGoal As Single = 0
        Dim bFitsWithin As Boolean = False

        If iWidth < iWidthPage And iHeight < iHeightPage Then
            'Image fits within the page
            bFitsWithin = True
            iWidthGoal = iWidth
            iHeightGoal = iHeight

        ElseIf iAspectRatio > iPageAspectRatio Then
            'Width is too big
            iWidthGoal = iWidthPage
            iHeightGoal = iWidthPage * (iHeight / iWidth)

        Else
            'Height is too big
            iWidthGoal = iHeightPage * (iWidth / iHeight)
            iHeightGoal = iHeightPage
        End If

        oPdfImage = iTextSharp.text.Image.GetInstance(oImage, System.Drawing.Imaging.ImageFormat.Png)
        oPdfImage.SetAbsolutePosition(1, 1)

        If iAspectRatio < 1 Then
            'Landscape image
            oPdfDoc.SetPageSize(iTextSharp.text.PageSize.LETTER)
        Else
            oPdfDoc.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate())
        End If

        oPdfDoc.NewPage()
        oPdfImage.ScaleAbsolute(iWidthGoal, iHeightGoal)
        oDirectContent.AddImage(oPdfImage)

    End Sub

    Private Sub AddBookmark(ByRef oPdfDoc As iTextSharp.text.Document, ByVal sBookmarkTitle As String)
        Dim oChapter As New iTextSharp.text.Chapter("", 0)
        oChapter.NumberDepth = 0
        oChapter.BookmarkTitle = sBookmarkTitle
        oPdfDoc.Add(oChapter)
    End Sub

    Private Function PadExt(ByVal s As String) As String
        s = UCase(s)
        If s.Length > 3 Then
            s = s.Substring(1, 3)
        End If
        Return s
    End Function

End Class
