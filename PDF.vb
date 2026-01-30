'Control for iTextSharp

Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports System.Runtime.InteropServices

<ProgId("IgorKrup.PDF")>
<Guid("6332A55E-EEFF-433B-9976-41A4A0420877")>
<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class PDF

    Public Function PageCount(ByVal sInPdf As String) As Integer
        Dim doc As PdfReader = Nothing
        Try
            doc = New PdfReader(sInPdf)
            Return doc.NumberOfPages
        Finally
            If doc IsNot Nothing Then doc.Close()
        End Try
    End Function

    Public Sub ExtractPage(ByVal sInFilePath As String, ByVal sOutFilePath As String, iPage As Integer)
        Dim oPdfReader As PdfReader = Nothing
        Dim oPdfDoc As Document = Nothing
        Dim oPdfWriter As PdfWriter = Nothing

        Try
            PdfReader.unethicalreading = True
            oPdfReader = New PdfReader(sInFilePath)
            oPdfDoc = New Document(oPdfReader.GetPageSizeWithRotation(iPage))
            oPdfWriter = PdfWriter.GetInstance(oPdfDoc, New System.IO.FileStream(sOutFilePath, System.IO.FileMode.Create))

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

    Public Function GetFileText(ByVal sInPdf As String) As String
        Dim sb As New System.Text.StringBuilder()

        Using reader As New PdfReader(sInPdf)
            For iPage As Integer = 1 To reader.NumberOfPages
                ' Extract text using a built-in strategy
                Dim text As String = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, iPage, New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy())
                sb.AppendLine(text)
            Next
        End Using

        Return sb.ToString()
    End Function

    Public Function GetPageText(ByVal sInPdf As String, iPage As Integer) As String
        Dim text As String = ""
        Using reader As New PdfReader(sInPdf)
            text = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, iPage, New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy())
        End Using
        Return text
    End Function

    Public Sub MergeFilesInFolder(ByVal sFolderPath As String,
                          ByVal sOutFilePath As String,
                          ByVal bResize As Boolean,
                          Optional sFileType As String = "All")

        Dim oOcrTempFiles As New ArrayList()
        Dim oFiles As String() = System.IO.Directory.GetFiles(sFolderPath)

        Dim oPdfDoc As New iTextSharp.text.Document()
        Dim oPdfWriter As PdfWriter = PdfWriter.GetInstance(oPdfDoc, New System.IO.FileStream(sOutFilePath, System.IO.FileMode.Create))
        oPdfDoc.Open()

        System.Array.Sort(Of String)(oFiles)

        For i As Integer = 0 To oFiles.Length - 1
            Dim sFromFilePath As String = oFiles(i)
            Dim oFileInfo As New System.IO.FileInfo(sFromFilePath)
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
                System.IO.File.Delete(sOutFilePath)
            Catch ex2 As Exception
            End Try
        End Try

    End Sub


    Private Sub AddPdf(ByVal sInFilePath As String, ByRef oPdfDoc As iTextSharp.text.Document,
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

    Private Sub AddImage(ByVal sInFilePath As String, ByRef oPdfDoc As iTextSharp.text.Document,
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
                    Dim oMemoryStream As New System.IO.MemoryStream()
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

    Private Sub AddImage2(ByRef oImage As System.Drawing.Image, ByRef oPdfDoc As iTextSharp.text.Document, ByRef oPdfWriter As PdfWriter)

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
    Public Function HasAnySignature(sInFilePath As String) As Boolean
        Using oPdfReader As New PdfReader(sInFilePath)
            Return HasAnySignature(oPdfReader)
        End Using
    End Function

    Private Function HasAnySignature(reader As PdfReader) As Boolean
        Dim af As AcroFields = reader.AcroFields
        If af IsNot Nothing Then
            Dim sigNames As IList(Of String) = af.GetSignatureNames()
            If sigNames IsNot Nothing AndAlso sigNames.Count > 0 Then Return True 'already signed
        End If

        'Also detect unsigned signature fields (blank signature boxes already present)
        Dim root As PdfDictionary = reader.Catalog
        Dim acroForm As PdfDictionary = root.GetAsDict(PdfName.ACROFORM)
        If acroForm IsNot Nothing Then
            Dim fields As PdfArray = acroForm.GetAsArray(PdfName.FIELDS)
            If fields IsNot Nothing Then
                If ContainsSignatureFieldRecursive(fields) Then Return True
            End If
        End If

        Return False
    End Function

    Private mSigSearchList As New ArrayList()

    Public Sub AddSignature(ByVal sInFilePath As String, ByVal sOutFilePath As String)

        If mSigSearchList.Count = 0 Then
            Throw New Exception("Call AddSigSearch() before using AddSignature() function")
        End If

        Using oPdfReader As New PdfReader(sInFilePath)

            Dim oPdfDoc As New Document()
            Using fs As New System.IO.FileStream(sOutFilePath, System.IO.FileMode.Create, System.IO.FileAccess.Write)
                Dim oPdfWriter As PdfWriter = PdfWriter.GetInstance(oPdfDoc, fs)

                oPdfDoc.Open()
                oPdfDoc.SetPageSize(PageSize.LEDGER.Rotate())

                Dim oDirectContent As PdfContentByte = oPdfWriter.DirectContent
                Dim iNumberOfPages As Integer = oPdfReader.NumberOfPages
                Dim iPage As Integer = 0

                Do While (iPage < iNumberOfPages)
                    iPage += 1
                    oPdfDoc.SetPageSize(oPdfReader.GetPageSizeWithRotation(iPage))
                    oPdfDoc.NewPage()

                    Dim oPdfImportedPage As PdfImportedPage = oPdfWriter.GetImportedPage(oPdfReader, iPage)
                    Dim iRotation As Integer = oPdfReader.GetPageRotation(iPage)

                    If (iRotation = 90) Or (iRotation = 270) Then
                        oDirectContent.AddTemplate(oPdfImportedPage, 0, -1.0F, 1.0F, 0, 0, oPdfReader.GetPageSizeWithRotation(iPage).Height)
                    Else
                        oDirectContent.AddTemplate(oPdfImportedPage, 1.0F, 0, 0, 1.0F, 0, 0)
                    End If

                    Dim oTextExtractor As New TextExtractor()
                    PdfTextExtractor.GetTextFromPage(oPdfReader, iPage, oTextExtractor)

                    For Each oSearch As SigSearch In mSigSearchList

                        'Dim iBottomMargin As Integer, iLeftMargin As Integer, iWidth As Integer, iHeight As Integer, sFind As String

                        Dim oRect As iTextSharp.text.Rectangle = oTextExtractor.Find(oSearch.find)

                        If oRect Is Nothing Then
                            ' Try to find text manually
                            Dim oLines As New Hashtable
                            For i = 0 To oTextExtractor.oPoints.Count - 1
                                Dim r As RectAndText = oTextExtractor.oPoints(i)

                                If oLines.ContainsKey(r.Rect.Top) Then
                                    oLines(r.Rect.Top) = CStr(oLines(r.Rect.Top)) & " " & r.Text
                                Else
                                    oLines(r.Rect.Top) = r.Text
                                End If

                                Dim sLine As String = CStr(oLines(r.Rect.Top))
                                If sLine.IndexOf(oSearch.find, StringComparison.Ordinal) <> -1 Then
                                    oRect = r.Rect
                                    Exit For
                                End If
                            Next
                        End If

                        If oRect IsNot Nothing Then
                            Dim iX As Integer = CInt(oRect.Left + oRect.Width + oSearch.leftMargin)
                            Dim iY As Integer = CInt(oRect.Bottom - oSearch.bottomMargin)

                            Dim field As PdfFormField = PdfFormField.CreateSignature(oPdfWriter)
                            field.SetWidget(New Rectangle(iX, iY, iX + oSearch.width, iY + oSearch.height), PdfAnnotation.HIGHLIGHT_OUTLINE)

                            'Make name unique; duplicates can break the form
                            field.FieldName = $"myEmptySignatureField_{iPage}_{Guid.NewGuid().ToString("N")}"
                            oPdfWriter.AddAnnotation(field)
                        End If
                    Next
                Loop

                oPdfDoc.Close()
            End Using
        End Using
    End Sub

    Private Function ContainsSignatureFieldRecursive(fields As PdfArray) As Boolean
        For i As Integer = 0 To fields.Size - 1
            Dim obj As PdfObject = fields.GetPdfObject(i)
            Dim dict As PdfDictionary = TryCast(PdfReader.GetPdfObject(obj), PdfDictionary)
            If dict Is Nothing Then Continue For

            'FT = /Sig means it's a signature field
            Dim ft As PdfName = dict.GetAsName(PdfName.FT)
            If PdfName.SIG.Equals(ft) Then Return True

            'Kids recursion
            Dim kids As PdfArray = dict.GetAsArray(PdfName.KIDS)
            If kids IsNot Nothing AndAlso kids.Size > 0 Then
                If ContainsSignatureFieldRecursive(kids) Then Return True
            End If
        Next
        Return False
    End Function

    Public Sub AddSigSearch(find As String, bottomMargin As Integer, leftMargin As Integer, width As Integer, height As Integer)
        Dim o As New SigSearch()
        o.find = find
        o.bottomMargin = bottomMargin
        o.leftMargin = leftMargin
        o.width = width
        o.height = height
        mSigSearchList.Add(o)
    End Sub

    Private Class SigSearch
        Public find As String
        Public bottomMargin As Integer
        Public leftMargin As Integer
        Public width As Integer
        Public height As Integer
    End Class

    Private Class TextExtractor
        Inherits LocationTextExtractionStrategy
        Implements iTextSharp.text.pdf.parser.ITextExtractionStrategy
        Public oPoints As IList(Of RectAndText) = New List(Of RectAndText)
        Public Overrides Sub RenderText(renderInfo As TextRenderInfo) 'Implements IRenderListener.RenderText
            'https://stackoverflow.com/questions/23909893/getting-coordinates-of-string-using-itextextractionstrategy-and-locationtextextr
            MyBase.RenderText(renderInfo)

            Dim bottomLeft As Vector = renderInfo.GetDescentLine().GetStartPoint()
            Dim topRight As Vector = renderInfo.GetAscentLine().GetEndPoint() 'GetBaseline

            Dim rect As Rectangle = New Rectangle(bottomLeft(Vector.I1), bottomLeft(Vector.I2), topRight(Vector.I1), topRight(Vector.I2))
            oPoints.Add(New RectAndText(rect, renderInfo.GetText()))
        End Sub

        Private Function GetLines() As Dictionary(Of Single, ArrayList)
            Dim oLines As New Dictionary(Of Single, ArrayList)
            For Each p As RectAndText In oPoints
                Dim iBottom = p.Rect.Bottom

                If oLines.ContainsKey(iBottom) = False Then
                    oLines(iBottom) = New ArrayList()
                End If

                oLines(iBottom).Add(p)
            Next

            Return oLines
        End Function

        Function Find(ByVal sFind As String) As iTextSharp.text.Rectangle
            Dim oLines As Dictionary(Of Single, ArrayList) = GetLines()

            For Each oEntry As KeyValuePair(Of Single, ArrayList) In oLines
                'Dim iBottom As Integer = oEntry.Key
                Dim oRectAndTexts As ArrayList = oEntry.Value
                Dim sLine As String = ""
                For Each p As RectAndText In oRectAndTexts
                    sLine += p.Text
                    If sLine.IndexOf(sFind) <> -1 Then
                        Return p.Rect
                    End If
                Next
            Next

            Return Nothing
        End Function

    End Class

    Private Class RectAndText
        Public Rect As iTextSharp.text.Rectangle
        Public Text As String
        Public Sub New(ByVal rect As iTextSharp.text.Rectangle, ByVal text As String)
            Me.Rect = rect
            Me.Text = text
        End Sub
    End Class

End Class

