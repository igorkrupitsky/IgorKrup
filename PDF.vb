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

End Class
