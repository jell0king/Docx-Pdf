Attribute VB_Name = "Module1"
Sub SaveAsDocxAndPdf()
    Dim doc As Document
    Dim docPath As String
    Dim docName As String
    Dim pdfPath As String
    Dim pdfName As String

    Set doc = ActiveDocument

    docPath = doc.Path
    docName = Left(doc.Name, InStrRev(doc.Name, ".") - 1)

    ' Define the paths for the DOCX and PDF files
    pdfPath = docPath & "\" & docName & ".pdf"
    docPath = docPath & "\" & docName & ".docx"

    ' Save as DOCX
    doc.SaveAs2 FileName:=docPath, FileFormat:=wdFormatXMLDocument

    ' Save as PDF
    doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF

    ' Confirm save
    MsgBox "Document saved as DOCX and PDF.", vbInformation
End Sub

