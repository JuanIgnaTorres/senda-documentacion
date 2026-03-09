Attribute VB_Name = "modWordBatchExporter"
Option Explicit

Dim pathFolder As String
Dim filesFolder() As String
Dim wdApp As Word.Application
Dim originalDoc As Document
Dim templateDoc As Document

Sub main()
    Dim i As Long
    Dim pdfFolder As String

    On Error GoTo CleanUp

    obtainFolder pathFolder ' Seleccionar carpeta
    obtainFiles pathFolder  ' Listar archivos de carpeta

    ' Crear instancia de Word una sola vez
    Set wdApp = New Word.Application
    wdApp.Visible = False

    ' Asegurar carpeta de salida "PDF BATCH"
    pdfFolder = pathFolder & "\PDF BATCH"
    If Dir(pdfFolder, vbDirectory) = "" Then MkDir pdfFolder

    ' Procesar cada archivo del array
    For i = LBound(filesFolder) To UBound(filesFolder)
        Dim fullDocPath As String
        Dim pdfPath As String

        fullDocPath = pathFolder & "\" & filesFolder(i)
        pdfPath = pdfFolder & "\" & Replace(filesFolder(i), ".docx", ".pdf")

        ' Llamar a la rutina elemental que exporta a PDF
        ExportDocToPDF fullDocPath, pdfPath
    Next i

    MsgBox "Conversión completada. PDFs en: " & vbCrLf & pdfFolder

    ' Limpieza normal
    wdApp.Quit
    Set wdApp = Nothing
    Exit Sub

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True

    If Not templateDoc Is Nothing Then
        templateDoc.Close SaveChanges:=False
        Set templateDoc = Nothing
    End If

    If Not originalDoc Is Nothing Then
        originalDoc.Close SaveChanges:=False
        Set originalDoc = Nothing
    End If

    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If

    If Err.Number <> 0 Then
        Dim errMsg As String
        errMsg = "Error en archivo " & IIf(i >= LBound(filesFolder) And i <= UBound(filesFolder), filesFolder(i), "") & ":" & vbNewLine & Err.Description
        MsgBox errMsg, vbExclamation, "Error"
    End If
End Sub

Sub obtainFolder(ByRef pPathFolder As String)
    Dim dlg As FileDialog

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

    With dlg
        .Title = "Selecciona una carpeta"
        .AllowMultiSelect = False

        If .Show = -1 Then
            pPathFolder = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. El macro se detendrá."
            End
        End If
    End With

    Debug.Print pPathFolder

    Set dlg = Nothing
End Sub

Sub obtainFiles(ByVal pPathFolder As String)
    Dim file As String
    Dim i As Long

    i = 0
    file = Dir(pPathFolder & "\*.docx")

    Do While file <> ""
        ReDim Preserve filesFolder(i)
        filesFolder(i) = file
        i = i + 1
        file = Dir()
    Loop

    If i = 0 Then
        MsgBox "No se encontraron archivos .docx en la carpeta seleccionada."
        End
    End If
End Sub

' Rutina elemental que exporta un documento a PDF usando la instancia global wdApp
Sub ExportDocToPDF(ByVal docPath As String, ByVal pdfPath As String)
    On Error GoTo ErrHandler

    Set originalDoc = wdApp.Documents.Open(FileName:=docPath, ReadOnly:=True, AddToRecentFiles:=False)
    originalDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
    originalDoc.Close SaveChanges:=False
    Set originalDoc = Nothing
    Exit Sub

ErrHandler:
    If Not originalDoc Is Nothing Then
        originalDoc.Close SaveChanges:=False
        Set originalDoc = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
