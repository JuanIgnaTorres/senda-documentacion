Attribute VB_Name = "modECUFixFile"
Option Explicit

	Dim pathTemplate As String
	Dim pathFolder As String
	Dim filesFolder() As String
	Dim tableRowsCount As Integer

	Dim title As String
	Dim normalSequenceRow As Integer
	Dim exceptionRow As Integer
	Dim postconditionRow As Integer

	Const templateNormalSequenceRow As Integer = 7
	Const templateExceptionRow As Integer = 11
	Const templatePostconditionRow As Integer = 15
	Const templateBaseRows As Integer = 3
	Const cellMarkerLen As Integer = 2

	Dim normalSequenceRowsCount As Integer
	Dim exceptionRowsCount As Integer

Sub main()
	On Error GoTo CleanUp
	
	Dim i As Integer
	Dim offset As Integer
	Dim originalDoc As Document
	Dim templateDoc As Document
	
	obtainFile pathTemplate ' Seleccionar plantilla
	
	obtainFolder pathFolder ' Seleccionar carpeta
	
	Application.ScreenUpdating = False
	
	obtainFiles pathFolder ' Listar archivos de carpeta
	
	For i = 0 To UBound(filesFolder) ' Todos los archivos de la carpeta
		Set originalDoc = Documents.Open(pathFolder & "\" & filesFolder(i), Visible:=False)
		
		Set templateDoc = Documents.Add(Template:=pathTemplate, Visible:=False) ' Copia de la plantilla
		
		obtainTitle title, originalDoc ' Obtener título del archivo original
		
		obtainTableRowsCount tableRowsCount, originalDoc ' Obtener cantidad de filas de tabla original
		
		obtainRow "Secuencia normal", normalSequenceRow, originalDoc ' Obtener encabezado secuencia normal tabla original
		
		If normalSequenceRow <> 7 Then
			MsgBox "Secuencia normal en fila incorrecta. El macro se detendrá."
			GoTo CleanUp
		End If
		
		obtainRow "Excepción", exceptionRow, originalDoc ' Obtener encabezado excepción tabla original
		
		normalSequenceRowsCount = exceptionRow - normalSequenceRow - 1 ' Calcular cantidad filas secuencia normal tabla original
		
		editRows normalSequenceRowsCount, templateNormalSequenceRow + (templateBaseRows - 1), templateExceptionRow, templateDoc ' Edita filas secuencia normal plantilla
		
		obtainRow "Postcondición:", postconditionRow, originalDoc ' Obtener postcondición tabla original
		
		If postconditionRow <> tableRowsCount - 2 Then
			MsgBox "Excepción en fila incorrecta. El macro se detendrá."
			GoTo CleanUp
		End If
		
		exceptionRowsCount = postconditionRow - exceptionRow - 1 ' Calcular cantidad filas excepción tabla original
		
		offset = normalSequenceRowsCount - templateBaseRows ' Calcular ajuste de filas
		
		editRows exceptionRowsCount, templateExceptionRow + (templateBaseRows - 1) + offset, templatePostconditionRow + offset, templateDoc ' Edita filas excepción plantilla 
		
		editTitle templateDoc, title ' Editar título plantilla
		
		editHeader templateDoc, originalDoc ' Editar cabezal planilla
		
		editBody templateDoc, originalDoc, normalSequenceRowsCount, templateNormalSequenceRow + 1 ' Editar secuencia normal plantilla
		
		editBody templateDoc, originalDoc, exceptionRowsCount, templateExceptionRow + offset ' Editar excepción plantilla
		
		editFooter templateDoc, originalDoc, postconditionRow ' Editar pie plantilla
		
		saveCopy templateDoc, originalDoc
		
		templateDoc.Close SaveChanges:=False
		
		Set templateDoc = Nothing
		
		originalDoc.Close SaveChanges:=False
		
		Set originalDoc = Nothing
		
	Next i
	
	Application.ScreenUpdating = True
	
	Exit Sub
	
	CleanUp:
		Application.ScreenUpdating = True

		' Cerrar templateDoc si sigue abierto
		If Not templateDoc Is Nothing Then
			templateDoc.Close SaveChanges:=False
			Set templateDoc = Nothing
		End If

		' Cerrar originalDoc si sigue abierto
		If Not originalDoc Is Nothing Then
			originalDoc.Close SaveChanges:=False
			Set originalDoc = Nothing
		End If

		If Err.Number <> 0 Then
			Dim errMsg As String
			errMsg = "Error: " & Err.Description
			On Error Resume Next
			errMsg = "Error en archivo " & filesFolder(i) & ":" & vbNewLine & Err.Description
			On Error GoTo 0
			MsgBox errMsg
		End If
	
End Sub

Sub obtainFile(ByRef pPathFile As String)
	Dim dlg As FileDialog
	Set dlg = Application.FileDialog(msoFileDialogFilePicker)
	
	With dlg
		.Title = "Selecciona un archivo"
		.AllowMultiSelect = False
		.Filters.Clear
		.Filters.Add "Documentos Word", "*.docx"
		
		If .Show = -1 Then
			pPathFile = .SelectedItems(1)
		Else
			MsgBox "No se seleccionó ningún archivo. El macro se detendrá."
			End
		End If
	End With
	
	Debug.Print pPathFile
	
	Set dlg = Nothing
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

Sub obtainFiles(ByVal pPathFolder As String )
	Dim file As String
	Dim i As Integer
	
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

Sub obtainTitle(ByRef pTitle As String, ByVal pDoc As Document)
	pTitle = pDoc.Paragraphs(1).Range.Text
	
	Debug.Print "Título: " & pTitle
End Sub

Sub obtainTableRowsCount(ByRef pTableRowsCount As Integer, ByVal pDoc As Document)
	pTableRowsCount = pDoc.Tables(1).Rows.Count
	
	Debug.Print "Cantidad de filas: " & pTableRowsCount
End Sub

Sub obtainRow(ByVal pSearchedWord As String, ByRef pTableRow As Integer, ByVal pDoc As Document)
	Dim tbl As Table
	Dim j As Integer

	Set tbl = pDoc.Tables(1)

	For j = 1 To tableRowsCount
		If InStr(tbl.Cell(j, 1).Range.Text, pSearchedWord) > 0 Then
			pTableRow = j
			Debug.Print "Fila de " & pSearchedWord & ": " & pTableRow
			Exit Sub
		End If
	Next j
	
	MsgBox "Palabra no encontrada: " & pSearchedWord
	Err.Raise vbObjectError + 1 
End Sub

Sub editRows(ByVal pRowsCount As Integer, ByVal pTemplateInsertRow As Integer, ByVal pTemplateDeleteRow As Integer, ByRef pTemplateDoc As Document)
	Dim templateTable As Table
	Dim i As Integer
	
	Set templateTable = pTemplateDoc.Tables(1)
	
	Select Case pRowsCount
		Case Is > templateBaseRows
			For i = 1 To pRowsCount - templateBaseRows
				templateTable.Rows.Add BeforeRow:=templateTable.Rows(pTemplateInsertRow)
			Next i

		Case 1 To 2
			For i = 1 To templateBaseRows - pRowsCount
				templateTable.Rows(pTemplateDeleteRow - i).Delete
			Next i

		Case 0
			MsgBox "No hay filas."
			Err.Raise vbObjectError + 1

		Case Else
			MsgBox "No hay que hacer cambios."
	End Select
End Sub

Sub editTitle(ByRef pTemplateDoc As Document, ByVal pTitle As String)
	Dim rng As Range
	
	Set rng = pTemplateDoc.Paragraphs(1).Range
	
	rng.End = rng.End - 1 ' Excluye marcador párrafo
	rng.Text = pTitle
	
	pTemplateDoc.Paragraphs(2).Range.Delete
End Sub

Sub editHeader(ByRef pTemplateDoc As Document, ByVal pOriginalDoc As Document)
	Dim i As Integer
	Dim j As Integer
	
	Dim originalTable As Table
	Dim templateTable As Table
	
	Set originalTable = pOriginalDoc.Tables(1)
	Set templateTable = pTemplateDoc.Tables(1)
	
	For i = 1 To 2
		For j = 1 To 2
			templateTable.Cell(i, j).Range.Text = Left(originalTable.Cell(i, j).Range.Text, Len(originalTable.Cell(i, j).Range.Text) - cellMarkerLen)
		Next j
	Next i
	
	For i = 3 To 5
		For j = 1 To 1
			templateTable.Cell(i, j).Range.Text = Left(originalTable.Cell(i, j).Range.Text, Len(originalTable.Cell(i, j).Range.Text) - cellMarkerLen)
		Next j
	Next i
	
	For i = 6 To 6
		For j = 1 To 2
			templateTable.Cell(i, j).Range.Text = Left(originalTable.Cell(i, j).Range.Text, Len(originalTable.Cell(i, j).Range.Text) - cellMarkerLen)
		Next j
	Next i

End Sub

Sub editBody(ByRef pTemplateDoc As Document, ByVal pOriginalDoc As Document, ByVal pRowsCount As Integer, ByVal pTemplateStartRow As Integer)
	Dim i As Integer
	Dim j As Integer
	
	Dim originalTable As Table
	Dim templateTable As Table
	
	Set originalTable = pOriginalDoc.Tables(1)
	Set templateTable = pTemplateDoc.Tables(1)
	
	For i = pTemplateStartRow To pTemplateStartRow + pRowsCount
		For j = 1 To 3
			templateTable.Cell(i, j).Range.Text = Left(originalTable.Cell(i, j).Range.Text, Len(originalTable.Cell(i, j).Range.Text) - cellMarkerLen)
		Next j
	Next i
	
End Sub

Sub editFooter(ByRef pTemplateDoc As Document, ByVal pOriginalDoc As Document, ByVal pTemplateStartRow As Integer)
	Dim i As Integer
	Dim j As Integer

	Dim originalTable As Table
	Dim templateTable As Table
	
	Set originalTable = pOriginalDoc.Tables(1)
	Set templateTable = pTemplateDoc.Tables(1)
	
	For i = pTemplateStartRow To pTemplateStartRow + 3
		For j = 1 To 1
			templateTable.Cell(i, j).Range.Text = Left(originalTable.Cell(i, j).Range.Text, Len(originalTable.Cell(i, j).Range.Text) - cellMarkerLen)
		Next j
	Next i
	
End Sub

Sub saveCopy(ByRef pTemplateDoc As Document, ByVal pOriginalDoc As Document)
	Dim destinyFolder As String
	
	destinyFolder = pOriginalDoc.Path & "\FIXED FILES"
	
	If Dir(destinyFolder, vbDirectory) = "" Then MkDir destinyFolder
	
	pTemplateDoc.SaveAs2 FileName:=destinyFolder & "\" & pOriginalDoc.Name, FileFormat:=wdFormatXMLDocument
	
End Sub