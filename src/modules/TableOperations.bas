Attribute VB_Name = "TableOperations"
Option Explicit

Public Const WORKSHEET_NAME As String = "Datasheet"
Public Const TABLE_NAME As String = "__datatable__"

' Borra los datos existentes en la tabla de entrada.
Public Sub ClearTableData()
  Dim ws As Worksheet
  Dim tbl As ListObject

  Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)
  Set tbl = ws.ListObjects(TABLE_NAME)
  If tbl Is Nothing Then
    MsgBox "No se encontró la tabla '" & TABLE_NAME & "' en la hoja '" & WORKSHEET_NAME & "'.", vbCritical
    Exit Sub
  End If

  DisableApplicationSettings True

  If Not tbl.DataBodyRange Is Nothing Then
    tbl.DataBodyRange.Delete
    MsgBox "Datos borrados de la tabla '" & TABLE_NAME & "'.", vbInformation
  Else
    MsgBox "La tabla '" & TABLE_NAME & "' ya está vacía.", vbInformation
  End If

  DisableApplicationSettings False
End Sub

' Importa datos desde un archivo de texto delimitado por tabulaciones a la tabla de entrada,
' extrayendo códigos ICD-11 en el proceso.
Public Sub PopulateTableFromTXT()
  Dim ws As Worksheet
  Dim tbl As ListObject
  Dim targetRange As Range

  Dim filePath As String
  Dim fileNum As Integer
  Dim fileContent As String
  Dim lines As Variant, lineParts As Variant, dataOut() As Variant
  Dim importedCols As Variant, data As Variant, result() As Variant
  Dim colCount As Long, startCol As Long, endCol As Long
  Dim i As Long, j As Long, k As Long

  Dim userResponse As VbMsgBoxResult

  importedCols = Array(3, 6, 12, 13, 15, 17, 19, 22, 23, 24, 25, 27, 28, 29, 30, 32, 34, 36, 38, 41, 42, 44, 46, 48, 52, 56, 130, 132, 134, 136)

  filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select text file to import")
  If filePath = "False" Then Exit Sub ' Usuario cancela la operación

  Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)
  Set tbl = ws.ListObjects(TABLE_NAME)
  If tbl Is Nothing Then
    MsgBox "No se encontró la tabla '" & TABLE_NAME & "' en la hoja '" & WORKSHEET_NAME & "'.", vbCritical
    Exit Sub
  End If

  ' Si la tabla ya tiene datos, pregunta al usuario si desea borrarlos antes de importar los nuevos datos
  If Not tbl.DataBodyRange Is Nothing Then
    userResponse = MsgBox("La tabla ya contiene datos. ¿Desea borrar los datos existentes antes de importar?", vbYesNoCancel + vbQuestion, "Borrar datos existentes")
    If userResponse = vbCancel Then Exit Sub
    If userResponse = vbYes Then
      tbl.DataBodyRange.Delete
    End If
  End If

  DisableApplicationSettings True

  ' Lee el contenido del archivo seleccionado en una variable de texto
  fileNum = FreeFile
  Open filePath For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
  Close #fileNum

  lines = Split(fileContent, vbCrLf)
  ReDim dataOut(1 To UBound(lines) + 1, 1 To UBound(importedCols) + 1)
  
  ' Procesa cada línea del archivo, extrayendo solo las columnas especificadas en importedCols
  ' y limpiando los datos para insertar en la tabla. También cuenta el número de filas importadas.
  k = 0
  For i = LBound(lines) To UBound(lines)
    Select Case i
      Case LBound(lines)
        ' Línea del encabezado
        ' No hacer nada
      
      Case Else
        If Len(Trim$(lines(i))) > 0 Then
          lineParts = Split(lines(i), vbTab)
          k = k + 1
          For j = LBound(importedCols) To UBound(importedCols)
            If UBound(lineParts) >= importedCols(j) - 1 Then
              dataOut(k, j + 1) = UCase(Trim$(lineParts(importedCols(j) - 1)))
            Else
              dataOut(k, j + 1) = ""
            End If
          Next j
        End If
    End Select
  Next i

  colCount = tbl.ListColumns.Count
  importedCols = UBound(importedCols) + 1

  ' Si el número de columnas importadas es menor que el número de columnas en la tabla,
  ' pregunta al usuario si desea continuar.
  If importedCols <> colCount Then
    userResponse = MsgBox( _
      "Los datos importados tienen " & importedCols & " columnas, pero la tabla '" & TABLE_NAME & "' tiene " & colCount & ". " & _
      "¿Desea continuar e insertar datos solo en las primeras " & importedCols & " columnas?", _
      vbYesNo + vbQuestion, "Desajuste de columnas")
    
    If userResponse = vbNo Then GoTo Cleanup
  End If

  ' Redimensiona la tabla para acomodar el número de filas importadas y
  ' luego inserta los datos procesados en la tabla.
  tbl.Resize tbl.Range.Resize(k + 1)

  Set targetRange = tbl.DataBodyRange.Resize(k, importedCols)
  targetRange.Value = dataOut

  startCol = 27
  endCol = 30

  data = tbl.DataBodyRange.Value
  ReDim result(1 To UBound(data, 1), 1 To endCol - startCol + 1)

  For i = 1 To UBound(data, 1)
    For j = startCol To endCol
      result(i, j - startCol + 1) = ExtractICD11Code(CStr(data(i, j)))
    Next j
  Next i

  tbl.DataBodyRange.Columns(31).Resize(, endCol - startCol + 1).Value = result
    
  MsgBox "Importación completa: " & k & " filas insertadas en '" & TABLE_NAME & "'.", vbInformation

Cleanup:
  DisableApplicationSettings False
End Sub

