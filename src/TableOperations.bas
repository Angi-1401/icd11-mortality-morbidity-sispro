Option Explicit

Public Const WORKSHEET_NAME As String = "Datasheet"
Public Const TABLE_NAME As String = "__datatable__"

Public Sub ClearTableData()
  Dim ws As Worksheet
  Dim tbl As ListObject

  ' Select the target worksheet and table
  Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)
  Set tbl = ws.ListObjects(TABLE_NAME)
  If tbl Is Nothing Then
    MsgBox "Table '" & TABLE_NAME & "' not found in worksheet '" & WORKSHEET_NAME & "'.", vbCritical
    Exit Sub
  End If

  Call DisableApplicationSettings(True)

  If Not tbl.DataBodyRange Is Nothing Then
    tbl.DataBodyRange.Delete
    MsgBox "All data cleared from table '" & TABLE_NAME & "'.", vbInformation
  Else
    MsgBox "Table '" & TABLE_NAME & "' is already empty.", vbInformation
  End If

  Call DisableApplicationSettings(False)
End Sub

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

  Dim userResponse As VbMsgBoxResult

  Const WORKSHEET_NAME As String = "Datasheet"
  Const TABLE_NAME As String = "__datatable__"

  ' Columns to import (1-based positions in the .txt file)
  importedCols = Array(3, 6, 12, 13, 15, 17, 19, 22, 23, 24, 25, 27, 28, 29, 30, 32, 34, 36, 38, 41, 42, 44, 46, 48, 52, 56, 130, 132, 134, 136)

  ' Prompt the user to select .txt file
  filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select text file to import")
  If filePath = "False" Then Exit Sub ' User cancelled

  ' Select the target worksheet and table
  Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)
  Set tbl = ws.ListObjects(TABLE_NAME)
  If tbl Is Nothing Then
    MsgBox "Table '" & TABLE_NAME & "' not found in worksheet '" & WORKSHEET_NAME & "'.", vbCritical
    Exit Sub
  End If

  If Not tbl.DataBodyRange Is Nothing Then
    userResponse = MsgBox("The table already contains data. Do you want to clear existing data before importing?", vbYesNoCancel + vbQuestion, "Clear Existing Data")
    If userResponse = vbCancel Then Exit Sub
    If userResponse = vbYes Then
      tbl.DataBodyRange.Delete
    End If
  End If

  Call DisableApplicationSettings(True)

  ' Read the entire file content
  fileNum = FreeFile
  Open filePath For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
  Close #fileNum

  ' Split content into lines
  lines = Split(fileContent, vbCrLf)
  ReDim dataOut(1 To UBound(lines) + 1, 1 To UBound(importedCols) + 1)

  ' Process each line
  Dim i As Long, j As Long, k As Long
  k = 0
  For i = LBound(lines) To UBound(lines)
    Select Case i
      Case LBound(lines)
        ' Header line
        ' Do nothing
      
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

  ' Handle column mismatch
  colCount = tbl.ListColumns.Count
  importedCols = UBound(importedCols) + 1

  If importedCols <> colCount Then
    userResponse = MsgBox( _
      "Imported data has " & importedCols & " columns, but table '" & TABLE_NAME & "' has " & colCount & ". "  & _
      "Do you want to continue and insert data only into the first " & importedCols & " columns?", _
      vbYesNo + vbQuestion, "Column Mismatch")
    
    If userResponse = vbNo Then GoTo Cleanup
  End If

  ' Resize table if needed
  tbl.Resize tbl.Range.Resize(k + 1) ' Keep same column count

  ' Insert data into the table
  Set targetRange = tbl.DataBodyRange.Resize(k, importedCols)
  targetRange.Value = dataOut

  startCol = 27
  endCol = 30

  data = tbl.DataBodyRange.Value
  ReDim result(1 To UBound(data, 1), 1 To endCol - startCol + 1)

  ' Process each row and specified columns in memory
  For i = 1 To UBound(data, 1)
    For j = startCol To endCol
      result(i, j - startCol + 1) = ExtractICD11Code(CStr(data(i, j)))
    Next j
  Next i

  ' Write results back to the table
  tbl.DataBodyRange.Columns(31).Resize(, endCol - startCol + 1).Value = result
    
  MsgBox "Import complete: " & k & " rows inserted into '" & TABLE_NAME & "'.", vbInformation

Cleanup:
  Call DisableApplicationSettings(False)
End Sub

Private Function DisableApplicationSettings(ByVal disable As Boolean)
  If disable Then
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
  Else
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
  End If
End Function