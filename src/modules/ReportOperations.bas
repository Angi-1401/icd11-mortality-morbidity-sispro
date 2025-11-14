Attribute VB_Name = "ReportOperations"
Option Explicit

Public Const INPUT_WORKSHEET_NAME As String = "Datasheet"
Public Const OUTPUT_WORKSHEET_NAME As String = "Reports"

Public Const TABLE_NAME As String = "__datatable__"
Public Const TARGET_COLUMN As Long = 35

Public Sub ClearReport()
  Dim wsOutput As Worksheet
  Dim i As Long
  
  Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_WORKSHEET_NAME)

  DisableApplicationSettings True

  For i = 1 To 4
    wsOutput.Range("_report" & i).ClearContents
  Next i

  DisableApplicationSettings False

  MsgBox "Report cleared successfully!", vbInformation
End Sub

Public Sub GenerateReport()
  Dim wsInput As Worksheet, wsOutput As Worksheet
  Dim tbl As ListObject

  Dim data As Variant
  Dim outRanges As Variant, filters As Variant
  Dim filterCol As Long, filterVal As Variant
  Dim mun As String

  Dim freqDict As Object
  Dim sortedKeys() As Variant
  Dim topArr As Variant
  Dim i As Long

  Dim frm As FrmProgress
  Dim TotalSteps As Long
  Dim CurrentStep As Long
  
  Set frm = New FrmProgress
  frm.Show vbModeless
  Set Utils.frm = frm
  
  On Error GoTo ErrHandler
  LogMessage "Starting GenerateReport..."
  
  Set wsInput = ThisWorkbook.Worksheets(INPUT_WORKSHEET_NAME)
  Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_WORKSHEET_NAME)
  Set tbl = wsInput.ListObjects(TABLE_NAME)
  
  If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
    LogMessage "Error: Input table '" & TABLE_NAME & "' is missing or empty!", LOG_ERROR
    Exit Sub
  End If
  
  data = tbl.DataBodyRange.Value
  LogMessage "Table loaded. Rows: " & UBound(data, 1) & ", Columns: " & UBound(data, 2)
  
  mun = wsOutput.Range("W2").Value
  filters = Array(Array(0, vbNullString), Array(12, "FEMENINO"), Array(12, "MASCULINO"), Array(6, mun))
  outRanges = Array("C6", "C34", "C62", "C90")
  
  If UBound(filters) <> UBound(outRanges) Then
    LogMessage "Error: Filters and output ranges mismatch!", LOG_ERROR
    Exit Sub
  End If
  
  TotalSteps = (UBound(filters) - LBound(filters) + 1) * 25
  CurrentStep = 0
  frm.TotalSteps = TotalSteps
  frm.CurrentStep = 0
  
  DisableApplicationSettings True
  
  For i = LBound(filters) To UBound(filters)
    filterCol = filters(i)(0)
    filterVal = filters(i)(1)
    LogMessage "Processing filter index " & i & ": filterCol=" & filterCol & ", filterVal=" & filterVal
    
    Set freqDict = BuildFilteredFrequencyDict(data, filterCol, filterVal, TARGET_COLUMN)
    LogMessage "Frequency dictionary count: " & freqDict.Count

    sortedKeys = freqDict.keys
    SortKeysByFrequencyDescending sortedKeys, freqDict

    topArr = GetTopNArray(sortedKeys, 25)
    WriteTopNToRange topArr, wsOutput.Range(outRanges(i)), 25
    LogMessage "Top N written to " & outRanges(i)

    frm.CurrentStep = frm.CurrentStep + 12
    frm.UpdateProgress frm.CurrentStep, frm.TotalSteps

    WriteICD11LabelsToRange topArr, wsOutput.Range(outRanges(i)).Offset(0, 2), frm
    LogMessage "ICD-11 labels written to " & wsOutput.Range(outRanges(i)).Offset(0, 2).Address

    frm.CurrentStep = frm.CurrentStep + 13
    frm.UpdateProgress frm.CurrentStep, frm.TotalSteps

    LogMessage "Filter index " & i & " processed completely."
  Next i

  DisableApplicationSettings False
  
  LogMessage "Report generation complete!"
  frm.UpdateProgress TotalSteps, TotalSteps
  MsgBox "Report generated successfully!", vbInformation
  
  Unload frm
  Exit Sub
  
ErrHandler:
  DisableApplicationSettings False
  MsgBox "Error in GenerateReport: " & Err.Description, vbCritical
  LogMessage "Error in GenerateReport: " & Err.Description, LOG_ERROR
  On Error Resume Next
  Unload frm
End Sub

Private Function BuildFilteredFrequencyDict(ByVal data As Variant, _
  ByVal filterCol As Long, ByVal filterVal As Variant, _
  ByVal targetCol As Long) As Object

  Dim r As Long, val As Variant
  Dim rowsCount As Long, colsCount As Long
  Dim dict As Object
  
  rowsCount = UBound(data, 1)
  colsCount = UBound(data, 2)
  Set dict = CreateObject("Scripting.Dictionary")
  
  LogMessage "Building frequency dictionary... TargetCol=" & targetCol & ", FilterCol=" & filterCol
  
  If targetCol > colsCount Or targetCol < 1 Then
    LogMessage "TARGET_COLUMN out of range! Max columns=" & colsCount, LOG_ERROR
    Set BuildFilteredFrequencyDict = dict
    Exit Function
  End If
  
  If filterCol <> 0 Then
    If filterCol > colsCount Or filterCol < 1 Then
      LogMessage "Error: Filter column out of range! Max columns=" & colsCount, LOG_ERROR
      Set BuildFilteredFrequencyDict = dict
      Exit Function
    End If
  End If
  
  For r = 1 To rowsCount
    val = data(r, targetCol)
    If Not IsError(val) Then
      If Trim(CStr(val)) <> "" Then
        If filterCol = 0 Then
          dict(val) = IIf(dict.Exists(val), dict(val) + 1, 1)
        Else
          Dim fVal As Variant
          fVal = data(r, filterCol)
          If Not IsError(fVal) Then
            If StrComp(CStr(fVal), CStr(filterVal), vbTextCompare) = 0 Then
              dict(val) = IIf(dict.Exists(val), dict(val) + 1, 1)
            End If
          End If
        End If
      End If
    End If
  Next r
  
  Set BuildFilteredFrequencyDict = dict
End Function

Private Function SortKeysByFrequencyDescending(ByRef keys As Variant, ByVal freqDict As Object)
  Dim i As Long, j As Long, tmp As Variant
  If Not IsArray(keys) Then Exit Function
  If UBound(keys) < LBound(keys) Then Exit Function
  
  For i = LBound(keys) To UBound(keys) - 1
    For j = i + 1 To UBound(keys)
      If freqDict(keys(i)) < freqDict(keys(j)) Then
        tmp = keys(i)
        keys(i) = keys(j)
        keys(j) = tmp
      End If
    Next j
  Next i
End Function

Private Function GetTopNArray(ByVal arr As Variant, Optional ByVal n As Long = 25) As Variant
  Dim lim As Long, i As Long
  Dim res() As Variant

  If Not IsArray(arr) Then
    GetTopNArray = Array()
    Exit Function
  End If

  If UBound(arr) < LBound(arr) Then
    GetTopNArray = Array()
    Exit Function
  End If

  If n <= 0 Then n = 25
  lim = WorksheetFunction.Min(n, UBound(arr) - LBound(arr) + 1)
  ReDim res(1 To lim)

  For i = 1 To lim
    res(i) = arr(LBound(arr) + i - 1)
  Next i

  GetTopNArray = res
End Function

Private Sub WriteTopNToRange(ByVal arr As Variant, ByVal startCell As Range, Optional ByVal n As Long = 25)
  Dim lim As Long, i As Long
  Dim outputArr() As Variant
  
  If Not IsArray(arr) Then Exit Sub
  If UBound(arr) < LBound(arr) Then Exit Sub
  
  If n <= 0 Then n = 25
  lim = WorksheetFunction.Min(n, UBound(arr) - LBound(arr) + 1)
  ReDim outputArr(1 To lim, 1 To 1)
  
  For i = 1 To lim
    outputArr(i, 1) = arr(LBound(arr) + i - 1)
  Next i
  
  startCell.Resize(lim, 1).Value = outputArr
End Sub

Private Sub WriteICD11LabelsToRange(ByVal arr As Variant, ByVal startCell As Range, ByRef frm As FrmProgress)
  Dim code As String, label As String
  Dim outputArr() As Variant
  Dim totalItems As Long
  Dim i As Long, idx As Long
  
  If Not IsArray(arr) Then Exit Sub
  If UBound(arr) < LBound(arr) Then Exit Sub
  
  totalItems = UBound(arr) - LBound(arr) + 1
  ReDim outputArr(1 To totalItems, 1 To 1)
  
  LogMessage "Starting ICD-11 label retrieval for " & totalItems & " codes..."
  
  idx = 1
  For i = LBound(arr) To UBound(arr)
    code = Trim(CStr(arr(i)))
    If code <> "" Then
      label = GetICD11CodeLabel(code)
    Else
      label = ""
    End If
    
    outputArr(idx, 1) = label
    idx = idx + 1
  Next i
  
  startCell.Resize(totalItems, 1).Value = outputArr
  LogMessage "ICD-11 label retrieval completed for " & totalItems & " codes."
End Sub

