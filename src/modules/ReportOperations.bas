Attribute VB_Name = "ReportOperations"
Option Explicit

Public Const INPUT_WORKSHEET_NAME As String = "Datasheet"
Public Const OUTPUT_WORKSHEET_NAME As String = "Reports"

Public Const TABLE_NAME As String = "__datatable__"
Public Const TARGET_COLUMN As Long = 35

' Borra los datos existentes en el reporte
Public Sub ClearReport()
  Dim wsOutput As Worksheet
  Dim i As Long
  
  Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_WORKSHEET_NAME)

  DisableApplicationSettings True

  For i = 1 To 4
    wsOutput.Range("_report" & i).ClearContents
  Next i

  DisableApplicationSettings False

  MsgBox "¡Reporte borrado exitosamente!", vbInformation
End Sub

' Genera el reporte de mortalidad y morbilidad basado en los datos de entrada y los filtros definidos.
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

  Dim frm As frmProgress
  Dim TotalSteps As Long
  Dim CurrentStep As Long
  
  ' Inicializa el formulario de progreso
  Set frm = New frmProgress
  frm.Show vbModeless
  Set Utils.frm = frm
  
  On Error GoTo ErrHandler
  LogMessage "Iniciando GenerateReport..."
  
  Set wsInput = ThisWorkbook.Worksheets(INPUT_WORKSHEET_NAME)
  Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_WORKSHEET_NAME)
  Set tbl = wsInput.ListObjects(TABLE_NAME)
  
  If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
    LogMessage "Error: ¡La tabla de entrada '" & TABLE_NAME & "' es inexistente o está vacía!", LOG_ERROR
    Exit Sub
  End If
  
  data = tbl.DataBodyRange.Value
  LogMessage "Datos cargados. Filas: " & UBound(data, 1) & ", Columnas: " & UBound(data, 2)
  
  mun = wsOutput.Range("W2").Value
  filters = Array(Array(0, vbNullString), Array(12, "FEMENINO"), Array(12, "MASCULINO"), Array(6, mun))
  outRanges = Array("C6", "C34", "C62", "C90")
  
  If UBound(filters) <> UBound(outRanges) Then
    LogMessage "Error: ¡El número de filtros no coincide con el número de rangos de salida!", LOG_ERROR
    Exit Sub
  End If
  
  TotalSteps = (UBound(filters) - LBound(filters) + 1) * 25
  CurrentStep = 0
  frm.TotalSteps = TotalSteps
  frm.CurrentStep = 0
  
  DisableApplicationSettings True
  
  ' Procesa cada filtro y genera el reporte correspondiente
  For i = LBound(filters) To UBound(filters)
    filterCol = filters(i)(0)
    filterVal = filters(i)(1)
    LogMessage "Procesando filtro índice " & i & ": filterCol=" & filterCol & ", filterVal=" & filterVal
    
    Set freqDict = BuildFilteredFrequencyDict(data, filterCol, filterVal, TARGET_COLUMN)
    LogMessage "Conteo del diccionario de frecuencias: " & freqDict.Count

    sortedKeys = freqDict.keys
    SortKeysByFrequencyDescending sortedKeys, freqDict

    topArr = GetTopNArray(sortedKeys, 25)
    WriteTopNToRange topArr, wsOutput.Range(outRanges(i)), 25
    LogMessage "Top N escrito en " & outRanges(i)

    frm.CurrentStep = frm.CurrentStep + 12
    frm.UpdateProgress frm.CurrentStep, frm.TotalSteps

    WriteICD11LabelsToRange topArr, wsOutput.Range(outRanges(i)).Offset(0, 1), frm
    LogMessage "Etiquetas ICD-11 escritas en " & wsOutput.Range(outRanges(i)).Offset(0, 1).Address

    frm.CurrentStep = frm.CurrentStep + 13
    frm.UpdateProgress frm.CurrentStep, frm.TotalSteps

    LogMessage "Índice de filtro " & i & " procesado completamente."
  Next i

  DisableApplicationSettings False
  
  LogMessage "¡Generación de reporte completa!"
  frm.UpdateProgress TotalSteps, TotalSteps
  MsgBox "¡Reporte generado exitosamente!", vbInformation
  
  Unload frm
  Exit Sub
  
ErrHandler:
  DisableApplicationSettings False
  MsgBox "Error en GenerateReport: " & Err.Description, vbCritical
  LogMessage "Error en GenerateReport: " & Err.Description, LOG_ERROR
  On Error Resume Next
  Unload frm
End Sub

' Construye un diccionario de frecuencias para los valores en targetCol,
' opcionalmente filtrando por filterCol = filterVal.
Private Function BuildFilteredFrequencyDict(ByVal data As Variant, _
  ByVal filterCol As Long, ByVal filterVal As Variant, _
  ByVal targetCol As Long) As Object

  Dim r As Long, val As Variant
  Dim rowsCount As Long, colsCount As Long
  Dim dict As Object
  
  rowsCount = UBound(data, 1)
  colsCount = UBound(data, 2)
  Set dict = CreateObject("Scripting.Dictionary")
  
  LogMessage "Construyendo diccionario de frecuencias... TargetCol=" & targetCol & ", FilterCol=" & filterCol
  
  If targetCol > colsCount Or targetCol < 1 Then
    LogMessage "¡TARGET_COLUMN fuera de rango! Máximo de columnas=" & colsCount, LOG_ERROR
    Set BuildFilteredFrequencyDict = dict
    Exit Function
  End If
  
  If filterCol <> 0 Then
    If filterCol > colsCount Or filterCol < 1 Then
      LogMessage "Error: ¡Columna de filtro fuera de rango! Máximo de columnas=" & colsCount, LOG_ERROR
      Set BuildFilteredFrequencyDict = dict
      Exit Function
    End If
  End If
  
  ' Recorre cada fila y cuenta las frecuencias de los valores en targetCol, aplicando el filtro si es necesario
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

' Ordena un array de claves basado en las frecuencias almacenadas en freqDict, de mayor a menor.
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

' Obtiene los primeros N elementos de un array, o todos si el array tiene menos de N elementos.
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

' Escribe un array de valores en una columna a partir de una celda inicial, limitando a N filas.
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

' Escribe las etiquetas ICD-11 correspondientes a un array de códigos en una columna
' a partir de una celda inicial, mostrando el progreso en el formulario.
Private Sub WriteICD11LabelsToRange(ByVal arr As Variant, ByVal startCell As Range, ByRef frm As frmProgress)
  Dim code As String, label As String
  Dim outputArr() As Variant
  Dim totalItems As Long
  Dim i As Long, idx As Long
  
  If Not IsArray(arr) Then Exit Sub
  If UBound(arr) < LBound(arr) Then Exit Sub
  
  totalItems = UBound(arr) - LBound(arr) + 1
  ReDim outputArr(1 To totalItems, 1 To 1)
  
  LogMessage "Iniciando bucle de recuperación de etiquetas ICD-11 para " & totalItems & " códigos..."
  
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
  LogMessage "Recuperación de etiquetas ICD-11 completada para " & totalItems & " códigos."
End Sub

