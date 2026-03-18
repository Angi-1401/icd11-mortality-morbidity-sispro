Attribute VB_Name = "ICD11"
Option Explicit

Public Const CLIENT_ID As String = "TU_CLIENT_ID_AQUI"
Public Const CLIENT_SECRET As String = "TU_CLIENT_SECRET_AQUI"

' Extrae un posible código ICD-11 de una cadena que contiene un guion antes del código.
Public Function ExtractICD11Code(ByVal str As String) As String
  Dim dashPos As Long
  Dim segment As String
  Dim i As Long
  Dim ch As String
  Dim code As String

  dashPos = InStrRev(str, "-")
  If dashPos > 0 Then
    segment = Mid$(str, dashPos + 1)
  Else
    Exit Function
  End If

  code = ""
  For i = 1 To Len(segment)
    ch = Mid$(segment, i, 1)
    ' Permite letras, números, espacios, ampersands, barras y puntos
    If ch Like "[A-Za-z0-9 &/\.]" Then
      code = code & ch
    End If
  Next i

  ' Reemplaza espacios por puntos
  code = Replace(code, " ", ".")
  ExtractICD11Code = code
End Function

' Recupera la etiqueta de un código ICD-11 usando el endpoint de descripción.
Public Function GetICD11CodeLabel(ByVal code As String) As String
  Dim http As Object
  Dim token As String
  Dim jsonResponse As String
  Dim url As String
  Dim label As String

  On Error GoTo ErrHandler

  token = GetICD11AccessToken()
  If token = "" Then
    LogMessage "GetICD11CodeLabel: Unable to retrieve access token.", LOG_ERROR
    GetICD11CodeLabel = "Error: Unable to retrieve access token."
    Exit Function
  End If

  If Not IsValidICD11Code(code) Then
    LogMessage "GetICD11CodeLabel: Invalid ICD-11 code format: " & code, LOG_WARNING
    GetICD11CodeLabel = "Error: Invalid ICD-11 code format."
    Exit Function
  End If

  ' Codifica en URL los caracteres problemáticos
  code = Replace(code, "&", "%26")
  code = Replace(code, "/", "%2F")
  code = Replace(code, " ", "%20")

  url = "https://id.who.int/icd/release/11/2025-01/mms/describe?code=" & code

  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  http.Open "GET", url, False
  http.setRequestHeader "Accept", "application/json"
  http.setRequestHeader "API-Version", "v2"
  http.setRequestHeader "Accept-Language", "es"
  http.setRequestHeader "Authorization", "Bearer " & token
  http.Send

  If http.Status <> 200 Then
    LogMessage "GetICD11CodeLabel: API request failed. Status=" & http.Status & " URL=" & url, LOG_ERROR
    GetICD11CodeLabel = "Error: API request failed with status " & http.Status
    Exit Function
  End If

  jsonResponse = http.responseText
  LogMessage "GetICD11CodeLabel: JSON response length=" & Len(jsonResponse), LOG_DEBUG

  label = ParseJSONValue(jsonResponse, """label"":""", """")
  If label <> "" Then
    label = DecodeUnicode(label)
    GetICD11CodeLabel = label
  Else
    LogMessage "GetICD11CodeLabel: Label not found in response for code " & code & ". JSON excerpt: " & Left(jsonResponse, 300), LOG_WARNING
    GetICD11CodeLabel = "Error: Label not found for code " & code
  End If

  Exit Function

ErrHandler:
  LogMessage "GetICD11CodeLabel Error: " & Err.Number & " - " & Err.Description, LOG_ERROR
  GetICD11CodeLabel = "Error: " & Err.Description
End Function

' Gets (and caches) the access token for the ICD API
Private Function GetICD11AccessToken() As String
  Static cachedToken As String
  Static tokenExpiry As Date

  Dim http As Object
  Dim postData As String
  Dim jsonResponse As String
  Dim token As String
  Dim expiresInText As String
  Dim expiresIn As Long

  Const TOKEN_URL As String = "https://icdaccessmanagement.who.int/connect/token"

  On Error GoTo ErrHandler

  ' Usa el token en caché si aún es válido
  If cachedToken <> "" And Now < tokenExpiry Then
    GetICD11AccessToken = cachedToken
    Exit Function
  End If

  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

  postData = "grant_type=client_credentials" & _
    "&client_id=" & CLIENT_ID & _
    "&client_secret=" & CLIENT_SECRET & _
    "&scope=icdapi_access"

  http.Open "POST", TOKEN_URL, False
  http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  http.Send postData

  If http.Status <> 200 Then
    LogMessage "GetICD11AccessToken: Token request failed with status " & http.Status & ". Response: " & Left(http.responseText, 300), LOG_ERROR
    GetICD11AccessToken = ""
    Exit Function
  End If

  jsonResponse = http.responseText
  LogMessage "GetICD11AccessToken: token response length=" & Len(jsonResponse), LOG_DEBUG

  token = ParseJSONValue(jsonResponse, """access_token"":""", """")
  expiresInText = ParseJSONValue(jsonResponse, """expires_in"":", ",")

  If token <> "" Then
    cachedToken = token
    On Error Resume Next
    expiresIn = CLng(Trim(expiresInText))
    If Err.Number <> 0 Then
      ' Si no se pudo analizar, usa 300 segundos por defecto.
      expiresIn = 300
      Err.Clear
    End If
    On Error GoTo 0
    tokenExpiry = DateAdd("s", expiresIn - 60, Now) ' 60s como buffer
    GetICD11AccessToken = cachedToken
  Else
    LogMessage "GetICD11AccessToken: access_token not found in response. JSON excerpt: " & Left(jsonResponse, 300), LOG_ERROR
    GetICD11AccessToken = ""
  End If

  Exit Function

ErrHandler:
  LogMessage "GetICD11AccessToken Error: " & Err.Number & " - " & Err.Description, LOG_ERROR
  GetICD11AccessToken = ""
End Function

' Valida el formato básico del código ICD-11 (letras/números, con segmentos separados por puntos)
Private Function IsValidICD11Code(ByVal code As String) As Boolean
  Dim regex As Object
  Set regex = CreateObject("VBScript.RegExp")

  regex.Pattern = "^[A-Za-z0-9]+(\.[A-Za-z0-9&/]+)*$"
  regex.IgnoreCase = True
  regex.Global = False

  IsValidICD11Code = regex.Test(code)
End Function

' Extrae el valor entre startTag y endTag. Devuelve cadena vacía si no se encuentra startTag.
Private Function ParseJSONValue(ByVal json As String, ByVal startTag As String, ByVal endTag As String) As String
  Dim startPos As Long
  Dim endPos As Long
  Dim raw As String

  startPos = InStr(1, json, startTag, vbTextCompare)
  If startPos = 0 Then Exit Function

  startPos = startPos + Len(startTag)
  If endTag = "" Then
    ' Si no hay endTag, tomamos hasta el final o hasta que el objeto se cierre.
    endPos = Len(json) + 1
  Else
    endPos = InStr(startPos, json, endTag, vbTextCompare)
    If endPos = 0 Then endPos = Len(json) + 1
  End If

  raw = Mid$(json, startPos, endPos - startPos)
  ' Elimina comillas de apertura/cierre si existen
  If Len(raw) >= 2 Then
    If Left$(raw, 1) = """" And Right$(raw, 1) = """" Then
      raw = Mid$(raw, 2, Len(raw) - 2)
    End If
  End If
  ' Reemplaza comillas escapadas
  raw = Replace(raw, "\""", """")
  ParseJSONValue = raw
End Function

' Reemplaza secuencias comunes \u00xx con caracteres acentuados
Private Function DecodeUnicode(ByVal str As String) As String
  If Len(str) = 0 Then
    DecodeUnicode = str
    Exit Function
  End If

  ' Caracteres comunes en español
  str = Replace(str, "\u00E1", "á")
  str = Replace(str, "\u00E9", "é")
  str = Replace(str, "\u00ED", "í")
  str = Replace(str, "\u00F3", "ó")
  str = Replace(str, "\u00FA", "ú")
  str = Replace(str, "\u00F1", "ñ")
  str = Replace(str, "\u00C1", "Á")
  str = Replace(str, "\u00C9", "É")
  str = Replace(str, "\u00CD", "Í")
  str = Replace(str, "\u00D3", "Ó")
  str = Replace(str, "\u00DA", "Ú")
  str = Replace(str, "\u00D1", "Ñ")

  ' Otros caracteres comunes
  str = Replace(str, "\\/", "/")
  str = Replace(str, "\u2013", "-") ' guion corto
  str = Replace(str, "\u2014", "-") ' guion largo

  DecodeUnicode = str
End Function