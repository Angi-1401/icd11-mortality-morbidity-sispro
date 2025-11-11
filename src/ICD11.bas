Option Explicit

Public Function ExtractICD11Code(ByVal str As String) As String
  Dim dashPos As Long
  Dim ampPos As Long
  Dim segment As String
  Dim char As String
  Dim code As String

  ' Find las dash into the input string and get the segment after it
  dashPos = InStrRev(str, "-")
  If dashPos > 0 Then
    segment = Mid$(str, dashPos + 1)
  Else
    Exit Function
  End If

  ' Remove every character that doesn't belong to ICD-11 code format
  Dim i As Integer
  For i = 1 To Len(segment)
    char = Mid$(segment, i, 1)
    If char Like "[A-Za-z0-9 &/]" Then
      code = code & char
    End If
  Next i

  ' Replace spaces with dots
  code = Replace(code, " ", ".")

  ExtractICD11Code = code
End Function

Public Function GetICD11CodeLabel(ByVal code As String) As String
  Dim http As Object
  Dim token As String
  Dim jsonResponse As String
  Dim url As String
  Dim label As String
  
  On Error Goto ErrHandler
  
  ' Retrieve access token
  token = GetICD11AccessToken()
  If token = "" Then
    GetICD11CodeLabel = "Error: Unable to retrieve access token."
    Exit Function
  End If
  
  ' Validate ICD-11 code format
  If Not IsValidICD11Code(code) Then
    GetICD11CodeLabel = "Error: Invalid ICD-11 code format."
    Exit Function
  End If
  
  ' Encode the ICD-11 code for URL
  code = Replace(code, "&", "%26")
  code = Replace(code, "/", "%2F")
  
  ' Build API request URL
  url = "https://id.who.int/icd/release/11/2025-01/mms/describe?code=" & code
  
  ' Make HTTP GET request to retrieve code details
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  http.Open "GET", url, False
  http.setRequestHeader "Accept", "application/json"
  http.setRequestHeader "API-Version", "v2"
  http.setRequestHeader "Accept-Language", "es"
  http.setRequestHeader "Authorization", "Bearer " & token
  http.Send
  
  If http.Status <> 200 Then
    GetICD11CodeLabel = "Error: API request failed with status " & http.Status
    Exit Function
  End If
  
  jsonResponse = http.responseText
  Debug.Print "JSON Response from GetICD11CodeLabel: " & jsonResponse
  
  ' Parse JSON response to extract the label
  label = ParseJSONValue(jsonResponse, """label"":""", """")
  
  If label <> "" Then
    ' Decode Unicode characters
    label = DecodeUnicode(label)
    GetICD11CodeLabel = label
  Else
    GetICD11CodeLabel = "Error: Label not found for code " & code
  End If

  Exit Function

ErrHandler:
  GetICD11CodeLabel = "Error: " & Err.Description
  Exit Function
End Function

Private Function GetICD11AccessToken() As String
  Static cachedToken As String
  Static tokenExpiry As Date

  Dim http As Object
  Dim postData As String
  Dim jsonResponse As String
  Dim token As String
  Dim expiresIn As Long

  Const TOKEN_URL As String = "https://icdaccessmanagement.who.int/connect/token"
  Const CLIENT_ID As String = "YOUR_CLIENT_ID_HERE"
  Const CLIENT_SECRET As String = "YOUR_CLIENT_SECRET_HERE"

  ' Check if cached token is still valid
  If cachedToken <> "" And Now < tokenExpiry Then
    GetICD11AccessToken = cachedToken
    Exit Function
  End If

  ' Make HTTP POST request to get new access token
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  postData = "grant_type=client_credentials" & _
    "&client_id=" & CLIENT_ID & _
    "&client_secret=" & CLIENT_SECRET & _
    "&scope=icdapi_access"
  
  http.Open "POST", TOKEN_URL, False
  http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  http.Send postData

  If http.Status <> 200 Then
    GetICD11AccessToken = ""
    Exit Function
  End If

  jsonResponse = http.responseText
  Debug.Print "JSON Response from GetICD11AccessToken: " & jsonResponse

  ' Parse JSON response to extract access token and expiry
  token = ParseJSONValue(jsonResponse, """access_token"":""", """")
  expiresIn = CLng(ParseJSONValue(jsonResponse, """expires_in"":", ","))

  If token <> "" Then
    cachedToken = token
    tokenExpiry = DateAdd("s", expiresIn - 60, Now) ' Subtract 60 seconds as buffer
    GetICD11AccessToken = cachedToken
  Else
    GetICD11AccessToken = ""
  End If
End Function

Private Function IsValidICD11Code(ByVal code As String) As Boolean
  Dim regex As Object
  Set regex = CreateObject("VBScript.RegExp")

  ' Define the ICD-11 code pattern
  regex.Pattern = "^[A-Za-z0-9]+(\.[A-Za-z0-9&/]+)*$"
  regex.IgnoreCase = True
  regex.Global = False

  ' Test if the input code matches the pattern
  IsValidICD11Code = regex.Test(code)
End Function

Private Function ParseJSONValue(ByVal json As String, ByVal startTag As String, ByVal endTag As String) As String
  Dim startPos As Long
  Dim endPos As Long

  ' Find the start position of the value
  startPos = InStr(json, startTag)
  If startPos = 0 Then Exit Function

  ' Define final positions
  startPos = startPos + Len(startTag)
  endPos = InStr(startPos, json, endTag)
  
  If endPos = 0 Then endPos = Len(json) + 1
  
  ' Extract the value
  ParseJSONValue = Mid$(json, startPos, endPos - startPos)
End Function

Private Function DecodeUnicode(ByVal str As String) As String
  ' Replace Unicode escape sequences with actual characters
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

  DecodeUnicode = str
End Function