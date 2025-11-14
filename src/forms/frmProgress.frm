VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "ICD-11 Mortality Morbidity Report for SISPRO"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TotalSteps As Long
Public CurrentStep As Long

Private Sub UserForm_Initialize()
  lblProgress.Width = 0
  lblPercent.Caption = "0%"
  
  txtLog.Value = ""
  With txtLog
    .SelStart = 0
    .SetFocus
  End With
  
  TotalSteps = 0
  CurrentStep = 0
End Sub

' Updates the progress bar
Public Sub UpdateProgress(ByVal current As Long, ByVal total As Long)
  Dim pct As Double
  If total <= 0 Then
    pct = 0
  Else
    pct = current / total
    If pct > 1 Then pct = 1 ' <-- Cap at 100%
  End If
  
  lblProgress.Width = fraProgress.Width * pct
  lblPercent.Caption = Format(pct, "0%")
  DoEvents
End Sub

' Adds a message to the log textbox with optional coloring
Public Sub AddLog(timeStamp As String, prefix As String, msg As String, level As LogLevel)
  Dim formattedMsg As String
  formattedMsg = timeStamp & " " & prefix & " " & msg & vbCrLf
  
  With txtLog
    .Text = .Text & formattedMsg
    .SetFocus
    .SelStart = Len(.Text)
    .SelLength = 0
  End With
  
  DoEvents
End Sub
