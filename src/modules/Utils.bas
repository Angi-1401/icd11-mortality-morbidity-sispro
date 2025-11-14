Attribute VB_Name = "Utils"
Option Explicit

Public frm As FrmProgress

Public Enum LogLevel
  LOG_INFO = 0
  LOG_WARNING = 1
  LOG_ERROR = 2
  LOG_DEBUG = 3
End Enum

' Logs messages with optional severity
Public Sub LogMessage(msg As String, Optional level As LogLevel = LOG_INFO)
  Dim prefix As String
  Dim timeStamp As String
  
  timeStamp = Format(Now, "hh:nn:ss")
  
  Select Case level
    Case LOG_INFO: prefix = "[INFO] "
    Case LOG_WARNING: prefix = "[WARN] "
    Case LOG_ERROR: prefix = "[ERROR]"
    Case LOG_DEBUG: prefix = "[DEBUG]"
    Case Else: prefix = "[INFO] "
  End Select
  
  Debug.Print timeStamp & " " & prefix & msg
  
  ' Add to progress form log if frm is assigned
  If Not frm Is Nothing Then
    frm.AddLog timeStamp, prefix, msg, level
  End If
End Sub

' Temporarily disables application settings for performance
Public Sub DisableApplicationSettings(ByVal disable As Boolean)
  If disable Then
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
  Else
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
  End If
End Sub
