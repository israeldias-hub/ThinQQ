Attribute VB_Name = "modCoreApi"
Option Explicit

Public Function Core_Ping() As String
    Core_Ping = "pong"
End Function

Public Function Core_Version() As String
    Core_Version = APP_VERSION
End Function
