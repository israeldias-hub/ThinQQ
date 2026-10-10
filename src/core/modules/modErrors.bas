Attribute VB_Name = "modErrors"
Option Explicit

Public Sub Errors_Log(msg As String)
    On Error Resume Next
    Dim f As Integer, dir_ As String
    dir_ = Environ$("LOCALAPPDATA") & "\ThinQQ\logs"
    Util_EnsureDir dir_
    f = FreeFile
    Open dir_ & "\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "core" & vbTab & msg
    Close #f
End Sub

' Usar em rótulos Fail: de funções públicas. Loga e avisa o usuário (sem caixa em build de teste).
Public Sub Errors_Handle(context As String)
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    Errors_Log context & vbTab & num & vbTab & desc
    Core_Notify "Ocorreu um erro no ThinQQ (" & context & "):" & vbLf & desc & vbLf & vbLf & _
                "Detalhes em %LOCALAPPDATA%\ThinQQ\logs\thinqq.log", vbExclamation
End Sub
