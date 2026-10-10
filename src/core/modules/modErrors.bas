Attribute VB_Name = "modErrors"
Option Explicit

' Nunca propaga erro (roda dentro de handlers Fail:): todo o corpo é protegido, inclusive Open/Print/Close.
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
    Errors_HandleErr context, Err.Number, Err.Description
End Sub

' Igual a Errors_Handle, com o erro já capturado (quando o Err foi limpo por uma limpeza de recursos).
Public Sub Errors_HandleErr(context As String, num As Long, desc As String)
    Errors_Log context & vbTab & num & vbTab & desc
    Core_Notify "Ocorreu um erro no ThinQQ (" & context & "):" & vbLf & desc & vbLf & vbLf & _
                "Detalhes em %LOCALAPPDATA%\ThinQQ\logs\thinqq.log", vbExclamation, False
End Sub