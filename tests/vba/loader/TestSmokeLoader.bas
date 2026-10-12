Attribute VB_Name = "TestSmokeLoader"
Option Explicit

Public Sub Test_Loader_Ping()
    AssertEqual "pong", Loader_Ping()
End Sub

' Prova que o pipeline de UserForm por JSON funciona (o formulário é consumido na Tarefa 13).
Public Sub Test_Loader_FormUpdate_Exists()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmUpdate")
    AssertTrue Not frm Is Nothing, "frmUpdate não foi criado pelo build"
    frm.Prepare "Título", "1.0.0 -> 1.1.0", "notas"
    AssertEqual "Título", frm.Controls("lblTitle").Caption
    Unload frm
End Sub
