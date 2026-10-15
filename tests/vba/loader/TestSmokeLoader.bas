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
    AssertFalse frm.Controls("chkPpam").Visible, "sem o suplemento do PowerPoint: caixa oculta"
    frm.Prepare "Título", "1.0.0 -> 1.1.0", "notas", True, True
    AssertTrue frm.Controls("chkPpam").Visible, "com o suplemento instalado: caixa visivel"
    AssertEqual (Settings_Get(KEY_PPAM_AUTO, "1") <> "0"), CBool(frm.Controls("chkPpam").Value), "caixa reflete a opcao gravada"
    Unload frm
End Sub
