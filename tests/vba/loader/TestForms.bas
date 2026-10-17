Attribute VB_Name = "TestForms"
Option Explicit

' Todo UserForm do loader (src\loader\forms) tem o "caption" do JSON como titulo da janela (e nao o "UserForm1" padrao do Excel).
Public Sub Test_Forms_Loader_TituloDaJanela()
    AssertEqual 1, CheckFormCaptions("loader"), "formularios do loader conferidos"
End Sub
