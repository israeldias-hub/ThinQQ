Attribute VB_Name = "TestForms"
Option Explicit

' Todo UserForm do core (src\core\forms) tem o "caption" do JSON como titulo da janela (e nao o "UserForm1" padrao do Excel).
Public Sub Test_Forms_Core_TituloDaJanela()
    AssertEqual 1, CheckFormCaptions("core"), "formularios do core conferidos"
End Sub
