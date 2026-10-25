Attribute VB_Name = "TestPptForms"
Option Explicit

' Todo UserForm do suplemento do PowerPoint (src\ppt\forms; 1.11.0: frmTqqLinks) tem o "caption" do JSON como titulo da janela (e nao o
' "UserForm1" padrao). Mesmo conferidor do Excel (tests\vba\common\modFormCheck.bas, importado no projeto de teste do .ppam).
Public Sub Test_Forms_Ppt_TituloDaJanela()
    AssertEqual 1, CheckFormCaptions("ppt"), "formularios do suplemento do PowerPoint conferidos"
End Sub
