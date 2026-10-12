VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble dos objetos do PowerPoint para os testes de modPowerPoint (ligacao tardia: so os membros que o codigo le). O mesmo tipo faz o papel
' de aplicativo, janela, modo de exibicao, selecao, apresentacao, slide e forma; membro sem valor (Nothing) faz a leitura falhar como no
' PowerPoint (erro 91). SlideIndex so existe num slide (SlideIndexValue > 0): mestre e layout nao tem (erro 438).
Public Windows As Collection
Public ActiveWindow As Object
Public Presentations As Collection
Public View As Object
Public Slide As Object
Public Selection As Object
Public SelectedSlide As Object
Public Presentation As Object
Public Slides As Collection
Public Chart As Object
Public ChartData As Object
Public IsLinked As Variant
Public SlideIndexValue As Long

Public Property Get SlideIndex() As Long
    If SlideIndexValue <= 0 Then Err.Raise 438, "TestPptFake", "nao e um slide"
    SlideIndex = SlideIndexValue
End Property

Public Function SlideRange(i As Long) As Object
    If SelectedSlide Is Nothing Then Err.Raise 91, "TestPptFake", "nada selecionado"
    Set SlideRange = SelectedSlide
End Function
