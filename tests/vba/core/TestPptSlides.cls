VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptSlides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble da colecao Slides do PowerPoint para os testes de modPowerPoint: Count, Item (membro padrao: Slides(i)) e Add(indice, layout), que
' cria um TestPptFake com SlideIndex e conta as insercoes.
Private mItems As New Collection
Public AddedCount As Long

Public Property Get Item(i As Long) As Object
Attribute Item.VB_UserMemId = 0
    Set Item = mItems(i)
End Property

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Sub Append(sl As Object)
    mItems.Add sl
End Sub

Public Function Add(index As Long, layout As Long) As Object
    Dim sl As TestPptFake
    Set sl = New TestPptFake
    sl.SlideIndexValue = index
    If index <= mItems.Count Then mItems.Add sl, Before:=index Else mItems.Add sl
    AddedCount = AddedCount + 1
    Set Add = sl
End Function
