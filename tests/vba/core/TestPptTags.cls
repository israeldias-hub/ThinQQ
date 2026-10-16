VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble das etiquetas (Tags) de uma apresentacao ou forma do PowerPoint: Item de etiqueta ausente devolve "" (como no PowerPoint), Add
' substitui o valor e conta as gravacoes (Writes), Delete apaga e conta (Deletes). Prova que o ciclo automatico nao regrava o indice a toa.
Private mD As Object
Public Writes As Long
Public Deletes As Long

Private Sub Class_Initialize()
    Set mD = CreateObject("Scripting.Dictionary")
    mD.CompareMode = 1
End Sub

Public Function Item(key As String) As String
    If mD.Exists(key) Then Item = CStr(mD(key))
End Function

Public Sub Add(key As String, value As String)
    mD(key) = value
    Writes = Writes + 1
End Sub

Public Sub Delete(key As String)
    If mD.Exists(key) Then mD.Remove key
    Deletes = Deletes + 1
End Sub

' Valor posto pelo teste (nao conta como gravacao do codigo testado).
Public Sub Preset(key As String, value As String)
    mD(key) = value
End Sub
