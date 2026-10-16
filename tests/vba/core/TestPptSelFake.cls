VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptSelFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble da selecao de uma janela do PowerPoint para o texto vinculado pelo botao do Excel (PptText_SelectionParts; o tipo da selecao vai a
' parte, porque "Type" e palavra reservada do VBA): ShapeRange (Collection: Count e (1)) e, no texto, TextRange.Start/Length (a mesma
' instancia).
Public ShapeRange As Collection
Public HasChildShapeRange As Boolean            ' selecao dentro de um grupo
Public ChildShapeRange As Collection
Public Start As Long
Public Length As Long

Public Property Get TextRange() As Object
    Set TextRange = Me
End Property
