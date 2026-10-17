VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptDataFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble de slide/apresentacao para Core_PptInsertChart sem PowerPoint (Parte G): o slide tem Parent (a apresentacao); a apresentacao tem
' Path, FullName, Name e Tags. So para caminhos que falham ANTES da colagem (nenhum outro membro do PowerPoint e lido).
Public Parent As Object
Public Path As String
Public FullName As String
Public Name As String
Public Tags As Object