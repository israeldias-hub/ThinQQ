VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Handlers de evento nunca propagam erro.
Public WithEvents App As Application

Private Sub App_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Watchers_Rehook
    Watchers_NotifySelection        ' atualiza o estado da faixa ao trocar de planilha/janela
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    Watchers_Rehook
    Watchers_NotifySelection
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    Watchers_Rehook
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Watchers_NotifySelection
End Sub
