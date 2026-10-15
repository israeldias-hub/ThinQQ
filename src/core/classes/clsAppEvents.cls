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
    DeltaPick_OnActivate Sh.Parent.Name, Sh.Name     ' delta por cliques armado noutra planilha: cancela (os observadores mudam de planilha)
    Watchers_Rehook
    Watchers_NotifySelection        ' atualiza o estado da faixa ao trocar de planilha/janela
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    DeltaPick_OnActivate Wb.Name, ""
    Watchers_Rehook
    Watchers_NotifySelection
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    Watchers_Rehook
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    DeltaPick_OnSheetClick                  ' clique numa celula cancela o delta por cliques (sai na 1a linha quando desarmado)
    Watchers_NotifySelection
End Sub

' A pasta do grafico com o delta por cliques armado vai fechar: desarma (moldura, Esc e OnTime) antes.
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    DeltaPick_OnBookClose Wb.Name
End Sub
