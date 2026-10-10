VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Um observador por gráfico ThinQQ da planilha ativa. Handlers de evento nunca propagam erro.
Public WithEvents Cht As Chart

Private Sub Cht_Activate()
    On Error Resume Next
    Watchers_NotifySelection
End Sub

Private Sub Cht_Deactivate()
    On Error Resume Next
    Watchers_NotifySelection
End Sub

Private Sub Cht_Calculate()
    On Error Resume Next
    Core_SafeRefresh Cht
End Sub

Private Sub Cht_Resize()
    On Error Resume Next
    Core_SafeRefresh Cht
End Sub
