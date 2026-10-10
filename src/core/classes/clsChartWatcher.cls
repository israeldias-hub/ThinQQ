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
' Um observador por gráfico ThinQQ da planilha ativa.
Public WithEvents Cht As Chart

Private Sub Cht_Activate()
    Watchers_NotifySelection
End Sub

Private Sub Cht_Deactivate()
    Watchers_NotifySelection
End Sub

Private Sub Cht_Calculate()
    Core_SafeRefresh Cht
End Sub

Private Sub Cht_Resize()
    Core_SafeRefresh Cht
End Sub
