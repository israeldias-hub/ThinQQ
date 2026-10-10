Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    MsgBox "ThinQQ loader " & APP_VERSION, vbInformation, APP_NAME
End Sub
