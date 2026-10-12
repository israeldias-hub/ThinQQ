Option Explicit

Private Sub Workbook_Open()
    Core_OnLoaded
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Core_OnUnload
End Sub
