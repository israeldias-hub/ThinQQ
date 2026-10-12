Option Explicit

Private Sub Workbook_Open()
    If BUILD_CONFIG = "Test" Then Exit Sub
    Loader_Startup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If BUILD_CONFIG = "Test" Then Exit Sub
    Loader_Shutdown
End Sub
