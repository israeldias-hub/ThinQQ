Attribute VB_Name = "TestSettings"
Option Explicit

Public Sub Test_Settings_RoundTrip()
    Settings_Set "TestKey", "valor"
    AssertEqual "valor", Settings_Get("TestKey")
    Settings_Delete "TestKey"
    AssertEqual "padrao", Settings_Get("TestKey", "padrao")
End Sub

Public Sub Test_Settings_UpdateSource_Default()
    Dim saved As String
    saved = Settings_Get(KEY_UPDATE_SOURCE)
    Settings_Delete KEY_UPDATE_SOURCE
    AssertEqual DEFAULT_UPDATE_SOURCE, Settings_UpdateSource()
    If Len(saved) > 0 Then Settings_Set KEY_UPDATE_SOURCE, saved
End Sub
