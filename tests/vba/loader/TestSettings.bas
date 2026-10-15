Attribute VB_Name = "TestSettings"
Option Explicit

Public Sub Test_Settings_RoundTrip()
    Settings_Set "TestKey", "valor"
    AssertEqual "valor", Settings_Get("TestKey")
    Settings_Delete "TestKey"
    AssertEqual "padrao", Settings_Get("TestKey", "padrao")
End Sub

' O build Test grava num ramo proprio do registro (ThinQQ.Test): um Excel morto no meio de um teste nunca deixa alteradas as configuracoes
' reais do atualizador do usuario (ThinQQ). O mesmo nome e usado pelo core (troca do loader).
Public Sub Test_Settings_TestBuild_UsesOwnRegistryApp()
    AssertEqual "ThinQQ.Test", Settings_AppName()
    AssertEqual "ThinQQ.Test", CStr(Loader_RunCore("Util_UpdaterRegApp")), "core e loader no mesmo ramo"
    Settings_Set "TestKey", "valor-teste"
    AssertEqual "valor-teste", GetSetting("ThinQQ.Test", "Updater", "TestKey", "")
    Settings_Delete "TestKey"
    AssertEqual "", GetSetting("ThinQQ.Test", "Updater", "TestKey", "")
End Sub

Public Sub Test_Settings_UpdateSource_Default()
    Dim saved As String
    saved = Settings_Get(KEY_UPDATE_SOURCE)
    Settings_Delete KEY_UPDATE_SOURCE
    Dim actual As String
    actual = Settings_UpdateSource()
    If Len(saved) > 0 Then Settings_Set KEY_UPDATE_SOURCE, saved
    AssertEqual DEFAULT_UPDATE_SOURCE, actual
End Sub
