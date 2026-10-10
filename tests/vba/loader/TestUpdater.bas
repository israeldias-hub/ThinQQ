Attribute VB_Name = "TestUpdater"
Option Explicit

Private mSaved(0 To 5) As String

Private Function SettingKeys() As Variant
    SettingKeys = Array(KEY_UPDATE_SOURCE, KEY_SKIPPED, KEY_LAST_CHECK, KEY_PENDING_LOADER, KEY_INTERVAL, KEY_LAST_NOTIFIED_CORE)
End Function

' Guarda as configuracoes do usuario e zera os ganchos do updater.
Private Sub Setup()
    Dim k As Variant, i As Long
    i = 0
    For Each k In SettingKeys()
        mSaved(i) = Settings_Get(CStr(k))
        i = i + 1
    Next
    Updater_TestFormResult = ""
    Updater_FormShownCount = 0
    Updater_TestSetBusy False
    Loader_LastMsg = ""
    Loader_SetRootOverride ""
End Sub

' Restaura tudo que o teste pode ter alterado (registro, raiz, ganchos).
Private Sub Cleanup()
    Dim k As Variant, i As Long
    Loader_SetRootOverride ""
    Updater_TestFormResult = ""
    Updater_TestSetBusy False
    Loader_LastMsg = ""
    i = 0
    For Each k In SettingKeys()
        If Len(mSaved(i)) > 0 Then Settings_Set CStr(k), mSaved(i) Else Settings_Delete CStr(k)
        i = i + 1
    Next
End Sub

' Chamado no tratador de erro dos testes: restaura e repropaga a falha original.
Private Sub FailCleanup()
    Dim n As Long, s As String, d As String
    n = Err.Number: s = Err.Source: d = Err.Description
    Cleanup
    Err.Raise n, s, d
End Sub

Private Function HttpBase() As String
    Dim p As String
    p = RepoPath() & "\tests\results\http-base.txt"
    If Len(Dir(p)) = 0 Then Err.Raise vbObjectError + 900, "TestUpdater", "servidor HTTP de fixtures ausente (Node não encontrado?)"
    HttpBase = Transport_ReadTextFile(p)
End Function

Private Function CoreFixture() As String
    CoreFixture = RepoPath() & "\tests\fixtures\release\ThinQQ.core.test.xlam"
End Function

Private Function Manifest(coreVer As String, coreUrl As String, sha As String, Optional loaderVer As String = "0.0.1") As String
    Manifest = "{""schemaVersion"":1,""coreVersion"":""" & coreVer & """,""coreUrl"":""" & Replace(coreUrl, "\", "\\") & _
        """,""coreSha256"":""" & sha & """,""loaderVersion"":""" & loaderVer & """,""loaderUrl"":"""",""loaderSha256"":"""",""tutorialUrl"":"""",""publishedAt"":""2099-01-01"",""notes"":""x""}"
End Function

Private Function LocalManifest(coreUrl As String, sha As String) As String
    LocalManifest = Manifest("9.9.9", coreUrl, sha)
End Function

' Grava um manifesto em arquivo temporario e o aponta como fonte de atualizacao.
Private Sub UseManifestFile(json As String)
    Dim p As String
    p = TempDir() & "\manifest.json"
    Transport_WriteTextFile p, json
    Settings_Set KEY_UPDATE_SOURCE, p
End Sub

Public Sub Test_Updater_ShouldCheckNow()
    AssertTrue Updater_ShouldCheckNow("", Now, 24)
    AssertTrue Updater_ShouldCheckNow("lixo", Now, 24)
    AssertTrue Updater_ShouldCheckNow("2026-01-01 10:00:00", DateSerial(2026, 1, 2) + TimeSerial(10, 0, 1), 24)
    AssertFalse Updater_ShouldCheckNow("2026-01-01 10:00:00", DateSerial(2026, 1, 1) + TimeSerial(20, 0, 0), 24)
    AssertTrue Updater_ShouldCheckNow("2026-01-01 10:00:00", DateSerial(2026, 1, 1) + TimeSerial(10, 0, 0), 0), "intervalo 0 sempre verifica"
    AssertTrue Updater_ShouldCheckNow("2030-01-01 10:00:00", DateSerial(2026, 1, 1), 24), "relogio no passado: verifica"
End Sub

Public Sub Test_Updater_Evaluate()
    Dim j As String
    j = "{""coreVersion"":""1.2.0"",""loaderVersion"":""1.0.0""}"
    AssertEqual "core", Updater_Evaluate(j, "1.1.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate(j, "1.2.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate(j, "1.3.0", "1.0.0")
    AssertEqual "both", Updater_Evaluate(j, "1.1.0", "0.9.0")
    AssertEqual "loader", Updater_Evaluate(j, "1.2.0", "0.9.0")
End Sub

Public Sub Test_Updater_Evaluate_MalformedJson_IsNone()
    AssertEqual "none", Updater_Evaluate("<html>erro</html>", "1.0.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate("", "1.0.0", "1.0.0")
End Sub

Public Sub Test_Updater_Evaluate_MissingOrInvalidVersion_IsNone()
    AssertEqual "none", Updater_Evaluate("{""loaderUrl"":""x""}", "1.0.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate("{""coreVersion"":""abc"",""loaderVersion"":""1.x""}", "1.0.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate("{""coreVersion"":"""",""loaderVersion"":""""}", "0.0.0", "0.0.0")
End Sub

Public Sub Test_Updater_FriendlyError()
    Dim m As String
    m = Updater_FriendlyError(-2147012889, "The server name or address could not be resolved")
    AssertTrue InStr(m, "Sem conexão") > 0, "erro cru de rede deve virar mensagem amigável"
    AssertTrue InStr(m, "resolved") = 0, "a descrição crua não deve vazar para o usuário"
    AssertTrue InStr(Updater_FriendlyError(vbObjectError + 1004, "x"), "demorou") > 0
    AssertTrue InStr(Updater_FriendlyError(vbObjectError + 1001, "HTTP 404"), "HTTP 404") > 0
    AssertEqual "Manifesto sem coreVersion válida.", Updater_FriendlyError(vbObjectError + 2000, "Manifesto sem coreVersion válida.")
End Sub

Public Sub Test_Updater_Apply_FromLocalPath_InstallsCore()
    Dim root As String, src As String, sha As String, ok As Boolean, dest As String
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    Settings_Set KEY_SKIPPED, "9.9.9"
    Settings_Delete KEY_LAST_NOTIFIED_CORE
    src = CoreFixture()
    sha = Transport_Sha256File(src)
    ok = Updater_Apply(LocalManifest(src, sha), "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok, "Updater_Apply deveria retornar True"
    dest = root & "\versions\9.9.9\" & Loader_CoreFileName()
    AssertTrue Len(Dir(dest)) > 0, "core não foi copiado para versions\9.9.9"
    AssertEqual "9.9.9", Loader_FindNewestCoreVersion(root & "\versions")
    AssertEqual dest, Loader_CoreWorkbook().FullName
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    AssertEqual "9.9.9", Settings_Get(KEY_LAST_NOTIFIED_CORE), "LastNotifiedCore deve ser gravado"
    AssertEqual "", Settings_Get(KEY_SKIPPED), "SkippedVersion deve ser limpo"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_BadSha_DoesNotInstall()
    Dim root As String, src As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    Settings_Set KEY_SKIPPED, "9.9.9"
    src = CoreFixture()
    ok = Updater_Apply(LocalManifest(src, "0000000000000000000000000000000000000000000000000000000000000000"), "core", False)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
    AssertTrue InStr(Loader_LastMsg, "Falha ao atualizar") > 0, "mensagem amigável esperada"
    AssertTrue InStr(Loader_LastMsg, "SHA-256") > 0
    AssertEqual "9.9.9", Settings_Get(KEY_SKIPPED), "falha não deve limpar SkippedVersion"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_InvalidCoreVersion_DoesNotInstall()
    Dim root As String, src As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    src = CoreFixture()
    ok = Updater_Apply(Manifest("abc", src, Transport_Sha256File(src)), "core", True)
    AssertFalse ok
    ok = Updater_Apply(Manifest("", src, ""), "core", True)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_MissingSource_FriendlyAndNothingInstalled()
    Dim root As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    ok = Updater_Apply(LocalManifest(root & "\nao-existe.xlam", ""), "core", False)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
    AssertTrue InStr(Loader_LastMsg, "Falha ao atualizar") > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_FromHttp()
    Dim root As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    AssertEqual "9.9.9", JsonGetString(json, "coreVersion")
    Loader_SetRootOverride root
    ok = Updater_Apply(json, "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_LoaderPart_SchedulesPending()
    Dim root As String, src As String, sha As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    src = RepoPath() & "\tests\fixtures\http\hello.txt"
    sha = Transport_Sha256File(src)
    json = "{""loaderVersion"":""9.9.9"",""loaderUrl"":""" & Replace(src, "\", "\\") & """,""loaderSha256"":""" & sha & """}"
    Loader_SetRootOverride root
    Settings_Delete KEY_PENDING_LOADER
    ok = Updater_Apply(json, "loader", True)
    Loader_SetRootOverride ""
    AssertTrue ok
    AssertEqual root & "\tmp\ThinQQ.xlam", Settings_Get(KEY_PENDING_LOADER)
    Settings_Delete KEY_PENDING_LOADER        ' o Core_ApplyPendingLoader agendado vira no-op
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_LoaderPart_BadSha_NotPending()
    Dim root As String, src As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    src = RepoPath() & "\tests\fixtures\http\hello.txt"
    json = "{""loaderVersion"":""9.9.9"",""loaderUrl"":""" & Replace(src, "\", "\\") & """,""loaderSha256"":""abc""}"
    Loader_SetRootOverride root
    Settings_Delete KEY_PENDING_LOADER
    ok = Updater_Apply(json, "loader", True)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Settings_Get(KEY_PENDING_LOADER)
    AssertEqual 0, Len(Dir(root & "\tmp\ThinQQ.xlam")), "arquivo com SHA errado deve ser apagado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_UnreachablePath_Silent_ShowsNothing()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, TempDir() & "\nao-existe.json"
    Settings_Delete KEY_LAST_CHECK
    Updater_Check True
    AssertEqual "", Loader_LastMsg
    AssertEqual 0, Updater_FormShownCount
    AssertEqual "", Settings_Get(KEY_LAST_CHECK), "falha não deve registrar LastCheck"
    AssertFalse Updater_IsBusy(), "flag de ocupado deve ser liberado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_UnreachablePath_Interactive_Friendly()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, TempDir() & "\nao-existe.json"
    Updater_Check False
    AssertTrue InStr(Loader_LastMsg, "Não foi possível verificar atualizações") > 0, "mensagem amigável esperada: " & Loader_LastMsg
    AssertTrue InStr(Loader_LastMsg, "Fonte:") > 0
    AssertEqual 0, Updater_FormShownCount
    AssertFalse Updater_IsBusy()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_ClosedHttpPort_Silent_ShowsNothing()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, "http://127.0.0.1:1/manifest.json"
    Updater_Check True
    AssertEqual "", Loader_LastMsg
    AssertFalse Updater_IsBusy()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_ClosedHttpPort_Interactive_Friendly()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, "http://127.0.0.1:1/manifest.json"
    Updater_Check False
    AssertTrue InStr(Loader_LastMsg, "Não foi possível verificar atualizações") > 0, "mensagem amigável esperada: " & Loader_LastMsg
    AssertFalse Updater_IsBusy()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_MalformedManifest_NoPromptNoInstall()
    On Error GoTo Fail
    Setup
    UseManifestFile "<html>erro 500</html>"
    Updater_TestFormResult = "update"
    Updater_Check False
    AssertEqual 0, Updater_FormShownCount
    AssertTrue InStr(Loader_LastMsg, "já está na versão mais recente") > 0, "manifesto sem versões = nada a instalar"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_UpToDate_Message()
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("0.0.1", "x", "")
    Updater_Check False
    AssertTrue InStr(Loader_LastMsg, "já está na versão mais recente") > 0
    AssertEqual 0, Updater_FormShownCount
    Loader_LastMsg = ""
    Updater_Check True
    AssertEqual "", Loader_LastMsg, "verificação silenciosa não fala"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_Silent_SkippedVersion_DoesNotPrompt()
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    Settings_Set KEY_SKIPPED, "9.9.9"
    Updater_TestFormResult = "update"
    Updater_Check True
    AssertEqual 0, Updater_FormShownCount, "versão ignorada não deve incomodar na verificação silenciosa"
    AssertEqual "", Loader_LastMsg
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_Interactive_SkippedVersion_StillPrompts()
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    Settings_Set KEY_SKIPPED, "9.9.9"
    Updater_TestFormResult = "later"
    Updater_Check False
    AssertEqual 1, Updater_FormShownCount, "verificação manual sempre mostra a atualização"
    AssertEqual "9.9.9", Settings_Get(KEY_SKIPPED)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_Silent_NewVersion_Prompts_Later_And_Skip()
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    Settings_Delete KEY_SKIPPED
    Updater_Check True
    AssertEqual 1, Updater_FormShownCount
    AssertEqual "", Settings_Get(KEY_SKIPPED), "later não ignora a versão"
    Updater_TestFormResult = "skip"
    Updater_Check True
    AssertEqual 2, Updater_FormShownCount
    AssertEqual "9.9.9", Settings_Get(KEY_SKIPPED)
    Updater_Check True
    AssertEqual 2, Updater_FormShownCount
    AssertTrue Len(Settings_Get(KEY_LAST_CHECK)) > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_Update_InstallsAndClearsSkip()
    Dim root As String, src As String
    On Error GoTo Fail
    Setup
    root = TempDir()
    src = CoreFixture()
    UseManifestFile Manifest("9.9.9", src, Transport_Sha256File(src))
    Settings_Set KEY_SKIPPED, "9.9.9"
    Loader_SetRootOverride root
    Updater_TestFormResult = "update"
    Updater_Check False
    Loader_SetRootOverride ""
    AssertEqual 1, Updater_FormShownCount
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0, "core instalado após 'update'"
    AssertEqual "", Settings_Get(KEY_SKIPPED)
    AssertTrue InStr(Loader_LastMsg, "atualizado com sucesso") > 0
    AssertFalse Updater_IsBusy()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Check_Reentrancy_SecondCallReturnsImmediately()
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    Settings_Delete KEY_LAST_CHECK
    Updater_TestSetBusy True
    Updater_Check False
    AssertEqual 0, Updater_FormShownCount
    AssertEqual "", Loader_LastMsg
    AssertEqual "", Settings_Get(KEY_LAST_CHECK), "chamada reentrante não faz nada"
    AssertTrue Updater_IsBusy(), "a chamada reentrante não deve liberar o flag do dono"
    Updater_TestSetBusy False
    Updater_Check False
    AssertEqual 1, Updater_FormShownCount
    AssertFalse Updater_IsBusy()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_StartupCheck_RespectsInterval()
    Dim stamp As String
    On Error GoTo Fail
    Setup
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Settings_Set KEY_LAST_CHECK, stamp
    Settings_Set KEY_INTERVAL, "24"
    Updater_StartupCheck
    AssertEqual 0, Updater_FormShownCount, "dentro do intervalo: não verifica"
    AssertEqual stamp, Settings_Get(KEY_LAST_CHECK)
    Settings_Set KEY_LAST_CHECK, "2020-01-01 10:00:00"
    Updater_StartupCheck
    AssertEqual 1, Updater_FormShownCount, "intervalo vencido: verifica"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_StartupCheck_Unreachable_IsSilentAndTrapped()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, TempDir() & "\nao-existe.json"
    Settings_Delete KEY_LAST_CHECK
    Updater_StartupCheck
    AssertEqual "", Loader_LastMsg
    AssertFalse Updater_IsBusy()
    Updater_CheckInteractive
    AssertTrue InStr(Loader_LastMsg, "Não foi possível verificar") > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub
