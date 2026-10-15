Attribute VB_Name = "TestUpdater"
Option Explicit

Private mSaved(0 To 14) As String
Private mStdCore As String          ' core de teste padrao (dist), restaurado no Cleanup

Private Function SettingKeys() As Variant
    SettingKeys = Array(KEY_UPDATE_SOURCE, KEY_SKIPPED, KEY_LAST_CHECK, KEY_PENDING_LOADER, KEY_INTERVAL, KEY_LAST_NOTIFIED_CORE, KEY_PENDING_LOADER_SHA, KEY_LAST_FAILURE, KEY_ALLOW_UNVERIFIED, KEY_SWAP_IN_PROGRESS, _
        KEY_PPAM_VERSION, KEY_PPAM_AUTO, KEY_PENDING_PPAM, KEY_PENDING_PPAM_SHA, KEY_PENDING_PPAM_VER)
End Function

' Zera os ganchos do suplemento do PowerPoint (Setup e Cleanup: nenhum teste herda estado de outro).
Private Sub ResetPpamHooks()
    Updater_TestSetPptRunning Empty
    Updater_TestSetPpamAuto Empty
    Updater_CancelPpamRetry
    Updater_PpamRetryScheduleCount = 0
End Sub

' Guarda as configuracoes do usuario e zera os ganchos do updater.
Private Sub Setup()
    Dim k As Variant, i As Long
    i = 0
    For Each k In SettingKeys()
        mSaved(i) = Settings_Get(CStr(k))
        i = i + 1
    Next
    Settings_Delete KEY_LAST_FAILURE
    Settings_Delete KEY_ALLOW_UNVERIFIED
    Settings_Delete KEY_SWAP_IN_PROGRESS
    Settings_Delete KEY_PPAM_AUTO
    Settings_Delete KEY_PENDING_PPAM: Settings_Delete KEY_PENDING_PPAM_SHA: Settings_Delete KEY_PENDING_PPAM_VER
    ResetPpamHooks
    Updater_TestFormResult = ""
    Updater_FormShownCount = 0
    Updater_PendingScheduleCount = 0
    Updater_TestSetBusy False
    Loader_LastMsg = ""
    Loader_SetRootOverride ""
    If Not Loader_CoreWorkbook() Is Nothing Then mStdCore = Loader_CoreWorkbook().FullName
End Sub

' Restaura tudo que o teste pode ter alterado (registro, raiz, ganchos).
Private Sub Cleanup()
    Dim k As Variant, i As Long
    Loader_SetRootOverride ""
    Updater_TestFormResult = ""
    Updater_TestSetBusy False
    ResetPpamHooks
    Loader_LastMsg = ""
    i = 0
    For Each k In SettingKeys()
        If Len(mSaved(i)) > 0 Then Settings_Set CStr(k), mSaved(i) Else Settings_Delete CStr(k)
        i = i + 1
    Next
    ' fecha o core trocado a quente (pasta temporaria) e volta ao core de teste padrao
    On Error Resume Next
    If Len(mStdCore) > 0 Then
        If Loader_CoreWorkbook() Is Nothing Then
            Loader_OpenCore mStdCore
        ElseIf StrComp(Loader_CoreWorkbook().FullName, mStdCore, vbTextCompare) <> 0 Then
            Loader_OpenCore mStdCore
        End If
    End If
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
    AssertTrue InStr(Updater_FriendlyError(vbObjectError + 1005, "duas causas"), "duas causas") > 0
    AssertEqual "Manifesto sem coreVersion válida.", Updater_FriendlyError(vbObjectError + 2000, "Manifesto sem coreVersion válida.")
End Sub

Public Sub Test_Updater_Apply_FromLocalPath_InstallsCore()
    Dim root As String, src As String, sha As String, ok As Boolean, dest As String
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    Settings_Set KEY_SKIPPED, "9.9.9|0.0.1"
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
    Settings_Set KEY_SKIPPED, "9.9.9|0.0.1"
    src = CoreFixture()
    ok = Updater_Apply(LocalManifest(src, "0000000000000000000000000000000000000000000000000000000000000000"), "core", False)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
    AssertTrue InStr(Loader_LastMsg, "Falha ao atualizar") > 0, "mensagem amigável esperada"
    AssertTrue InStr(Loader_LastMsg, "SHA-256") > 0
    AssertEqual "9.9.9|0.0.1", Settings_Get(KEY_SKIPPED), "falha não deve limpar SkippedVersion"
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
    ok = Updater_Apply(LocalManifest(root & "\nao-existe.xlam", String$(64, "a")), "core", False)
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
    AssertEqual sha, Settings_Get(KEY_PENDING_LOADER_SHA), "SHA esperado gravado ao lado do caminho pendente"
    AssertEqual 1, Updater_PendingScheduleCount, "troca do loader agendada (core carregado e novo o bastante)"
    Settings_Delete KEY_PENDING_LOADER
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
    AssertEqual 0, Updater_PendingScheduleCount
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
    Settings_Set KEY_SKIPPED, "9.9.9|0.0.1"
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
    Settings_Set KEY_SKIPPED, "9.9.9|0.0.1"
    Updater_TestFormResult = "later"
    Updater_Check False
    AssertEqual 1, Updater_FormShownCount, "verificação manual sempre mostra a atualização"
    AssertEqual "9.9.9|0.0.1", Settings_Get(KEY_SKIPPED)
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
    AssertEqual "9.9.9|0.0.1", Settings_Get(KEY_SKIPPED)
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
    Settings_Set KEY_SKIPPED, "9.9.9|0.0.1"
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

Public Sub Test_Updater_ShouldCheckNow_RecentFailure_RetriesAfterOneHour()
    Dim t As Date
    t = DateSerial(2026, 3, 1) + TimeSerial(12, 0, 0)
    AssertFalse Updater_ShouldCheckNow("", t, 24, "2026-03-01 11:30:00"), "falha ha 30 min: não insiste"
    AssertTrue Updater_ShouldCheckNow("", t, 24, "2026-03-01 10:30:00"), "falha ha 90 min: tenta de novo"
    AssertFalse Updater_ShouldCheckNow("2026-03-01 08:00:00", t, 24, "2026-03-01 08:00:00"), "sucesso recente e intervalo 24 h"
    AssertTrue Updater_ShouldCheckNow("2026-02-27 08:00:00", t, 24, "lixo"), "falha inválida é ignorada"
End Sub

Public Sub Test_Updater_Check_Failure_RecordsLastFailure_And_Startup_BacksOff()
    On Error GoTo Fail
    Setup
    Settings_Set KEY_UPDATE_SOURCE, TempDir() & "\nao-existe.json"
    Settings_Delete KEY_LAST_CHECK
    Updater_StartupCheck
    AssertTrue Len(Settings_Get(KEY_LAST_FAILURE)) > 0, "falha registrada"
    UseManifestFile Manifest("9.9.9", CoreFixture(), "")
    Updater_StartupCheck
    AssertEqual 0, Updater_FormShownCount, "logo após falha o startup não tenta de novo"
    Settings_Set KEY_LAST_FAILURE, "2020-01-01 10:00:00"
    Updater_StartupCheck
    AssertEqual 1, Updater_FormShownCount, "após 1 h tenta de novo"
    AssertEqual "", Settings_Get(KEY_LAST_FAILURE), "sucesso limpa a falha"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Evaluate_SchemaVersionNewer_IsNone()
    AssertEqual "none", Updater_Evaluate("{""schemaVersion"":2,""coreVersion"":""9.0.0"",""loaderVersion"":""9.0.0""}", "1.0.0", "1.0.0")
    AssertEqual "core", Updater_Evaluate("{""schemaVersion"":1,""coreVersion"":""9.0.0""}", "1.0.0", "1.0.0")
End Sub

Public Sub Test_Updater_VersionLine_And_SkipKey()
    Dim j As String
    j = "{""coreVersion"":""1.2.0"",""loaderVersion"":""1.1.0"",""publishedAt"":""2026-10-10""}"
    AssertTrue InStr(Updater_VersionLine("core", "1.0.0", j), "Nova: 1.2.0") > 0
    AssertTrue InStr(Updater_VersionLine("loader", "1.0.0", j), "carregador") > 0, "loader-only deve dizer que é o carregador"
    AssertTrue InStr(Updater_VersionLine("loader", "1.0.0", j), "1.1.0") > 0
    AssertTrue InStr(Updater_VersionLine("both", "1.0.0", j), "Carregador") > 0
    AssertEqual "1.2.0|1.1.0", Updater_SkipKey(j)
End Sub

' Mesma versao do core em uso: recusa sem tocar no arquivo em execucao.
Public Sub Test_Updater_Apply_SameVersionAsRunningCore_IsRefusedUntouched()
    Dim root As String, vd As String, running As String, bad As String, f As Integer, lenBefore As Long, stampBefore As Date, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    vd = root & "\versions\9.9.9"
    MkDir root & "\versions": MkDir vd
    running = vd & "\" & Loader_CoreFileName()
    FileCopy CoreFixture(), running
    Loader_SetRootOverride root
    AssertTrue Loader_OpenCore(running)
    lenBefore = FileLen(running)
    stampBefore = FileDateTime(running)
    bad = TempDir() & "\corrompido.xlam"
    f = FreeFile: Open bad For Output As #f: Print #f, "corrompido": Close #f
    ok = Updater_Apply(Manifest("9.9.9", bad, Transport_Sha256File(bad)), "core", False)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertTrue InStr(Loader_LastMsg, "já está instalada") > 0, "mensagem amigável: " & Loader_LastMsg
    AssertEqual running, Loader_CoreWorkbook().FullName
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    AssertEqual lenBefore, FileLen(running), "arquivo do core em uso intacto"
    AssertTrue stampBefore = FileDateTime(running), "arquivo do core em uso não foi regravado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_NoSha_IsRefusedByDefault()
    Dim root As String, ok As Boolean, json As String
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    ok = Updater_Apply(Manifest("9.9.9", CoreFixture(), ""), "core", False)
    AssertFalse ok
    AssertTrue InStr(Loader_LastMsg, "não informa o hash") > 0, "mensagem amigável esperada: " & Loader_LastMsg
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
    json = "{""loaderVersion"":""9.9.9"",""loaderUrl"":""" & Replace(RepoPath() & "\tests\fixtures\http\hello.txt", "\", "\\") & """,""loaderSha256"":""""}"
    Settings_Delete KEY_PENDING_LOADER
    ok = Updater_Apply(json, "loader", True)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Settings_Get(KEY_PENDING_LOADER)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_NoSha_InstallsWhenAllowUnverified()
    Dim root As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    Settings_Set KEY_ALLOW_UNVERIFIED, "1"
    ok = Updater_Apply(Manifest("9.9.9", CoreFixture(), ""), "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok, "com AllowUnverified=1 instala sem SHA"
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_Apply_ShaInUpperCase_Matches()
    Dim root As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = TempDir()
    Loader_SetRootOverride root
    ok = Updater_Apply(Manifest("9.9.9", CoreFixture(), UCase$(Transport_Sha256File(CoreFixture()))), "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' SHA confere, mas o arquivo nao e um core valido: nada fica instalado e o core anterior volta a responder.
Public Sub Test_Updater_Apply_CorruptButShaMatches_RollsBack()
    Dim root As String, bad As String, prev As String, ok As Boolean, f As Integer
    On Error GoTo Fail
    Setup
    root = TempDir()
    bad = TempDir() & "\core-corrompido.xlam"
    f = FreeFile: Open bad For Output As #f: Print #f, "isto nao e um xlam": Close #f
    prev = Loader_CoreWorkbook().FullName
    Loader_SetRootOverride root
    Settings_Set KEY_SKIPPED, "x"
    ok = Updater_Apply(Manifest("9.9.9", bad, Transport_Sha256File(bad)), "core", False)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual 0, Len(Dir(root & "\versions\9.9.9", vbDirectory)), "pasta da versão nova apagada"
    AssertTrue Not Loader_CoreWorkbook() Is Nothing, "core anterior recarregado"
    AssertEqual prev, Loader_CoreWorkbook().FullName
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    AssertTrue InStr(Loader_LastMsg, "Falha ao atualizar") > 0
    AssertEqual "x", Settings_Get(KEY_SKIPPED), "falha não mexe nas configurações"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' Core mais novo corrompido: o startup cai para a versao valida anterior.
Public Sub Test_Loader_OpenBestCore_FallsBackWhenNewestIsCorrupt()
    Dim root As String, vd As String, f As Integer, c As Collection
    On Error GoTo Fail
    Setup
    root = TempDir()
    vd = root & "\versions"
    MkDir vd: MkDir vd & "\1.0.0": MkDir vd & "\9.9.9": MkDir vd & "\2.0.0"
    FileCopy CoreFixture(), vd & "\1.0.0\" & Loader_CoreFileName()
    f = FreeFile: Open vd & "\9.9.9\" & Loader_CoreFileName() For Output As #f: Print #f, "corrompido": Close #f
    f = FreeFile: Open vd & "\2.0.0\" & Loader_CoreFileName() For Output As #f: Print #f, "corrompido": Close #f
    Set c = Loader_ListCoreVersions(vd)
    AssertEqual 3, c.Count
    AssertEqual "9.9.9", c(1)
    AssertEqual "2.0.0", c(2)
    AssertEqual "1.0.0", c(3)
    AssertEqual "1.0.0", Loader_OpenBestCore(vd)
    AssertEqual vd & "\1.0.0\" & Loader_CoreFileName(), Loader_CoreWorkbook().FullName
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    Loader_CloseCore
    Kill vd & "\1.0.0\" & Loader_CoreFileName()
    AssertEqual "", Loader_OpenBestCore(vd)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Loader_Shutdown_KeepsCoreOpen_WhileSwapFlagIsFresh()
    On Error GoTo Fail
    Setup
    AssertTrue Not Loader_CoreWorkbook() Is Nothing
    Settings_Set KEY_SWAP_IN_PROGRESS, Format$(Now - 1 / 1440, "yyyy-mm-dd hh:nn:ss")
    AssertTrue Loader_SwapInProgress()
    AssertFalse Loader_ShouldCloseCoreOnShutdown(), "troca em andamento: não fecha o core"
    Loader_Shutdown
    AssertTrue Not Loader_CoreWorkbook() Is Nothing, "core continua aberto"
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Loader_SwapFlag_StaleOrInvalid_IsIgnoredAndCleared()
    On Error GoTo Fail
    Setup
    AssertTrue Loader_ShouldCloseCoreOnShutdown(), "sem flag: fecha normalmente"
    Settings_Set KEY_SWAP_IN_PROGRESS, Format$(Now - 10 / 1440, "yyyy-mm-dd hh:nn:ss")
    AssertTrue Loader_ShouldCloseCoreOnShutdown(), "flag com 10 min é obsoleta"
    AssertEqual "", Settings_Get(KEY_SWAP_IN_PROGRESS), "flag obsoleta é limpa"
    Settings_Set KEY_SWAP_IN_PROGRESS, "lixo"
    AssertTrue Loader_ShouldCloseCoreOnShutdown()
    AssertEqual "", Settings_Get(KEY_SWAP_IN_PROGRESS)
    Settings_Set KEY_SWAP_IN_PROGRESS, Format$(Now + 1, "yyyy-mm-dd hh:nn:ss")
    AssertTrue Loader_ShouldCloseCoreOnShutdown(), "flag no futuro (relógio) não vale"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Loader_OpenCore_SamePath_ReregistersLoader()
    On Error GoTo Fail
    Setup
    Loader_RunCore "Core_RegisterLoader", "Outro.xlam", "C:\outro\Outro.xlam"
    AssertEqual "C:\outro\Outro.xlam", CStr(Loader_RunCore("Core_LoaderPath"))
    AssertTrue Loader_OpenCore(Loader_CoreWorkbook().FullName), "mesmo caminho: sucesso sem reabrir"
    AssertEqual ThisWorkbook.FullName, CStr(Loader_RunCore("Core_LoaderPath")), "core reapontado para este loader"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Public Sub Test_Updater_SchedulePendingLoader_SkippedWhileSwapInProgress()
    Dim d As String, f As Integer
    On Error GoTo Fail
    Setup
    d = TempDir()
    f = FreeFile: Open d & "\p.xlam" For Output As #f: Print #f, "x": Close #f
    Settings_Set KEY_PENDING_LOADER, d & "\p.xlam"
    Settings_Set KEY_SWAP_IN_PROGRESS, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Updater_SchedulePendingLoader 1
    AssertEqual 0, Updater_PendingScheduleCount, "durante a troca não agenda uma segunda aplicação"
    Settings_Delete KEY_SWAP_IN_PROGRESS
    Updater_SchedulePendingLoader 1
    AssertEqual 1, Updater_PendingScheduleCount
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' ---------- Suplemento do PowerPoint (ThinQQ.ppam; spec 2026-10-14 secao 4.9) ----------

' Raiz temporaria com um suplemento "instalado" (conteudo antigo) e a versao instalada 1.0.0.
Private Function PpamRoot(Optional withPpam As Boolean = True) As String
    Dim root As String
    root = TempDir()
    If withPpam Then Transport_WriteTextFile root & "\ThinQQ.ppam", "ppam-antigo"
    Loader_SetRootOverride root
    Settings_Set KEY_PPAM_VERSION, "1.0.0"
    PpamRoot = root
End Function

Private Function PpamJson(url As String, sha As String, ver As String) As String
    PpamJson = "{""schemaVersion"":1,""coreVersion"":""0.0.1"",""loaderVersion"":""0.0.1"",""pptAddinVersion"":""" & ver & """,""pptAddinUrl"":""" & Replace(url, "\", "\\") & """,""pptAddinSha256"":""" & sha & """}"
End Function

Private Function FixturePpam() As String
    FixturePpam = RepoPath() & "\tests\fixtures\release\ThinQQ.ppam"
End Function

' (5) Nunca instala o suplemento para quem nao o tem: sem <raiz>\ThinQQ.ppam nada e baixado nem criado.
Public Sub Test_Updater_Ppam_NotInstalled_NeverInstalls()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot(False)
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    AssertFalse Updater_PpamInstalled(), "sem o arquivo: nao instalado"
    AssertFalse Updater_PpamNeeded(json)
    AssertEqual "", Updater_PpamUpdateNow(json), "nada a fazer"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam")), "nada criado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' Build de Teste sem raiz de teste: a raiz real (%LOCALAPPDATA%\ThinQQ) nunca conta como instalada nem recebe pendencia aplicada.
Public Sub Test_Updater_Ppam_TestBuild_NeverTouchesRealRoot()
    On Error GoTo Fail
    Setup
    Loader_SetRootOverride ""
    AssertFalse Updater_PpamInstalled(), "sem raiz de teste: nunca instalado"
    Settings_Set KEY_PENDING_PPAM, "1": Settings_Set KEY_PENDING_PPAM_SHA, String$(64, "a"): Settings_Set KEY_PENDING_PPAM_VER, "9.9.9"
    Updater_ApplyPendingPpam
    AssertEqual "1", Settings_Get(KEY_PENDING_PPAM), "a pendencia da instalacao real nao e tocada pelos testes"
    AssertEqual 0, Updater_PpamRetryScheduleCount
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (1)(2)(7) Instalado e mais antigo, PowerPoint fechado: baixa pelo servidor HTTP de teste, confere o SHA-256 e troca atomicamente; grava a
' versao; nao sobra .old, .new nem temporario.
Public Sub Test_Updater_Ppam_ReplacesWhenUnlocked_FromHttp()
    Dim root As String, json As String, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    AssertEqual "9.9.9", JsonGetString(json, "pptAddinVersion")
    AssertTrue Updater_PpamNeeded(json)
    st = Updater_PpamUpdateNow(json)
    AssertEqual "replaced", st
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual "9.9.9", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.old")) + Len(Dir(root & "\ThinQQ.ppam.new")) + Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam"))
    AssertFalse Updater_PpamNeeded(json), "ja atualizado"
    AssertEqual 0, Updater_PpamRetryScheduleCount, "trocado: nada agendado"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (2) Arquivo em uso (aberto com trava exclusiva, como o PowerPoint faz com um arquivo aberto para edicao: medido): fica .new pendente,
' agenda nova tentativa (uma so: nunca empilha); liberado o arquivo, a aplicacao pendente reconfere o SHA-256 e troca.
Public Sub Test_Updater_Ppam_Locked_StagesThenAppliesPending()
    Dim root As String, json As String, f As Integer, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning True
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    f = FreeFile
    Open root & "\ThinQQ.ppam" For Binary Access Read Lock Read Write As #f
    st = Updater_PpamUpdateNow(json)
    AssertEqual "staged", st
    AssertTrue Len(Dir(root & "\ThinQQ.ppam.new")) > 0
    AssertEqual "1", Settings_Get(KEY_PENDING_PPAM)
    AssertEqual "9.9.9", Settings_Get(KEY_PENDING_PPAM_VER)
    AssertEqual 1, Updater_PpamRetryScheduleCount, "nova tentativa agendada"
    AssertTrue InStr(Updater_PpamStatusText(st), "quando o PowerPoint for fechado") > 0
    AssertFalse Updater_PpamNeeded(json), "a mesma versao ja pendente nao e baixada de novo"
    Updater_ApplyPendingPpam
    AssertEqual "1", Settings_Get(KEY_PENDING_PPAM), "ainda em uso: continua pendente"
    AssertEqual 1, Updater_PpamRetryScheduleCount, "tentativa ja agendada: nao empilha outra"
    Updater_PpamRetryTick
    AssertEqual 2, Updater_PpamRetryScheduleCount, "a tentativa disparada e ainda em uso reagenda uma"
    Close #f
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "o arquivo em uso nao foi tocado"
    Updater_ApplyPendingPpam
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM): AssertEqual "9.9.9", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new")) + Len(Dir(root & "\ThinQQ.ppam.old"))
    Updater_PpamRetryTick
    AssertEqual 2, Updater_PpamRetryScheduleCount, "sem pendencia: nao reagenda"
    Cleanup
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    FailCleanup
End Sub

' (7) SHA-256 divergente: nada trocado, temporario apagado. .new adulterado depois de pendente: recusado e apagado.
Public Sub Test_Updater_Ppam_HashMismatch_AndTamperedPending()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = PpamJson(FixturePpam(), String$(64, "0"), "9.9.9")
    AssertEqual "", Updater_PpamUpdateNow(json), "SHA errado: nada"
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual 0, Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam")), "temporario apagado"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new")), "nada pendente"
    Transport_WriteTextFile root & "\ThinQQ.ppam.new", "adulterado"
    Settings_Set KEY_PENDING_PPAM, "1": Settings_Set KEY_PENDING_PPAM_SHA, Transport_Sha256File(FixturePpam()): Settings_Set KEY_PENDING_PPAM_VER, "9.9.9"
    Updater_ApplyPendingPpam
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "pendente adulterado nao entra"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new")), "e apagado"
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' Pendencia que nao e mais nova que a instalada (ex.: o usuario reinstalou pelo ZIP): descartada, nunca rebaixa o suplemento.
Public Sub Test_Updater_Ppam_PendingNotNewer_IsDiscarded()
    Dim root As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Settings_Set KEY_PPAM_VERSION, "9.9.9"
    FileCopy FixturePpam(), root & "\ThinQQ.ppam.new"
    Settings_Set KEY_PENDING_PPAM, "1": Settings_Set KEY_PENDING_PPAM_SHA, Transport_Sha256File(FixturePpam()): Settings_Set KEY_PENDING_PPAM_VER, "9.9.9"
    Updater_ApplyPendingPpam
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "nao rebaixa"
    AssertEqual 0, Len(Dir(root & "\ThinQQ.ppam.new"))
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (4) Opcao desligada: nada; (2) PowerPoint aberto mas arquivo livre (suplemento carregado sem trava): troca e pede reinicio.
Public Sub Test_Updater_Ppam_AutoOff_And_RestartNotice()
    Dim root As String, json As String, st As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    Updater_TestSetPpamAuto "0"
    AssertFalse Updater_PpamNeeded(json), "opcao desligada"
    Updater_TestSetPpamAuto Empty
    Settings_Set KEY_PPAM_AUTO, "0"
    AssertFalse Updater_PpamNeeded(json), "opcao desligada no registro"
    Settings_Delete KEY_PPAM_AUTO
    AssertTrue Updater_PpamNeeded(json), "padrao: ligada"
    Updater_TestSetPptRunning True
    st = Updater_PpamUpdateNow(json)
    AssertEqual "replaced-restart", st
    AssertTrue InStr(Updater_PpamStatusText(st), "feche e abra o PowerPoint") > 0
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (3) Ordem: com core + suplemento, o .ppam e baixado e conferido ANTES da troca do core; .ppam indisponivel nao derruba o core, mas o
' resultado diz que o suplemento fica para a proxima verificacao (nunca em silencio).
Public Sub Test_Updater_Apply_CoreOk_PpamFails_IsReported()
    Dim root As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    json = Replace(json, "/release/ThinQQ.ppam", "/release/nao-existe.ppam")
    ok = Updater_Apply(json, "core", True)
    AssertTrue ok, "o core e atualizado"
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "suplemento intocado"
    AssertEqual "1.0.0", Settings_Get(KEY_PPAM_VERSION)
    AssertTrue InStr(Updater_LastPpamNote(), "próxima verificação") > 0, Updater_LastPpamNote()
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (3) Core + suplemento pelo servidor HTTP: os dois trocados e o aviso de sucesso cita o suplemento.
Public Sub Test_Updater_Apply_CoreAndPpam_FromHttp()
    Dim root As String, json As String, ok As Boolean
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    ok = Updater_Apply(json, "core", False)
    AssertTrue ok
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    AssertEqual "9.9.9", Settings_Get(KEY_PPAM_VERSION)
    AssertTrue InStr(Loader_LastMsg, "atualizado com sucesso") > 0, Loader_LastMsg
    AssertTrue InStr(Loader_LastMsg, "Suplemento do PowerPoint atualizado.") > 0, Loader_LastMsg
    AssertEqual 0, Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam"))
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (3) Falha do core: o .ppam ja baixado e apagado e o suplemento nao e trocado (nunca "suplemento novo + core velho").
Public Sub Test_Updater_Apply_CoreFails_PpamUntouched()
    Dim root As String, json As String, ok As Boolean, bad As String, f As Integer
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    bad = TempDir() & "\core-corrompido.xlam"
    f = FreeFile: Open bad For Output As #f: Print #f, "isto nao e um xlam": Close #f
    json = Manifest("9.9.9", bad, Transport_Sha256File(bad))
    json = Left$(json, Len(json) - 1) & ",""pptAddinVersion"":""9.9.9"",""pptAddinUrl"":""" & Replace(FixturePpam(), "\", "\\") & """,""pptAddinSha256"":""" & Transport_Sha256File(FixturePpam()) & """}"
    AssertTrue Updater_PpamNeeded(json)
    ok = Updater_Apply(json, "core", True)
    AssertFalse ok
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "suplemento intocado"
    AssertEqual "1.0.0", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam")) + Len(Dir(root & "\ThinQQ.ppam.new")), "baixado apagado"
    AssertEqual "", Settings_Get(KEY_PENDING_PPAM)
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' (1)(4) So o suplemento atrasado (core em dia): a verificacao silenciosa o atualiza sem abrir o formulario; a linha de versao do formulario
' cita o suplemento quando ele entra junto.
Public Sub Test_Updater_Check_PpamOnly_Silent_And_VersionLine()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9")
    json = Replace(json, """coreVersion"":""0.0.1""", """coreVersion"":""" & Loader_CurrentCoreVersion() & """")
    UseManifestFile json
    Updater_FormShownCount = 0
    Updater_Check True
    AssertEqual 0, Updater_FormShownCount, "so o suplemento: sem formulario"
    AssertEqual "", Loader_LastMsg, "verificacao silenciosa nao fala"
    AssertEqual "fake-ppam-9.9.9", Transport_ReadTextFile(root & "\ThinQQ.ppam")
    Settings_Set KEY_PPAM_VERSION, "1.0.0"
    AssertTrue InStr(Updater_VersionLine("core", "1.0.0", Replace(json, """coreVersion"":""" & Loader_CurrentCoreVersion() & """", """coreVersion"":""9.9.9""")), "Suplemento do PowerPoint: 1.0.0 > 9.9.9") > 0
    Updater_Check False
    AssertEqual 0, Updater_FormShownCount
    AssertTrue InStr(Loader_LastMsg, "Suplemento do PowerPoint atualizado.") > 0, "verificacao manual mostra o resultado: " & Loader_LastMsg
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

Private Function LogText(root As String) As String
    If Len(Dir(root & "\logs\thinqq.log")) > 0 Then LogText = Transport_ReadTextFile(root & "\logs\thinqq.log")
End Function

' Manifesto de esquema mais novo (ignorado pelo avaliador): o suplemento tambem e ignorado, em todos os caminhos (Check, Apply, VersionLine).
Public Sub Test_Updater_Ppam_SchemaVersionNewer_IsIgnored()
    Dim root As String, json As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    Updater_TestSetPptRunning False
    json = Replace(PpamJson(FixturePpam(), Transport_Sha256File(FixturePpam()), "9.9.9"), """schemaVersion"":1", """schemaVersion"":2")
    AssertFalse Updater_PpamNeeded(json), "schemaVersion 2: nada"
    AssertEqual "", Updater_PpamUpdateNow(json)
    UseManifestFile json
    Updater_Check False
    AssertEqual "ppam-antigo", Transport_ReadTextFile(root & "\ThinQQ.ppam"), "nada trocado"
    AssertEqual "1.0.0", Settings_Get(KEY_PPAM_VERSION)
    AssertEqual 0, Len(Dir(Loader_TmpDir() & "\ThinQQ.ppam")) + Len(Dir(root & "\ThinQQ.ppam.new")), "nada baixado"
    AssertTrue InStr(Loader_LastMsg, "já está na versão mais recente") > 0, Loader_LastMsg
    AssertTrue InStr(LogText(root), "schemaVersion") > 0, "registrado no log"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

' Manifesto com pptAddinUrl mas hash ausente, curto ou malformado: nada, e o motivo vai para o log.
Public Sub Test_Updater_Ppam_BadHash_IsLogged()
    Dim root As String
    On Error GoTo Fail
    Setup
    root = PpamRoot()
    AssertFalse Updater_PpamNeeded(PpamJson(FixturePpam(), "", "9.9.9")), "sem hash"
    AssertFalse Updater_PpamNeeded(PpamJson(FixturePpam(), "abc", "9.9.9")), "hash curto"
    AssertFalse Updater_PpamNeeded(PpamJson(FixturePpam(), String$(64, "z"), "9.9.9")), "hash nao hexadecimal"
    AssertTrue InStr(LogText(root), "pptAddinSha256") > 0, "motivo registrado no log"
    AssertTrue Updater_PpamNeeded(PpamJson(FixturePpam(), UCase$(Transport_Sha256File(FixturePpam())), "9.9.9")), "hash valido (maiusculas)"
    Cleanup
    Exit Sub
Fail:
    FailCleanup
End Sub

