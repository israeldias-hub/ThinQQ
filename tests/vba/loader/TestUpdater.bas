Attribute VB_Name = "TestUpdater"
Option Explicit

Private mSaved(0 To 9) As String
Private mStdCore As String          ' core de teste padrao (dist), restaurado no Cleanup

Private Function SettingKeys() As Variant
    SettingKeys = Array(KEY_UPDATE_SOURCE, KEY_SKIPPED, KEY_LAST_CHECK, KEY_PENDING_LOADER, KEY_INTERVAL, KEY_LAST_NOTIFIED_CORE, KEY_PENDING_LOADER_SHA, KEY_LAST_FAILURE, KEY_ALLOW_UNVERIFIED, KEY_SWAP_IN_PROGRESS)
End Function

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

