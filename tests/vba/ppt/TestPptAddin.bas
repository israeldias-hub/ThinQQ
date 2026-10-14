Attribute VB_Name = "TestPptAddin"
Option Explicit
' Funcoes do suplemento do PowerPoint (projeto de teste .pptm aberto numa instancia propria e sem janela).

Private Function NewPres() As Object
    Set NewPres = Application.Presentations.Add(0)
    NewPres.Slides.Add 1, 12
End Function

Private Sub ClosePres(p As Object)
    On Error Resume Next
    p.Saved = -1
    p.Close
End Sub

Public Sub Test_PptTqq_OnOpenAction()
    AssertEqual "ask", PptTqq_OnOpenAction("ask", False, 2, False)
    AssertEqual "sync", PptTqq_OnOpenAction("always", False, 2, False)
    AssertEqual "", PptTqq_OnOpenAction("never", False, 2, False)
    AssertEqual "", PptTqq_OnOpenAction("ask", False, 0, False), "sem graficos marcados"
    AssertEqual "", PptTqq_OnOpenAction("always", True, 2, False), "somente leitura"
    AssertEqual "", PptTqq_OnOpenAction("always", False, 2, True), "ja tratada nesta sessao"
End Sub

Public Sub Test_PptTqq_NormalizeModeAndVersions()
    AssertEqual "always", PptTqq_NormalizeMode(" Always ")
    AssertEqual "ask", PptTqq_NormalizeMode("qualquer")
    AssertEqual "1.10.0", PptTqq_NewestVersion("1.2.0|1.10.0|1.3.1|x|1.2|")
    AssertEqual "", PptTqq_NewestVersion("logs|tmp|")
    AssertEqual "C:\a\b.xlsx", PptTqq_SourcePath("C:\a\b.xlsx!Dados!x", "")
    AssertEqual "C:\t.xlsx", PptTqq_SourcePath("", "C:\t.xlsx")
End Sub

Public Sub Test_PptTqq_ApiCompatible()
    AssertTrue PptTqq_ApiCompatible(1): AssertTrue PptTqq_ApiCompatible(2), "core mais novo mantem a API 1"
    AssertFalse PptTqq_ApiCompatible(0), "core sem Core_PptApiVersion (anterior a 1.4.0): so os dados"
End Sub

Public Sub Test_PptTqq_ErrorTexts()
    AssertEqual PPT_MSG_BUSY, PptTqq_ErrorText(-2147418111, "x")
    AssertEqual PPT_MSG_NO_EXCEL, PptTqq_ErrorText(429, "x")
    AssertTrue InStr(PptTqq_ErrorText(70, "Permissão negada"), "Permissão negada (70)") > 0
    AssertTrue InStr(PptTqq_AskText(3), "3 gráfico(s) do ThinQQ") > 0
End Sub

Private Function NewTagged() As Object
    Dim p As Object, s As Object
    Set p = NewPres()
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "abc123"
    Set NewTagged = p
End Function

' Primeira vista marca a apresentacao mesmo sem graficos (colar depois nao pergunta no proximo WindowActivate); fechar esquece (reabrir
' depois que os dados mudaram pergunta de novo). O teste chama PptTqq_HandleOpen (o projeto de teste nao tem janela; a guarda de janela
' de PptTqq_OnOpen tem teste proprio).
Public Sub Test_PptTqq_HandleOpen_FirstSightAndReopen()
    Dim p As Object, s As Object, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    AssertEqual 0, PptTqq_CountTagged(p)
    PptTqq_ClearLastMsg
    PptTqq_HandleOpen p
    AssertEqual "", PptTqq_LastMsg(), "sem graficos marcados: nada"
    s.Tags.Add PPT_TAG_CHART, "abc123"
    AssertEqual 1, PptTqq_CountTagged(p)
    PptTqq_HandleOpen p
    AssertEqual "", PptTqq_LastMsg(), "grafico colado depois da primeira vista: nao pergunta"
    PptTqq_OnClose p
    PptTqq_HandleOpen p
    AssertEqual PptTqq_AskText(1), PptTqq_LastMsg(), "fechada e reaberta: pergunta de novo (Test responde Nao)"
    PptTqq_ClearLastMsg
    PptTqq_HandleOpen p
    AssertEqual "", PptTqq_LastMsg(), "uma vez enquanto aberta"
Cleanup:
    en = Err.Number: ed = Err.Description          ' antes de ClosePres: o On Error dela zera o Err
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Guarda: abertura sem janela (automacao) ou PowerPoint invisivel nao pergunta nem atualiza, e nem marca como vista.
Public Sub Test_PptTqq_OnOpen_WindowlessDoesNothing()
    Dim p As Object, en As Long, ed As String
    AssertTrue PptTqq_ShouldHandle(1, True)
    AssertFalse PptTqq_ShouldHandle(0, True), "sem janela"
    AssertFalse PptTqq_ShouldHandle(1, False), "PowerPoint invisivel"
    Set p = NewTagged()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    AssertEqual 0, CLng(p.Windows.Count), "apresentacao do teste sem janela"
    PptTqq_OnOpen p
    AssertEqual 1, PptTqq_OnOpenCallsForTest()
    AssertEqual "", PptTqq_LastMsg(), "sem janela: nenhuma pergunta"
    PptTqq_HandleOpen p
    AssertEqual PptTqq_AskText(1), PptTqq_LastMsg(), "a guarda nao marcou como vista"
Cleanup:
    en = Err.Number: ed = Err.Description
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Auto_Open trata as apresentacoes ja abertas (PptTqq_ScanOpen chama PptTqq_OnOpen em cada uma; a guarda vale).
Public Sub Test_PptTqq_ScanOpen_VisitsOpenPresentations()
    Dim p As Object, en As Long, ed As String
    Set p = NewTagged()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_ScanOpen
    AssertEqual CLng(Application.Presentations.Count), PptTqq_OnOpenCallsForTest(), "uma chamada por apresentacao aberta"
    AssertEqual "", PptTqq_LastMsg(), "todas sem janela: nada"
Cleanup:
    en = Err.Number: ed = Err.Description
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Eventos ligados sem registro: a classe recebe AfterPresentationOpen e PresentationCloseFinal deste PowerPoint. Sem janela (o PowerPoint do
' teste e invisivel) a guarda nao pergunta; a pergunta e conferida chamando PptTqq_HandleOpen com a apresentacao aberta pelo evento.
Public Sub Test_PptTqq_Events_OpenCloseReopen()
    Dim ev As clsPptEvents, p As Object, p2 As Object, f As String, en As Long, ed As String
    f = Environ$("TEMP") & "\thinqq_test_ppt_ev_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd() * 100000) & ".pptx"
    Set p = NewTagged()
    On Error GoTo Cleanup
    p.SaveAs f, 24                                  ' ppSaveAsOpenXMLPresentation
    p.Close
    Set p = Nothing
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    Set ev = New clsPptEvents
    Set ev.App = Application
    AssertTrue ev.App Is Application, "eventos ligados a este PowerPoint"
    Set p2 = Application.Presentations.Open(f, 0, 0, 0)
    AssertEqual 1, PptTqq_OnOpenCallsForTest(), "AfterPresentationOpen chegou ao suplemento"
    AssertEqual "", PptTqq_LastMsg(), "sem janela: nenhuma pergunta"
    PptTqq_HandleOpen p2
    AssertEqual PptTqq_AskText(1), PptTqq_LastMsg(), "apresentacao aberta: pergunta"
    p2.Saved = -1
    p2.Close
    Set p2 = Nothing
    Set p2 = Application.Presentations.Open(f, 0, 0, 0)
    AssertEqual 2, PptTqq_OnOpenCallsForTest(), "reaberta: AfterPresentationOpen de novo"
    PptTqq_ClearLastMsg
    PptTqq_HandleOpen p2
    AssertEqual PptTqq_AskText(1), PptTqq_LastMsg(), "PresentationCloseFinal esqueceu a apresentacao: pergunta de novo"
Cleanup:
    en = Err.Number: ed = Err.Description
    If Not ev Is Nothing Then Set ev.App = Nothing
    Set ev = Nothing
    If Not p2 Is Nothing Then ClosePres p2
    If Not p Is Nothing Then ClosePres p
    PptTqq_ResetTestHooks
    On Error Resume Next
    If Len(Dir$(f)) > 0 Then Kill f
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Descoberta do core como o loader: so pastas x.y.z COM o arquivo do core, da mais nova para a mais velha. Pura.
Public Sub Test_PptTqq_CoreCandidates()
    AssertEqual "9.9.9|1.10.0|1.2.0", PptTqq_CoreCandidates("1.2.0:1|1.10.0:1|1.3.1:0|x:1|1.2:1|9.9.9:1|")
    AssertEqual "", PptTqq_CoreCandidates("")
    AssertEqual "", PptTqq_CoreCandidates("logs:0|1.4.0:0|")
    AssertEqual "1.4.0", PptTqq_CoreCandidates("1.4.0:1")
End Sub

' A mesma regra lendo pastas reais (temporarias): pasta parcial sem o arquivo fica de fora.
Public Sub Test_PptTqq_ListCoreCandidates_SkipsPartialFolder()
    Dim fso As Object, d As String, en As Long, ed As String, f As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    d = Environ$("TEMP") & "\thinqq_test_ppt_versions_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd() * 100000)
    On Error GoTo Cleanup
    fso.CreateFolder d
    fso.CreateFolder d & "\9.9.8"
    fso.CreateFolder d & "\9.9.9"
    fso.CreateFolder d & "\logs"
    f = FreeFile
    Open d & "\9.9.8\ThinQQ.core" & BUILD_SUFFIX & ".xlam" For Output As #f
    Close #f
    AssertEqual "9.9.8", PptTqq_ListCoreCandidates(d), "9.9.9 sem o arquivo (parcial) fica de fora"
    AssertEqual "", PptTqq_ListCoreCandidates(d & "\nao-existe")
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If fso.FolderExists(d) Then fso.DeleteFolder d, True
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Resposta do Core_Ping: 1 = core, 0 = sem core, -1 = Excel ocupado (nao cai para CreateObject). Pura.
Public Sub Test_PptTqq_PingClass()
    AssertEqual 1, PptTqq_PingClass(0, "pong")
    AssertEqual 0, PptTqq_PingClass(1004, "")
    AssertEqual 0, PptTqq_PingClass(0, "x")
    AssertEqual -1, PptTqq_PingClass(-2147418111, ""), "RPC_E_CALL_REJECTED"
    AssertEqual -1, PptTqq_PingClass(-2147417846, ""), "RPC_E_RETRYLATER"
End Sub

Public Sub Test_PptTqq_SyncWithoutTaggedAndDataOnlyMissing()
    Dim p As Object, s As Object, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_Sync p, True
    AssertEqual PPT_MSG_NONE, PptTqq_LastMsg()
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "abc123"
    s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\pasta-ausente.xlsx"
    AssertEqual "Dados atualizados em 0 gráfico(s); 1 sem a pasta de origem (ficaram como estavam).", PptTqq_DataOnly(p)
Cleanup:
    en = Err.Number: ed = Err.Description
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub
