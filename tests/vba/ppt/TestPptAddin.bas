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
    AssertEqual "ask", PptTqq_OnOpenAction("ask", False, 2, False, False)
    AssertEqual "sync", PptTqq_OnOpenAction("always", False, 2, False, False)
    AssertEqual "", PptTqq_OnOpenAction("never", False, 2, False, False)
    AssertEqual "", PptTqq_OnOpenAction("ask", False, 0, False, False), "sem graficos marcados"
    AssertEqual "", PptTqq_OnOpenAction("always", True, 2, False, False), "somente leitura"
    AssertEqual "", PptTqq_OnOpenAction("always", False, 2, True, False), "ja tratada nesta sessao"
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

' Revisao final (I1): a pergunta ao abrir mostra as pastas de origem (caminho completo, no maximo 5 e "..."), mantendo as duas frases fixas.
Public Sub Test_PptTqq_AskText_ListsSources()
    Dim c As New Collection, i As Long, s As String
    s = PptTqq_AskText(2)
    AssertEqual 0, InStr(s, "Pastas de origem"), "sem pastas conhecidas: sem a lista"
    For i = 1 To 7
        c.Add "\\srv" & i & "\dados\pasta" & i & ".xlsx"
    Next
    s = PptTqq_AskText(7, c)
    AssertTrue InStr(s, " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?") > 0, "frase fixa 1"
    AssertTrue InStr(s, "Para não perguntar mais: aba ThinQQ > Ao abrir.") > 0, "frase fixa 2"
    AssertTrue InStr(s, "Pastas de origem:") > 0, "lista"
    AssertTrue InStr(s, "\\srv1\dados\pasta1.xlsx") > 0 And InStr(s, "\\srv5\dados\pasta5.xlsx") > 0, "as 5 primeiras, caminho completo"
    AssertEqual 0, InStr(s, "pasta6.xlsx"), "no maximo 5"
    AssertTrue InStr(s, vbLf & "...") > 0, "mais de 5: reticencias"
    AssertTrue InStr(s, "Pastas de origem:") < InStr(s, "Para não perguntar mais"), "a lista vem antes da dica"
End Sub

' Pastas de origem distintas (sem diferenciar maiusculas) dos graficos marcados: SourceFullName do vinculo ou a etiqueta TQQ_BOOK.
Public Sub Test_PptTqq_CountTagged_CollectsDistinctSources()
    Dim p As Object, s As Object, srcs As New Collection, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): s.Tags.Add PPT_TAG_CHART, "a1": s.Tags.Add PPT_TAG_BOOK, "C:\Dados\A.xlsx"
    Set s = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): s.Tags.Add PPT_TAG_CHART, "a2": s.Tags.Add PPT_TAG_BOOK, "c:\dados\a.xlsx"
    Set s = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50): s.Tags.Add PPT_TAG_CHART, "a3": s.Tags.Add PPT_TAG_BOOK, "\\srv\x\B.xlsx"
    Set s = p.Slides(1).Shapes.AddShape(1, 190, 10, 50, 50): s.Tags.Add PPT_TAG_BOOK, "C:\nao-marcado.xlsx"
    AssertEqual 3, PptTqq_CountTagged(p, srcs)
    AssertEqual 2, srcs.Count, "pastas distintas"
    AssertEqual "C:\Dados\A.xlsx", CStr(srcs(1))
    AssertEqual "\\srv\x\B.xlsx", CStr(srcs(2))
    AssertEqual 3, PptTqq_CountTagged(p), "sem a colecao: so conta"
Cleanup:
    en = Err.Number: ed = Err.Description
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Origem remota (rede \\ ou endereco http/https) ainda nao aberta no Excel: "Atualizar sempre" volta a perguntar. Pura.
Public Sub Test_PptTqq_RemoteSource_AlwaysFallsBackToAsk()
    Dim srcs As New Collection, names As New Collection
    AssertTrue PptTqq_IsRemotePath("\\srv\dados\a.xlsx")
    AssertTrue PptTqq_IsRemotePath("https://contoso.sharepoint.com/x/a.xlsx")
    AssertTrue PptTqq_IsRemotePath("HTTP://intranet/a.xlsx")
    AssertFalse PptTqq_IsRemotePath("C:\Dados\a.xlsx")
    AssertFalse PptTqq_IsRemotePath("")
    srcs.Add "C:\Dados\a.xlsx"
    AssertFalse PptTqq_HasRemoteSource(srcs, names), "so local"
    srcs.Add "https://contoso.sharepoint.com/x/b.xlsx"
    AssertTrue PptTqq_HasRemoteSource(srcs, names), "https nao aberta"
    names.Add "HTTPS://contoso.sharepoint.com/x/B.xlsx"
    AssertFalse PptTqq_HasRemoteSource(srcs, names), "https ja aberta no Excel"
    AssertFalse PptTqq_HasRemoteSource(Nothing, names), "sem lista"
    AssertEqual "ask", PptTqq_OnOpenAction("always", False, 2, False, True), "sempre + origem remota: pergunta"
    AssertEqual "ask", PptTqq_OnOpenAction("ask", False, 2, False, True)
    AssertEqual "", PptTqq_OnOpenAction("never", False, 2, False, True)
    AssertEqual "", PptTqq_OnOpenAction("always", True, 2, False, True), "somente leitura continua sem nada"
End Sub

' Modo so dados (sem o core): mesma regra do core (PptSync_ReachableAmong): aberta no Excel ou, se nao for http(s), arquivo no disco.
Public Sub Test_PptTqq_ReachableAmong()
    Dim names As New Collection, f As String, h As Integer
    AssertFalse PptTqq_ReachableAmong("https://contoso.sharepoint.com/x/a.xlsx", names), "https nao aberta"
    names.Add "https://contoso.sharepoint.com/x/A.xlsx"
    AssertTrue PptTqq_ReachableAmong("https://contoso.sharepoint.com/x/a.xlsx", names), "https aberta no Excel"
    AssertFalse PptTqq_ReachableAmong("C:\nao\existe\a.xlsx", names), "local inexistente"
    AssertFalse PptTqq_ReachableAmong("", names), "vazio"
    f = Environ$("TEMP") & "\thinqq_test_ppt_reach_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd() * 100000) & ".xlsx"
    h = FreeFile: Open f For Output As #h: Close #h
    AssertTrue PptTqq_ReachableAmong(f, New Collection), "local existente"
    Kill f
End Sub

' A pergunta ao abrir cita a pasta de origem do grafico marcado.
Public Sub Test_PptTqq_HandleOpen_AskShowsSource()
    Dim p As Object, s As Object, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "abc123": s.Tags.Add PPT_TAG_BOOK, "\\srv\dados\vendas.xlsx"
    PptTqq_HandleOpen p
    AssertTrue InStr(PptTqq_LastMsg(), "Pastas de origem:" & vbLf & "\\srv\dados\vendas.xlsx") > 0, PptTqq_LastMsg()
Cleanup:
    en = Err.Number: ed = Err.Description
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Re-revisao (I1): barras normais e mistas, file:, \\?\UNC\ e WebDAV contam como remotas; \\?\C:\ (unidade local com prefixo longo) e
' letra de unidade (inclusive mapeada, que nao da para distinguir) contam como locais. Pura.
Public Sub Test_PptTqq_IsRemotePath_Variants()
    AssertTrue PptTqq_IsRemotePath("//srv/share/x.xlsx"), "barras normais"
    AssertTrue PptTqq_IsRemotePath("\/srv\share\x.xlsx"), "barras mistas"
    AssertTrue PptTqq_IsRemotePath("/\srv/share/x.xlsx"), "barras mistas (2)"
    AssertTrue PptTqq_IsRemotePath("file://srv/share/x.xlsx"), "file: com servidor"
    AssertTrue PptTqq_IsRemotePath("FILE:///C:/dados/x.xlsx"), "file: (sempre remota: conservador)"
    AssertTrue PptTqq_IsRemotePath("\\?\UNC\srv\share\x.xlsx"), "\\?\UNC\"
    AssertTrue PptTqq_IsRemotePath("\\srv@SSL\DavWWWRoot\x.xlsx"), "WebDAV"
    AssertTrue PptTqq_IsRemotePath("https://contoso.sharepoint.com/x.xlsx"), "https"
    AssertTrue PptTqq_IsRemotePath("http:\\intranet\x.xlsx"), "http com barras invertidas"
    AssertTrue PptTqq_IsRemotePath("  \\srv\x.xlsx"), "espacos antes"
    AssertFalse PptTqq_IsRemotePath("\\?\C:\dados\x.xlsx"), "\\?\C:\ e unidade local"
    AssertFalse PptTqq_IsRemotePath("C:/dados/x.xlsx"), "letra de unidade com barras normais"
    AssertFalse PptTqq_IsRemotePath("Z:\mapeada\x.xlsx"), "unidade mapeada: indistinguivel (limite documentado)"
End Sub

' Caracteres de controle no caminho viram espaco (nunca criam linha nova na pergunta); repetidos colapsam.
Public Sub Test_PptTqq_AskText_SanitizesControlChars()
    Dim c As New Collection, s As String
    c.Add "C:\a" & vbLf & "Sim, atualize" & vbCr & vbTab & Chr$(127) & ".xlsx"
    s = PptTqq_AskText(1, c)
    AssertEqual 5, Len(s) - Len(Replace(s, vbLf, "")), "so as quebras da propria pergunta (2 antes da lista, 1 do caminho, 2 antes da dica): " & s
    AssertTrue InStr(s, "C:\a Sim, atualize .xlsx") > 0, s
    AssertEqual 0, InStr(s, vbCr), "sem CR"
End Sub

' Caminho muito longo: cortado no meio (inicio e fim, no maximo 150), remoto listado primeiro, a dica e a frase da pergunta sempre ficam e
' a pergunta inteira cabe em 900 caracteres (o MsgBox corta perto de 1024).
Public Sub Test_PptTqq_AskText_LongPathsKeepRemoteAndHint()
    Dim c As New Collection, s As String, i As Long, ln As Variant, longest As Long
    c.Add "C:\" & String$(400, "a") & "\fim.xlsx"
    c.Add "\\srv\dados\remota.xlsx"
    s = PptTqq_AskText(2, c)
    AssertTrue InStr(s, "Pastas de origem:" & vbLf & "\\srv\dados\remota.xlsx") > 0, "remota primeiro: " & s
    AssertTrue InStr(s, "Para não perguntar mais: aba ThinQQ > Ao abrir.") > 0, "dica"
    AssertTrue InStr(s, " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?") > 0, "frase da pergunta"
    AssertTrue InStr(s, "\fim.xlsx") > 0 And InStr(s, "C:\aaa") > 0 And InStr(s, "...") > 0, "comeco e fim do caminho longo, com reticencias"
    For Each ln In Split(s, vbLf)
        If Len(ln) > longest Then longest = Len(ln)
    Next
    AssertTrue longest <= 150, "nenhuma linha passa de 150 (" & longest & ")"
    Set c = New Collection
    For i = 1 To 5
        c.Add "C:\" & i & String$(246, "b") & ".xlsx"
    Next
    s = PptTqq_AskText(5, c)
    AssertTrue InStr(s, "Para não perguntar mais: aba ThinQQ > Ao abrir.") > 0, "5 caminhos de 250: dica fica"
    AssertTrue Len(s) <= 900, "cabe em 900 (" & Len(s) & ")"
    AssertTrue Right$(s, Len("Para não perguntar mais: aba ThinQQ > Ao abrir.")) = "Para não perguntar mais: aba ThinQQ > Ao abrir.", "a dica fecha a pergunta"
End Sub
