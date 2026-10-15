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

' ---- 1.6.0 (C3): Atualizar automaticamente / Atualizar uma vez ----

' Contrato com o core (TestPptSync.Test_PptIdx_Literal tem as MESMAS cadeias).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String, bad As Variant, g As String, c As Collection
    a = "256" & Chr$(31) & "tqq_abc" & Chr$(31) & "C:\a\b.xlsx"
    b = "257" & Chr$(31) & "tqq_def" & Chr$(31) & "C:\a\c.xlsx"
    s = PptTqq_IdxSet("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & Chr$(30) & b, s
    AssertEqual a & Chr$(30) & b, PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptTqq_IdxSet(s, 256, "tqq_abc", "", False)
    AssertEqual "", PptTqq_IdxSet(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptTqq_IdxParse(s & Chr$(30) & "x" & Chr$(31) & "y" & Chr$(31) & "z" & Chr$(30) & Chr$(30)).Count
    AssertEqual a, PptTqq_IdxEntry(256, "tqq_abc", "C:\a\b.xlsx")
    ' total para lixo da apresentacao: SlideID so com 1 a 9 algarismos; nada levanta erro
    For Each bad In Array("9e9", "1E10", "3d9", "&H10", "(5)", " 1", "-1", "1 ", "+1", "1234567890", "", "1.5", "1,5", ChrW$(1633))
        g = g & CStr(bad) & Chr$(31) & "n" & Chr$(31) & "" & Chr$(30)
    Next
    AssertEqual 2, PptTqq_IdxParse(g & s).Count, "SlideID invalido ignorado"
    AssertEqual 0, PptTqq_IdxParse(Chr$(31) & Chr$(31) & Chr$(31) & Chr$(30) & Chr$(31)).Count, "so separadores"
    AssertEqual 0, PptTqq_IdxParse("").Count
    AssertEqual 1, PptTqq_IdxParse("0" & Chr$(31) & "n" & Chr$(31)).Count, "pasta vazia e SlideID 0 sao validos"
    Set c = PptTqq_IdxParse("999999999" & Chr$(31) & "n" & Chr$(31) & "b")
    AssertEqual 1, c.Count, "9 algarismos"
    AssertEqual 999999999, IdxSlide(c, 1)
    ' separadores nunca entram num campo
    AssertEqual "1" & Chr$(31) & "a_b" & Chr$(31) & "c_d", PptTqq_IdxEntry(1, "a" & Chr$(30) & "b", "c" & Chr$(31) & "d")
    AssertEqual 1, PptTqq_IdxParse(PptTqq_IdxSet("", 1, "x" & Chr$(30) & "y" & Chr$(31) & "z", "p" & Chr$(30), True)).Count, "nome com separadores: uma entrada"
    AssertEqual "", PptTqq_IdxSet(PptTqq_IdxSet("", 1, "x" & Chr$(30) & "y", "", True), 1, "x" & Chr$(30) & "y", "", False), "desligar acha o nome saneado"
End Sub

' Tetos do indice (iguais aos do core): entradas alem de PPT_IDX_MAX_ENTRIES ignoradas; etiqueta acima de PPT_IDX_MAX_LEN recusada inteira.
Public Sub Test_PptIdx_Caps()
    Dim a() As String, i As Long, big As String, c As Collection
    ReDim a(0 To 2499)
    For i = 0 To 2499
        a(i) = PptTqq_IdxEntry(i + 1, "tqq_" & i, "")
    Next
    Set c = PptTqq_IdxParse(Join(a, Chr$(30)))
    AssertEqual 2000, c.Count, "2500 entradas: as 2000 primeiras"
    AssertEqual 1, IdxSlide(c, 1)
    AssertEqual 2000, IdxSlide(c, 2000)
    AssertEqual 2000, PptTqq_IdxParse(PptTqq_IdxSet(Join(a, Chr$(30)), 9999, "nova", "", True)).Count
    For i = 0 To 2499
        a(i) = PptTqq_IdxEntry(i + 1, "tqq_" & i, "C:\" & String$(480, "a") & ".xlsx")
    Next
    big = Join(a, Chr$(30))
    AssertTrue Len(big) > 1200000, CStr(Len(big))
    AssertEqual 0, PptTqq_IdxParse(big).Count, "etiqueta de 1,2 MB recusada"
    AssertEqual 0, PptTqq_IdxParse(Left$(big, PPT_IDX_MAX_LEN + 1)).Count, "um caractere acima do teto"
    AssertTrue PptTqq_IdxParse(Left$(big, PPT_IDX_MAX_LEN)).Count > 0, "no teto ainda le"
End Sub

Private Function IdxSlide(c As Collection, i As Long) As Long
    Dim v As Variant
    v = c.Item(i)
    IdxSlide = CLng(v(0))
End Function

' Ligar/desligar sem Excel: etiqueta na forma, indice na apresentacao, nome repetido no slide renomeado.
Public Sub Test_PptTqq_SetAutoShapes_TagsAndIndex()
    Dim p As Object, s1 As Object, s2 As Object, c As New Collection, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s1 = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): s1.Name = "tqq_x": s1.Tags.Add PPT_TAG_CHART, "abc123": s1.Tags.Add PPT_TAG_BOOK, "C:\t\a.xlsx"
    Set s2 = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): s2.Name = "tqq_x": s2.Tags.Add PPT_TAG_CHART, "abc123": s2.Tags.Add PPT_TAG_BOOK, "C:\t\a.xlsx"
    c.Add s1: c.Add s2
    AssertFalse PptTqq_ShapesAuto(c)
    PptTqq_SetAutoShapes c, True
    AssertEqual "1", s1.Tags(PPT_TAG_AUTO): AssertEqual "1", s2.Tags(PPT_TAG_AUTO)
    AssertTrue s1.Name <> s2.Name, "nomes unicos no slide"
    AssertEqual 2, PptTqq_IdxParse(p.Tags(PPT_TAG_INDEX)).Count
    AssertTrue PptTqq_ShapesAuto(c)
    PptTqq_SetAutoShapes c, False
    AssertEqual "", p.Tags(PPT_TAG_INDEX)
    AssertFalse PptTqq_ShapesAuto(c)
    AssertEqual PptTqq_IdxEntry(p.Slides(1).SlideID, s1.Name, "") & Chr$(30) & PptTqq_IdxEntry(p.Slides(1).SlideID, s2.Name, ""), PptTqq_EntriesOf(c)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' "Atualizar uma vez" sem Excel (gancho com Nothing): aviso de Excel ausente, nada quebra.
Public Sub Test_PptTqq_SyncEntries_NoExcel()
    Dim p As Object
    Set p = NewTagged()
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ClearLastMsg
    PptTqq_SyncEntries p, PptTqq_IdxEntry(p.Slides(1).SlideID, p.Slides(1).Shapes(1).Name, ""), True
    AssertEqual PPT_MSG_NO_EXCEL, PptTqq_LastMsg()
    PptTqq_ResetTestHooks
    ClosePres p
End Sub

' Medido (revisao C1, item 13): Shapes(nome) do PowerPoint acha a forma IGNORANDO a caixa (com "tqq_A" e "tqq_a" no slide, os dois nomes
' acham a primeira). Por isso o nome unico compara sem diferenciar maiusculas: ligar as duas deixa cada entrada do indice achar a sua forma.
Public Sub Test_PptTqq_SetAutoShapes_CaseOnlyNamesBecomeUnique()
    Dim p As Object, sl As Object, s1 As Object, s2 As Object, c As New Collection, e As Variant, ids As String, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set sl = p.Slides(1)
    Set s1 = sl.Shapes.AddShape(1, 10, 10, 50, 50): s1.Name = "tqq_A": s1.Tags.Add PPT_TAG_CHART, "c1"
    Set s2 = sl.Shapes.AddShape(1, 70, 10, 50, 50): s2.Name = "tqq_a": s2.Tags.Add PPT_TAG_CHART, "c2"
    AssertEqual CLng(s1.Id), CLng(sl.Shapes("tqq_a").Id), "medido: a busca pelo nome ignora a caixa"
    AssertEqual CLng(s1.Id), CLng(sl.Shapes("TQQ_A").Id), "medido: a busca pelo nome ignora a caixa (2)"
    c.Add s1: c.Add s2
    PptTqq_SetAutoShapes c, True
    AssertTrue StrComp(CStr(s1.Name), CStr(s2.Name), vbTextCompare) <> 0, "nomes unicos sem diferenciar maiusculas: " & s1.Name & " / " & s2.Name
    For Each e In PptTqq_IdxParse(CStr(p.Tags(PPT_TAG_INDEX)))
        ids = ids & CStr(p.Slides.FindBySlideID(CLng(e(0))).Shapes(CStr(e(1))).Id) & ","
    Next
    AssertEqual CStr(s1.Id) & "," & CStr(s2.Id) & ",", ids, "cada entrada do indice acha a sua forma"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    Set s1 = Nothing: Set s2 = Nothing: Set sl = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' A API e conferida por recurso: a sincronizacao ao abrir continua com a API 1 (cores 1.5.x); "Atualizar uma vez"/automatico pedem a 2.
Public Sub Test_PptTqq_ApiPerFeature()
    AssertEqual 1, PPT_REQUIRED_API, "a API exigida para a sincronizacao ao abrir nao subiu"
    AssertEqual 2, PPT_API_SHAPES
    AssertFalse PptTqq_ShapesApiOk(1), "core 1.5.x: sem Core_PptSyncShapes"
    AssertTrue PptTqq_ShapesApiOk(2): AssertTrue PptTqq_ShapesApiOk(3)
    AssertFalse PptTqq_ShapesApiOk(0)
End Sub

Private Function NewFakeXl(api As Variant) As TestXlFake
    Dim f As New TestXlFake
    f.Api = api
    Set NewFakeXl = f
End Function

' Core com a API 2: "Atualizar uma vez" chama Core_PptSyncShapes com as entradas e mostra o resumo (inclusive o texto de ocupado do core).
Public Sub Test_PptTqq_SyncEntries_Api2_CallsSyncShapes()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewTagged()
    On Error GoTo Cleanup
    Set f = NewFakeXl(2)
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest f, ""
    PptTqq_ClearLastMsg
    PptTqq_SyncEntries p, "E1", True
    AssertEqual "selecionados", PptTqq_LastMsg()
    AssertEqual "E1", f.LastEntries
    AssertEqual 0, InStr(f.Calls, "Core_PptSyncPresentation"), f.Calls
    f.ShapesResult = "O PowerPoint está ocupado."
    PptTqq_SyncEntries p, "E2", True
    AssertEqual "O PowerPoint está ocupado.", PptTqq_LastMsg(), "texto do core exibido como veio"
    f.FailOn = "Core_PptSyncShapes"
    PptTqq_SyncEntries p, "E3", True
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível atualizar") > 0, PptTqq_LastMsg()
    f.FailOn = "": f.ShapesResult = ""
    PptTqq_SyncEntries p, "E4", True
    AssertEqual "selecionados", PptTqq_LastMsg(), "depois da falha a reentrada foi liberada"
    AssertEqual "E4", f.LastEntries
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Core 1.5.x (API 1): "Atualizar uma vez" atualiza a apresentacao inteira e diz para atualizar o ThinQQ; ligar o automatico grava a marca e
' avisa; a sincronizacao ao abrir/botao continua chamando Core_PptSyncPresentation sem aviso de core antigo.
Public Sub Test_PptTqq_SyncEntries_OldCore_WholePresentation()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewTagged()
    On Error GoTo Cleanup
    Set f = NewFakeXl(1)
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest f, ""
    PptTqq_ClearLastMsg
    PptTqq_SyncEntries p, "E1", True
    AssertEqual "inteira" & vbLf & PPT_MSG_SEL_OLDCORE, PptTqq_LastMsg()
    AssertEqual 0, InStr(f.Calls, "Core_PptSyncShapes"), f.Calls
    c.Add p.Slides(1).Shapes(1)
    f.Calls = ""
    PptTqq_ToggleAuto c, True
    AssertEqual "1", CStr(p.Slides(1).Shapes(1).Tags(PPT_TAG_AUTO)), "a marca foi gravada"
    AssertEqual PPT_MSG_AUTO_OLDCORE, PptTqq_LastMsg(), "ligar com core antigo: so a marca e o aviso"
    AssertEqual 0, InStr(f.Calls, "Core_PptSync"), "ligar com core antigo nao atualiza nada: " & f.Calls
    f.Calls = ""
    PptTqq_Sync p, True
    AssertEqual "inteira", PptTqq_LastMsg(), "sincronizacao da apresentacao: API 1 basta"
    AssertTrue InStr(f.Calls, "Core_PptSyncPresentation") > 0, f.Calls
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Ligar o automatico ja atualiza uma vez so as formas dadas (Core_PptSyncShapes); desligar so tira a marca (nenhuma chamada ao Excel).
Public Sub Test_PptTqq_ToggleAuto_OnSyncsOff_NoExcel()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewTagged()
    On Error GoTo Cleanup
    Set f = NewFakeXl(2)
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest f, ""
    c.Add p.Slides(1).Shapes(1)
    PptTqq_ToggleAuto c, True
    AssertTrue PptTqq_ShapesAuto(c)
    AssertEqual PptTqq_EntriesOf(c), f.LastEntries
    AssertEqual "selecionados", PptTqq_LastMsg()
    f.Calls = "": PptTqq_ClearLastMsg
    PptTqq_ToggleAuto c, False
    AssertFalse PptTqq_ShapesAuto(c)
    AssertEqual "", f.Calls, "desligar nao chama o Excel"
    AssertEqual "", PptTqq_LastMsg()
    AssertEqual "", CStr(p.Tags(PPT_TAG_INDEX))
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Fix C3 (I1): "Atualizar uma vez" deixa os nomes unicos ANTES de montar as entradas (grafico copiado e colado no mesmo slide tem o mesmo
' nome; a busca do core por nome ignora a caixa): cada entrada acha a SUA forma.
Public Sub Test_PptTqq_UpdateSelected_DuplicateNames_DistinctEntries()
    Dim p As Object, sl As Object, s1 As Object, s2 As Object, s3 As Object, f As TestXlFake, c As New Collection, e As Variant, ids As String, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set sl = p.Slides(1)
    Set s1 = sl.Shapes.AddShape(1, 10, 10, 50, 50): s1.Name = "tqq_x": s1.Tags.Add PPT_TAG_CHART, "c1"
    Set s2 = sl.Shapes.AddShape(1, 70, 10, 50, 50): s2.Name = "tqq_x": s2.Tags.Add PPT_TAG_CHART, "c2"
    Set s3 = sl.Shapes.AddShape(1, 130, 10, 50, 50): s3.Name = "TQQ_X": s3.Tags.Add PPT_TAG_CHART, "c3"
    Set f = NewFakeXl(2)
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest f, ""
    c.Add s2: c.Add s3
    PptTqq_UpdateSelected c
    AssertEqual "selecionados", PptTqq_LastMsg()
    For Each e In PptTqq_IdxParse(f.LastEntries)
        ids = ids & CStr(p.Slides.FindBySlideID(CLng(e(0))).Shapes(CStr(e(1))).Id) & ","
    Next
    AssertEqual CStr(s2.Id) & "," & CStr(s3.Id) & ",", ids, "entradas resolvem para as formas selecionadas: " & f.LastEntries
    AssertEqual "", CStr(s2.Tags(PPT_TAG_AUTO)), "atualizar uma vez nao liga o automatico"
    PptTqq_ClearLastMsg
    PptTqq_UpdateSelected New Collection
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s1 = Nothing: Set s2 = Nothing: Set s3 = Nothing: Set sl = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Fix C3 (m1): ligar o automatico sem Excel ou com Excel sem o core so grava a marca e da a dica de abrir o Excel com o ThinQQ (nada de
' atualizar, nada de Excel novo); "Atualizar uma vez" com core sem Core_PptApiVersion: so os dados e UM aviso (sem empilhar).
Public Sub Test_PptTqq_ToggleAuto_NoCoreOrNoExcel_TagOnly()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewTagged()
    On Error GoTo Cleanup
    c.Add p.Slides(1).Shapes(1)
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ToggleAuto c, True
    AssertTrue PptTqq_ShapesAuto(c), "marca gravada sem Excel"
    AssertEqual PPT_MSG_AUTO_NOCORE, PptTqq_LastMsg(), "sem Excel"
    Set f = NewFakeXl(2): f.NoCore = True
    PptTqq_SetExcelForTest f, "", True
    PptTqq_ToggleAuto c, False
    PptTqq_ToggleAuto c, True
    AssertTrue PptTqq_ShapesAuto(c)
    AssertEqual PPT_MSG_AUTO_NOCORE, PptTqq_LastMsg(), "Excel sem o core"
    AssertEqual 0, InStr(f.Calls, "Core_PptSync"), f.Calls
    Set f = NewFakeXl(Empty)
    PptTqq_SetExcelForTest f, ""
    PptTqq_SyncEntries p, "E1", True
    AssertEqual PptTqq_DataOnly(p, New Collection) & vbLf & PPT_MSG_OLD_CORE, PptTqq_LastMsg(), "core sem a API: um aviso so"
    AssertEqual 1, (Len(f.Calls) - Len(Replace(f.Calls, "Core_PptApiVersion", ""))) / Len("Core_PptApiVersion"), "API lida uma vez: " & f.Calls
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Fix C3 (m6): core com a API 2 e a atualizacao automatica desligada no Excel: ligar avisa onde liga-la. (m2): falha ao gravar a marca
' tem texto proprio.
Public Sub Test_PptTqq_ToggleAuto_AutoOffNote_AndTagFailure()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection, bad As New Collection
    Set p = NewTagged()
    On Error GoTo Cleanup
    c.Add p.Slides(1).Shapes(1)
    Set f = NewFakeXl(2): f.AutoOn = False
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest f, ""
    PptTqq_ToggleAuto c, True
    AssertEqual "selecionados" & vbLf & PPT_MSG_AUTO_OFF, PptTqq_LastMsg()
    f.AutoOn = True
    PptTqq_ToggleAuto c, True
    AssertEqual "selecionados", PptTqq_LastMsg(), "ligada no Excel: sem a nota"
    bad.Add New TestXlFake
    PptTqq_ToggleAuto bad, True
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível marcar o gráfico como automático: ") = 1, PptTqq_LastMsg()
    PptTqq_ToggleAuto bad, False
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível voltar o gráfico para a atualização manual: ") = 1, PptTqq_LastMsg()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Fix C3 (m4): o aviso de Excel ocupado nao cita um botao (vale para Atualizar graficos, Atualizar uma vez e Atualizar automaticamente).
Public Sub Test_PptTqq_BusyText_Generic()
    AssertEqual 0, InStr(PPT_MSG_BUSY, "Atualizar gráficos"), PPT_MSG_BUSY
    AssertTrue InStr(PPT_MSG_BUSY, "tente de novo") > 0, PPT_MSG_BUSY
End Sub

' Usado por tests/ps/Test-PptAddin.ps1 (ponta a ponta, passo a2): "Atualizar uma vez" de uma forma dada, como o botao com ela selecionada.
Public Sub E2e_UpdateShape(shp As Object)
    Dim c As New Collection
    c.Add shp
    PptTqq_UpdateSelected c
End Sub

' Faixa sem selecao (o PowerPoint do teste nao tem janela): botoes desligados, toggle solto, clique avisa para selecionar.
Public Sub Test_PptRx_ElemCallbacks_NoSelection()
    Dim en As Variant, pr As Variant
    AssertEqual 0, PptTqq_SelectedTagged().Count
    PptRx_GetElemEnabled Nothing, en
    AssertFalse CBool(en)
    PptRx_GetAuto Nothing, pr
    AssertFalse CBool(pr)
    PptTqq_ClearLastMsg
    PptRx_OnUpdateSel Nothing
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg()
    PptTqq_ClearLastMsg
    PptRx_OnAuto Nothing, True
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg()
    PptRx_InvalidateElem
End Sub
