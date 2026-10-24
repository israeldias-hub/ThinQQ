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
    AssertTrue InStr(s, "Para não perguntar mais: Vínculos > Ao abrir a apresentação.") > 0, "frase fixa 2"
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
    AssertTrue InStr(s, "Para não perguntar mais: Vínculos > Ao abrir a apresentação.") > 0, "dica"
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
    AssertTrue InStr(s, "Para não perguntar mais: Vínculos > Ao abrir a apresentação.") > 0, "5 caminhos de 250: dica fica"
    AssertTrue Len(s) <= 900, "cabe em 900 (" & Len(s) & ")"
    AssertTrue Right$(s, Len("Para não perguntar mais: Vínculos > Ao abrir a apresentação.")) = "Para não perguntar mais: Vínculos > Ao abrir a apresentação.", "a dica fecha a pergunta"
End Sub

' ---- 1.6.0 (C3): Atualizar automaticamente / Atualizar uma vez ----

' Contrato com o core (TestPptSync.Test_PptIdx_Literal tem as MESMAS cadeias).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String, bad As Variant, g As String, c As Collection
    a = "256" & "|" & "tqq_abc" & "|" & "C:\a\b.xlsx"
    b = "257" & "|" & "tqq_def" & "|" & "C:\a\c.xlsx"
    s = PptTqq_IdxSet("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & vbLf & b, s
    AssertEqual a & vbLf & b, PptTqq_IdxSet(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptTqq_IdxSet(s, 256, "tqq_abc", "", False)
    AssertEqual "", PptTqq_IdxSet(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptTqq_IdxParse(s & vbLf & "x" & "|" & "y" & "|" & "z" & vbLf & vbLf).Count
    AssertEqual a, PptTqq_IdxEntry(256, "tqq_abc", "C:\a\b.xlsx")
    ' total para lixo da apresentacao: SlideID so com 1 a 9 algarismos; nada levanta erro
    For Each bad In Array("9e9", "1E10", "3d9", "&H10", "(5)", " 1", "-1", "1 ", "+1", "1234567890", "", "1.5", "1,5", ChrW$(1633))
        g = g & CStr(bad) & "|" & "n" & "|" & "" & vbLf
    Next
    AssertEqual 2, PptTqq_IdxParse(g & s).Count, "SlideID invalido ignorado"
    AssertEqual 0, PptTqq_IdxParse("|" & "|" & "|" & vbLf & "|").Count, "so separadores"
    AssertEqual 0, PptTqq_IdxParse("").Count
    AssertEqual 1, PptTqq_IdxParse("0" & "|" & "n" & "|").Count, "pasta vazia e SlideID 0 sao validos"
    Set c = PptTqq_IdxParse("999999999" & "|" & "n" & "|" & "b")
    AssertEqual 1, c.Count, "9 algarismos"
    AssertEqual 999999999, IdxSlide(c, 1)
    ' separadores nunca entram num campo
    AssertEqual "1" & "|" & "a_b" & "|" & "c_d", PptTqq_IdxEntry(1, "a" & vbLf & "b", "c" & "|" & "d")
    AssertEqual "2|a_b_c_d|e_f", PptTqq_IdxEntry(2, "a" & Chr$(30) & "b" & Chr$(31) & "c" & vbTab & "d", "e" & vbCr & "f"), "o que o .pptx salvo muda"
    AssertEqual 1, PptTqq_IdxParse(PptTqq_IdxSet("", 1, "x" & vbLf & "y" & "|" & "z", "p" & vbLf, True)).Count, "nome com separadores: uma entrada"
    AssertEqual "", PptTqq_IdxSet(PptTqq_IdxSet("", 1, "x" & vbLf & "y", "", True), 1, "x" & vbLf & "y", "", False), "desligar acha o nome saneado"
End Sub

' Tetos do indice (iguais aos do core): entradas alem de PPT_IDX_MAX_ENTRIES ignoradas; etiqueta acima de PPT_IDX_MAX_LEN recusada inteira.
Public Sub Test_PptIdx_Caps()
    Dim a() As String, i As Long, big As String, c As Collection
    ReDim a(0 To 2499)
    For i = 0 To 2499
        a(i) = PptTqq_IdxEntry(i + 1, "tqq_" & i, "")
    Next
    Set c = PptTqq_IdxParse(Join(a, vbLf))
    AssertEqual 2000, c.Count, "2500 entradas: as 2000 primeiras"
    AssertEqual 1, IdxSlide(c, 1)
    AssertEqual 2000, IdxSlide(c, 2000)
    AssertEqual 2000, PptTqq_IdxParse(PptTqq_IdxSet(Join(a, vbLf), 9999, "nova", "", True)).Count
    For i = 0 To 2499
        a(i) = PptTqq_IdxEntry(i + 1, "tqq_" & i, "C:\" & String$(480, "a") & ".xlsx")
    Next
    big = Join(a, vbLf)
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
    AssertEqual PptTqq_IdxEntry(p.Slides(1).SlideID, s1.Name, "") & vbLf & PptTqq_IdxEntry(p.Slides(1).SlideID, s2.Name, ""), PptTqq_EntriesOf(c)
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
    AssertEqual 3, PPT_API_DATA, "criar/editar pelo PowerPoint (1.8.0) pede a API 3"
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
    AssertEqual PPT_MSG_SELECT_LINKED, PptTqq_LastMsg()
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
    AssertEqual PPT_MSG_SELECT_LINKED, PptTqq_LastMsg()
    PptTqq_ClearLastMsg
    PptRx_OnAuto Nothing, True
    AssertEqual PPT_MSG_SELECT_LINKED, PptTqq_LastMsg()
    PptTqq_ClearLastMsg
    PptRx_OnEditData Nothing
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "Editar dados sem selecao"
    PptRx_InvalidateElem
End Sub

' ---- 1.8.0 (Parte G, G4): criar graficos pelo PowerPoint ----

' Mesmas cadeias de TestPptData.Test_PptData_PathFor (contrato core x .ppam).
Public Sub Test_PptTqq_DataPathFor()
    AssertEqual "C:\a\Vendas - dados ThinQQ.xlsx", PptTqq_DataPathFor("C:\a\Vendas.pptx")
    AssertEqual "C:\a\b.c - dados ThinQQ.xlsx", PptTqq_DataPathFor("C:\a\b.c.pptx")
    AssertEqual "\\srv\x\R - dados ThinQQ.xlsx", PptTqq_DataPathFor("\\srv\x\R.pptx")
    AssertEqual "", PptTqq_DataPathFor("https://empresa.sharepoint.com/x/R.pptx")
    AssertEqual "", PptTqq_DataPathFor("Apresentação1")
    AssertEqual "a - dados ThinQQ.xlsx", PptTqq_FileNameOf("\\srv\x\a - dados ThinQQ.xlsx")
    AssertEqual "b.xlsx", PptTqq_FileNameOf("C:/x/b.xlsx")
    AssertEqual "", PptTqq_FileNameOf("C:\x\")
End Sub

Public Sub Test_PptTqq_InsertChartIn_UnsavedAndNoExcel()
    Dim p As Object, f As String, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual PPT_MSG_SAVE_FIRST, PptTqq_LastMsg(), "apresentacao nunca salva"
    f = Environ$("TEMP") & "\tqq-ppam-g4.pptx"
    p.SaveAs f
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ClearLastMsg
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual PPT_MSG_NO_EXCEL_DATA, PptTqq_LastMsg(), "sem Excel"
    AssertEqual 0, p.Slides(1).Shapes.Count
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
    End If
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Reapontar: sem a etiqueta ou com a pasta de dados no lugar, nada; forma que nao e grafico vinculado: nada reapontado (erro so registrado).
Public Sub Test_PptTqq_RelinkData_Guards()
    Dim p As Object, s As Object, f As String, d As String, k As Integer, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    AssertEqual 0, PptTqq_RelinkData(p), "sem TQQ_DATABOOK"
    f = Environ$("TEMP") & "\tqq-relink.pptx": p.SaveAs f
    d = Environ$("TEMP") & "\tqq-relink - dados ThinQQ.xlsx"
    k = FreeFile: Open d For Output As #k: Close #k
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\tqq-relink - dados ThinQQ.xlsx"
    p.Tags.Add PPT_TAG_DATABOOK, "C:\nao\existe\tqq-relink - dados ThinQQ.xlsx"
    AssertEqual 0, PptTqq_RelinkData(p), "forma sem vinculo: nao reaponta"
    AssertEqual "C:\nao\existe\tqq-relink - dados ThinQQ.xlsx", CStr(s.Tags(PPT_TAG_BOOK)), "etiqueta intacta"
    p.Tags.Add PPT_TAG_DATABOOK, d
    AssertEqual 0, PptTqq_RelinkData(p), "pasta de dados no lugar: barato, nada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    Set s = Nothing
    ClosePres p
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
    End If
    If Len(d) > 0 Then
        If Len(Dir$(d)) > 0 Then Kill d
    End If
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Apresentacao nova salva em %TEMP% (o chamador apaga o arquivo).
Private Function NewSavedPres(nm As String) As Object
    Dim p As Object
    Set p = NewPres()
    p.SaveAs Environ$("TEMP") & "\" & nm & ".pptx"
    Set NewSavedPres = p
End Function

Private Sub KillPres(p As Object)
    Dim f As String
    On Error Resume Next
    f = CStr(p.FullName)
    ClosePres p
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
    End If
    Err.Clear
End Sub

' Excel sem o core (nao criado por nos) e core sem a API 3: recusas proprias, sem chamar Core_PptInsertChart nem criar slide.
Public Sub Test_PptTqq_InsertChartIn_NoCoreAndOldCore()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewSavedPres("tqq-ppam-g4a")
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(3): f.NoCore = True
    PptTqq_SetExcelForTest f, ""
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_DATA_NOCORE, PptTqq_LastMsg(), "Excel sem o ThinQQ"
    Set f = NewFakeXl(2)
    PptTqq_SetExcelForTest f, ""
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_OLD_CORE_DATA, PptTqq_LastMsg(), "core 1.6/1.7 (API 2)"
    AssertEqual 0, InStr(f.Calls, "Core_PptInsertChart"), f.Calls
    AssertEqual 1, p.Slides.Count, "nenhum slide criado nas recusas"
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_OLD_CORE_DATA, PptTqq_LastMsg(), "a reentrada foi liberada na recusa"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    KillPres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Revisao G4 (I3): Excel CRIADO nesta chamada sem um core utilizavel (nenhum core na pasta de versoes, ou core sem a API 3) recebe Quit na
' hora e nao fica guardado (o clique seguinte nao acha Excel); o do usuario (nao criado) nunca recebe Quit.
Public Sub Test_PptTqq_GetEditExcel_CreatedWithoutUsableCore_Quits()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewSavedPres("tqq-ppam-g4d")
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(3): f.NoCore = True
    PptTqq_SetExcelForTest f, "", True
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_DATA_NOCORE, PptTqq_LastMsg(), "sem core instalado"
    AssertEqual 1, f.QuitCount, "o Excel criado recebeu Quit"
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_NO_EXCEL_DATA, PptTqq_LastMsg(), "nada guardado depois do Quit"
    Set f = NewFakeXl(2)
    PptTqq_SetExcelForTest f, "", True
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_OLD_CORE_DATA, PptTqq_LastMsg(), "core sem a API 3"
    AssertEqual 1, f.QuitCount, "o Excel criado com core antigo recebeu Quit"
    AssertEqual 0, InStr(f.Calls, "Core_PptInsertChart"), f.Calls
    Set f = NewFakeXl(2)
    PptTqq_SetExcelForTest f, "", False
    PptTqq_InsertChartIn p, Nothing, "stacked"
    AssertEqual PPT_MSG_OLD_CORE_DATA, PptTqq_LastMsg()
    AssertEqual 0, f.QuitCount, "o Excel do usuario nunca recebe Quit"
    AssertEqual 1, p.Slides.Count, "nenhum slide criado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    KillPres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Revisao final (M2, M8): Excel criado nesta chamada com core API 3 mas o core RECUSA (inserir ou editar): Quit, nenhum Excel vazio; deu certo:
' fica (sem Quit). Apresentacao Nothing (Modo de Exibicao Protegido): "Abra uma apresentacao".
Public Sub Test_PptTqq_CreatedExcel_QuitOnRefusal_KeptOnSuccess()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewSavedPres("tqq-ppam-g4e")
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_InsertChartIn Nothing, Nothing, "stacked"
    AssertEqual PPT_MSG_NO_PRES, PptTqq_LastMsg(), "sem apresentacao ativa"
    Set f = NewFakeXl(3): f.InsertResult = "A pasta de dados está aberta por outra pessoa."
    PptTqq_SetExcelForTest f, "", True
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual "A pasta de dados está aberta por outra pessoa.", PptTqq_LastMsg()
    AssertEqual 1, f.QuitCount, "recusa do core: o Excel criado agora recebeu Quit"
    Set f = NewFakeXl(3): f.InsertAddsShape = True
    PptTqq_SetExcelForTest f, "", True
    PptTqq_InsertChartIn p, p.Slides(1), "stacked"
    AssertEqual "edit:Pasta - Excel", PptTqq_LastMsg()
    AssertEqual 0, f.QuitCount, "deu certo: o Excel fica com o usuario"
    c.Add p.Slides(1).Shapes(1)
    Set f = NewFakeXl(3): f.EditResult = "!Gráfico não encontrado na pasta de dados (apagado ou recriado)."
    PptTqq_SetExcelForTest f, "", True
    PptTqq_EditShapes p, c
    AssertEqual "Gráfico não encontrado na pasta de dados (apagado ou recriado).", PptTqq_LastMsg()
    AssertEqual 1, f.QuitCount, "Editar dados recusado num Excel criado agora: Quit"
    Set f = NewFakeXl(3): f.EditResult = "!x"
    PptTqq_SetExcelForTest f, "", False
    PptTqq_EditShapes p, c
    AssertEqual 0, f.QuitCount, "o Excel do usuario nunca recebe Quit"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set c = Nothing
    KillPres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Revisao final (I1, M3): na abertura (probeRemote False) um caminho antigo REMOTO nunca e testado (reaponta para o candidato local sem tocar a
' rede); um caminho antigo local e testado; um vinculo cujo arquivo nao termina em " - dados ThinQQ.xlsx" nunca e reapontado (nem testado).
' A forma do teste nao e um grafico vinculado: a tentativa de gravar o vinculo fica no registro (prova que o reapontar foi tentado).
Public Sub Test_PptTqq_RelinkData_NoRemoteProbeOnOpen_OnlyDataBooks()
    Dim p As Object, s As Object, f As String, d As String, o As String, k As Integer, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    f = Environ$("TEMP") & "\tqq-relink3.pptx": p.SaveAs f
    d = Environ$("TEMP") & "\tqq-relink3 - dados ThinQQ.xlsx"
    o = Environ$("TEMP") & "\Orcamento-tqq.xlsx"
    k = FreeFile: Open d For Output As #k: Close #k
    k = FreeFile: Open o For Output As #k: Close #k
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, "\\tqq-host-que-nao-existe\s\tqq-relink3 - dados ThinQQ.xlsx"
    p.Tags.Add PPT_TAG_DATABOOK, "\\tqq-host-que-nao-existe\s\tqq-relink3 - dados ThinQQ.xlsx"
    PptTqq_ResetTestHooks
    AssertEqual 0, PptTqq_RelinkData(p, False)
    AssertEqual 0, PptTqq_RelinkProbesForTest(), "abertura: o caminho remoto antigo nao foi testado"
    AssertTrue InStr(PptTqq_LastLog(), "PptTqq_RelinkData:") = 1, "reapontar para o candidato local foi tentado: " & PptTqq_LastLog()
    PptTqq_ResetTestHooks
    s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\tqq-relink3 - dados ThinQQ.xlsx"
    AssertEqual 0, PptTqq_RelinkData(p, False)
    AssertEqual 1, PptTqq_RelinkProbesForTest(), "caminho antigo local: testado"
    PptTqq_ResetTestHooks
    s.Tags.Add PPT_TAG_BOOK, "C:\nao\existe\Orcamento-tqq.xlsx"
    AssertEqual 0, PptTqq_RelinkData(p, True)
    AssertEqual 0, PptTqq_RelinkProbesForTest(), "pasta que nao e de dados do ThinQQ: nem testada"
    AssertEqual "", PptTqq_LastLog(), "nem tentada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing
    ClosePres p
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then Kill f
    End If
    If Len(d) > 0 Then
        If Len(Dir$(d)) > 0 Then Kill d
    End If
    If Len(o) > 0 Then
        If Len(Dir$(o)) > 0 Then Kill o
    End If
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Core com a API 3: insere no slide dado (ou num slide novo no fim), abre os dados da forma nova (Core_PptEditData com a entrada dela) e
' mostra o aviso de primeira vez UMA vez por sessao; texto de recusa do core exibido como veio, sem abrir os dados.
Public Sub Test_PptTqq_InsertChartIn_Api3_InsertsAndOpensData()
    Dim p As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewSavedPres("tqq-ppam-g4b")
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(3): f.InsertAddsShape = True
    PptTqq_SetExcelForTest f, ""
    PptTqq_InsertChartIn p, p.Slides(1), "waterfall"
    AssertEqual "waterfall", f.LastKind
    AssertEqual 1, f.InsertSlideIndex
    AssertTrue InStr(f.Calls, "Core_PptInsertChart;Core_PptEditData;") > 0, f.Calls
    c.Add p.Slides(1).Shapes(1)
    AssertEqual PptTqq_EntriesOf(c), f.LastEntries, "Editar dados da forma nova"
    AssertEqual "hint;edit;", PptTqq_TrailForTest(), "aviso de primeira vez ANTES de abrir os dados (revisao I2)"
    AssertEqual "edit:Pasta - Excel", PptTqq_LastMsg()
    AssertTrue InStr(PptTqq_LastHintForTest(), "Gráfico criado. Clique em OK e digite os dados") = 1, PptTqq_LastHintForTest()
    AssertTrue InStr(PptTqq_LastHintForTest(), "'tqq-ppam-g4b - dados ThinQQ.xlsx'") > 0, "aviso cita o arquivo de dados: " & PptTqq_LastHintForTest()
    AssertTrue InStr(PptTqq_LastHintForTest(), "ThinQQ > Empacotar para enviar") > 0, "aviso cita o Empacotar (1.9.0): " & PptTqq_LastHintForTest()
    PptTqq_InsertChartIn p, Nothing, "pie"
    AssertEqual 2, p.Slides.Count, "sem slide: um slide novo no fim"
    AssertEqual 2, f.InsertSlideIndex
    AssertEqual "hint;edit;edit;", PptTqq_TrailForTest(), "aviso de primeira vez so uma vez por sessao"
    f.Calls = "": f.InsertResult = "Tipo de gráfico desconhecido: xyz"
    PptTqq_InsertChartIn p, p.Slides(1), "xyz"
    AssertEqual "Tipo de gráfico desconhecido: xyz", PptTqq_LastMsg()
    AssertEqual 0, InStr(f.Calls, "Core_PptEditData"), f.Calls
    PptTqq_InsertChartIn p, Nothing, "xyz"
    AssertEqual 2, p.Slides.Count, "recusa do core: o slide em branco criado aqui foi apagado (revisao M2)"
    f.InsertResult = "": f.FailOn = "Core_PptInsertChart"
    PptTqq_InsertChartIn p, Nothing, "line"
    AssertEqual 2, p.Slides.Count, "falha: o slide em branco criado aqui foi apagado"
    PptTqq_InsertChartIn p, p.Slides(1), "line"
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível criar o gráfico: falha simulada") = 1, PptTqq_LastMsg()
    AssertEqual 2, p.Slides.Count, "slide dado: nunca apagado"
    f.FailOn = ""
    PptTqq_InsertChartIn p, p.Slides(1), "line"
    AssertEqual "line", f.LastKind, "reentrada liberada depois da falha"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set c = Nothing
    KillPres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Editar dados: sem forma = selecionar; core sem a API 3 = atualizar; "!" do core = a mensagem sem o "!"; legenda = abre (no build Test
' registra "edit:" + legenda em vez do AppActivate).
Public Sub Test_PptTqq_EditShapes_Paths()
    Dim p As Object, s As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewSavedPres("tqq-ppam-g4c")
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "c1"
    PptTqq_ResetTestHooks
    PptTqq_EditShapes p, New Collection
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg()
    c.Add s
    Set f = NewFakeXl(2)
    PptTqq_SetExcelForTest f, ""
    PptTqq_EditShapes p, c
    AssertEqual PPT_MSG_OLD_CORE_DATA, PptTqq_LastMsg()
    Set f = NewFakeXl(3): f.EditResult = "!Pasta de dados não encontrada: C:\x.xlsx"
    PptTqq_SetExcelForTest f, ""
    PptTqq_EditShapes p, c
    AssertEqual "Pasta de dados não encontrada: C:\x.xlsx", PptTqq_LastMsg()
    f.EditResult = "Vendas - dados ThinQQ.xlsx - Excel"
    PptTqq_EditShapes p, c
    AssertEqual "edit:Vendas - dados ThinQQ.xlsx - Excel", PptTqq_LastMsg()
    AssertEqual PptTqq_EntriesOf(c), f.LastEntries
    f.FailOn = "Core_PptEditData"
    PptTqq_EditShapes p, c
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível abrir os dados do gráfico: falha simulada") = 1, PptTqq_LastMsg()
    f.FailOn = "": f.EditResult = "cap2"
    PptTqq_EditShapes p, c
    AssertEqual "edit:cap2", PptTqq_LastMsg(), "reentrada liberada depois da falha"
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_EditShapes p, c
    AssertEqual PPT_MSG_NO_EXCEL_DATA, PptTqq_LastMsg()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing: Set c = Nothing
    KillPres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Salvar a apresentacao: so com TQQ_DATABOOK e core com a API 3 chama Core_PptDataSave (nunca cria Excel; sem Excel: nada).
Public Sub Test_PptTqq_OnSave_OnlyWithDataBook()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(3)
    PptTqq_SetExcelForTest f, ""
    PptTqq_OnSave p
    AssertEqual "", f.Calls, "sem pasta de dados: nem chama o Excel"
    p.Tags.Add PPT_TAG_DATABOOK, "C:\x\a - dados ThinQQ.xlsx"
    PptTqq_OnSave p
    AssertTrue InStr(f.Calls, "Core_PptDataSave;") > 0, f.Calls
    Set f = NewFakeXl(2)
    PptTqq_SetExcelForTest f, ""
    PptTqq_OnSave p
    AssertEqual 0, InStr(f.Calls, "Core_PptDataSave"), "core sem a API 3: " & f.Calls
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_ClearLastMsg
    PptTqq_OnSave p
    AssertEqual "", PptTqq_LastMsg(), "evento de salvar nunca mostra mensagem"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Usado por tests/ps/Test-PptAddin.ps1 (ponta a ponta, passo c): "Editar dados" de uma forma dada, como o botao com ela selecionada.
Public Sub E2e_EditShape(shp As Object)
    Dim c As New Collection
    c.Add shp
    PptTqq_EditShapes shp.Parent.Parent, c
End Sub

' ---- 1.9.0 (Parte G-II, G8): grupo "Grafico do ThinQQ" ----

Public Sub Test_PptTqq_OptValueAndCap()
    Dim o As String
    o = "labels=1;totals=0;connectors=1;numfmt=pct0;legend=right;caps=labels,totals,cagr,"
    AssertEqual "1", PptTqq_OptValue(o, "labels")
    AssertEqual "pct0", PptTqq_OptValue(o, "numfmt")
    AssertEqual "", PptTqq_OptValue(o, "nada")
    AssertTrue PptTqq_OptCap(o, "cagr")
    AssertFalse PptTqq_OptCap(o, "diff")
    AssertFalse PptTqq_OptCap(o, "tal"), "nao confunde prefixo (totals)"
    AssertEqual "", PptTqq_OptValue("", "labels"), "forma sem TQQ_OPTS (colada por um core anterior)"
    AssertFalse PptTqq_OptCap("", "labels")
    AssertFalse PptTqq_OptCap(o, ""), "controle sem tag"
    AssertEqual "", PptTqq_OptValue(o, ""), "chave vazia"
    AssertEqual 4, PPT_API_CMD, "comandos pelo PowerPoint (1.9.0) pedem a API 4"
End Sub

Public Sub Test_PptTqq_PickTwo_TestHook()
    Dim a As Long, b As Long
    PptTqq_SetPickForTest 1, 3
    AssertTrue PptTqq_PickTwo("A" & Chr$(31) & "B" & Chr$(31) & "C", a, b)
    AssertEqual 1, a: AssertEqual 3, b
    PptTqq_SetPickForTest 3, 3
    AssertFalse PptTqq_PickTwo("A" & Chr$(31) & "B" & Chr$(31) & "C", a, b), "de = ate: recusado"
    PptTqq_SetPickForTest 1, 4
    AssertFalse PptTqq_PickTwo("A" & Chr$(31) & "B" & Chr$(31) & "C", a, b), "ate alem das categorias"
    PptTqq_SetPickForTest 1, 2
    AssertFalse PptTqq_PickTwo("A", a, b), "uma categoria so: nada a escolher"
    AssertEqual 0, a
    PptTqq_ResetTestHooks
    AssertFalse PptTqq_PickTwo("A" & Chr$(31) & "B", a, b), "gancho zerado: cancelado"
    AssertEqual 0, a: AssertEqual 0, b
End Sub

' Comandos (G8): sem selecao; sem Excel; Excel criado sem o core (Quit); core sem a API 4 (nenhum Core_PptCommand); API 4: comando, valor e
' entradas chegam ao core; texto do core exibido como veio; falha com texto proprio e reentrada liberada; o Excel do usuario nunca recebe Quit.
Public Sub Test_PptTqq_CommandShapes_Paths()
    Dim p As Object, s As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "c1"
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_CommandShapes p, New Collection, "toggle", "labels"
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "sem selecao"
    c.Add s
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_CommandShapes p, c, "toggle", "labels"
    AssertEqual PPT_MSG_NO_EXCEL_DATA, PptTqq_LastMsg(), "sem Excel"
    Set f = NewFakeXl(4): f.NoCore = True
    PptTqq_SetExcelForTest f, "", True
    PptTqq_CommandShapes p, c, "toggle", "labels"
    AssertEqual PPT_MSG_DATA_NOCORE, PptTqq_LastMsg(), "Excel sem o ThinQQ"
    AssertEqual 1, f.QuitCount, "o Excel criado recebeu Quit"
    Set f = NewFakeXl(3)
    PptTqq_SetExcelForTest f, ""
    PptTqq_CommandShapes p, c, "toggle", "labels"
    AssertEqual PPT_MSG_OLD_CORE_CMD, PptTqq_LastMsg(), "core 1.8.0 (API 3)"
    AssertEqual 0, InStr(f.Calls, "Core_PptCommand"), f.Calls
    Set f = NewFakeXl(4)
    PptTqq_SetExcelForTest f, ""
    PptTqq_ClearLastMsg
    PptTqq_CommandShapes p, c, "toggle", "labels"
    AssertEqual "toggle:labels", f.LastCmd
    AssertEqual PptTqq_EntriesOf(c), f.LastEntries
    AssertEqual "", PptTqq_LastMsg(), "deu certo: sem aviso"
    f.CommandResult = "Os gráficos selecionados não foram encontrados na apresentação (renomeados?): selecione-os de novo."
    PptTqq_CommandShapes Nothing, c, "numfmt", "pct0"
    AssertEqual "numfmt:pct0", f.LastCmd, "apresentacao Nothing: a da forma"
    AssertEqual f.CommandResult, PptTqq_LastMsg(), "texto do core exibido como veio"
    f.CommandResult = "": f.FailOn = "Core_PptCommand"
    PptTqq_CommandShapes p, c, "clear", ""
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível aplicar: falha simulada") = 1, PptTqq_LastMsg()
    f.FailOn = ""
    PptTqq_ClearLastMsg
    PptTqq_CommandShapes p, c, "clear", ""
    AssertEqual "clear:", f.LastCmd, "reentrada liberada depois da falha"
    AssertEqual "", PptTqq_LastMsg()
    AssertEqual 0, f.QuitCount, "o Excel do usuario nunca recebe Quit"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing: Set c = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' CAGR e diferenca (G8): categorias do core -> escolha de/ate -> "de,ate"; cancelado = nada (nem Core_PptCommand nem aviso); resposta
' invalida = aviso sem chamar o core; uma categoria so = aviso; sem categorias (pasta indisponivel) = o core e chamado com "0,0" e o motivo
' dele e exibido; nomes repetidos no slide ficam unicos antes de montar as entradas.
Public Sub Test_PptTqq_CommandShapes_Arrows()
    Dim p As Object, s As Object, s2 As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): s.Name = "tqq_y": s.Tags.Add PPT_TAG_CHART, "c1"
    Set s2 = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): s2.Name = "tqq_y": s2.Tags.Add PPT_TAG_CHART, "c2"
    c.Add s2
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(4)
    PptTqq_SetExcelForTest f, ""
    f.CatsResult = "2020" & Chr$(31) & "2021" & Chr$(31) & "2022"
    PptTqq_SetPickForTest 1, 3
    PptTqq_ClearLastMsg
    PptTqq_CommandShapes p, c, "cagr", ""
    AssertTrue InStr(f.Calls, "Core_PptCategories;Core_PptCommand;") > 0, f.Calls
    AssertEqual "cagr:1,3;", f.CmdLog
    AssertEqual "", PptTqq_LastMsg()
    AssertTrue StrComp(CStr(s.Name), CStr(s2.Name), vbTextCompare) <> 0, "nomes unicos no slide: " & s2.Name
    AssertEqual PptTqq_EntriesOf(c), f.LastEntries
    f.CmdLog = ""
    PptTqq_SetPickForTest 0, 0
    PptTqq_CommandShapes p, c, "diff", ""
    AssertEqual "", f.CmdLog, "cancelado: nada"
    AssertEqual "", PptTqq_LastMsg()
    PptTqq_SetPickForTest 3, 2
    PptTqq_CommandShapes p, c, "diff", ""
    AssertEqual "", f.CmdLog, "de > ate: o core nao e chamado"
    AssertEqual PPT_MSG_CATS, PptTqq_LastMsg()
    PptTqq_SetPickForTest 2, 3
    PptTqq_CommandShapes p, c, "diff", ""
    AssertEqual "diff:2,3;", f.CmdLog
    f.CmdLog = "": f.CatsResult = "2020"
    PptTqq_ClearLastMsg
    PptTqq_CommandShapes p, c, "diff", ""
    AssertEqual "", f.CmdLog, "uma categoria so"
    AssertEqual PPT_MSG_CATS, PptTqq_LastMsg()
    f.CatsResult = "": f.CommandResult = "Abra a pasta de origem no Excel para mudar este gráfico (o ThinQQ não altera pastas fechadas do usuário)."
    PptTqq_CommandShapes p, c, "cagr", ""
    AssertEqual "cagr:0,0;", f.CmdLog, "sem categorias: o core diz o motivo"
    AssertEqual f.CommandResult, PptTqq_LastMsg()
    f.CmdLog = "": f.CommandResult = ""
    PptTqq_CommandShapes p, c, "cagr", ""
    AssertEqual PPT_MSG_CATS, PptTqq_LastMsg(), "sem categorias e sem motivo: o aviso das categorias"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing: Set s2 = Nothing: Set c = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Faixa sem selecao: o grupo "Grafico do ThinQQ" fica solto e desligado, e cada clique avisa para selecionar (controle sem tag: nada quebra).
Public Sub Test_PptRx_ChartCallbacks_NoSelection()
    Dim v As Variant
    PptTqq_ResetTestHooks
    AssertEqual "", PptTqq_SelectedOpts()
    v = True: PptRx_GetOpt Nothing, v
    AssertFalse CBool(v)
    v = True: PptRx_GetOptEnabled Nothing, v
    AssertFalse CBool(v)
    v = "x": PptRx_GetOptItem Nothing, v
    AssertEqual "", CStr(v)
    PptTqq_ClearLastMsg
    PptRx_OnToggleOpt Nothing, True
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "alternancia"
    PptTqq_ClearLastMsg
    PptRx_OnOptItem Nothing, "numfmtpct0", 3
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "lista"
    PptTqq_ClearLastMsg
    PptRx_OnArrow Nothing
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "seta"
    PptTqq_ClearLastMsg
    PptRx_OnClearAnn Nothing
    AssertEqual PPT_MSG_SELECT, PptTqq_LastMsg(), "remover setas"
    PptRx_InvalidateElem
End Sub

' Revisao final G-II: (M2) a alternancia manda o estado do botao ("labels:1"), nao "inverta"; (M5) lista com valor fora dos itens (legenda
' "lineend" do Excel) ou sem selecao: o 1o item, sempre um id que existe; (M3) sem Excel e na falha a faixa e invalidada (o botao volta).
Public Sub Test_PptTqq_ChartGroup_FinalReview()
    Dim p As Object, s As Object, f As TestXlFake, en As Long, ed As String, c As New Collection
    AssertEqual "labels:1", PptRx_ToggleArg("labels", True)
    AssertEqual "totals:0", PptRx_ToggleArg("totals", False)
    AssertEqual "legendright", PptTqq_OptItemId("legend", "right")
    AssertEqual "legendnone", PptTqq_OptItemId("legend", "none")
    AssertEqual "legendright", PptTqq_OptItemId("legend", "lineend"), "legenda do Excel fora dos itens: o 1o item"
    AssertEqual "legendright", PptTqq_OptItemId("legend", "inbar")
    AssertEqual "legendright", PptTqq_OptItemId("legend", ""), "sem selecao"
    AssertEqual "numfmtpct1", PptTqq_OptItemId("numfmt", "pct1")
    AssertEqual "numfmtint", PptTqq_OptItemId("numfmt", "xyz")
    AssertEqual "", PptTqq_OptItemId("", "x"), "controle sem tag"
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Tags.Add PPT_TAG_CHART, "c1"
    c.Add s
    PptTqq_ResetTestHooks
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_CommandShapes p, c, "toggle", "labels:1"
    AssertEqual PPT_MSG_NO_EXCEL_DATA, PptTqq_LastMsg()
    AssertEqual 1, PptTqq_InvalidatesForTest(), "sem Excel: faixa invalidada"
    Set f = NewFakeXl(4): f.FailOn = "Core_PptCommand"
    PptTqq_SetExcelForTest f, ""
    PptTqq_CommandShapes p, c, "toggle", "labels:1"
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível aplicar: ") = 1, PptTqq_LastMsg()
    AssertEqual 2, PptTqq_InvalidatesForTest(), "falha: faixa invalidada"
    f.FailOn = ""
    PptTqq_CommandShapes p, c, "toggle", "labels:0"
    AssertEqual "toggle:labels:0", f.LastCmd, "o valor do botao chega ao core"
    AssertEqual 3, PptTqq_InvalidatesForTest(), "sucesso: faixa invalidada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing: Set c = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' ---- 1.9.0 (Parte G-II, G9): "Empacotar para enviar" ----

Public Sub Test_PptTqq_Package_CopiesBoth()
    Dim p As Object, f As String, d As String, dest As String, k As Integer, en As Long, ed As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set p = NewPres()
    On Error GoTo Cleanup
    f = Environ$("TEMP") & "\tqq-pack.pptx": p.SaveAs f
    d = Environ$("TEMP") & "\tqq-pack - dados ThinQQ.xlsx"
    k = FreeFile: Open d For Output As #k: Print #k, "x": Close #k
    p.Tags.Add PPT_TAG_DATABOOK, d
    PptTqq_ResetTestHooks
    dest = PptTqq_Package(p)
    AssertEqual Environ$("TEMP") & "\tqq-pack (ThinQQ)", dest
    AssertTrue fso.FileExists(dest & "\tqq-pack.pptx"), "apresentacao copiada"
    AssertTrue fso.FileExists(dest & "\tqq-pack - dados ThinQQ.xlsx"), "pasta de dados copiada"
    AssertTrue InStr(PptTqq_LastMsg(), "Pronto: ") = 1, PptTqq_LastMsg()
    dest = PptTqq_Package(p)
    AssertEqual Environ$("TEMP") & "\tqq-pack (ThinQQ)", dest, "de novo: a mesma pasta, arquivos substituidos"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    On Error Resume Next
    fso.DeleteFolder Environ$("TEMP") & "\tqq-pack (ThinQQ)", True
    Kill f: Kill d
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Empacotar junta as pastas de dados DISTINTAS das formas TQQ_DATA (revisao G4, M8) e a marca, mas so as que estao NA PASTA da apresentacao
' (revisao final G-II, I1): a de outra pasta (Salvar como para outra pasta, ou um deck de terceiros apontando para a pasta de dados de OUTRA
' apresentacao) nao e copiada nem sondada no disco e e citada em linha propria; idem a de rede; nunca uma pasta do usuario (nome sem
' " - dados ThinQQ.xlsx"); a ausente e citada; o "Pronto" lista cada arquivo copiado; a copiada que esta aberta no Excel com alteracoes nao
' salvas e citada (M8). Apresentacao nunca salva: o aviso de salvar, nada criado. Sem apresentacao: o aviso de abrir uma.
Public Sub Test_PptTqq_Package_GathersDataBooks()
    Dim p As Object, s As Object, f As String, d1 As String, d2 As String, u As String, dest As String, k As Integer, en As Long, ed As String
    Dim fso As Object, base As String, m As String, xf As TestXlFake, wf As TestXlFake
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = Environ$("TEMP") & "\tqq-pack2-" & Format$(Now, "hhnnss")
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    AssertEqual "", PptTqq_Package(p), "nunca salva"
    AssertEqual PPT_MSG_PACK_SAVE_FIRST, PptTqq_LastMsg()
    AssertEqual "", PptTqq_Package(Nothing)
    AssertEqual PPT_MSG_NO_PRES, PptTqq_LastMsg()
    fso.CreateFolder base
    fso.CreateFolder base & "\A"
    f = base & "\B.pptx": p.SaveAs f
    d1 = base & "\A\A - dados ThinQQ.xlsx"
    d2 = base & "\B - dados ThinQQ.xlsx"
    u = base & "\Orcamento.xlsx"
    k = FreeFile: Open d1 For Output As #k: Print #k, "a": Close #k
    k = FreeFile: Open d2 For Output As #k: Print #k, "b": Close #k
    k = FreeFile: Open u For Output As #k: Print #k, "u": Close #k
    p.Tags.Add PPT_TAG_DATABOOK, d2
    Set s = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, d1
    Set s = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, UCase$(d1)
    Set s = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, d2
    Set s = p.Slides(1).Shapes.AddShape(1, 190, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, u
    Set s = p.Slides(1).Shapes.AddShape(1, 250, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, base & "\Sumiu - dados ThinQQ.xlsx"
    Set s = p.Slides(1).Shapes.AddShape(1, 310, 10, 50, 50): s.Tags.Add PPT_TAG_DATA, "1": s.Tags.Add PPT_TAG_BOOK, "\\tqq-host-que-nao-existe\s\Rede - dados ThinQQ.xlsx"
    Set s = p.Slides(1).Shapes.AddShape(1, 370, 10, 50, 50): s.Tags.Add PPT_TAG_CHART, "c1": s.Tags.Add PPT_TAG_BOOK, base & "\Colado - dados ThinQQ.xlsx"
    Set s = Nothing
    Set xf = NewFakeXl(3)                                  ' Excel aberto: a pasta de dados B aberta com alteracoes nao salvas
    Set wf = New TestXlFake: wf.FullName = d2: wf.Saved = False
    xf.Workbooks.Add wf
    PptTqq_SetExcelForTest xf, ""
    dest = PptTqq_Package(p)
    AssertEqual base & "\B (ThinQQ)", dest
    AssertTrue fso.FileExists(dest & "\B.pptx"), "apresentacao"
    AssertTrue fso.FileExists(dest & "\B - dados ThinQQ.xlsx"), "pasta de dados marcada, na pasta da apresentacao"
    AssertFalse fso.FileExists(dest & "\A - dados ThinQQ.xlsx"), "pasta de dados de OUTRA pasta: nao copiada"
    AssertFalse fso.FileExists(dest & "\Orcamento.xlsx"), "pasta do usuario: nunca"
    AssertEqual 2, fso.GetFolder(dest).Files.Count, "so a apresentacao e a pasta de dados ao lado"
    AssertEqual 2, PptTqq_PackProbesForTest(), "so as da pasta da apresentacao foram conferidas no disco (B e Sumiu); a de outra pasta e a de rede nao"
    m = PptTqq_LastMsg()
    AssertTrue InStr(m, "Pronto: a apresentação e 1 pasta(s) de dados do ThinQQ foram copiadas para '") = 1, m
    AssertTrue InStr(m, "': B.pptx, B - dados ThinQQ.xlsx. Envie a pasta ") > 0, "Pronto lista cada arquivo copiado: " & m
    AssertTrue InStr(m, vbLf & "Não copiada(s): fora da pasta da apresentação: A - dados ThinQQ.xlsx, Rede - dados ThinQQ.xlsx") > 0, "de fora citadas em linha propria: " & m
    AssertTrue InStr(m, "Não encontrada(s) (os gráficos ficam com os últimos dados): Sumiu - dados ThinQQ.xlsx") > 0, "ausente citada: " & m
    AssertTrue InStr(m, "aberta(s) no Excel com alterações não salvas") > 0 And InStr(m, "levá-las): B - dados ThinQQ.xlsx") > 0, "aberta e alterada citada (M8): " & m
    AssertEqual 0, InStr(m, "Orcamento"), "pasta do usuario nem citada: " & m
    AssertEqual 0, InStr(m, "Colado - dados"), "forma sem TQQ_DATA (colada do Excel) nao entra: " & m
    AssertTrue InStr(m, "Gráficos colados do Excel continuam ligados") > 0, m
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing
    ClosePres p
    On Error Resume Next
    fso.DeleteFolder base, True
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' ---- 1.9.0 (Parte E, E3): Texto vinculado ao Excel ----

' Sem janela (projeto de teste): nada selecionado. O callback avisa para selecionar uma caixa de texto.
Public Sub Test_PptTqq_SelectionTextRange_NoWindow()
    Dim s As Object, a As Long, b As Long, g As Boolean
    PptTqq_ResetTestHooks
    AssertFalse PptTqq_SelectionTextRange(s, a, b, g)
    AssertTrue s Is Nothing
    AssertFalse g, "sem janela: nao e grupo"
    AssertFalse PptTqq_SelInGroup(New Collection), "objeto sem HasChildShapeRange: nao e grupo"
    PptTqq_ClearLastMsg
    PptRx_OnLinkText Nothing
    AssertEqual PPTTEXT_MSG_NOTEXT_PPT, PptTqq_LastMsg()
    AssertEqual 5, PPT_API_TEXT
    AssertEqual 1, PPT_REQUIRED_API, "a API exigida ao abrir nao subiu"
End Sub

' Caminhos: sem Excel aberto ou sem o ThinQQ nele (nunca cria um Excel; o do usuario nunca recebe Quit), core antigo (sem chamar), sucesso
' (o trecho chega ao core; aviso com o lembrete de salvar), recusa do core (como veio), cortado (com o lembrete), falha (texto proprio;
' reentrada liberada).
Public Sub Test_PptTqq_LinkText_Paths()
    Dim p As Object, s As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddTextbox(1, 10, 10, 200, 50)
    s.TextFrame.TextRange.Text = "n=X"
    PptTqq_ResetTestHooks
    PptTqq_ClearLastMsg
    PptTqq_LinkText Nothing, 1, 0
    AssertEqual PPTTEXT_MSG_NOTEXT_PPT, PptTqq_LastMsg(), "sem forma"
    PptTqq_SetExcelForTest Nothing, ""
    PptTqq_LinkText s, 3, 1
    AssertEqual PPT_MSG_OPEN_EXCEL, PptTqq_LastMsg(), "sem Excel aberto: nunca cria um"
    Set f = NewFakeXl(5): f.NoCore = True
    PptTqq_SetExcelForTest f, "", True                          ' mesmo marcado como criavel: o vincular nunca cria nem abre o core
    PptTqq_LinkText s, 3, 1
    AssertEqual PPT_MSG_OPEN_EXCEL, PptTqq_LastMsg(), "Excel sem o ThinQQ"
    AssertEqual 0, f.QuitCount, "o Excel do usuario nunca recebe Quit"
    AssertEqual 0, InStr(f.Calls, "Core_PptLinkText"), f.Calls
    Set f = NewFakeXl(4)
    PptTqq_SetExcelForTest f, ""
    PptTqq_LinkText s, 3, 1
    AssertEqual PPT_MSG_OLD_CORE_TEXT, PptTqq_LastMsg(), "core 1.9.0 sem a API 5"
    AssertEqual 0, InStr(f.Calls, "Core_PptLinkText"), f.Calls
    Set f = NewFakeXl(5)
    PptTqq_SetExcelForTest f, ""
    PptTqq_LinkText s, 3, 1
    AssertEqual "3:1", f.LastLink
    AssertEqual CStr(s.Name), CStr(f.LinkShape.Name)
    AssertEqual PPT_MSG_TEXT_OK & " " & PPT_MSG_TEXT_SAVE, PptTqq_LastMsg(), "vinculado: com o lembrete de salvar"
    f.LinkResult = "Salve a pasta de trabalho antes de vincular: o vínculo aponta para o arquivo salvo."
    PptTqq_LinkText s, 1, 0
    AssertEqual f.LinkResult, PptTqq_LastMsg(), "recusa do core como veio"
    AssertEqual "1:0", f.LastLink, "cursor"
    f.LinkResult = "Texto vinculado à célula Vendas.xlsx!Dados!B2."
    PptTqq_LinkText s, 3, 1
    AssertEqual f.LinkResult & " " & PPT_MSG_TEXT_SAVE, PptTqq_LastMsg(), "core 1.9.0: a celula e a pasta, com o lembrete (M9)"
    f.LinkResult = "Texto vinculado à célula Vendas.xlsx!Dados!B2. O valor da célula tem mais de 255 caracteres: o slide mostra só os primeiros 255."
    PptTqq_LinkText s, 1, 3
    AssertEqual f.LinkResult & " " & PPT_MSG_TEXT_SAVE, PptTqq_LastMsg(), "cortado: com o lembrete"
    f.LinkResult = "": f.FailOn = "Core_PptLinkText"
    PptTqq_LinkText s, 1, 3
    AssertTrue InStr(PptTqq_LastMsg(), "Não foi possível vincular o texto: falha simulada") = 1, PptTqq_LastMsg()
    f.FailOn = ""
    PptTqq_ClearLastMsg
    PptTqq_LinkText s, 2, 1
    AssertEqual "2:1", f.LastLink, "reentrada liberada depois da falha"
    AssertEqual 0, f.QuitCount
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    Set s = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' 1.9.0: uma apresentacao so com textos vinculados (TQQ_TEXT = "1") conta para "Ao abrir" e "Atualizar graficos" (antes so TQQ_CHART);
' TQQ_TEXT diferente de "1" nao conta.
Public Sub Test_PptTqq_CountTagged_LinkedText()
    Dim p As Object, s As Object, srcs As New Collection, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set s = p.Slides(1).Shapes.AddTextbox(1, 10, 10, 200, 50): s.Tags.Add PPT_TAG_TEXT, "1": s.Tags.Add PPT_TAG_BOOK, "C:\Dados\T.xlsx"
    Set s = p.Slides(1).Shapes.AddTextbox(1, 10, 80, 200, 50): s.Tags.Add PPT_TAG_TEXT, "0": s.Tags.Add PPT_TAG_BOOK, "C:\Dados\U.xlsx"
    AssertEqual 1, PptTqq_CountTagged(p, srcs)
    AssertEqual 1, srcs.Count
    AssertEqual "C:\Dados\T.xlsx", CStr(srcs(1)), "a pasta vem de TQQ_BOOK (texto nao tem vinculo nativo)"
Cleanup:
    en = Err.Number: ed = Err.Description
    Set s = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' 1.10.0: tabela vinculada (TQQ_TABLE = "1") tambem conta, com a pasta de TQQ_BOOK; TQQ_TABLE diferente de "1" nao conta.
Public Sub Test_PptTqq_CountTagged_LinkedTable()
    Dim p As Object, s As Object, srcs As New Collection, en As Long, ed As String
    Set p = NewPres()
    On Error GoTo Cleanup
    AssertEqual "TQQ_TABLE", PPT_TAG_TABLE
    Set s = p.Slides(1).Shapes.AddTable(2, 2, 10, 10, 200, 50): s.Tags.Add PPT_TAG_TABLE, "1": s.Tags.Add PPT_TAG_BOOK, "C:\Dados\Tab.xlsx"
    Set s = p.Slides(1).Shapes.AddTable(2, 2, 10, 80, 200, 50): s.Tags.Add PPT_TAG_TABLE, "x": s.Tags.Add PPT_TAG_BOOK, "C:\Dados\V.xlsx"
    AssertEqual 1, PptTqq_CountTagged(p, srcs)
    AssertEqual 1, srcs.Count
    AssertEqual "C:\Dados\Tab.xlsx", CStr(srcs(1)), "a pasta vem de TQQ_BOOK"
Cleanup:
    en = Err.Number: ed = Err.Description
    Set s = Nothing
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' ---- 1.11.0 (U2): bandeira, todos os tipos, silencio no sucesso, abertura ----

Public Sub Test_PptTqq_IsPlainSuccess()
    AssertTrue PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 1.")
    AssertTrue PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 0. Sem mudança: 2.")
    AssertFalse PptTqq_IsPlainSuccess("Gráficos do ThinQQ atualizados: 0." & vbLf & "Pasta de origem não encontrada ou sem acesso: a.xlsx.")
    AssertFalse PptTqq_IsPlainSuccess(PPT_MSG_BUSY)
    AssertFalse PptTqq_IsPlainSuccess("")
End Sub

Public Sub Test_PptTqq_ShowResult_SilentOnSuccess()
    PptTqq_ClearLastMsg
    PptTqq_ShowResult "Gráficos do ThinQQ atualizados: 1."
    AssertEqual "", PptTqq_LastMsg(), "sucesso puro: sem caixa (como o think-cell)"
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptTqq_LastResult()
    PptTqq_ShowResult "Gráficos do ThinQQ atualizados: 0." & vbLf & "Falha ao atualizar: 1 (veja o log do ThinQQ)."
    AssertTrue InStr(PptTqq_LastMsg(), "Falha ao atualizar") > 0, "problema: caixa"
    PptTqq_ShowResult "!" & PPT_MSG_BUSY
    AssertEqual PPT_MSG_BUSY, PptTqq_LastMsg(), "'!' = aviso de exclamacao, sem o '!'"
End Sub

Public Sub Test_PptTqq_LinkedShapes_AndAnyNew()
    Dim p As Object, sl As Object, a As Object, b As Object, c As Object, d As Object, col As New Collection, t As Single, i As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    Set sl = p.Slides(1)
    Set a = sl.Shapes.AddShape(1, 10, 10, 50, 50): a.Tags.Add PPT_TAG_CHART, "abc"
    Set b = sl.Shapes.AddShape(1, 70, 10, 50, 50): b.Tags.Add PPT_TAG_TEXT, "1"
    Set c = sl.Shapes.AddShape(1, 130, 10, 50, 50): c.Tags.Add PPT_TAG_TABLE, "1"
    Set d = sl.Shapes.AddShape(1, 190, 10, 50, 50)
    AssertTrue PptTqq_IsLinkedShape(a) And PptTqq_IsLinkedShape(b) And PptTqq_IsLinkedShape(c)
    AssertFalse PptTqq_IsLinkedShape(d)
    col.Add a: col.Add b
    AssertFalse PptTqq_AnyNew(col)
    b.Tags.Add PPT_TAG_STATE, "new"
    AssertTrue PptTqq_AnyNew(col), "texto vinculado com dados novos"
    b.Tags.Add PPT_TAG_STATE, "x"
    AssertFalse PptTqq_AnyNew(col), "so 'new' e bandeira"
    AssertEqual 0, PptTqq_SelectedLinked().Count, "sem janela: nada selecionado"
    For i = 1 To 196
        sl.Shapes.AddShape 1, 1, 1, 2, 2
    Next
    t = Timer
    For i = 1 To 200
        PptTqq_AnyNew col
    Next
    PptTqq_Log "P-U2: AnyNew x200 = " & Format$((Timer - t) * 1000, "0") & " ms"
    AssertTrue (Timer - t) < 1, "P-U2: 200 leituras acima de 5 ms cada: " & PptTqq_LastLog()
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Abertura (1.11.0): bandeiras gravadas no arquivo saem, TQQ_LINKIDX e reconstruido com os elementos vinculados (formato do core), Saved fica.
Public Sub Test_PptTqq_PrepareOnOpen_ClearsFlags_BuildsIndex_KeepsSaved()
    Dim p As Object, a As Object, b As Object, c As Object, idx As String
    Set p = NewPres()
    On Error GoTo Cleanup
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_STATE, "new"
    a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "txt_b": b.Tags.Add PPT_TAG_TEXT, "1": b.Tags.Add PPT_TAG_BOOK, "C:\x\b.xlsx"
    Set c = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50)
    p.Saved = -1
    AssertEqual 2, PptTqq_PrepareOnOpen(p)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_STATE), "bandeira do arquivo sai"
    idx = PptTqq_Tag(p, PPT_TAG_LINKIDX)
    AssertEqual PptTqq_IdxEntry(CLng(p.Slides(1).SlideID), "tqq_a", "C:\x\a.xlsx") & vbLf & PptTqq_IdxEntry(CLng(p.Slides(1).SlideID), "txt_b", "C:\x\b.xlsx"), idx
    AssertTrue CLng(p.Saved) <> 0, "Saved preservado"
    ' a abertura (HandleOpen) tambem limpa: bandeira gravada de novo sai na primeira vista
    a.Tags.Add PPT_TAG_STATE, "new"
    PptTqq_ResetTestHooks
    PptTqq_HandleOpen p
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_STATE), "HandleOpen chama PptTqq_PrepareOnOpen"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptTqq_LinkStates_ApiAndCore()
    Dim p As Object, f As TestXlFake, en As Long, ed As String
    Set p = NewTagged()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(5)
    PptTqq_SetExcelForTest f, ""
    AssertEqual "!" & PPT_MSG_OLD_CORE_STATES, PptTqq_LinkStates(p, "", False)
    Set f = NewFakeXl(6)
    PptTqq_SetExcelForTest f, ""
    AssertEqual "1|a|b|chart:stacked|manual|ok", PptTqq_LinkStates(p, "E", True)
    AssertEqual "E", f.LastEntries
    AssertTrue CBool(f.LastVerify)
    f.NoCore = True
    AssertEqual "!" & PPT_MSG_NOCORE_STATES, PptTqq_LinkStates(p, "", False)
    PptTqq_SetExcelForTest Nothing, ""
    AssertEqual "!" & PPT_MSG_NO_EXCEL, PptTqq_LinkStates(p, "", False)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptRx_LinkedCallbacks_NoSelection()
    Dim v As Variant
    v = True: PptRx_GetLinkedEnabled Nothing, v
    AssertFalse CBool(v)
    v = True: PptRx_GetNewVisible Nothing, v
    AssertFalse CBool(v), "sem selecao: sem bandeira"
    v = False: PptRx_GetUpdVisible Nothing, v
    AssertTrue CBool(v), "sem selecao: o Atualizar normal aparece (desligado)"
    AssertEqual "never", PPT_DEFAULT_OPENMODE
    AssertEqual "c", PptRx_ItemPrefix("cddNumFmt"), "aba contextual: itens com prefixo c"
    AssertEqual "", PptRx_ItemPrefix("ddPptNumFmt")
End Sub

' ---- 1.11.0 (U3): dialogo Vinculos (modPptLinks; o formulario so e conferido pela legenda) ----

Public Sub Test_PptForms_LinksCaption()
    Dim f As Object
    Set f = VBA.UserForms.Add("frmTqqLinks")
    AssertEqual "Vínculos do ThinQQ", f.Caption
    AssertEqual 5, CLng(f.Controls("lstLinks").ColumnCount)
    Unload f
    AssertEqual 0, PptLinks_FormCount(), "descarregado"
End Sub

Public Sub Test_PptLinks_Labels()
    AssertEqual "Dados novos", PptLinks_StateLabel("new")
    AssertEqual "Atualizado", PptLinks_StateLabel("ok")
    AssertEqual "Pasta não encontrada", PptLinks_StateLabel("missing")
    AssertEqual "Não encontrado na pasta", PptLinks_StateLabel("nosource")
    AssertEqual "Sem vínculo", PptLinks_StateLabel("unlinked")
    AssertEqual "Não verificado (pasta fechada)", PptLinks_StateLabel("unknown")
    AssertEqual "Não verificado", PptLinks_StateLabel("")
    AssertEqual "Gráfico", PptLinks_KindLabel("chart")
    AssertEqual "Texto", PptLinks_KindLabel("text")
    AssertEqual "Tabela", PptLinks_KindLabel("table")
    AssertEqual "Automático", PptLinks_ModeLabel("auto")
    AssertEqual "Manual", PptLinks_ModeLabel("manual")
    AssertEqual "b.xlsx", PptLinks_FileName("C:\a\b.xlsx")
    AssertEqual "c.xlsx", PptLinks_FileName("https://x/y/c.xlsx")
    AssertEqual 0, PptLinks_OpenModeIndex("never")
    AssertEqual 1, PptLinks_OpenModeIndex("ask")
    AssertEqual 2, PptLinks_OpenModeIndex("always")
    AssertEqual "never", PptLinks_OpenModeOf(0)
    AssertEqual "always", PptLinks_OpenModeOf(2)
End Sub

Public Sub Test_PptLinks_CollectFilterAndCache()
    Dim p As Object, a As Object, b As Object, c As Object, rows As Collection, sid As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    PptLinks_ResetTestHooks
    sid = CLng(p.Slides(1).SlideID)
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    a.Tags.Add PPT_TAG_STATE, "new"
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "txt_b": b.Tags.Add PPT_TAG_TEXT, "1": b.Tags.Add PPT_TAG_BOOK, "C:\x\b.xlsx"
    b.Tags.Add PPT_TAG_AUTO, "1"
    Set c = p.Slides(1).Shapes.AddShape(1, 130, 10, 50, 50)
    Set rows = PptLinks_Collect(p)
    AssertEqual 2, rows.Count
    AssertEqual "tqq_a", CStr(rows(1)(2))
    AssertEqual "chart", CStr(rows(1)(3))
    AssertEqual "manual", CStr(rows(1)(4))
    AssertEqual "new", CStr(rows(1)(5))
    AssertEqual "auto", CStr(rows(2)(4))
    AssertEqual "", CStr(rows(2)(5)), "sem verificar"
    AssertEqual 1, PptLinks_Filter(rows, 1).Count, "dados novos"
    AssertEqual 1, PptLinks_Filter(rows, 3).Count, "automaticos"
    AssertEqual 0, PptLinks_Filter(rows, 4).Count, "com problema"
    AssertEqual 1, PptLinks_Remember(sid & "|txt_b|C:\x\b.xlsx|text|auto|missing")
    Set rows = PptLinks_Collect(p)
    AssertEqual "missing", CStr(rows(2)(5)), "estado do core lembrado na sessao"
    AssertEqual 1, PptLinks_Filter(rows, 4).Count
    AssertEqual 0, PptLinks_Remember("lixo" & vbLf & "1|2"), "linha invalida ignorada"
    AssertTrue InStr(PptLinks_HeadText(rows), "2 elemento(s)") > 0, PptLinks_HeadText(rows)
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Acoes sem Excel: manual/automatico gravam a marca e o indice; desvincular tira so as etiquetas TQQ_* (a forma fica); sem selecao: aviso.
Public Sub Test_PptLinks_Do_ModesAndUnlink_NoExcel()
    Dim p As Object, a As Object, rows As Collection, msg As String
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc": a.Tags.Add PPT_TAG_BOOK, "C:\x\a.xlsx"
    Set rows = PptLinks_Collect(p)
    AssertEqual PPTLINKS_MSG_SELECT, PptLinks_Do(p, "update", New Collection)
    msg = PptLinks_Do(p, "manual", rows)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_AUTO)
    msg = PptLinks_Do(p, "auto", rows)
    AssertEqual "1", PptTqq_Tag(a, PPT_TAG_AUTO)
    AssertTrue InStr(PptTqq_Tag(p, PPT_TAG_INDEX), "tqq_a") > 0, "no indice dos automaticos"
    AssertTrue Len(msg) > 0, "sem Excel: a dica de abrir o Excel"
    msg = PptLinks_Do(p, "unlink", rows)
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_CHART), "desvinculado"
    AssertEqual "", PptTqq_Tag(a, PPT_TAG_AUTO)
    AssertEqual "", PptTqq_Tag(p, PPT_TAG_INDEX), "fora do indice dos automaticos"
    AssertEqual "", PptTqq_Tag(p, PPT_TAG_LINKIDX), "fora do indice de todos"
    AssertEqual 1, p.Slides(1).Shapes.Count, "a forma continua no slide"
    AssertEqual "Desvinculados: 1.", msg
    AssertTrue InStr(PptTqq_LastMsg(), "vínculo do próprio PowerPoint") > 0, "a pergunta diz que o vinculo nativo fica: " & PptTqq_LastMsg()
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

Public Sub Test_PptLinks_Do_Check_UsesCoreAndRemembers()
    Dim p As Object, a As Object, f As TestXlFake, rows As Collection, msg As String, sid As Long
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    sid = CLng(p.Slides(1).SlideID)
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc"
    Set f = NewFakeXl(6)
    f.StatesResult = sid & "|tqq_a|C:\x\a.xlsx|chart:stacked|manual|nosource"
    PptTqq_SetExcelForTest f, ""
    msg = PptLinks_Do(p, "check", PptLinks_Collect(p))
    AssertTrue CBool(f.LastVerify), "Verificar abre pastas fechadas (somente leitura, no core)"
    AssertEqual "", f.LastEntries, "todos"
    AssertEqual "nosource", CStr(PptLinks_Collect(p)(1)(5))
    AssertTrue InStr(msg, "Verificados: 1") = 1, msg
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Atualizar pelo dialogo: Core_PptSyncShapes so com as linhas dadas (forcado) e, depois, os estados delas pelo core (sem abrir pasta fechada);
' o estado velho da linha (lembrado antes) e esquecido quando o core nao responde.
Public Sub Test_PptLinks_Do_Update_SyncsAndRefreshesStates()
    Dim p As Object, a As Object, b As Object, f As TestXlFake, msg As String, sid As Long, one As New Collection
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    sid = CLng(p.Slides(1).SlideID)
    Set a = p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a": a.Tags.Add PPT_TAG_CHART, "abc"
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "tqq_b": b.Tags.Add PPT_TAG_CHART, "def"
    PptLinks_Remember sid & "|tqq_a|C:\x\a.xlsx|chart:stacked|manual|missing" & vbLf & sid & "|tqq_b|C:\x\a.xlsx|chart:stacked|manual|missing", p
    one.Add PptLinks_Collect(p)(1)
    Set f = NewFakeXl(6)
    f.ShapesResult = "Gráficos do ThinQQ atualizados: 1."
    f.StatesResult = sid & "|tqq_a|C:\x\a.xlsx|chart:stacked|manual|ok"
    PptTqq_SetExcelForTest f, ""
    msg = PptLinks_Do(p, "update", one)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", msg
    AssertTrue InStr(f.Calls, "Core_PptSyncShapes;") > 0 And InStr(f.Calls, "Core_PptLinkStates;") > InStr(f.Calls, "Core_PptSyncShapes;"), f.Calls
    AssertFalse CBool(f.LastVerify), "depois de atualizar: estados sem abrir pasta fechada"
    AssertTrue InStr(f.LastEntries, "tqq_a") > 0 And InStr(f.LastEntries, "tqq_b") = 0, "so a linha dada: " & f.LastEntries
    AssertEqual "ok", CStr(PptLinks_Collect(p)(1)(5)), "estado novo do core"
    AssertEqual "missing", CStr(PptLinks_Collect(p)(2)(5)), "a outra linha fica como estava"
    PptTqq_SetExcelForTest Nothing, ""
    PptLinks_Do p, "update", one
    AssertEqual "", CStr(PptLinks_Collect(p)(1)(5)), "sem Excel: o estado velho da linha atualizada e esquecido (Nao verificado)"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Abrir no Excel vale para graficos: so texto/tabela na selecao = aviso na linha de status, sem chamar o Excel.
Public Sub Test_PptLinks_Do_Edit_TextOnly_NoExcel()
    Dim p As Object, b As Object, f As TestXlFake
    Set p = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    Set b = p.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "txt_b": b.Tags.Add PPT_TAG_TEXT, "1"
    Set f = NewFakeXl(6)
    PptTqq_SetExcelForTest f, ""
    AssertEqual PPTLINKS_MSG_EDIT_CHART, PptLinks_Do(p, "edit", PptLinks_Collect(p))
    AssertEqual "", f.Calls, "o Excel nao e chamado"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' O cache de estados e por apresentacao (duas apresentacoes tem o mesmo SlideID 256 e o mesmo nome de forma) e sai ao fechar.
Public Sub Test_PptLinks_StatesPerPresentation_ForgetOnClose()
    Dim p As Object, q As Object, sid As Long
    Set p = NewPres()
    Set q = NewPres()
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    p.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50).Name = "tqq_a"
    q.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50).Name = "tqq_a"
    p.Slides(1).Shapes("tqq_a").Tags.Add PPT_TAG_CHART, "abc"
    q.Slides(1).Shapes("tqq_a").Tags.Add PPT_TAG_CHART, "abc"
    sid = CLng(p.Slides(1).SlideID)
    AssertEqual sid, CLng(q.Slides(1).SlideID), "premissa: mesmo SlideID"
    AssertEqual 1, PptLinks_Remember(sid & "|tqq_a|C:\x\a.xlsx|chart:stacked|manual|missing", p)
    AssertEqual "missing", CStr(PptLinks_Collect(p)(1)(5))
    AssertEqual "", CStr(PptLinks_Collect(q)(1)(5)), "a outra apresentacao nao herda o estado"
    PptTqq_OnClose p
    AssertEqual "", CStr(PptLinks_Collect(p)(1)(5)), "fechar esquece os estados da apresentacao"
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    PptLinks_ResetTestHooks
    ClosePres p
    ClosePres q
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Sem piscar: a assinatura das linhas so muda quando algo visivel muda (modo, estado, nome, slide, origem). Chave da linha = SlideID|nome.
Public Sub Test_PptLinks_SignatureAndRowKey()
    Dim r1 As New Collection, r2 As New Collection
    r1.Add Array(1, 256, "tqq_A", "chart", "manual", "", "C:\a.xlsx")
    r2.Add Array(1, 256, "tqq_A", "chart", "manual", "", "C:\a.xlsx")
    AssertEqual PptLinks_Signature(r1), PptLinks_Signature(r2)
    Set r2 = New Collection
    r2.Add Array(1, 256, "tqq_A", "chart", "manual", "new", "C:\a.xlsx")
    AssertTrue PptLinks_Signature(r1) <> PptLinks_Signature(r2), "estado mudou: redesenha"
    AssertEqual "256|tqq_a", PptLinks_RowKey(r1(1))
    AssertEqual "", PptLinks_Signature(New Collection)
End Sub

' O evento de selecao nunca carrega o formulario (nada aberto: nada acontece) e Collect e barato: 220 formas, 20 vinculadas, 5 leituras < 1 s.
Public Sub Test_PptLinks_SelChange_NoForm_AndCollectCost()
    Dim p As Object, i As Long, t As Single, s As Object
    Set p = NewPres()
    On Error GoTo Cleanup
    PptLinks_ResetTestHooks
    PptLinks_OnSelChange
    AssertEqual 0, PptLinks_FormCount(), "a selecao nao abre o dialogo"
    For i = 1 To 220
        Set s = p.Slides(1).Shapes.AddShape(1, 1, 1, 2, 2)
        If i Mod 11 = 0 Then s.Tags.Add PPT_TAG_CHART, "c" & i
    Next
    t = Timer
    For i = 1 To 5
        AssertEqual 20, PptLinks_Collect(p).Count
    Next
    PptTqq_Log "U3: Collect 220 formas x5 = " & Format$((Timer - t) * 1000, "0") & " ms"
    AssertTrue (Timer - t) < 1, "Collect lento: " & PptTqq_LastLog()
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptLinks_ResetTestHooks
    ClosePres p
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Opcoes do core pelo dialogo (API 2): leitura so com o Excel ja aberto e o core; gravacao pelo Core_PptSetFlag.
Public Sub Test_PptTqq_CoreFlag_Fake()
    Dim f As TestXlFake, avail As Boolean
    On Error GoTo Cleanup
    PptTqq_ResetTestHooks
    Set f = NewFakeXl(6)
    PptTqq_SetExcelForTest f, ""
    AssertTrue PptTqq_CoreFlag("autoon", avail): AssertTrue avail
    f.AutoOn = False
    AssertFalse PptTqq_CoreFlag("autoon", avail): AssertTrue avail
    AssertEqual "Opção gravada no Excel.", PptTqq_SetCoreFlag("autonew", True)
    AssertEqual "autonew=1", f.LastFlag
    f.Api = 1
    AssertFalse PptTqq_CoreFlag("autoon", avail): AssertFalse avail, "core antigo"
    AssertEqual PPT_MSG_OLD_CORE_STATES, PptTqq_SetCoreFlag("autonew", True)
    f.Api = 6: f.NoCore = True
    AssertFalse PptTqq_CoreFlag("autoon", avail): AssertFalse avail, "sem o core"
    PptTqq_SetExcelForTest Nothing, ""
    AssertFalse PptTqq_CoreFlag("autoon", avail): AssertFalse avail, "sem Excel"
    AssertEqual PPT_MSG_NOCORE_STATES, PptTqq_SetCoreFlag("autonew", True)
Cleanup:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    PptTqq_ResetTestHooks
    If en <> 0 Then Err.Raise en, "TestPptAddin", ed
End Sub

' Ponta a ponta (Test-PptAddin.ps1): Verificar pelo dialogo com o core de verdade. Devolve "linha de status|estado da 1a linha".
Public Function E2e_LinksCheck(pres As Object) As String
    Dim msg As String
    PptLinks_ResetTestHooks
    msg = PptLinks_Do(pres, "check", PptLinks_Collect(pres))
    E2e_LinksCheck = msg & "|" & CStr(PptLinks_Collect(pres)(1)(5))
    PptLinks_ResetTestHooks
End Function
