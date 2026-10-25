Attribute VB_Name = "TestPowerPoint"
Option Explicit
' True so no teste com duble (janelas falsas): o pulo nao vai ao log, onde o run-tests conta os testes de PowerPoint real pulados.
Private mFakeWindows As Boolean

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2024, 10, 5): ws.Range("A3:C3").Value = Array(2025, 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

' Processos POWERPNT.EXE (consulta WMI; o moniker winmgmts leva duas barras invertidas, montadas com Chr(92)).
Private Function PowerPointProcessCount() As Long
    PowerPointProcessCount = GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count
End Function

' False se a area de transferencia do Windows nao pode ser aberta agora (outro programa a mantem aberta): a colagem real nao tem como rodar.
Private Function ClipboardUsable(cht As Chart) As Boolean
    On Error Resume Next
    cht.CopyPicture xlScreen, xlPicture
    ClipboardUsable = (Err.Number = 0)
    Err.Clear
End Function

Public Sub Test_Ppt_FitSize()
    Dim w As Double, h As Double
    Ppt_FitSize 460, 250, 960, 540, w, h
    AssertAlmost 768, w, 0.001: AssertAlmost 250# * 768 / 460, h, 0.001, "80% da largura, proporcao mantida"
    Ppt_FitSize 460, 280, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "o grafico padrao (460 x 280) e limitado pela altura": AssertAlmost 460# * 432 / 280, w, 0.001
    Ppt_FitSize 200, 600, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "limitado por 80% da altura": AssertAlmost 200# * 432 / 600, w, 0.001
    Ppt_FitSize 0, 0, 960, 540, w, h
    AssertAlmost 768, w, 0.001, "tamanho invalido: a caixa inteira (80% do slide)": AssertAlmost 432, h, 0.001
    Ppt_FitSize -460, 280, 960, 540, w, h
    AssertAlmost 768, w, 0.001, "largura negativa: a caixa inteira": AssertAlmost 432, h, 0.001
    Ppt_FitSize 460, 280, 0, 540, w, h
    AssertAlmost 432, h, 0.001, "slide de largura 0: vale o 16:9 padrao (960 x 540)": AssertAlmost 460# * 432 / 280, w, 0.001
    Ppt_FitSize 460, 280, -960, -540, w, h
    AssertAlmost 432, h, 0.001, "slide de tamanho negativo: vale o 16:9 padrao": AssertTrue w > 0
End Sub

' Duble de uma apresentacao com n slides (SlideIndex 1..n).
Private Function FakePres(n As Long) As TestPptFake
    Dim p As TestPptFake, sl As TestPptFake, i As Long
    Set p = New TestPptFake
    Set p.Slides = New TestPptSlides
    For i = 1 To n
        Set sl = New TestPptFake: sl.SlideIndexValue = i
        p.Slides.Append sl
    Next
    Set FakePres = p
End Function

' Onda B (B1.1): com janela, o destino vem da apresentacao DA JANELA ATIVA: o slide do modo Normal, senao o selecionado (Classificador de
' slides, Estrutura de topicos), senao o ultimo dela (mestre: View.Slide devolve o mestre, que nao e slide); nunca Presentations(1), que pode
' ser outra apresentacao do usuario. So sem janela vale a primeira apresentacao. Nada e criado quando ha apresentacao.
Public Sub Test_Ppt_TargetSlide_UsesActiveWindowPresentation()
    Dim app As TestPptFake, win As TestPptFake, presA As TestPptFake, presB As TestPptFake, sel As TestPptFake, vw As TestPptFake, master As TestPptFake
    Dim created As Object, added As Object, presE As TestPptFake, got As Object
    Ppt_ResetTestHooks                                            ' sem destino fixo nem instancia injetada de um teste anterior
    Set presA = FakePres(2): Set presB = FakePres(3)
    Set win = New TestPptFake: Set win.Presentation = presB      ' View = Nothing: View.Slide falha (Classificador de slides)
    Set app = New TestPptFake
    Set app.Windows = New Collection: app.Windows.Add win
    Set app.ActiveWindow = win
    Set app.Presentations = New Collection: app.Presentations.Add presA: app.Presentations.Add presB
    AssertTrue Ppt_TargetSlide(app, created) Is presB.Slides(3), "ultimo slide da apresentacao da janela ativa (nao Presentations(1))"
    AssertTrue created Is Nothing, "nada criado"
    Set sel = New TestPptFake: Set sel.SelectedSlide = presB.Slides(2)
    Set win.Selection = sel
    AssertTrue Ppt_TargetSlide(app) Is presB.Slides(2), "o slide selecionado na janela"
    Set master = New TestPptFake                                  ' sem SlideIndex: mestre/layout
    Set vw = New TestPptFake: Set vw.Slide = master
    Set win.View = vw
    AssertTrue Ppt_TargetSlide(app) Is presB.Slides(2), "View.Slide no modo mestre nao vale como slide"
    Set vw.Slide = presB.Slides(1)
    AssertTrue Ppt_TargetSlide(app) Is presB.Slides(1), "modo Normal: o slide da janela"
    Set win.View = Nothing: Set win.Selection = Nothing
    Set app.ActiveWindow = Nothing                                ' sem janela ativa: a primeira janela, ainda nao Presentations(1)
    AssertTrue Ppt_TargetSlide(app) Is presB.Slides(3), "janela sem foco: a apresentacao dela"
    Set app.Windows = New Collection                              ' sem janela: a primeira apresentacao
    AssertTrue Ppt_TargetSlide(app, created, added) Is presA.Slides(2), "sem janela: o ultimo slide de Presentations(1)"
    AssertTrue created Is Nothing
    AssertTrue added Is Nothing, "apresentacao com slides: nenhum slide acrescentado"
    Set presE = FakePres(0)                                       ' rodada 1 (Importante 3): apresentacao aberta sem slides
    Set app.Presentations = New Collection: app.Presentations.Add presE
    Set got = Ppt_TargetSlide(app, created, added)
    AssertEqual 1, presE.Slides.AddedCount, "um slide acrescentado"
    AssertTrue added Is got, "o slide acrescentado volta para quem chama (apagado se falhar depois)"
    AssertTrue created Is Nothing, "a apresentacao nao foi criada pela chamada"
End Sub

' Janela do Modo de Exibicao Protegido (um .pptx baixado ou anexo que o usuario abre no PowerPoint do teste) nao entra em Windows nem em
' Presentations: o teste a conta como janela do usuario (para antes de chamar o core) e nao encerra a instancia (Quit fecharia a janela
' dele). Contagem ilegivel: na duvida, e do usuario.
Public Sub Test_Ppt_ScratchHelpers_RespectProtectedView()
    Dim app As Object, keep As TestPptFake, own As Collection
    mFakeWindows = True
    Set keep = New TestPptFake
    Set keep.Windows = New Collection: Set keep.Presentations = New Collection
    Set keep.ProtectedViewWindows = New Collection: keep.ProtectedViewWindows.Add New TestPptFake
    Set app = keep
    AssertTrue UserWindowAppeared(app), "janela protegida conta como janela do usuario"
    Set own = New Collection
    StopScratchPowerPoint app, own
    AssertFalse keep.QuitCalled, "com janela protegida aberta a instancia nao e encerrada"
    Set keep.ProtectedViewWindows = Nothing                       ' leitura falha (erro 91)
    Set app = keep
    AssertTrue UserWindowAppeared(app), "contagem ilegivel: na duvida, do usuario"
    Set own = New Collection
    StopScratchPowerPoint app, own
    AssertFalse keep.QuitCalled, "contagem ilegivel: nao encerra"
    Set keep.ProtectedViewWindows = New Collection
    Set app = keep
    AssertFalse UserWindowAppeared(app), "sem janelas: segue o teste"
    Set own = New Collection
    StopScratchPowerPoint app, own
    AssertTrue keep.QuitCalled, "sem janelas nem apresentacoes: encerra"
    mFakeWindows = False
End Sub

' Onda B (B1.3): CreateObject anexa a um PowerPoint ja aberto que a ROT nao ve (aberto so para incorporacao, sem apresentacoes): so e "nosso"
' (encerrado numa falha) o PowerPoint aberto quando nao havia NENHUM processo POWERPNT; WMI indisponivel (-1) = na duvida, nao e nosso.
Public Sub Test_Ppt_CreatedOnlyWithoutPowerPointProcess()
    Dim a As Long, b As Long
    a = PowerPointProcessCount(): b = Ppt_ProcessCount()
    If a <> b Then WaitSeconds 1: a = PowerPointProcessCount(): b = Ppt_ProcessCount()     ' um POWERPNT entrando ou saindo entre as duas leituras
    AssertEqual a, b, "processos POWERPNT por WMI"
    AssertTrue Ppt_CreatedByUs(0, New Collection), "nenhum processo antes: nosso"
    AssertFalse Ppt_CreatedByUs(1, New Collection), "havia um processo: CreateObject anexou a ele"
    AssertFalse Ppt_CreatedByUs(-1, New Collection), "contagem indisponivel: nao e nosso"
    AssertFalse Ppt_CreatedByUs(0, Nothing), "sem aplicativo"
End Sub

' Onda B (B1.4): estado do vinculo do grafico colado (lido sob On Error) e o texto do aviso; PowerPoint ocupado (RPC_E_CALL_REJECTED) ganha
' aviso proprio de "tente de novo".
Public Sub Test_Ppt_LinkStateAndTexts()
    Dim shp As TestPptFake, ch As TestPptFake, cd As TestPptFake, t As String
    Set cd = New TestPptFake: cd.IsLinked = -1
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set shp = New TestPptFake: Set shp.Chart = ch
    AssertEqual 1, Ppt_LinkState(shp), "vinculado"
    cd.IsLinked = 0
    AssertEqual 0, Ppt_LinkState(shp), "colado sem vinculo"
    Set shp.Chart = Nothing
    AssertEqual -1, Ppt_LinkState(shp), "leitura falha: desconhecido"
    AssertEqual -1, Ppt_LinkState(Nothing), "sem forma: desconhecido"
    AssertEqual "Gráfico copiado para o slide 2 de 'Pres' como gráfico vinculado.", Ppt_DoneText(2, "Pres", True, 1)
    AssertEqual "Gráfico copiado para o slide 2 de 'Pres' como gráfico vinculado.", Ppt_DoneText(2, "Pres", True, -1), "desconhecido: como antes"
    t = Ppt_DoneText(2, "Pres", True, 0)
    AssertTrue InStr(t, "slide 2 de 'Pres'") > 0 And InStr(t, "sem vínculo") > 0, t
    AssertEqual "Gráfico copiado para o slide 1 de 'P' como imagem.", Ppt_DoneText(1, "P", False, -1)
    AssertEqual PPT_ERR_BUSY, Ppt_ErrorText(-2147418111, "x"), "RPC_E_CALL_REJECTED"
    AssertTrue Ppt_IsKnownFailure(-2147418111), "PowerPoint ocupado: aviso proprio"
    AssertTrue InStr(PPT_ERR_BUSY, "tente de novo") > 0, PPT_ERR_BUSY
End Sub

' Review Focus 5 (parte 1): pasta nao salva e vinculo.
Public Sub Test_Ppt_LinkedPrecheck_UnsavedWorkbook()
    Dim wb As Workbook, p As String
    Set wb = NewTempWorkbook()
    AssertEqual PPT_ERR_UNSAVED, Ppt_LinkedPrecheck(wb)
    p = TempDir() & "\ppt.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    AssertEqual "", Ppt_LinkedPrecheck(wb), "pasta salva"
    CloseTemp wb
    AssertEqual PPT_ERR_NOT_INSTALLED, Ppt_ErrorText(429, "x")
    AssertTrue InStr(Ppt_ErrorText(70, "Permissão negada"), "Permissão negada") > 0
    AssertEqual PPT_ERR_CLIPBOARD, Ppt_ErrorText(-2147221040, "x"), "CLIPBRD_E_CANT_OPEN: mensagem propria"
    AssertTrue Ppt_IsKnownFailure(429) And Ppt_IsKnownFailure(-2147221040) And Not Ppt_IsKnownFailure(70), "falhas esperadas viram aviso proprio"
End Sub

Public Sub Test_Core_CopyToPowerPoint_Unsaved_Linked_NoticeBeforePowerPoint()
    Dim wb As Workbook, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual PPT_ERR_UNSAVED, Core_LastNotice(), "vinculado com pasta nao salva: so o aviso"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "xyz"
    AssertEqual "", Core_LastNotice(), "modo desconhecido: nada (so log)"
    rng.Worksheet.Range("A1").Select
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual "", Core_LastNotice(), "sem grafico: nada"
    Teardown wb
End Sub

' Seguranca dos testes com PowerPoint real (Onda B, rodada 1): o PowerPoint e de instancia unica, entao um PowerPoint ou .pptx que o usuario
' abra enquanto o teste roda cai NA INSTANCIA DO TESTE. Por isso: (1) o destino das colagens e fixado no slide do proprio teste
' (Ppt_PinTargetForTest) ou forcado a uma apresentacao nova (forceNew), nunca a janela ativa; (2) antes de cada chamada, uma janela na
' instancia (o teste so cria apresentacoes sem janela) encerra o teste com o motivo no log; (3) so as apresentacoes que o teste ou a chamada
' criaram (pelo objeto, nunca por indice) sao fechadas, e o PowerPoint so recebe Quit sem nenhuma apresentacao aberta.

' Review Focus 5 (parte 2): colagem real. Roda SO quando nenhum PowerPoint esta aberto (aplicativo de instancia unica: CreateObject anexaria
' ao do usuario) e o PowerPoint existe; senao sai verde registrando o motivo. Apresentacao sem janela; a instancia e encerrada ao final.
Public Sub Test_Ppt_RealPaste_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, shp As Object
    Dim own As New Collection
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    If Not ClipboardUsable(c) Then Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "imagem colada"
    AssertEqual 13, slide.Shapes(1).Type, "msoPicture"
    AssertAlmost 432, CDbl(slide.Shapes(1).Height), 2, "limitado pela altura (432 pt = 80% de 540): largura ~710 (" & slide.Shapes(1).Width & ")"
    AssertTrue InStr(Core_LastNotice(), "slide 1") > 0 And InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    AssertTrue LinkFrame_Find(c) Is Nothing, "X2b: imagem nao ganha moldura"
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(2)
    AssertEqual -1, shp.HasChart, "gráfico do PowerPoint"
    AssertEqual -1, CLng(shp.Chart.ChartData.IsLinked), "vinculado a pasta salva"
    AssertTrue shp.Chart.Shapes.Count >= 2, "formas da seta CAGR preservadas: " & shp.Chart.Shapes.Count
    AssertTrue InStr(Core_LastNotice(), "gráfico vinculado") > 0, Core_LastNotice()
    ' X2b (1.13.0): Copiar como grafico vinculado (vinculo mantido) tambem marca o intervalo de origem com a moldura
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "X2b: colado vinculado ganha moldura"
    LinkFrame_SetState c, False                       ' revisao final 1.13.0 (M3): moldura vermelha de antes volta a verde ao colar de novo
    AssertEqual rng.Address, LinkFrame_Find(c).TopLeftCell.Resize(rng.Rows.Count, rng.Columns.Count).Address, "no intervalo de origem"
    rng.Worksheet.Range("B2").Value = 11
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertTrue InStr(Core_LastNotice(), PPT_NOTE_DIRTY) > 0, "aviso de alteracoes nao salvas"
    AssertEqual 1, CountFramesOf(rng.Worksheet), "colar de novo nao duplica a moldura"
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB, "M3: colou de novo = verde"
    ' X2b: com a moldura, Alterar dados avisa que o slide precisa ser colado de novo
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(c, rng.Worksheet.Range("A1:B3")), "Alterar dados"
    AssertEqual LINKFRAME_NOTE_REBUILT, Core_LastNotice(), "aviso: colar de novo"
    AssertEqual 0, CountFramesOf(rng.Worksheet), "a moldura sai"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing     ' sem referencias: o processo do teste pode terminar depois do Quit
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' So cria o PowerPoint se nenhum estiver aberto (nem na ROT nem como processo POWERPNT: um aberto so para incorporacao nao aparece na ROT).
' Devolve Nothing (e registra o motivo) se pulou. Desarma os ganchos de um teste anterior.
' A nota da colagem com alteracoes nao salvas diz a verdade medida: pasta ABERTA no Excel = o vinculo le os valores nao salvos;
' fechada = o arquivo salvo (a frase antiga dizia que o PowerPoint sempre le o arquivo salvo).
Public Sub Test_Ppt_NoteDirty_OpenWorkbookReadsUnsavedValues()
    AssertTrue InStr(1, PPT_NOTE_DIRTY, "aberta no Excel", vbTextCompare) > 0, "cita a pasta aberta no Excel"
    AssertTrue InStr(1, PPT_NOTE_DIRTY, "fechada", vbTextCompare) > 0, "cita a pasta fechada"
    AssertTrue InStr(1, PPT_NOTE_DIRTY, "nas atualizações do vínculo", vbTextCompare) = 0, "sem a afirmacao antiga"
End Sub

Private Function StartScratchPowerPoint() As Object
    Dim app As Object
    mFakeWindows = False
    Ppt_ResetTestHooks
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Function
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

' True (e registra) se apareceu uma janela na instancia do teste: o teste so cria apresentacoes SEM janela, entao a janela e do usuario (que
' abriu o PowerPoint ou um .pptx durante o teste). Conta tambem as janelas do Modo de Exibicao Protegido (fora de Windows); contagem
' ilegivel conta como janela. Quem chama encerra o teste sem chamar o core de novo.
Private Function UserWindowAppeared(app As Object) As Boolean
    Dim n As Long, npv As Long
    n = -1: npv = -1
    On Error Resume Next
    n = app.Windows.Count
    npv = app.ProtectedViewWindows.Count
    On Error GoTo 0
    UserWindowAppeared = (n <> 0 Or npv <> 0)
    If UserWindowAppeared And Not mFakeWindows Then Errors_Log "TestPowerPoint: resto do teste pulado (uma janela apareceu na instancia do teste: PowerPoint do usuario)"
End Function

' True se ha uma apresentacao aberta com esse nome na instancia.
Private Function PresOpen(app As Object, nm As String) As Boolean
    Dim p As Object
    On Error Resume Next
    For Each p In app.Presentations
        If CStr(p.Name) = nm Then PresOpen = True: Exit Function
    Next
End Function

' Encerra o PowerPoint criado pelo teste: desarma os ganchos (inclusive o destino fixo), fecha SO as apresentacoes de own (as que o teste ou
' a chamada testada criaram, registradas pelo objeto) que continuam sem janela (com janela nao e do teste), solta as referencias e so da
' Quit se nao sobrou nenhuma apresentacao nem janela do Modo de Exibicao Protegido (contagem ilegivel: nao encerra); espera (ate 10 s) o
' processo terminar, para o teste seguinte nao pular por um POWERPNT que ainda esta saindo.
Private Sub StopScratchPowerPoint(app As Object, own As Collection)
    Dim p As Variant, nw As Long, left_ As Long, t As Single, npv As Long
    On Error Resume Next
    Ppt_ResetTestHooks
    If app Is Nothing Then Exit Sub
    For Each p In own
        nw = -1
        nw = p.Windows.Count
        If nw = 0 Then
            p.Saved = -1
            p.Close
        End If
    Next
    Set p = Nothing
    Set own = New Collection
    left_ = -1: npv = -1
    left_ = app.Presentations.Count
    npv = app.ProtectedViewWindows.Count
    If left_ = 0 And npv = 0 Then
        app.Quit
        Set app = Nothing
        t = Timer
        Do While PowerPointProcessCount() > 0
            If Timer < t Or Timer - t > 10 Then Exit Do
            WaitSeconds 0.25
        Loop
    Else
        Errors_Log "TestPowerPoint: o PowerPoint do teste ficou aberto (" & left_ & " apresentacao(oes) e " & npv & " janela(s) protegida(s) que nao sao do teste)"
        Set app = Nothing
    End If
End Sub

Private Sub WaitSeconds(s As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer - t < s
        DoEvents
    Loop
End Sub

' Sem apresentacao aberta (gancho "slide"): a apresentacao criada pela chamada que falhou e fechada; o gancho nao vaza para a chamada
' seguinte (nem numa saida antecipada); sem falha, a chamada cria a apresentacao e cola no slide 1. Nao usa a area de transferencia ate o gancho.
' Destino forcado a uma apresentacao nova (forceNew): uma janela do usuario na instancia nunca vira destino; a apresentacao criada e
' identificada pelo gancho Ppt_LastCreatedPres*ForTest, nunca por indice.
Public Sub Test_Ppt_NoPresentation_FailureClosesWhatItCreated_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, en As Long, ed As String, own As New Collection, created As Object, nm As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Ppt_PinTargetForTest Nothing, True
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src3.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Ppt_SetFailForTest "slide"
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "image"
    nm = Ppt_LastCreatedPresNameForTest()
    AssertTrue Len(nm) > 0, "a chamada criou uma apresentacao"
    AssertFalse PresOpen(app, nm), "a apresentacao criada pela chamada que falhou foi fechada (" & nm & ")"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (slide)") > 0, Core_LastNotice()
    Ppt_SetFailForTest "slide"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "image"
    Set created = Ppt_LastCreatedPresForTest()
    nm = Ppt_LastCreatedPresNameForTest()
    If PresOpen(app, nm) Then own.Add created         ' criada por esta chamada (pelo objeto): fechada no fim
    If Core_LastNotice() = PPT_ERR_CLIPBOARD Then
        Errors_Log "TestPowerPoint: parte da colagem pulada (area de transferencia do Windows indisponivel)"
        AssertFalse PresOpen(app, nm), "sem area de transferencia: a apresentacao criada e fechada"
    Else
        AssertTrue PresOpen(app, nm), "apresentacao criada e mantida"
        AssertEqual 1, created.Slides(1).Shapes.Count, "imagem no slide novo"
        AssertTrue InStr(Core_LastNotice(), "slide 1 de '" & nm & "'") > 0, Core_LastNotice()
    End If
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Set created = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Falha no ajuste (gancho "place"): a forma colada sai do slide e o aviso explica; o gancho nao vaza; a copia do Excel (CutCopyMode) e
' cancelada. Com a area de transferencia indisponivel o teste valida so o aviso proprio e que nada foi colado. Destino fixo no slide do teste.
Public Sub Test_Ppt_FailureAfterPaste_RemovesPastedShape_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, own As New Collection
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src2.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ppt_SetAppForTest app
    If Not ClipboardUsable(c) Then
        Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel); validando so o aviso"
        Core_ClearLastNotice
        If UserWindowAppeared(app) Then GoTo Cleanup
        Core_CopyToPowerPoint "image"
        AssertEqual 0, slide.Shapes.Count, "nada colado"
        AssertEqual PPT_ERR_CLIPBOARD, Core_LastNotice()
        GoTo Cleanup
    End If
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "image"
    AssertEqual 0, slide.Shapes.Count, "imagem colada e removida quando o ajuste falha"
    AssertTrue InStr(Core_LastNotice(), "Não foi possível colar no PowerPoint") > 0 And InStr(Core_LastNotice(), "falha injetada (place)") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, slide.Shapes.Count, "grafico colado e removido quando o ajuste falha"
    AssertTrue LinkFrame_Find(c) Is Nothing, "X2b: colagem que falhou nao deixa moldura"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado na falha"
    Core_ClearLastNotice
    Ppt_SetFailForTest "after"                       ' Onda B (B1.4): falha DEPOIS de colar e ajustar tambem tira a forma do slide
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, slide.Shapes.Count, "falha depois de colar: a forma colada sai do slide"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (after)") > 0, Core_LastNotice()
    AssertTrue LinkFrame_Find(c) Is Nothing, "X2b: falha depois de colar nao deixa moldura"
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "gancho desarmado: a chamada seguinte cola"
    AssertTrue InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico vinculado colado"
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "X2b: colagem vinculada que deu certo ganha moldura"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado apos colar o grafico"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Onda B, rodada 1 (Importante 3): apresentacao aberta SEM slides (duble, sem PowerPoint real e sem area de transferencia): a chamada
' acrescenta um slide e, numa falha depois (gancho "target", logo apos escolher o destino), apaga o slide que acrescentou; uma apresentacao com
' slides nao ganha nem perde slide. O gancho de destino fixo fica desarmado.
Public Sub Test_Ppt_AddedSlideRemovedOnFailure_Fake()
    Dim wb As Workbook, rng As Range, app As TestPptFake, presE As TestPptFake, presF As TestPptFake, en As Long, ed As String
    On Error GoTo Cleanup
    Ppt_ResetTestHooks
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set presE = FakePres(0)
    Set app = New TestPptFake
    Set app.Windows = New Collection
    Set app.Presentations = New Collection: app.Presentations.Add presE
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Ppt_SetFailForTest "target"
    Core_CopyToPowerPoint "image"
    AssertEqual 1, presE.Slides.AddedCount, "slide acrescentado a apresentacao vazia"
    AssertTrue presE.Slides(1).Deleted, "a falha depois apaga o slide que a chamada acrescentou"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (target)") > 0, Core_LastNotice()
    Set presF = FakePres(2)
    Set app.Presentations = New Collection: app.Presentations.Add presF
    Ppt_SetFailForTest "target"
    Core_CopyToPowerPoint "image"
    AssertEqual 0, presF.Slides.AddedCount, "apresentacao com slides: nenhum slide novo"
    AssertFalse presF.Slides(2).Deleted, "o slide do usuario nao e apagado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Ppt_ResetTestHooks
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' D2 (revisao I1): a moldura so fica se a colagem vinculada manteve o vinculo (ou a leitura falhou: o aviso diz vinculado); imagem sempre.
Public Sub Test_Ppt_PasteKept_UnlinkedPasteDropsFrame()
    AssertTrue Ppt_PasteKept(True, 1)
    AssertTrue Ppt_PasteKept(True, -1), "leitura falhou: o aviso diz vinculado"
    AssertFalse Ppt_PasteKept(True, 0), "colado sem vinculo: sem moldura"
    AssertTrue Ppt_PasteKept(False, -1), "imagem"
    ' X2b (1.13.0): Copiar como grafico vinculado desenha a moldura so com o vinculo mantido; imagem nunca
    AssertTrue Ppt_FrameAfterCopy(True, 1)
    AssertTrue Ppt_FrameAfterCopy(True, -1), "leitura falhou: o aviso diz vinculado"
    AssertFalse Ppt_FrameAfterCopy(True, 0), "colado sem vinculo: sem moldura"
    AssertFalse Ppt_FrameAfterCopy(False, 1), "imagem: sem moldura"
    AssertFalse Ppt_FrameAfterCopy(False, -1), "imagem: sem moldura"
End Sub

Private Function CountFramesOf(ws As Worksheet) As Long
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then CountFramesOf = CountFramesOf + 1
    Next
End Function
