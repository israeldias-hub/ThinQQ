Attribute VB_Name = "TestPowerPoint"
Option Explicit

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
    Set p.Slides = New Collection
    For i = 1 To n
        Set sl = New TestPptFake: sl.SlideIndexValue = i
        p.Slides.Add sl
    Next
    Set FakePres = p
End Function

' Onda B (B1.1): com janela, o destino vem da apresentacao DA JANELA ATIVA: o slide do modo Normal, senao o selecionado (Classificador de
' slides, Estrutura de topicos), senao o ultimo dela (mestre: View.Slide devolve o mestre, que nao e slide); nunca Presentations(1), que pode
' ser outra apresentacao do usuario. So sem janela vale a primeira apresentacao. Nada e criado quando ha apresentacao.
Public Sub Test_Ppt_TargetSlide_UsesActiveWindowPresentation()
    Dim app As TestPptFake, win As TestPptFake, presA As TestPptFake, presB As TestPptFake, sel As TestPptFake, vw As TestPptFake, master As TestPptFake
    Dim created As Object
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
    AssertTrue Ppt_TargetSlide(app, created) Is presA.Slides(2), "sem janela: o ultimo slide de Presentations(1)"
    AssertTrue created Is Nothing
End Sub

' Onda B (B1.3): CreateObject anexa a um PowerPoint ja aberto que a ROT nao ve (aberto so para incorporacao, sem apresentacoes): so e "nosso"
' (encerrado numa falha) o PowerPoint aberto quando nao havia NENHUM processo POWERPNT; WMI indisponivel (-1) = na duvida, nao e nosso.
Public Sub Test_Ppt_CreatedOnlyWithoutPowerPointProcess()
    AssertEqual PowerPointProcessCount(), Ppt_ProcessCount(), "processos POWERPNT por WMI"
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

' Review Focus 5 (parte 2): colagem real. Roda SO quando nenhum PowerPoint esta aberto (aplicativo de instancia unica: CreateObject anexaria
' ao do usuario) e o PowerPoint existe; senao sai verde registrando o motivo. Apresentacao sem janela; a instancia e encerrada ao final.
Public Sub Test_Ppt_RealPaste_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, shp As Object
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Sub
    ' a ROT (GetObject) nao ve um PowerPoint aberto so para incorporacao: confere tambem os processos (regra global: nenhum POWERPNT)
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Sub
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    If Not ClipboardUsable(c) Then Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "imagem colada"
    AssertEqual 13, slide.Shapes(1).Type, "msoPicture"
    AssertAlmost 432, CDbl(slide.Shapes(1).Height), 2, "limitado pela altura (432 pt = 80% de 540): largura ~710 (" & slide.Shapes(1).Width & ")"
    AssertTrue InStr(Core_LastNotice(), "slide 1") > 0 And InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(2)
    AssertEqual -1, shp.HasChart, "gráfico do PowerPoint"
    AssertEqual -1, CLng(shp.Chart.ChartData.IsLinked), "vinculado a pasta salva"
    AssertTrue shp.Chart.Shapes.Count >= 2, "formas da seta CAGR preservadas: " & shp.Chart.Shapes.Count
    AssertTrue InStr(Core_LastNotice(), "gráfico vinculado") > 0, Core_LastNotice()
    rng.Worksheet.Range("B2").Value = 11
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertTrue InStr(Core_LastNotice(), PPT_NOTE_DIRTY) > 0, "aviso de alteracoes nao salvas"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Ppt_SetAppForTest Nothing
    If Not pres Is Nothing Then pres.Saved = -1: pres.Close                         ' Saved: fecha sem a pergunta de salvar (travaria o teste)
    If Not app Is Nothing Then If app.Presentations.Count = 0 Then app.Quit      ' alguem anexou a esta instancia e abriu algo: nao encerra
    Set app = Nothing
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Mesmas regras de seguranca do teste anterior (so cria o PowerPoint se nenhum estiver aberto). Devolve Nothing (e registra o motivo) se pulou.
Private Function StartScratchPowerPoint() As Object
    Dim app As Object
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Function
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

' Encerra o PowerPoint criado pelo teste: fecha SO as apresentacoes que o teste (ou a chamada testada) criou, guardadas em own (o usuario
' pode ter aberto um .pptx nesta instancia enquanto o teste rodava: o PowerPoint e de instancia unica) e so da Quit se nao sobrou nenhuma.
Private Sub StopScratchPowerPoint(app As Object, own As Collection)
    Dim p As Variant
    On Error Resume Next
    Ppt_SetFailForTest ""
    Ppt_SetAppForTest Nothing
    If app Is Nothing Then Exit Sub
    For Each p In own
        p.Saved = -1
        p.Close
    Next
    If app.Presentations.Count = 0 Then app.Quit
End Sub

' Sem apresentacao aberta (gancho "slide"): a apresentacao criada pela chamada que falhou e fechada; o gancho nao vaza para a chamada
' seguinte (nem numa saida antecipada); sem falha, a chamada cria a apresentacao e cola no slide 1. Nao usa a area de transferencia ate o gancho.
Public Sub Test_Ppt_NoPresentation_FailureClosesWhatItCreated_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, en As Long, ed As String, own As New Collection
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src3.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Ppt_SetAppForTest app
    AssertEqual 0, app.Presentations.Count, "ponto de partida: sem apresentacao"
    Core_ClearLastNotice
    Ppt_SetFailForTest "slide"
    Core_CopyToPowerPoint "image"
    AssertEqual 0, app.Presentations.Count, "a apresentacao criada pela chamada que falhou foi fechada"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (slide)") > 0, Core_LastNotice()
    Ppt_SetFailForTest "slide"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    If Core_LastNotice() = PPT_ERR_CLIPBOARD Then
        Errors_Log "TestPowerPoint: parte da colagem pulada (area de transferencia do Windows indisponivel)"
        AssertEqual 0, app.Presentations.Count, "sem area de transferencia: a apresentacao criada e fechada"
    Else
        AssertEqual 1, app.Presentations.Count, "apresentacao criada"
        own.Add app.Presentations(1)                 ' criada pela chamada (havia 0): fechada no fim
        AssertEqual 1, app.Presentations(1).Slides(1).Shapes.Count, "imagem no slide novo"
        AssertTrue InStr(Core_LastNotice(), "slide 1 de '" & app.Presentations(1).Name & "'") > 0, Core_LastNotice()
    End If
Cleanup:
    en = Err.Number: ed = Err.Description
    StopScratchPowerPoint app, own
    Set app = Nothing
    On Error Resume Next
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Falha no ajuste (gancho "place"): a forma colada sai do slide e o aviso explica; o gancho nao vaza; a copia do Excel (CutCopyMode) e
' cancelada. Com a area de transferencia indisponivel o teste valida so o aviso proprio e que nada foi colado.
Public Sub Test_Ppt_FailureAfterPaste_RemovesPastedShape_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, own As New Collection
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src2.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ppt_SetAppForTest app
    If Not ClipboardUsable(c) Then
        Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel); validando so o aviso"
        Core_ClearLastNotice
        Core_CopyToPowerPoint "image"
        AssertEqual 0, slide.Shapes.Count, "nada colado"
        AssertEqual PPT_ERR_CLIPBOARD, Core_LastNotice()
        GoTo Cleanup
    End If
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "image"
    AssertEqual 0, slide.Shapes.Count, "imagem colada e removida quando o ajuste falha"
    AssertTrue InStr(Core_LastNotice(), "Não foi possível colar no PowerPoint") > 0 And InStr(Core_LastNotice(), "falha injetada (place)") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, slide.Shapes.Count, "grafico colado e removido quando o ajuste falha"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado na falha"
    Core_ClearLastNotice
    Ppt_SetFailForTest "after"                       ' Onda B (B1.4): falha DEPOIS de colar e ajustar tambem tira a forma do slide
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, slide.Shapes.Count, "falha depois de colar: a forma colada sai do slide"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (after)") > 0, Core_LastNotice()
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "gancho desarmado: a chamada seguinte cola"
    AssertTrue InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico vinculado colado"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado apos colar o grafico"
Cleanup:
    en = Err.Number: ed = Err.Description
    StopScratchPowerPoint app, own
    Set app = Nothing
    On Error Resume Next
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub
