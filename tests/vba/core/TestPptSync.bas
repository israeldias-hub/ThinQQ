Attribute VB_Name = "TestPptSync"
Option Explicit
' Parte A (1.4.0): motor de sincronizacao dos graficos ThinQQ colados no PowerPoint (modPptSync). Os testes com PowerPoint real so rodam sem
' nenhum POWERPNT aberto (aplicativo de instancia unica: CreateObject anexaria ao do usuario); senao saem verdes registrando o motivo.
Private mAutoWasActive As Boolean               ' AutoBegin/AutoEnd: os eventos do Excel estavam ligados antes do teste do motor automatico

Private Sub Teardown(wb As Workbook)
    PptSync_ResetTestHooks
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

Private Function ClipboardUsable(cht As Chart) As Boolean
    On Error Resume Next
    cht.CopyPicture xlScreen, xlPicture
    ClipboardUsable = (Err.Number = 0)
    Err.Clear
End Function

Private Sub WaitSeconds(s As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer - t < s
        DoEvents
    Loop
End Sub

' So cria o PowerPoint se nenhum processo POWERPNT existe, conferido ANTES de qualquer GetObject (nunca toca o PowerPoint do usuario);
' contagem desconhecida (WMI falhou: -1) tambem pula. Nothing (e o motivo no log) se pulou.
Private Function StartScratchPowerPoint() As Object
    Dim app As Object, n As Long
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    n = Ppt_ProcessCount()
    If n <> 0 Then Errors_Log "TestPptSync: pulado (processo POWERPNT existente ou contagem indisponivel: " & n & ")": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPptSync: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

Private Function UserWindowAppeared(app As Object) As Boolean
    Dim n As Long, npv As Long
    n = -1: npv = -1
    On Error Resume Next
    n = app.Windows.Count
    npv = app.ProtectedViewWindows.Count
    On Error GoTo 0
    UserWindowAppeared = (n <> 0 Or npv <> 0)
    If UserWindowAppeared Then Errors_Log "TestPptSync: resto do teste pulado (uma janela apareceu na instancia do teste: PowerPoint do usuario)"
End Function

' Fecha so as apresentacoes do teste (sem janela) e encerra o PowerPoint se nada mais ficou aberto; espera o processo sair (ate 10 s).
Private Sub StopScratchPowerPoint(app As Object, own As Collection)
    Dim p As Variant, nw As Long, left_ As Long, t As Single, npv As Long
    On Error Resume Next
    Ppt_ResetTestHooks
    If app Is Nothing Then Exit Sub
    For Each p In own
        nw = -1
        nw = p.Windows.Count
        If nw = 0 Then p.Saved = -1: p.Close
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
        Do While Ppt_ProcessCount() > 0
            If Timer < t Or Timer - t > 10 Then Exit Do
            WaitSeconds 0.25
        Loop
    Else
        Errors_Log "TestPptSync: o PowerPoint do teste ficou aberto (" & left_ & " apresentacao(oes), " & npv & " janela(s) protegida(s))"
        Set app = Nothing
    End If
End Sub

' Regra de tres propria no grafico DO SLIDE (eixo de valores = Axes(2) do PowerPoint).
Private Function PptYOf(pc As Object, v As Double) As Double
    Dim ax As Object
    Set ax = pc.Axes(2)
    PptYOf = pc.PlotArea.InsideTop + pc.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function PptHasShape(pc As Object, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = pc.Shapes(nm)
    PptHasShape = Not s Is Nothing
End Function

Private Function PptSeriesValue(pc As Object, seriesName As String, i As Long) As Double
    Dim k As Long, v As Variant
    For k = 1 To pc.SeriesCollection.Count
        If CStr(pc.SeriesCollection(k).Name) = seriesName Then
            v = pc.SeriesCollection(k).Values
            PptSeriesValue = CDbl(v(LBound(v) + i - 1))
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 1, "TestPptSync", "serie nao encontrada no grafico do slide: " & seriesName
End Function

' PowerPoint proprio com um slide fixado como destino e pasta SALVA (fileName) com grafico empilhado e seta CAGR. False (motivo no log) se
' o teste deve ser pulado; erros sobem ao chamador (que limpa no Cleanup: app, own e wb podem estar parcialmente montados).
Private Function PasteSetup(ByRef app As Object, own As Collection, ByRef slide As Object, ByRef wb As Workbook, ByRef rng As Range, ByRef c As Chart, fileName As String) As Boolean
    Dim pres As Object
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Function
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\" & fileName, xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    If Not ClipboardUsable(c) Then Errors_Log "TestPptSync: colagem pulada (area de transferencia indisponivel)": Exit Function
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then Exit Function
    PasteSetup = True
End Function

Public Sub Test_PptSync_SourcePath()
    AssertEqual "C:\a\b.xlsx", PptSync_SourcePath("C:\a\b.xlsx", "")
    AssertEqual "C:\a\b.xlsx", PptSync_SourcePath("C:\a\b.xlsx!Dados![b.xlsx]Dados Gráfico 1", "C:\x.xlsx"), "OLE: o caminho vai ate o primeiro '!'"
    AssertEqual "C:\x.xlsx", PptSync_SourcePath("", "C:\x.xlsx"), "sem vinculo legivel: a etiqueta TQQ_BOOK"
    AssertEqual "", PptSync_SourcePath("   ", "")
End Sub

Public Sub Test_PptSync_Fingerprint_ChangesWithDataSizeAndShapes()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, a As String, b As String
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432
    a = PptSync_Fingerprint(c.Parent, shp)
    AssertEqual 64, Len(a), "SHA-256 em hexadecimal"
    AssertEqual a, PptSync_Fingerprint(c.Parent, shp), "sem mudanca: igual"
    rng.Worksheet.Range("B3").Value = 21
    Application.Calculate
    b = PptSync_Fingerprint(c.Parent, shp)
    AssertTrue a <> b, "dado novo muda a impressao digital"
    shp.Width = 500
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "tamanho do slide muda a impressao digital"
    b = PptSync_Fingerprint(c.Parent, shp)
    Ann_AddCagr c, 1, 2
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "seta nova muda a impressao digital"
    ' revisao final E (I2): o titulo ligado a uma celula entra na impressao digital
    rng.Worksheet.Range("H1").Value = "Titulo 1"
    LinkTitle_Apply c, rng.Worksheet.Range("H1")
    b = PptSync_Fingerprint(c.Parent, shp)
    rng.Worksheet.Range("H1").Value = "Titulo 2"
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "titulo mudou: impressao digital muda"
    Teardown wb
End Sub

' Falha no meio (gancho): o grafico do Excel volta ao tamanho original e a pasta continua "salva"; o erro sobe.
Public Sub Test_PptSync_FailureRestoresChartSize()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, w0 As Double, h0 As Double, pt As Variant, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    w0 = c.Parent.Width: h0 = c.Parent.Height
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432      ' sem Chart: Ppt_LinkState = -1 (sem Update)
    For Each pt In Array("update", "snapshot", "apply")
        wb.Saved = True
        PptSync_SetFailForTest CStr(pt)
        On Error Resume Next
        PptSync_SyncShape shp, c.Parent
        en = Err.Number: ed = Err.Description
        On Error GoTo 0
        AssertEqual vbObjectError + 9992, en, "a falha injetada sobe (" & pt & ")"      ' o numero do gancho, nao o 91 do duble sem Chart
        AssertTrue InStr(ed, "falha injetada (" & pt & ")") > 0, "falha do ponto certo (" & pt & "): " & ed
        AssertAlmost w0, c.Parent.Width, 0.01, "largura restaurada (" & pt & ")"
        AssertAlmost h0, c.Parent.Height, 0.01, "altura restaurada (" & pt & ")"
        AssertTrue wb.Saved, "pasta continua salva (" & pt & ")"
        AssertTrue Application.EnableEvents, "eventos religados (" & pt & ")"
    Next
    Teardown wb
End Sub

' Review Focus 1: a colagem ajusta o grafico ao slide (709,7 x 432) e a seta CAGR fica sobre as colunas do SLIDE (regra de tres com a area
' interna e a escala do grafico do PowerPoint), nao esticada; etiquetas gravadas; o grafico do Excel volta ao tamanho dele.
Public Sub Test_PptSync_PastedChart_AnnotationsAlignedToSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, shp As Object, pc As Object, ln As Object
    Dim own As New Collection, en As Long, ed As String, w0 As Double, h0 As Double, extra As Double, y1 As Double, y2 As Double, x1 As Double, x2 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-a1.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    w0 = c.Parent.Width: h0 = c.Parent.Height
    If Not ClipboardUsable(c) Then Errors_Log "TestPptSync: colagem pulada (area de transferencia indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 1, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(1)
    AssertEqual Registry_IdFromChart(c), PptSync_Tag(shp, PPTSYNC_TAG_CHART)
    AssertEqual wb.FullName, PptSync_Tag(shp, PPTSYNC_TAG_BOOK)
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_VER)
    AssertEqual 64, Len(PptSync_Tag(shp, PPTSYNC_TAG_FP)), "impressao digital gravada"
    AssertAlmost 432, CDbl(shp.Height), 2, "ajustado pela altura do slide"
    Set pc = shp.Chart
    If Registry_GetOption(c, "totals", "0") = "1" Then extra = 18          ' ANN_TOTALS_EXTRA (colunas, totais ligados)
    y1 = PptYOf(pc, 15) - 10 - extra: y2 = PptYOf(pc, 26) - 10 - extra      ' topo das pilhas: 10+5 e 20+6; ANN_ARROW_GAP = 10
    If y1 < 22 Then y1 = 22
    If y2 < 22 Then y2 = 22
    x1 = pc.PlotArea.InsideLeft + pc.PlotArea.InsideWidth * 0.5 / 2: x2 = pc.PlotArea.InsideLeft + pc.PlotArea.InsideWidth * 1.5 / 2
    Set ln = pc.Shapes("tqq_ann_cagr_1")
    AssertAlmost Application.Min(y1, y2), CDbl(ln.Top), 1.5, "topo da seta no slide"
    AssertAlmost Application.Max(y1, y2), CDbl(ln.Top + ln.Height), 1.5, "base da seta no slide"
    AssertAlmost x1, CDbl(ln.Left), 1.5, "inicio da seta no centro da 1a coluna do slide"
    AssertAlmost x2, CDbl(ln.Left + ln.Width), 1.5, "fim da seta no centro da 2a coluna do slide"
    AssertEqual c.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text, CStr(pc.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost CDbl(c.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Font.Size), CDbl(pc.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Font.Size), 0.01, "fonte do rotulo no slide nao esticada"
    AssertAlmost w0, c.Parent.Width, 0.01, "grafico do Excel volta ao tamanho original"
    AssertAlmost h0, c.Parent.Height, 0.01
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set ln = Nothing: Set pc = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 2: dado novo (pasta ABERTA e nao salva), seta CAGR apagada e seta de diferenca nova no Excel; a sincronizacao atualiza os
' valores no slide, apaga a seta antiga, clona a nova (com o formato) e mantem a animacao da forma do slide.
Public Sub Test_PptSync_DataChangeNewArrowAndAnimation_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, shp As Object, pc As Object
    Dim own As New Collection, en As Long, ed As String, w0 As Double, seq As Object
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-a2.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    w0 = c.Parent.Width
    If Not ClipboardUsable(c) Then Errors_Log "TestPptSync: colagem pulada (area de transferencia indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, PptSync_UpdateCountForTest(), "a colagem nao chama LinkFormat.Update (os dados acabaram de ser colados)"
    Set shp = slide.Shapes(1)
    Set seq = slide.TimeLine.MainSequence
    seq.AddEffect shp, 10                                       ' msoAnimEffectFade
    rng.Worksheet.Range("B3").Value = 40                        ' sem salvar
    Application.Calculate
    Ann_RemoveAll c
    Ann_AddDifference c, 1, 2
    PptSync_SyncShape shp, c.Parent
    AssertEqual 1, PptSync_UpdateCountForTest(), "a sincronizacao chama LinkFormat.Update (pasta aberta)"
    Set pc = shp.Chart
    AssertAlmost 40, PptSeriesValue(pc, "Sul", 2), 0.0001, "valor nao salvo lido pelo vinculo (pasta aberta)"
    AssertFalse PptHasShape(pc, "tqq_ann_cagr_1"), "seta apagada no Excel some do slide"
    AssertTrue PptHasShape(pc, "tqq_ann_diff_1") And PptHasShape(pc, "tqq_ann_diff_1_lvl") And PptHasShape(pc, "tqq_ann_diff_1_lbl"), "seta nova clonada"
    AssertEqual c.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text, CStr(pc.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost c.Shapes("tqq_ann_diff_1").Line.Weight, CDbl(pc.Shapes("tqq_ann_diff_1").Line.Weight), 0.01, "formato da linha clonado"
    AssertEqual CLng(c.Shapes("tqq_ann_diff_1").Line.EndArrowheadStyle), CLng(pc.Shapes("tqq_ann_diff_1").Line.EndArrowheadStyle)
    AssertEqual -1, CLng(pc.Shapes("tqq_ann_diff_1_lbl").Fill.Visible), "caixa do rotulo clonado com preenchimento"
    AssertEqual CLng(c.Shapes("tqq_ann_diff_1_lbl").Fill.ForeColor.RGB), CLng(pc.Shapes("tqq_ann_diff_1_lbl").Fill.ForeColor.RGB), "cor do preenchimento clonada"
    AssertEqual CLng(c.Shapes("tqq_ann_diff_1_lbl").Line.ForeColor.RGB), CLng(pc.Shapes("tqq_ann_diff_1_lbl").Line.ForeColor.RGB), "borda do rotulo clonada"
    AssertEqual 1, CLng(seq.Count), "animacao mantida"
    AssertEqual CStr(shp.Name), CStr(seq.Item(1).Shape.Name)
    AssertAlmost w0, c.Parent.Width, 0.01
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set seq = Nothing: Set pc = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Colagem cuja sincronizacao falha (ganchos "update", "snapshot", "apply"): a colagem vale, as etiquetas ficam, a impressao digital NAO e
' gravada (a proxima atualizacao nao pode pular o grafico como "sem mudanca") e o grafico do Excel volta ao tamanho dele.
Public Sub Test_PptSync_PasteSyncFailure_NoFingerprint_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, slide As Object, c As Chart, shp As Object, pt As Variant
    Dim own As New Collection, en As Long, ed As String, w0 As Double, h0 As Double, n As Long
    On Error GoTo Cleanup
    If Not PasteSetup(app, own, slide, wb, rng, c, "pptsync-fp.xlsx") Then GoTo Cleanup
    w0 = c.Parent.Width: h0 = c.Parent.Height
    For Each pt In Array("update", "snapshot", "apply")
        n = n + 1
        PptSync_SetFailForTest CStr(pt)
        Core_CopyToPowerPoint "chart"
        AssertEqual n, CLng(slide.Shapes.Count), "a colagem continua valendo (" & pt & ")"
        Set shp = slide.Shapes(slide.Shapes.Count)
        AssertEqual Registry_IdFromChart(c), PptSync_Tag(shp, PPTSYNC_TAG_CHART), "etiquetas gravadas (" & pt & ")"
        AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_FP), "sem impressao digital: a sincronizacao falhou (" & pt & ")"
        AssertAlmost w0, c.Parent.Width, 0.01, "largura restaurada (" & pt & ")"
        AssertAlmost h0, c.Parent.Height, 0.01, "altura restaurada (" & pt & ")"
    Next
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set slide = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Foto da geometria do grafico do USUARIO (Excel): area interna e caixa de cada tqq_ann_*, depois de forcar o layout.
Private Function ExcelGeometry(c As Chart) As Collection
    Dim out As New Collection, i As Long, s As Shape
    PptSync_ForceLayout c
    out.Add Array("PlotArea", CDbl(c.PlotArea.InsideLeft), CDbl(c.PlotArea.InsideTop), CDbl(c.PlotArea.InsideWidth), CDbl(c.PlotArea.InsideHeight)), "PlotArea"
    For i = 1 To c.Shapes.Count
        Set s = c.Shapes(i)
        If Left$(s.Name, Len(ANN_PREFIX)) = ANN_PREFIX Then out.Add Array(s.Name, CDbl(s.Left), CDbl(s.Top), CDbl(s.Width), CDbl(s.Height)), s.Name
    Next
    Set ExcelGeometry = out
End Function

Private Sub AssertSameGeometry(before As Collection, c As Chart, stage As String)
    Dim now_ As Collection, g As Variant, h As Variant, k As Long
    Set now_ = ExcelGeometry(c)
    AssertEqual before.Count, now_.Count, "mesmas formas (" & stage & ")"
    For Each g In before
        h = now_(CStr(g(0)))
        For k = 1 To 4
            AssertAlmost CDbl(g(k)), CDbl(h(k)), 0.5, g(0) & " campo " & k & " (" & stage & ")"
        Next
    Next
End Sub

' O grafico do Excel do usuario fica como estava (area interna e setas, 0,5 pt) depois da colagem e de mais duas sincronizacoes: a volta ao
' tamanho original refaz o layout como a ida.
Public Sub Test_PptSync_ExcelChartGeometryUnchanged_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, slide As Object, c As Chart, shp As Object, geo As Collection
    Dim own As New Collection, en As Long, ed As String
    On Error GoTo Cleanup
    If Not PasteSetup(app, own, slide, wb, rng, c, "pptsync-geo.xlsx") Then GoTo Cleanup
    Set geo = ExcelGeometry(c)
    Core_CopyToPowerPoint "chart"
    AssertEqual 1, CLng(slide.Shapes.Count), "grafico colado"
    AssertSameGeometry geo, c, "depois da colagem"
    Set shp = slide.Shapes(1)
    PptSync_SyncShape shp, c.Parent
    PptSync_SyncShape shp, c.Parent
    AssertSameGeometry geo, c, "depois de mais duas sincronizacoes"
    AssertTrue Application.EnableEvents, "eventos religados depois das sincronizacoes"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set slide = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_PptSync_SourceReachable()
    Dim wb As Workbook, p As String
    Set wb = NewTempWorkbook()
    p = TempDir() & "\pptsync-reach.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    AssertTrue PptSync_SourceReachable(wb.FullName), "pasta aberta (e no disco)"
    AssertFalse PptSync_SourceReachable(TempDir() & "\pptsync-nao-existe.xlsx"), "arquivo inexistente"
    AssertFalse PptSync_SourceReachable("https://contoso.sharepoint.com/sites/x/Documentos/nao-aberta.xlsx"), "https: so conta pasta aberta"
    AssertFalse PptSync_SourceReachable(""), "vazio"
    AssertFalse PptSync_SourceReachable("   "), "so espacos"
    Teardown wb
End Sub

' LinkFormat.Update so com a pasta de origem no disco ou aberta (Update com a origem inexistente TRAVA o PowerPoint, spec 3.1) e so quando
' pedido (updateData). Duble: forma com grafico vinculado e LinkFormat que conta as chamadas; o gancho "update" para logo depois do passo.
Public Sub Test_PptSync_UpdateGuard_Fake()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, ch As TestPptFake, cd As TestPptFake, lf As TestPptFake, src As Variant
    Dim expected As Variant, k As Long, en As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-guard.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set cd = New TestPptFake: cd.IsLinked = True
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set lf = New TestPptFake
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432: Set shp.Chart = ch: Set shp.LinkFormat = lf
    AssertEqual 1, Ppt_LinkState(shp), "duble vinculado"
    src = Array(TempDir() & "\pptsync-nao-existe.xlsx", "https://contoso.sharepoint.com/x/nao-aberta.xlsx", wb.FullName)
    expected = Array(0, 0, 1)
    For k = 0 To 2
        lf.SourceFullName = CStr(src(k)): lf.UpdateCount = 0
        PptSync_SetFailForTest "update"
        On Error Resume Next
        PptSync_SyncShape shp, c.Parent
        en = Err.Number
        On Error GoTo 0
        AssertEqual vbObjectError + 9992, en, "parou no gancho update (" & src(k) & ")"
        AssertEqual CLng(expected(k)), lf.UpdateCount, "Update chamado? " & src(k)
    Next
    lf.SourceFullName = wb.FullName: lf.UpdateCount = 0
    PptSync_SetFailForTest "update"
    On Error Resume Next
    PptSync_SyncShape shp, c.Parent, False
    On Error GoTo 0
    AssertEqual 0, lf.UpdateCount, "updateData = False: sem Update"
    Teardown wb
End Sub

' PptSync_ForceLayout nao usa Dir$ (zeraria o laco Dir$ de quem chama) e apaga o PNG temporario (nome unico).
Public Sub Test_PptSync_ForceLayout_KeepsCallerDirLoopAndCleansUp()
    Dim wb As Workbook, rng As Range, c As Chart, d As String, a As String, b As String, f As Integer, nm As Variant, left_ As String
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    d = TempDir()
    For Each nm In Array("a.txt", "b.txt")
        f = FreeFile
        Open d & "\" & nm For Output As #f
        Print #f, "x"
        Close #f
    Next
    a = Dir$(d & "\*.txt")
    PptSync_ForceLayout c
    b = Dir$()
    AssertTrue Len(a) > 0 And Len(b) > 0 And a <> b, "o laco Dir$ do chamador continua (" & a & ", " & b & ")"
    left_ = Dir$(Environ$("TEMP") & "\thinqq_pptsync*.png")
    AssertEqual "", left_, "sem PNG temporario sobrando"
    Teardown wb
End Sub

Public Sub Test_PptSync_SummaryText()
    AssertEqual "Gráficos do ThinQQ atualizados: 2.", PptSync_SummaryText(2, 0, "", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 3.", PptSync_SummaryText(0, 3, "", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 1." & vbLf & "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): a.xlsx, b.xlsx. Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem.", PptSync_SummaryText(1, 0, "a.xlsx, b.xlsx", 0, 0, 0)
    AssertEqual "Gráficos do ThinQQ atualizados: 0." & vbLf & "Gráfico não encontrado na pasta de origem (apagado ou recriado): 2." & vbLf & "Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): 1." & vbLf & "Falha ao atualizar: 1 (veja o log do ThinQQ).", PptSync_SummaryText(0, 0, "", 2, 1, 1)
End Sub

Public Sub Test_Core_PptApiVersion()
    AssertEqual 7, Core_PptApiVersion(), "contrato com o suplemento do PowerPoint (spec 4.9; 2 = 1.6.0; 3 = 1.8.0; 4 = 1.9.0 comandos; 5 = 1.9.0 texto vinculado; 6 = 1.11.0 estados; 7 = 1.12.0 estilos)"
End Sub

Public Sub Test_Core_PptSyncPresentation_NeverRaises()
    Dim s As String
    s = Core_PptSyncPresentation(Nothing, True)
    AssertTrue InStr(s, "Não foi possível atualizar os gráficos do ThinQQ") = 1, s
End Sub

Public Sub Test_Core_PptSyncOpen_WithoutPowerPoint_Notifies()
    Ppt_ResetTestHooks                                         ' build Test sem instancia injetada: Ppt_GetApp devolve Nothing
    Core_ClearLastNotice
    Core_PptSyncOpen
    AssertEqual PPTSYNC_NO_PPT, Core_LastNotice()
End Sub

' Cola um grafico de colunas empilhadas com seta CAGR da pasta salva em p (uma apresentacao do teste, sem janela). Devolve a forma colada.
Private Function PasteLinked(app As Object, pres As Object, wb As Workbook, rng As Range, p As String) As Object
    Dim slide As Object, c As Chart
    Set slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Ppt_PinTargetForTest slide
    wb.SaveAs p, xlOpenXMLWorkbook
    rng.Worksheet.Activate
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    Ppt_SetAppForTest app
    Core_CopyToPowerPoint "chart"
    Set PasteLinked = slide.Shapes(slide.Shapes.Count)
End Function

' Sonda P-U6 (1.13.0, X2), medida NEGATIVA e fixada aqui: o grafico vinculado do slide guarda as proprias formulas de serie (celulas da tabela
' e do bloco na hora da colagem) e o LinkFormat.Update rele ESSAS celulas, nao o grafico do Excel pelo nome. Depois de "Alterar dados" (bloco
' novo, tabela com uma linha a mais) o slide continua nas celulas antigas: por isso Alterar dados mantem o id novo e o aviso
' LINKFRAME_NOTE_REBUILT. (Medido tambem com o id mantido: o Update devolveu "atualizados: 1." mas a serie de totais do slide ficou vazia, e
' Series.Formula do grafico do slide recusa escrita: "O metodo 'Formula' do objeto 'Series' falhou".) Se um Office futuro mudar isso, este
' teste avisa.
Public Sub Test_Rebuild_SlideKeepsOldCells_PU6_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String
    Dim f0 As String, c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pu6.xlsx")
    Set c = rng.Worksheet.ChartObjects(1).Chart
    f0 = c.SeriesCollection(1).Formula
    AssertEqual f0, CStr(s1.Chart.SeriesCollection(1).Formula), "o slide guarda a formula da serie do Excel"
    rng.Worksheet.Range("A4:C4").Value = Array("T3", 99, 8)
    AssertTrue Core_RebuildFromRange(c, rng.Worksheet.Range("A1:C4"))
    Set c = rng.Worksheet.ChartObjects(1).Chart
    AssertTrue c.SeriesCollection(1).Formula <> f0, "o grafico do Excel passou a ler a tabela nova"
    s1.LinkFormat.Update
    AssertEqual f0, CStr(s1.Chart.SeriesCollection(1).Formula), "P-U6 negativa: o slide continua nas celulas antigas"
    AssertEqual 2, UBound(s1.Chart.SeriesCollection(1).Values) - LBound(s1.Chart.SeriesCollection(1).Values) + 1, "sem a linha nova"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' 1.12.0 (S5): a versao do estilo (opcao stylever, na linha de metadados do bloco) entra na impressao digital: trocar o estilo marca o
' elemento do slide como "dados novos" (bandeira/automatico). Sem estilo a impressao nao muda (decks antigos nao viram "novos").
Public Sub Test_PptSync_Fingerprint_ChangesWithStyle()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, a As String, b As String
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432
    a = PptSync_Fingerprint(c.Parent, shp)
    Style_ApplyNamed c, "Cascata MC", True
    AssertTrue Len(Registry_GetOption(c, "stylever", "")) > 0, "stylever gravado"
    b = PptSync_Fingerprint(c.Parent, shp)
    AssertTrue a <> b, "estilo aplicado muda a impressao digital"
    AssertEqual b, PptSync_Fingerprint(c.Parent, shp), "sem mudanca: igual"
    Style_ApplyNamed c, "Cascata DDE", True
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "outro estilo muda a impressao digital"
    b = PptSync_Fingerprint(c.Parent, shp)
    Registry_SetOption c, "stylever", "00000000"
    AssertTrue b <> PptSync_Fingerprint(c.Parent, shp), "versao do estilo (texto editado) muda a impressao digital"
    Teardown wb
End Sub

' 1.12.0 (S5): a formatacao do estilo (cor das series e hachura da 1a barra) chega ao grafico do slide na atualizacao; comando "style" pelo PowerPoint.
Public Sub Test_PptSync_StyleFormat_ReachesSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptstyle.xlsx")
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Tons de cinza")
    AssertEqual "", msg
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Cascata MC")
    AssertEqual "", msg
    Set c = rng.Worksheet.ChartObjects(1).Chart
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual RGB(112, 173, 71), CLng(s1.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB), "serie1 do estilo no slide"
    AssertTrue InStr(Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Nao Existe"), "Estilo não encontrado") > 0
    AssertFalse wb.Saved, "pasta do usuario: alterada e NAO salva pelo comando (o usuario decide)"
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", ""), "estilo inexistente nao mexe no grafico"
    ' revisao final 1.12.0 (I3): tirar o estilo (paleta embutida) tambem chega ao slide
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Tons de cinza")
    AssertEqual "", msg
    AssertEqual "", Registry_GetOption(c, "style", "")
    AssertTrue CLng(c.SeriesCollection(1).Format.Fill.ForeColor.RGB) <> RGB(112, 173, 71), "Excel: a paleta cinza"
    AssertEqual CLng(c.SeriesCollection(1).Format.Fill.ForeColor.RGB), CLng(s1.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB), "slide: a cor da paleta cinza (nao ficou o verde do estilo)"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Revisao final 1.12.0 (I4): rosca com estilo de 3 cores -> as 3 fatias do slide com as 3 cores; linha com estilo -> a cor da linha no slide.
Public Sub Test_PptSync_StyleFormat_PieSlicesAndLines_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, own As New Collection, en As Long, ed As String
    Dim c As Chart, ws As Worksheet, slide As Object, k As Long, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook()
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Canal", "Vendas"): ws.Range("A2:B2").Value = Array("Loja", 55)
    ws.Range("A3:B3").Value = Array("Online", 30): ws.Range("A4:B4").Value = Array("Atacado", 15)
    ws.Range("D1:F1").Value = Array("Ano", "Sul", "Norte"): ws.Range("D2:F2").Value = Array("T1", 10, 5): ws.Range("D3:F3").Value = Array("T2", 20, 6)
    wb.SaveAs TempDir() & "\pptstyle-pie.xlsx", xlOpenXMLWorkbook
    ' rosca
    Set slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Ppt_PinTargetForTest slide
    ws.Range("A1:B4").Select
    Core_Insert "doughnut"
    Ppt_SetAppForTest app
    Core_CopyToPowerPoint "chart"
    Set s1 = slide.Shapes(slide.Shapes.Count)
    ' linha
    Set slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Ppt_PinTargetForTest slide
    ws.Range("D1:F3").Select
    Core_Insert "line"
    Ppt_SetAppForTest app
    Core_CopyToPowerPoint "chart"
    Set s2 = slide.Shapes(slide.Shapes.Count)
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), "style", "Cascata MC")
    AssertEqual "", msg
    msg = Core_PptCommand(pres, PptIdx_Entry(CLng(s2.Parent.SlideID), CStr(s2.Name), ""), "style", "Cascata MC")
    AssertEqual "", msg
    Set c = ws.ChartObjects(1).Chart
    For k = 1 To 3
        AssertEqual CLng(c.SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB), CLng(s1.Chart.SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB), "fatia " & k & " do slide com a cor do Excel"
    Next
    AssertEqual RGB(&HD9, &HD9, &HD9), CLng(s1.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB), "fatia do meio: serie2 do estilo"
    Set c = ws.ChartObjects(2).Chart
    AssertEqual RGB(112, 173, 71), CLng(c.SeriesCollection(1).Format.Line.ForeColor.RGB), "Excel: linha 1 com serie1"
    AssertEqual RGB(112, 173, 71), CLng(s2.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB), "slide: a cor da linha 1"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set s2 = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 3: pasta de origem inexistente (salva com outro nome e a original apagada): nenhum LinkFormat.Update (travaria), o grafico fica,
' o resumo diz o arquivo; tudo em menos de 10 s.
Public Sub Test_PptSync_MissingSource_DoesNotHang_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim a As String, s As String, t As Single, v0 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    a = TempDir() & "\pptsync-orig.xlsx"
    Set shp = PasteLinked(app, pres, wb, rng, a)
    v0 = PptSeriesValue(shp.Chart, "Sul", 1)
    wb.SaveAs TempDir() & "\pptsync-outro.xlsx", xlOpenXMLWorkbook
    Kill a
    t = Timer
    s = PptSync_Presentation(pres, True)
    AssertTrue Timer - t < 10, "sem travar (" & Format$(Timer - t, "0.0") & " s)"
    AssertTrue InStr(s, "Pasta de origem não encontrada ou sem acesso") > 0 And InStr(s, "pptsync-orig.xlsx") > 0, s
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 0.") = 1, s
    AssertAlmost v0, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "grafico ficou como estava"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 2: pasta fechada com dado novo salvo: aberta somente leitura, sincronizada e fechada (nenhuma pasta a mais aberta no fim).
Public Sub Test_PptSync_ClosedSource_OpenedReadOnlyAndClosed_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim s As String, n0 As Long, sec0 As Long
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-fechada.xlsx")
    rng.Worksheet.Range("B2").Value = 33
    wb.Save
    Watchers_Stop
    CloseTemp wb: Set wb = Nothing
    n0 = Application.Workbooks.Count
    sec0 = Application.AutomationSecurity
    s = PptSync_Presentation(pres, True)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertAlmost 33, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "dado novo da pasta fechada"
    AssertEqual n0, Application.Workbooks.Count, "a pasta aberta pela sincronizacao foi fechada"
    AssertEqual sec0, CLng(Application.AutomationSecurity), "AutomationSecurity restaurada depois de abrir a pasta"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Impressao digital: sem mudanca o grafico e pulado ("Sem mudança"), com dado novo e atualizado; forcar atualiza sempre. Core_PptSyncOpen
' (instancia injetada) cita a apresentacao no aviso.
Public Sub Test_PptSync_Fingerprint_SkipsUnchanged_AndCoreOpen_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String, s As String
    Dim c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-fp.xlsx")
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 1.", s
    rng.Worksheet.Range("C3").Value = 9
    Application.Calculate
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", s
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptSync_Presentation(pres, True), "forcar: atualiza mesmo sem mudanca"
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_PptSyncOpen
    AssertTrue InStr(Core_LastNotice(), "'" & CStr(pres.Name) & "': Gráficos do ThinQQ atualizados: 1.") > 0, Core_LastNotice()
    ' revisao final E (I2): titulo ligado a uma celula; mudar a celula chega ao slide sem forcar (como no ciclo automatico)
    Set c = rng.Worksheet.ChartObjects(rng.Worksheet.ChartObjects.Count).Chart
    rng.Worksheet.Range("H1").Value = "Receita 2026"
    LinkTitle_Apply c, rng.Worksheet.Range("H1")
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", s, "titulo ligado: atualiza sem forcar"
    rng.Worksheet.Range("H1").Value = "Receita 2027"
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", s, "celula do titulo mudou: atualiza sem forcar"
    AssertEqual "Receita 2027", CStr(shp.Chart.ChartTitle.Text), "o titulo chegou ao slide (copiado: o vinculo nativo nao o traz)"
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 1.", PptSync_Presentation(pres, False)
    LinkTitle_Apply c, Nothing
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptSync_Presentation(pres, False), "ligacao removida: atualiza"
    AssertFalse CBool(shp.Chart.HasTitle), "o titulo posto pelo ThinQQ sai do slide"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set c = Nothing: Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Adendo M8: caminho de pasta so "existe" se estiver aberta neste Excel (FullName igual, inclusive endereco https do OneDrive) ou for um caminho
' local/de rede existente no disco; https nao aberto nao e provavel (limite documentado: a pasta tem de estar aberta).
Public Sub Test_PptSync_FileExists()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.SaveAs TempDir() & "\pptsync-exists.xlsx", xlOpenXMLWorkbook
    AssertTrue PptSync_FileExists(wb.FullName), "arquivo no disco e aberto"
    AssertFalse PptSync_FileExists(TempDir() & "\pptsync-nao-existe.xlsx"), "inexistente"
    AssertFalse PptSync_FileExists("https://contoso.sharepoint.com/sites/x/Documentos/nao-aberta.xlsx"), "https nao aberto: nao provavel"
    AssertFalse PptSync_FileExists(""), "vazio"
    Teardown wb
End Sub

' Revisao final (C1, seguranca): a pasta de origem vem da apresentacao (de quem a fez). Aberta pela sincronizacao, ela nunca roda macros: uma
' .xlsm com Workbook_BeforeClose que grava um arquivo marcador, citada por uma forma marcada (TQQ_BOOK = ela, TQQ_CHART desconhecido) numa
' apresentacao falsa, e aberta e fechada (com os eventos ligados) sem o marcador aparecer; a AutomationSecurity do Excel volta ao valor anterior.
Public Sub Test_PptSync_HostileSourceMacros_NeverRun_Fake()
    Dim wb As Workbook, dir_ As String, xlsm As String, marker As String, cm As Object, pres As TestPptFake, sl As TestPptFake, shp As TestPptFake
    Dim ch As TestPptFake, cd As TestPptFake, tags As Object, sec0 As Long, prevEv As Boolean, s As String, n0 As Long, code As String
    dir_ = TempDir()
    xlsm = dir_ & "\pptsync-hostil.xlsm": marker = dir_ & "\marcador.txt"
    Set wb = NewTempWorkbook()
    code = "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    code = code & "    Dim f As Integer: f = FreeFile: Open """ & marker & """ For Output As #f: Print #f, ""macro rodou"": Close #f" & vbCrLf
    code = code & "End Sub"
    Set cm = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    cm.AddFromString code
    Set cm = Nothing
    wb.SaveAs xlsm, 52                                              ' xlOpenXMLWorkbookMacroEnabled
    prevEv = Application.EnableEvents: Application.EnableEvents = False
    wb.Close SaveChanges:=False                                     ' o proprio teste fecha sem disparar o BeforeClose
    Application.EnableEvents = prevEv
    Set wb = Nothing
    If Len(Dir(marker)) > 0 Then Kill marker
    AssertTrue Len(Dir(xlsm)) > 0, "a .xlsm hostil foi gravada"
    Set cd = New TestPptFake: cd.IsLinked = True
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add PPTSYNC_TAG_CHART, "tqq_nao_existe"
    tags.Add PPTSYNC_TAG_BOOK, xlsm
    Set shp = New TestPptFake: shp.Width = 400: shp.Height = 300: Set shp.Chart = ch: Set shp.Tags = tags
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp
    Set pres = New TestPptFake: Set pres.Slides = New Collection: pres.Slides.Add sl
    Application.EnableEvents = True
    sec0 = Application.AutomationSecurity
    n0 = Application.Workbooks.Count
    s = PptSync_Presentation(pres, True)
    Application.EnableEvents = prevEv
    AssertEqual 0, Len(Dir(marker)), "a macro da pasta de origem NAO rodou (" & s & ")"
    AssertEqual sec0, CLng(Application.AutomationSecurity), "AutomationSecurity restaurada"
    AssertEqual n0, Application.Workbooks.Count, "a pasta aberta pela sincronizacao foi fechada"
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 0.") = 1, s
End Sub

' Variante (re-revisao): so as macros desligadas ja bastam. Com o gancho de teste que deixa os eventos LIGADOS no fechamento, a macro
' Workbook_BeforeClose da .xlsm hostil continua sem rodar (a AutomationSecurity 3 do Open vale para a pasta inteira).
Public Sub Test_PptSync_HostileSourceMacros_MacrosOffAlone_Fake()
    Dim wb As Workbook, dir_ As String, xlsm As String, marker As String, cm As Object, pres As TestPptFake, sl As TestPptFake, shp As TestPptFake
    Dim ch As TestPptFake, cd As TestPptFake, tags As Object, prevEv As Boolean, s As String, code As String
    dir_ = TempDir()
    xlsm = dir_ & "\pptsync-hostil2.xlsm": marker = dir_ & "\marcador2.txt"
    Set wb = NewTempWorkbook()
    code = "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    code = code & "    Dim f As Integer: f = FreeFile: Open """ & marker & """ For Output As #f: Print #f, ""macro rodou"": Close #f" & vbCrLf
    code = code & "End Sub"
    Set cm = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    cm.AddFromString code
    Set cm = Nothing
    wb.SaveAs xlsm, 52
    prevEv = Application.EnableEvents: Application.EnableEvents = False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Set wb = Nothing
    If Len(Dir(marker)) > 0 Then Kill marker
    Set cd = New TestPptFake: cd.IsLinked = True
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add PPTSYNC_TAG_CHART, "tqq_nao_existe"
    tags.Add PPTSYNC_TAG_BOOK, xlsm
    Set shp = New TestPptFake: shp.Width = 400: shp.Height = 300: Set shp.Chart = ch: Set shp.Tags = tags
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp
    Set pres = New TestPptFake: Set pres.Slides = New Collection: pres.Slides.Add sl
    Application.EnableEvents = True
    PptSync_KeepEventsOnCloseForTest
    s = PptSync_Presentation(pres, True)
    PptSync_ResetTestHooks
    Application.EnableEvents = prevEv
    AssertEqual 0, Len(Dir(marker)), "eventos ligados no fechamento e mesmo assim a macro NAO rodou (" & s & ")"
End Sub

' Re-revisao da A1: o resultado "Update ignorado por origem inalcancavel" e visivel a quem chama (updateSkipped); com a pasta alcancavel ou
' updateData = False, False. O Update so roda quando nao foi ignorado.
Public Sub Test_PptSync_UpdateSkipped_IsReported_Fake()
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, ch As TestPptFake, cd As TestPptFake, lf As TestPptFake, src As Variant
    Dim expSkip As Variant, expUpd As Variant, k As Long, sk As Boolean
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptsync-skipped.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set cd = New TestPptFake: cd.IsLinked = True
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set lf = New TestPptFake
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432: Set shp.Chart = ch: Set shp.LinkFormat = lf
    src = Array(TempDir() & "\pptsync-nao-existe.xlsx", "https://contoso.sharepoint.com/x/nao-aberta.xlsx", wb.FullName)
    expSkip = Array(True, True, False)
    expUpd = Array(0, 0, 1)
    For k = 0 To 2
        lf.SourceFullName = CStr(src(k)): lf.UpdateCount = 0
        sk = Not CBool(expSkip(k))                                  ' valor inicial oposto: a rotina tem de gravar o resultado
        PptSync_SetFailForTest "update"
        On Error Resume Next
        PptSync_SyncShape shp, c.Parent, True, sk
        On Error GoTo 0
        AssertEqual CBool(expSkip(k)), sk, "updateSkipped (" & src(k) & ")"
        AssertEqual CLng(expUpd(k)), lf.UpdateCount, "Update chamado? " & src(k)
    Next
    lf.SourceFullName = TempDir() & "\pptsync-nao-existe.xlsx": lf.UpdateCount = 0
    sk = True
    PptSync_SetFailForTest "update"
    On Error Resume Next
    PptSync_SyncShape shp, c.Parent, False, sk
    On Error GoTo 0
    AssertFalse sk, "updateData = False: nada foi ignorado"
    Teardown wb
End Sub

' Re-revisao da A1: Update ignorado por origem inalcancavel (gancho "unreachable": a pasta some entre a conferencia e o Update): o grafico
' NAO conta como atualizado, a TQQ_FP NAO e gravada (a proxima atualizacao nao pode pular o grafico como "sem mudanca") e o resumo o lista
' em "Pasta de origem nao encontrada".
Public Sub Test_PptSync_UpdateSkipped_NoFingerprint_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim s As String, fp0 As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-skip.xlsx")
    fp0 = PptSync_Tag(shp, PPTSYNC_TAG_FP)
    AssertTrue Len(fp0) > 0, "a colagem gravou a impressao digital"
    rng.Worksheet.Range("C3").Value = 9
    Application.Calculate
    PptSync_ResetTestHooks
    PptSync_SetFailForTest "unreachable"
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 0.") = 1, s
    AssertTrue InStr(s, "Pasta de origem não encontrada ou sem acesso") > 0 And InStr(s, "pptsync-skip.xlsx") > 0, s
    AssertEqual 0, PptSync_UpdateCountForTest(), "nenhum LinkFormat.Update"
    AssertEqual fp0, PptSync_Tag(shp, PPTSYNC_TAG_FP), "TQQ_FP intacta: o Update foi ignorado"
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", s, "o grafico nao e pulado como sem mudanca"
    AssertTrue PptSync_Tag(shp, PPTSYNC_TAG_FP) <> fp0, "agora a impressao digital reflete o dado novo"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Cola de novo o grafico ativo do Excel em um slide novo da apresentacao pres (segunda forma com a mesma etiqueta TQQ_CHART). Devolve a forma.
Private Function PasteAgain(app As Object, pres As Object) As Object
    Dim slide As Object
    Set slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Ppt_PinTargetForTest slide
    Ppt_SetAppForTest app
    Core_CopyToPowerPoint "chart"
    Set PasteAgain = slide.Shapes(slide.Shapes.Count)
End Function

' Fix 1 (importante 1): pasta de origem protegida por senha (existe, mas nao abre): o Open nao pode abrir a caixa modal de senha (travaria o
' Application.Run); vai para a lista "sem acesso" e os outros graficos continuam sincronizando. O watchdog do teste pega um travamento.
Public Sub Test_PptSync_PasswordProtectedSource_ListedAsNoAccess_WhenAvailable()
    Dim wb As Workbook, wb2 As Workbook, rng As Range, rng2 As Range, app As Object, pres As Object, shp As Object, shp2 As Object
    Dim own As New Collection, en As Long, ed As String, s As String, p2 As String, t As Single, n0 As Long
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-aberta.xlsx")
    Set wb2 = NewTempWorkbook(): Set rng2 = Fixture(wb2)
    p2 = TempDir() & "\pptsync-senha.xlsx"
    Set shp2 = PasteLinked(app, pres, wb2, rng2, p2)
    Application.DisplayAlerts = False
    wb2.SaveAs p2, xlOpenXMLWorkbook, Password:="segredo"
    Application.DisplayAlerts = True
    Watchers_Stop
    CloseTemp wb2: Set wb2 = Nothing
    rng.Worksheet.Range("C3").Value = 9
    n0 = Application.Workbooks.Count
    t = Timer
    s = PptSync_Presentation(pres, True)
    AssertTrue Timer - t < 60, "sem travar (" & Format$(Timer - t, "0.0") & " s)"
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertTrue InStr(s, "Pasta de origem não encontrada ou sem acesso") > 0 And InStr(s, "pptsync-senha.xlsx") > 0, s
    AssertEqual n0, Application.Workbooks.Count, "nenhuma pasta a mais aberta"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.DisplayAlerts = True
    Set shp = Nothing: Set shp2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb2 Is Nothing Then Teardown wb2
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Fix 1 (importante 2): cada grafico e isolado em TODOS os passos (achar o grafico, impressao digital, gravar a TQQ_FP): a falha de um conta
' como "Falha" (ou, se so a gravacao da TQQ_FP falha depois de sincronizar, como atualizado, sem TQQ_FP velha) e o laco continua.
Public Sub Test_PptSync_PerChartIsolation_Hooks_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, shp2 As Object, own As New Collection, en As Long, ed As String
    Dim s As String, pt As Variant, expected As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-iso.xlsx")
    Set shp2 = PasteAgain(app, pres)
    AssertTrue Len(PptSync_Tag(shp, PPTSYNC_TAG_FP)) > 0 And Len(PptSync_Tag(shp2, PPTSYNC_TAG_FP)) > 0, "as duas colagens gravaram a impressao digital"
    For Each pt In Array("findchart", "fingerprint", "tagfp")
        PptSync_ResetTestHooks
        PptSync_SetFailForTest CStr(pt)
        s = PptSync_Presentation(pres, True)
        If CStr(pt) = "tagfp" Then
            expected = "Gráficos do ThinQQ atualizados: 2."
        Else
            expected = "Gráficos do ThinQQ atualizados: 1." & vbLf & "Falha ao atualizar: 1 (veja o log do ThinQQ)."
        End If
        AssertEqual expected, s, "falha em '" & pt & "': o outro grafico continua"
        AssertTrue Len(PptSync_Tag(shp2, PPTSYNC_TAG_FP)) > 0, "a segunda forma foi sincronizada (" & pt & ")"
    Next
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_FP), "TQQ_FP que nao pode ser regravada e removida (nao fica velha)"
    PptSync_ResetTestHooks
    s = PptSync_Presentation(pres, False)
    AssertEqual "Gráficos do ThinQQ atualizados: 1. Sem mudança: 1.", s, "a sem TQQ_FP e refeita, a outra e pulada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set shp2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Fix 1 (menor 10): a impressao digital gravada depois de uma atualizacao nao forcada e a que a proxima chamada calcula (o segundo
' PptSync_Presentation sem forcar diz "Sem mudanca"). Escolha: a TQQ_FP e calculada DEPOIS do PptSync_SyncShape tambem aqui (como na
' colagem, PptSyncAfterPaste): ela descreve o estado que o slide acabou de receber; a comparacao usa a de ANTES da sincronizacao.
Public Sub Test_PptSync_SecondNonForcedCall_ReportsNoChange_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String, fp1 As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-fp2.xlsx")
    rng.Worksheet.Range("C3").Value = 9
    Application.Calculate
    AssertEqual "Gráficos do ThinQQ atualizados: 1.", PptSync_Presentation(pres, False), "dado novo: atualiza"
    fp1 = PptSync_Tag(shp, PPTSYNC_TAG_FP)
    AssertTrue Len(fp1) > 0, "TQQ_FP gravada"
    AssertEqual "Gráficos do ThinQQ atualizados: 0. Sem mudança: 1.", PptSync_Presentation(pres, False), "segunda chamada sem forcar: sem mudanca"
    AssertEqual fp1, PptSync_Tag(shp, PPTSYNC_TAG_FP), "TQQ_FP intacta"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Fix 1 (menores 3 e 5): classificacao pura do estado do vinculo + caminho e a lista de pastas do aviso.
Public Sub Test_PptSync_Classify()
    AssertEqual PPTSYNC_C_PROCEED, PptSync_Classify(1, "C:\a\b.xlsx")
    AssertEqual PPTSYNC_C_UNLINKED, PptSync_Classify(0, "C:\a\b.xlsx"), "colado sem vinculo"
    AssertEqual PPTSYNC_C_UNKNOWN, PptSync_Classify(-1, "C:\a\b.xlsx"), "estado ilegivel: falha, nunca sincroniza"
    AssertEqual PPTSYNC_C_NOPATH, PptSync_Classify(1, ""), "sem SourceFullName e sem TQQ_BOOK"
End Sub

Public Sub Test_PptSync_MissingText()
    AssertEqual "a.xlsx, b.xlsx", PptSync_MissingText(Array("C:\x\a.xlsx", "C:\y\b.xlsx"))
    AssertEqual "Plano.xlsx", PptSync_MissingText(Array("https://contoso.sharepoint.com/sites/x/Documentos/Plano.xlsx")), "https: nome depois da ultima barra"
    AssertEqual "C:\x\a.xlsx, C:\y\A.xlsx, b.xlsx", PptSync_MissingText(Array("C:\x\a.xlsx", "C:\y\A.xlsx", "C:\z\b.xlsx")), "mesmo nome em pastas diferentes: caminho completo"
    AssertEqual "(origem desconhecida)", PptSync_MissingText(Array("")), "sem origem"
    AssertEqual "(origem desconhecida), a.xlsx", PptSync_MissingText(Array("", "C:\x\a.xlsx"))
End Sub

' M8 (menor 10): pasta aberta no Excel com endereco https conta como alcancavel (um Workbook https real e impossivel nos testes: a regra pura
' PptSync_ReachableAmong recebe a lista de FullName); https nao aberto nao.
Public Sub Test_PptSync_ReachableAmong_Https()
    Dim names As New Collection
    names.Add "C:\x\local.xlsx"
    names.Add "https://contoso.sharepoint.com/sites/x/Documentos/Aberta.xlsx"
    AssertTrue PptSync_ReachableAmong("https://contoso.sharepoint.com/sites/x/Documentos/Aberta.xlsx", names), "https aberto"
    AssertTrue PptSync_ReachableAmong("HTTPS://CONTOSO.SHAREPOINT.COM/sites/x/Documentos/aberta.xlsx", names), "sem diferenciar maiusculas"
    AssertFalse PptSync_ReachableAmong("https://contoso.sharepoint.com/sites/x/Documentos/Outra.xlsx", names), "https nao aberto"
    AssertFalse PptSync_ReachableAmong("", names), "vazio"
End Sub

' Ruling A3 (spec 4.4): o botao atualiza TODAS as apresentacoes abertas, inclusive somente leitura (so na memoria; so a atualizacao automatica ao abrir
' pula somente leitura, no suplemento). Apresentacao com o grafico vinculado salva, fechada e reaberta somente leitura (sem titulo, sem janela);
' o dado muda na pasta; Core_PptSyncOpen atualiza o grafico em memoria e o resumo cita a apresentacao.
Public Sub Test_Core_PptSyncOpen_ReadOnlyPresentation_IsUpdatedInMemory_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, ro As Object, shp As Object, own As New Collection, en As Long, ed As String
    Dim p As String, m As String, nm As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then own.Add pres: GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-ro-data.xlsx")
    p = TempDir() & "\pptsync-ro.pptx"
    pres.SaveAs p
    Set shp = Nothing
    pres.Close
    Set pres = Nothing
    Set ro = app.Presentations.Open(p, -1, 0, 0)               ' somente leitura, sem titulo, sem janela
    own.Add ro
    If UserWindowAppeared(app) Then GoTo Cleanup
    AssertTrue CLng(ro.ReadOnly) <> 0, "a apresentacao do teste esta somente leitura"
    Set shp = ro.Slides(1).Shapes(ro.Slides(1).Shapes.Count)
    AssertAlmost 10, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "dado antes"
    rng.Worksheet.Range("B2").Value = 77
    Application.Calculate
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_PptSyncOpen
    nm = CStr(ro.Name)
    m = Core_LastNotice()
    AssertTrue InStr(m, "'" & nm & "': Gráficos do ThinQQ atualizados: 1.") > 0, m
    AssertAlmost 77, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "somente leitura: atualizada em memoria"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set ro = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Fix 1 (menor 6): erro estrutural na apresentacao N (gancho "presentation", pula a 1a) nao descarta o resumo da 1..N-1: a linha de erro entra.
Public Sub Test_Core_PptSyncOpen_RaiseInOnePresentation_KeepsOthers_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, pres2 As Object, shp As Object, shp2 As Object
    Dim own As New Collection, en As Long, ed As String, n As String, n2 As String, m As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set pres2 = app.Presentations.Add(0): own.Add pres2
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptsync-open2.xlsx")
    Set shp2 = PasteAgain(app, pres2)
    n = CStr(pres.Name): n2 = CStr(pres2.Name)
    Ppt_SetAppForTest app
    PptSync_ResetTestHooks
    PptSync_SetFailAfterForTest "presentation", 1
    Core_ClearLastNotice
    Core_PptSyncOpen
    m = Core_LastNotice()
    AssertTrue InStr(m, "'" & n & "': Gráficos do ThinQQ atualizados: 1.") = 1, m
    AssertTrue InStr(m, "'" & n2 & "': Não foi possível atualizar os gráficos do ThinQQ") > 0, m
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set shp2 = Nothing: Set pres = Nothing: Set pres2 = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Parte C (1.6.0): indice das formas automaticas, PptSync_Items, API 2 ----

' Contrato com o .ppam (TestPptAddin.Test_PptIdx_Literal tem as MESMAS cadeias): entradas por vbLf, campos por "|" (Chr(30)/Chr(31) somem
' do .pptx salvo: Test_PptSync_IndexPersistsSaveReopen_WhenAvailable).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String, bad As Variant, g As String, c As Collection
    a = "256" & "|" & "tqq_abc" & "|" & "C:\a\b.xlsx"
    b = "257" & "|" & "tqq_def" & "|" & "C:\a\c.xlsx"
    s = PptIdx_Set("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & vbLf & b, s
    AssertEqual a & vbLf & b, PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptIdx_Set(s, 256, "tqq_abc", "", False), "desligar tira a entrada"
    AssertEqual "", PptIdx_Set(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptIdx_Parse(s & vbLf & "x" & "|" & "y" & "|" & "z" & vbLf & vbLf).Count, "entrada invalida e vazia ignoradas"
    AssertEqual a, PptIdx_Entry(256, "tqq_abc", "C:\a\b.xlsx")
    ' total para lixo da apresentacao: SlideID so com 1 a 9 algarismos; nada levanta erro
    For Each bad In Array("9e9", "1E10", "3d9", "&H10", "(5)", " 1", "-1", "1 ", "+1", "1234567890", "", "1.5", "1,5", ChrW$(1633))
        g = g & CStr(bad) & "|" & "n" & "|" & "" & vbLf
    Next
    AssertEqual 2, PptIdx_Parse(g & s).Count, "SlideID invalido ignorado"
    AssertEqual 0, PptIdx_Parse("|" & "|" & "|" & vbLf & "|").Count, "so separadores"
    AssertEqual 0, PptIdx_Parse("").Count
    AssertEqual 1, PptIdx_Parse("0" & "|" & "n" & "|").Count, "pasta vazia e SlideID 0 sao validos"
    Set c = PptIdx_Parse("999999999" & "|" & "n" & "|" & "b")
    AssertEqual 1, c.Count, "9 algarismos"
    AssertEqual 999999999, IdxSlide(c, 1)
    ' separadores nunca entram num campo
    AssertEqual "1" & "|" & "a_b" & "|" & "c_d", PptIdx_Entry(1, "a" & vbLf & "b", "c" & "|" & "d")
    AssertEqual "2|a_b_c_d|e_f", PptIdx_Entry(2, "a" & Chr$(30) & "b" & Chr$(31) & "c" & vbTab & "d", "e" & vbCr & "f"), "o que o .pptx salvo muda"
    AssertEqual 1, PptIdx_Parse(PptIdx_Set("", 1, "x" & vbLf & "y" & "|" & "z", "p" & vbLf, True)).Count, "nome com separadores: uma entrada"
    AssertEqual "", PptIdx_Set(PptIdx_Set("", 1, "x" & vbLf & "y", "", True), 1, "x" & vbLf & "y", "", False), "desligar acha o nome saneado"
End Sub

' Tetos do indice: entradas alem de PPTIDX_MAX_ENTRIES ignoradas; etiqueta acima de PPTIDX_MAX_LEN recusada inteira.
Public Sub Test_PptIdx_Caps()
    Dim a() As String, i As Long, big As String, c As Collection
    ReDim a(0 To 2499)
    For i = 0 To 2499
        a(i) = PptIdx_Entry(i + 1, "tqq_" & i, "")
    Next
    Set c = PptIdx_Parse(Join(a, vbLf))
    AssertEqual 2000, c.Count, "2500 entradas: as 2000 primeiras"
    AssertEqual 1, IdxSlide(c, 1)
    AssertEqual 2000, IdxSlide(c, 2000)
    AssertEqual 2000, PptIdx_Parse(PptIdx_Set(Join(a, vbLf), 9999, "nova", "", True)).Count
    For i = 0 To 2499
        a(i) = PptIdx_Entry(i + 1, "tqq_" & i, "C:\" & String$(480, "a") & ".xlsx")
    Next
    big = Join(a, vbLf)
    AssertTrue Len(big) > 1200000, CStr(Len(big))
    AssertEqual 0, PptIdx_Parse(big).Count, "etiqueta de 1,2 MB recusada"
    AssertEqual 0, PptIdx_Parse(Left$(big, PPTIDX_MAX_LEN + 1)).Count, "um caractere acima do teto"
    AssertTrue PptIdx_Parse(Left$(big, PPTIDX_MAX_LEN)).Count > 0, "no teto ainda le"
End Sub

Public Sub Test_PptSync_ItemKind_ByTags()
    Dim f As TestPptFake
    Set f = New TestPptFake
    AssertEqual "", PptSync_ItemKind(f), "duble sem Tags: nenhum tipo (sem erro)"
End Sub

Public Sub Test_PptAuto_Flags_TestBuildInMemory()
    Dim en As Long, ed As String
    On Error GoTo Cleanup
    PptAuto_ResetTestHooks
    AssertTrue PptAuto_Enabled(), "padrao: ligado"
    AssertFalse PptAuto_NewLinksAuto(), "padrao: vinculo nasce manual (como o think-cell)"
    Core_PptSetFlag "autonew", True
    AssertTrue Core_PptGetFlag("autonew")
    Core_PptSetFlag "autoon", False
    AssertFalse Core_PptGetFlag("autoon")
    AssertFalse Core_PptGetFlag("outra"), "chave desconhecida: False"
    PptAuto_ResetTestHooks
    AssertTrue PptAuto_Enabled()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    PptAuto_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_PptSync_SetAuto_IndexAndRebuild_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String, idx As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_ResetTestHooks
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-idx.xlsx")
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "nasce manual"
    PptSync_SetAuto shp, True
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    AssertEqual PptIdx_Entry(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp)), idx
    PptSync_TagDelete pres, PPTSYNC_TAG_INDEX
    PptSync_Presentation pres, True
    AssertEqual idx, PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "sincronizacao completa reconstroi o indice"
    PptSync_SetAuto shp, False
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    AssertEqual "", PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "desligar tira do indice"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' "Atualizar uma vez" so nos selecionados: duas colagens da mesma pasta, dado novo, so a primeira entrada sincronizada.
Public Sub Test_PptSync_SyncShapes_OnlyListed_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, own As New Collection, en As Long, ed As String
    Dim v2 As Double, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-sel.xlsx")
    Set s2 = PasteAgain(app, pres)
    v2 = PptSeriesValue(s2.Chart, "Sul", 1)
    rng.Worksheet.Range("B2").Value = 42
    Application.Calculate
    msg = Core_PptSyncShapes(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue InStr(msg, "Gráficos do ThinQQ atualizados: 1.") = 1, msg
    AssertAlmost 42, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001, "selecionado atualizado"
    AssertAlmost v2, PptSeriesValue(s2.Chart, "Sul", 1), 0.0001, "o outro ficou como estava"
    AssertEqual PPTSYNC_NOT_FOUND, Core_PptSyncShapes(pres, PptIdx_Entry(999999, "nada", ""), True), "entrada que nao existe"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s1 = Nothing: Set s2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' AutoNew ligado: a colagem ja nasce automatica e no indice.
Public Sub Test_PptSync_PasteWithAutoNew_TagsAndIndexes_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, shp As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_SetFlagsForTest Empty, True
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set shp = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-new.xlsx")
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO)
    AssertTrue InStr(PptSync_Tag(pres, PPTSYNC_TAG_INDEX), CStr(shp.Name)) > 0
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set shp = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Correcao 1 da C1 (revisao) ----

Private Function IdxSlide(c As Collection, i As Long) As Long
    Dim v As Variant
    v = c.Item(i)
    IdxSlide = CLng(v(0))
End Function

' Duble de forma do ThinQQ vinculada (para PptSync_Items sem PowerPoint): TQQ_CHART chartId, TQQ_BOOK book, vinculo legivel e ativo.
Private Function FakeLinkedShape(chartId As String, book As String) As TestPptFake
    Dim cd As TestPptFake, ch As TestPptFake, shp As TestPptFake, tags As Object
    Set cd = New TestPptFake: cd.IsLinked = True
    Set ch = New TestPptFake: Set ch.ChartData = cd
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add PPTSYNC_TAG_CHART, chartId
    tags.Add PPTSYNC_TAG_BOOK, book
    Set shp = New TestPptFake: shp.Width = 400: shp.Height = 300: Set shp.Chart = ch: Set shp.Tags = tags
    Set FakeLinkedShape = shp
End Function

' A mesma forma listada duas vezes entra uma vez so; entradas de slide inexistente ou forma inexistente sao ignoradas.
Public Sub Test_PptSync_ResolveEntries_Dedup_Fake()
    Dim shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, e As String
    Set shp = FakeLinkedShape("tqq_x", "C:\nao\existe.xlsx")
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_x"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 256
    Set pres = New TestPptFake: Set pres.Slides = slides
    e = PptIdx_Entry(256, "tqq_x", "")
    AssertEqual 1, PptSync_ResolveEntries(pres, e).Count
    AssertEqual 1, PptSync_ResolveEntries(pres, e & vbLf & e & vbLf & PptIdx_Entry(256, "tqq_x", "C:\outra.xlsx")).Count, "mesma forma uma vez so"
    AssertEqual 0, PptSync_ResolveEntries(pres, PptIdx_Entry(257, "tqq_x", "") & vbLf & PptIdx_Entry(256, "outra", "")).Count
End Sub

' PowerPoint ocupado no meio da sincronizacao (PptSync_SyncOne): sobe ate a API, que devolve PPT_ERR_BUSY na hora (sem laco), e a pasta de
' origem aberta pela sincronizacao e fechada. Vale para a API 2 (Core_PptSyncShapes) e a API 1 (Core_PptSyncPresentation).
Public Sub Test_PptSync_Busy_PropagatesToApi_ClosesBook_Fake()
    Dim wb As Workbook, p As String, shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, presA As TestPptFake, presB As TestPptFake
    Dim n0 As Long, s As String, t As Single, en As Long, ed As String
    On Error GoTo Cleanup
    p = TempDir() & "\pptsync-ocupado.xlsx"
    Set wb = NewTempWorkbook()
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set shp = FakeLinkedShape("tqq_x", p)
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_x"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 256
    Set presA = New TestPptFake: Set presA.Slides = slides
    Set presB = New TestPptFake: Set presB.Slides = New Collection: presB.Slides.Add sl
    n0 = Application.Workbooks.Count
    s = Core_PptSyncShapes(presA, PptIdx_Entry(256, "tqq_x", ""), True)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 0.") = 1, "sem gancho: a pasta abre e o grafico nao existe nela (" & s & ")"
    AssertEqual n0, Application.Workbooks.Count, "sem gancho: pasta fechada"
    PptSync_SetFailNumberForTest "findchart", -2147418111
    t = Timer
    s = Core_PptSyncShapes(presA, PptIdx_Entry(256, "tqq_x", ""), True)
    AssertEqual PPT_ERR_BUSY, s, "API 2"
    AssertTrue Timer - t < 10, "sem laco"
    AssertEqual n0, Application.Workbooks.Count, "API 2: pasta aberta pela sincronizacao foi fechada"
    PptSync_SetFailNumberForTest "findchart", -2147417846
    s = Core_PptSyncPresentation(presB, True)
    AssertEqual PPT_ERR_BUSY, s, "API 1"
    AssertEqual n0, Application.Workbooks.Count, "API 1: pasta fechada"
    PptSync_SetFailForTest "findchart"
    s = Core_PptSyncPresentation(presB, True)
    AssertTrue InStr(s, "Falha ao atualizar: 1") > 0, "outro erro continua contado por forma (" & s & ")"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_ResetTestHooks
    Application.StatusBar = False
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Indice reconstruido com PowerPoint real (sem Excel): entrada velha sai, nomes repetidos viram unicos ("_2"), nome com separadores e saneado,
' uma forma que falha (duble sem Parent) nao derruba as outras; apresentacao sem formas do ThinQQ perde o indice velho.
Public Sub Test_PptSync_RebuildIndex_StaleUniqueIsolated_WhenAvailable()
    Dim app As Object, pres As Object, pres2 As Object, sl As Object, a As Object, b As Object, c As Object, own As New Collection
    Dim f As TestPptFake, items As New Collection, idx As String, en As Long, ed As String, nmC As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set sl = pres.Slides.Add(1, 12)
    Set a = sl.Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_x"
    Set b = sl.Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "tqq_x"
    Set c = sl.Shapes.AddShape(1, 130, 10, 50, 50): c.Name = "tqq_c"
    On Error Resume Next
    c.Name = "tqq" & Chr$(30) & "c" & Chr$(31) & "d"
    nmC = CStr(c.Name)
    Err.Clear
    On Error GoTo Cleanup
    Errors_Log "TestPptSync: nome com separadores aceito pelo PowerPoint: " & CStr(InStr(nmC, Chr$(30)) > 0)
    a.Tags.Add PPTSYNC_TAG_CHART, "c1": a.Tags.Add PPTSYNC_TAG_BOOK, "C:\t\a.xlsx": a.Tags.Add PPTSYNC_TAG_AUTO, "1"
    b.Tags.Add PPTSYNC_TAG_CHART, "c2": b.Tags.Add PPTSYNC_TAG_AUTO, "1"
    c.Tags.Add PPTSYNC_TAG_CHART, "c3": c.Tags.Add PPTSYNC_TAG_AUTO, "1"
    pres.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Entry(999, "velha", "")
    Set f = New TestPptFake: Set f.Tags = CreateObject("Scripting.Dictionary"): f.Tags.Add PPTSYNC_TAG_AUTO, "1"
    items.Add f: items.Add a: items.Add b: items.Add c
    PptSync_RebuildIndex pres, items
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    AssertEqual 3, PptIdx_Parse(idx).Count, idx
    AssertEqual 0, InStr(idx, "velha"), "entrada velha saiu"
    AssertEqual "tqq_x_2", CStr(a.Name), "nome repetido: _2"
    AssertEqual "tqq_x", CStr(b.Name)
    AssertEqual 0, InStr(CStr(c.Name), Chr$(30)) + InStr(CStr(c.Name), Chr$(31)), "sem separadores no nome"
    AssertEqual PptIdx_Entry(CLng(sl.SlideID), CStr(a.Name), "C:\t\a.xlsx") & vbLf & PptIdx_Entry(CLng(sl.SlideID), "tqq_x", "") & vbLf & PptIdx_Entry(CLng(sl.SlideID), CStr(c.Name), ""), idx
    Set pres2 = app.Presentations.Add(0): own.Add pres2
    pres2.Slides.Add 1, 12
    pres2.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Entry(256, "velha", "")
    AssertEqual PPTSYNC_NONE, PptSync_Presentation(pres2, True)
    AssertEqual "", PptSync_Tag(pres2, PPTSYNC_TAG_INDEX), "sem formas do ThinQQ: indice velho apagado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set items = New Collection: Set f = Nothing     ' solta as formas do slide ANTES do Quit (senao o processo do PowerPoint fica vivo)
    Set a = Nothing: Set b = Nothing: Set c = Nothing: Set sl = Nothing: Set pres = Nothing: Set pres2 = Nothing
    StopScratchPowerPoint app, own
    PptSync_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Medido na C3: Shapes(nome) do PowerPoint ignora a caixa ("tqq_A" e "tqq_a" no slide: os dois nomes acham a primeira). O nome unico do indice
' compara sem diferenciar maiusculas: depois de reconstruir, cada entrada resolve a sua forma (duas formas, nao uma).
Public Sub Test_PptSync_EnsureUniqueName_CaseOnly_WhenAvailable()
    Dim app As Object, pres As Object, sl As Object, a As Object, b As Object, own As New Collection
    Dim items As New Collection, got As Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set sl = pres.Slides.Add(1, 12)
    Set a = sl.Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_A"
    Set b = sl.Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "tqq_a"
    Errors_Log "TestPptSync: Shapes(""tqq_a"") acha a forma ""tqq_A"" (busca ignora a caixa): " & CStr(CLng(sl.Shapes("tqq_a").Id) = CLng(a.Id))
    a.Tags.Add PPTSYNC_TAG_CHART, "c1": a.Tags.Add PPTSYNC_TAG_AUTO, "1"
    b.Tags.Add PPTSYNC_TAG_CHART, "c2": b.Tags.Add PPTSYNC_TAG_AUTO, "1"
    items.Add a: items.Add b
    PptSync_RebuildIndex pres, items
    AssertTrue StrComp(CStr(a.Name), CStr(b.Name), vbTextCompare) <> 0, "nomes unicos sem diferenciar maiusculas: " & a.Name & " / " & b.Name
    Set got = PptSync_ResolveEntries(pres, PptSync_Tag(pres, PPTSYNC_TAG_INDEX))
    AssertEqual 2, got.Count, "cada entrada acha a sua forma"
    AssertEqual CLng(a.Id), CLng(got(1).Id)
    AssertEqual CLng(b.Id), CLng(got(2).Id)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set items = New Collection: Set got = Nothing
    Set a = Nothing: Set b = Nothing: Set sl = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    PptSync_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Importante 4 (revisao final da C): o indice TQQ_AUTOIDX sobrevive a Salvar como .pptx, fechar e reabrir no PowerPoint real (as etiquetas
' vao para ppt/tags/tagN.xml), inclusive a entrada de uma forma cujo nome tinha separadores (saneado ao ligar). Medido na revisao: com os
' separadores antigos Chr(30)/Chr(31) o indice voltava sem eles (43 -> 38 caracteres) e uma etiqueta "x<Tab>y<LF>z<CR>w|v" voltava
' "x y<LF>z<LF>w|v" (Tab vira espaco, CR vira LF; LF e "|" ficam). As sondas das etiquetas de forma ficam registradas no log.
Public Sub Test_PptSync_IndexPersistsSaveReopen_WhenAvailable()
    Dim app As Object, pres As Object, p2 As Object, sl As Object, a As Object, b As Object, c As Object, own As New Collection, nmC As String
    Dim dir_ As String, tmp As String, idx As String, idx2 As String, probe As String, probe2 As String, got As Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set sl = pres.Slides.Add(1, 12)
    Set a = sl.Shapes.AddShape(1, 10, 10, 50, 50): a.Name = "tqq_a"
    Set b = sl.Shapes.AddShape(1, 70, 10, 50, 50): b.Name = "tqq_b"
    Set c = sl.Shapes.AddShape(1, 130, 10, 50, 50): c.Name = "tqq_c"
    On Error Resume Next
    c.Name = "tqq|c" & vbLf & "d" & Chr$(30) & "e"              ' o PowerPoint pode recusar parte disto: o nome que ficou e saneado ao ligar
    If CStr(c.Name) = "tqq_c" Then c.Name = "tqq|c" & Chr$(30) & "e"
    Err.Clear
    On Error GoTo Cleanup
    nmC = CStr(c.Name)
    Errors_Log "TestPptSync: nome hostil aceito pelo PowerPoint: " & Replace(Replace(nmC, vbLf, "<LF>"), Chr$(30), "<30>")
    AssertTrue nmC <> "tqq_c", "premissa: o nome hostil foi aceito (" & nmC & ")"
    AssertTrue nmC <> PptIdx_SafeField(nmC), "premissa: o nome tem caracteres a sanear"
    a.Tags.Add PPTSYNC_TAG_CHART, "c1": a.Tags.Add PPTSYNC_TAG_BOOK, "C:\t\a.xlsx"
    b.Tags.Add PPTSYNC_TAG_CHART, "c2": b.Tags.Add PPTSYNC_TAG_BOOK, "C:\t\b.xlsx"
    c.Tags.Add PPTSYNC_TAG_CHART, "c3": c.Tags.Add PPTSYNC_TAG_BOOK, "C:\t\c|d.xlsx"
    probe = "x" & Chr$(30) & "y" & Chr$(31) & "z"
    b.Tags.Add "TQQ_PROBE", probe
    b.Tags.Add "TQQ_PROBE2", "x" & vbTab & "y" & vbLf & "z" & vbCr & "w|v"
    PptSync_SetAuto a, True
    PptSync_SetAuto b, True
    PptSync_SetAuto c, True
    nmC = CStr(c.Name)
    AssertEqual nmC, PptIdx_SafeField(nmC), "nome saneado ao ligar"
    AssertTrue InStr(nmC, "tqq_c") = 1, "o nome saneado vem do hostil (" & nmC & ")"
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    AssertEqual 3, PptIdx_Parse(idx).Count, "antes de salvar"
    dir_ = TempDir()
    tmp = dir_ & "\pptidx-persist.pptx"
    pres.SaveAs tmp
    Set a = Nothing: Set b = Nothing: Set c = Nothing: Set sl = Nothing
    pres.Close
    Set pres = Nothing
    Set own = New Collection
    Set p2 = app.Presentations.Open(tmp, -1, 0, 0): own.Add p2
    idx2 = PptSync_Tag(p2, PPTSYNC_TAG_INDEX)
    probe2 = PptSync_Tag(p2.Slides(1).Shapes("tqq_b"), "TQQ_PROBE")
    Errors_Log "TestPptSync: TQQ_AUTOIDX depois de salvar e reabrir igual ao gravado: " & CStr(idx2 = idx) & " (" & Len(idx) & "/" & Len(idx2) & " caracteres); etiqueta de forma com Chr(30)/Chr(31) igual: " & CStr(probe2 = probe) & " (" & Len(probe2) & " de " & Len(probe) & ")"
    Errors_Log "TestPptSync: etiqueta de forma x<TAB>y<LF>z<CR>w|v depois de reabrir: " & Replace(Replace(Replace(PptSync_Tag(p2.Slides(1).Shapes("tqq_b"), "TQQ_PROBE2"), vbTab, "<TAB>"), vbLf, "<LF>"), vbCr, "<CR>")
    AssertEqual idx, idx2, "indice igual depois de reabrir"
    AssertEqual 3, PptIdx_Parse(idx2).Count, "tres entradas depois de reabrir"
    Set got = PptSync_ResolveEntries(p2, idx2)
    AssertEqual 3, got.Count, "as tres entradas acham as formas"
    AssertEqual "tqq_a", CStr(got(1).Name)
    AssertEqual "tqq_b", CStr(got(2).Name)
    AssertEqual nmC, CStr(got(3).Name), "a forma do nome saneado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set got = Nothing: Set a = Nothing: Set b = Nothing: Set c = Nothing: Set sl = Nothing: Set pres = Nothing: Set p2 = Nothing
    StopScratchPowerPoint app, own
    If Len(tmp) > 0 Then Kill tmp
    PptSync_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Parte C2: motor automatico ----

' Estado limpo do motor: nada agendado, fila vazia, opcoes e relogio padrao, sem OnTime real (o teste chama PptAuto_Tick).
Private Sub AutoReset()
    PptAuto_Cancel
    PptAuto_ResetTestHooks
    PptAuto_SetNoTimerForTest True
End Sub

' Comeco de um teste do motor: guarda se os eventos do Excel estavam ligados (outro modulo pode te-los ligado) e os desliga, para nenhum
' evento marcar a pasta sem o teste pedir; AutoEnd devolve o estado de antes.
Private Sub AutoBegin()
    mAutoWasActive = Watchers_IsActive()
    Watchers_Stop
    AutoReset
End Sub

Private Sub AutoEnd()
    AutoReset
    Watchers_Stop
    If mAutoWasActive Then Watchers_Start
    mAutoWasActive = False
End Sub

' Pasta SALVA com grafico do ThinQQ (planilha _ThinQQ). Devolve a pasta; rng = dados.
Private Function AutoBook(fileName As String, ByRef rng As Range) As Workbook
    Dim wb As Workbook
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    wb.SaveAs TempDir() & "\" & fileName, xlOpenXMLWorkbook
    Set AutoBook = wb
End Function

' Duble de apresentacao com uma forma automatica (slide 256, forma "tqq_x") no indice, apontando para book.
Private Function FakeAutoPres(shp As TestPptFake, book As String) As TestPptFake
    Dim sl As TestPptFake, slides As TestPptFake, pres As TestPptFake
    shp.Tags.Add PPTSYNC_TAG_AUTO, "1"
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_x"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 256
    Set pres = New TestPptFake: Set pres.Slides = slides
    Set pres.Tags = CreateObject("Scripting.Dictionary")
    pres.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Entry(256, "tqq_x", book)
    Set FakeAutoPres = pres
End Function

' Duble do aplicativo com as apresentacoes dadas (sem janela ativa).
Private Function FakeApp(pres As TestPptFake) As TestPptFake
    Dim app As TestPptFake
    Set app = New TestPptFake: Set app.Presentations = New Collection
    If Not pres Is Nothing Then app.Presentations.Add pres
    Set FakeApp = app
End Function

' So pastas SALVAS com graficos do ThinQQ (_ThinQQ) agendam; desligado nada agenda.
Public Sub Test_PptAuto_OnChange_Filters()
    Dim wb As Workbook, wb2 As Workbook, rng As Range
    AutoBegin
    Set wb2 = NewTempWorkbook()
    PptAuto_OnChange wb2
    AssertEqual 0, PptAuto_PendingCountForTest(), "pasta nova sem ThinQQ e sem arquivo"
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    PptAuto_OnChange wb
    AssertEqual 0, PptAuto_PendingCountForTest(), "pasta nao salva (o vinculo exige o arquivo)"
    wb.SaveAs TempDir() & "\pptauto-filter.xlsx", xlOpenXMLWorkbook
    PptAuto_SetFlagsForTest False, Empty
    PptAuto_OnChange wb
    AssertEqual 0, PptAuto_PendingCountForTest(), "desligado"
    PptAuto_SetFlagsForTest Empty, Empty
    PptAuto_OnChange wb
    AssertEqual 1, PptAuto_PendingCountForTest()
    AssertTrue PptAuto_IsScheduledForTest(), "ciclo agendado"
    PptAuto_OnChange wb
    AssertEqual 1, PptAuto_PendingCountForTest(), "a mesma pasta conta uma vez"
    wb2.SaveAs TempDir() & "\pptauto-filter-sem.xlsx", xlOpenXMLWorkbook
    PptAuto_OnChange wb2
    AssertEqual 1, PptAuto_PendingCountForTest(), "pasta salva sem graficos do ThinQQ"
    AutoReset
    AssertFalse PptAuto_IsScheduledForTest(), "Cancel desagenda"
    AssertEqual 0, PptAuto_PendingCountForTest()
    CloseTemp wb2
    Teardown wb
    AutoEnd
End Sub

' Eventos reais do Excel: digitar na pasta salva agenda o ciclo.
Public Sub Test_PptAuto_SheetChangeEvent_Schedules()
    Dim wb As Workbook, rng As Range
    AutoBegin
    Set wb = AutoBook("pptauto-evt.xlsx", rng)
    Watchers_Start
    PptAuto_Cancel: PptAuto_SetNoTimerForTest True
    rng.Worksheet.Range("B2").Value = 77
    AssertEqual 1, PptAuto_PendingCountForTest(), "SheetChange/SheetCalculate marcaram a pasta"
    AssertTrue PptAuto_IsScheduledForTest()
    AutoReset
    AssertFalse PptAuto_IsScheduledForTest()
    Teardown wb
    AutoEnd
End Sub

' Sem PowerPoint injetado (build Test: Ppt_GetApp devolve so a instancia injetada) o ciclo esvazia a fila sem erro; opcao desligada
' depois da mudanca: o ciclo so esvazia a fila.
Public Sub Test_PptAuto_Tick_NoPowerPoint_ClearsQueue()
    Dim wb As Workbook, rng As Range
    AutoBegin
    Ppt_ResetTestHooks
    Set wb = AutoBook("pptauto-noppt.xlsx", rng)
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertEqual 0, PptAuto_PendingCountForTest()
    AssertEqual 0, PptAuto_RetriesForTest()
    AssertFalse PptAuto_IsScheduledForTest(), "nada reagendado"
    PptAuto_OnChange wb
    PptAuto_SetFlagsForTest False, Empty
    PptAuto_Tick
    AssertEqual 0, PptAuto_PendingCountForTest(), "desligado no meio: fila esvaziada"
    AssertFalse PptAuto_IsScheduledForTest()
    Teardown wb
    AutoEnd
End Sub

' PowerPoint ocupado (gancho): nova tentativa agendada, fila mantida; depois de PPTAUTO_MAX_RETRIES desiste e esvazia.
Public Sub Test_PptAuto_Busy_Retries()
    Dim wb As Workbook, rng As Range, i As Long
    AutoBegin
    Set wb = AutoBook("pptauto-busy.xlsx", rng)
    PptAuto_OnChange wb
    PptAuto_SetBusyForTest 99
    PptAuto_Tick
    AssertEqual 1, PptAuto_RetriesForTest()
    AssertEqual 1, PptAuto_PendingCountForTest(), "fila mantida para a nova tentativa"
    AssertTrue PptAuto_IsScheduledForTest()
    For i = 2 To PPTAUTO_MAX_RETRIES + 1
        PptAuto_Tick
    Next
    AssertEqual 0, PptAuto_PendingCountForTest(), "desistiu"
    AssertEqual 0, PptAuto_RetriesForTest()
    AssertFalse PptAuto_IsScheduledForTest(), "depois de desistir nada fica agendado"
    Teardown wb
    AutoEnd
End Sub

' Review Focus 1: eventos durante o ciclo (o proprio refresh) nao marcam a pasta; um disparo durante o ciclo nao entra e deixa o ciclo
' reagendado (a marca nunca fica sem OnTime). O "ciclo em andamento" e simulado pelo gancho PptAuto_SetInTickForTest (o mesmo indicador
' que PptAuto_Tick liga); o caminho real (refresh do proprio ciclo) e coberto por Test_PptAuto_Tick_UpdatesOnlyAutoShapes_WhenAvailable.
Public Sub Test_PptAuto_NoReentryDuringTick()
    Dim wb As Workbook, rng As Range
    AutoBegin
    Set wb = AutoBook("pptauto-reentry.xlsx", rng)
    PptAuto_SetInTickForTest True
    PptAuto_OnChange wb
    AssertEqual 0, PptAuto_PendingCountForTest(), "evento do proprio ciclo ignorado"
    PptAuto_SetInTickForTest False
    PptAuto_OnChange wb
    AssertEqual 1, PptAuto_PendingCountForTest()
    PptAuto_SetInTickForTest True
    PptAuto_Tick
    AssertEqual 1, PptAuto_PendingCountForTest(), "ciclo em andamento: o segundo disparo nao entra"
    AssertTrue PptAuto_IsScheduledForTest(), "e a marca continua agendada"
    PptAuto_SetInTickForTest False
    Teardown wb
    AutoEnd
End Sub

' OnTime REAL (opcao do teste): reagendar em outro segundo cancela o OnTime antigo (cancelar o horario antigo de novo falha: nada empilhou)
' e deixa UM OnTime vivo (o cancelamento do horario vivo da certo uma vez e a segunda falha); Core_OnUnload cancela o OnTime real (o
' OnTime reabriria o core descarregado). Nada fica pendente depois do teste.
Public Sub Test_PptAuto_RealTimer_NoStacking_AndUnload()
    Dim wb As Workbook, rng As Range, en As Long, ed As String, d1 As Date, d2 As Date, d3 As Date, d4 As Date, s0 As Date, t As Single
    On Error GoTo Cleanup
    AutoBegin
    Set wb = AutoBook("pptauto-timer.xlsx", rng)
    PptAuto_SetNoTimerForTest False
    PptAuto_OnChange wb
    d1 = PptAuto_DueForTest()
    AssertTrue d1 <> 0, "agendado"
    s0 = Now: t = Timer
    Do While Now <= s0                                      ' proximo segundo do relogio (o OnTime vence em 2 s: nao dispara aqui)
        If Timer < t Or Timer - t > 1.5 Then Exit Do
        DoEvents
    Loop
    PptAuto_OnChange wb
    d2 = PptAuto_DueForTest()
    AssertTrue d2 > d1, "o horario novo e de outro segundo"
    PptAuto_OnChange wb
    d3 = PptAuto_DueForTest()
    AssertEqual 0, PptAuto_TimerFailuresForTest(), "reagendar cancelou o OnTime vivo"
    AssertFalse PptAuto_CancelAtForTest(d1), "o OnTime antigo ja tinha sido cancelado (nada empilhou)"
    If d3 <> d2 Then AssertFalse PptAuto_CancelAtForTest(d2), "o do meio tambem"
    AssertTrue PptAuto_CancelAtForTest(d3), "o OnTime vivo existe"
    AssertFalse PptAuto_CancelAtForTest(d3), "uma vez so"
    PptAuto_SetNoTimerForTest False
    PptAuto_OnChange wb
    d4 = PptAuto_DueForTest()
    AssertTrue d4 <> 0
    Core_OnUnload
    AssertFalse PptAuto_IsScheduledForTest(), "Core_OnUnload desagenda"
    AssertEqual 0, PptAuto_PendingCountForTest()
    AssertEqual 0, PptAuto_TimerFailuresForTest(), "cancelamento pelo mesmo horario"
    AssertFalse PptAuto_CancelAtForTest(d4), "o OnTime real foi cancelado pelo Core_OnUnload"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' So as formas automaticas das pastas sujas (sem diferenca de caixa no caminho), uma vez cada; o vinculo real da forma manda sobre a pasta do
' indice (entrada velha apontando para a pasta suja nao sincroniza outra pasta); sem TQQ_AUTO fica fora. Importante 1 da revisao final: a
' entrada de uma pasta que nao mudou nao custa nenhuma chamada ao PowerPoint (nem FindBySlideID); a da pasta suja continua resolvida.
Public Sub Test_PptAuto_ItemsFor_OnlyDirtyBooks_Fake()
    Dim shp As TestPptFake, pres As TestPptFake, d As Object
    Set shp = FakeLinkedShape("tqq_x", "C:\t\a.xlsx")
    Set pres = FakeAutoPres(shp, "C:\t\a.xlsx")
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    d.Add "C:\t\b.xlsx", True
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count, "outra pasta"
    AssertEqual 0, pres.Slides.FindBySlideIDCalls, "outra pasta: nenhuma chamada ao PowerPoint"
    d.Add "c:\T\A.xlsx", True
    AssertEqual 1, PptAuto_ItemsFor(pres, d).Count, "pasta suja"
    AssertEqual 1, pres.Slides.FindBySlideIDCalls, "a entrada da pasta suja e resolvida"
    pres.Tags.Item(PPTSYNC_TAG_INDEX) = PptIdx_Entry(256, "tqq_x", "C:\t\a.xlsx") & vbLf & PptIdx_Entry(256, "tqq_x", "C:\t\a.xlsx")
    AssertEqual 1, PptAuto_ItemsFor(pres, d).Count, "a mesma forma uma vez so"
    d.Remove "c:\T\A.xlsx"
    pres.Tags.Item(PPTSYNC_TAG_INDEX) = PptIdx_Entry(256, "tqq_x", "C:\t\b.xlsx")
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count, "indice velho: a forma esta vinculada a a.xlsx, nao a b.xlsx"
    d.Add "C:\t\a.xlsx", True
    pres.Tags.Item(PPTSYNC_TAG_INDEX) = PptIdx_Entry(256, "tqq_x", "C:\t\a.xlsx")
    shp.Tags.Remove PPTSYNC_TAG_AUTO
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count, "sem TQQ_AUTO"
    shp.Tags.Item(PPTSYNC_TAG_AUTO) = "1"                   ' (a leitura do Dictionary ja recriou a chave vazia)
    shp.Tags.Remove PPTSYNC_TAG_CHART
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count, "so TQQ_AUTO (sem TQQ_CHART): nao e do ThinQQ, nunca chega a PptSync_Items"
End Sub

' Menor (b) da revisao final: o indice so e regravado quando uma entrada sai (velha ou repetida, inclusive a mesma forma com outra caixa no
' nome); um ciclo sem entrada velha nao grava nada; a entrada de uma pasta que nao mudou fica como esta (nao e lida).
Public Sub Test_PptAuto_ItemsFor_IndexRewriteOnlyWhenRemoved_Fake()
    Dim shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, tg As TestPptTags, d As Object, eA As String, eB As String
    Set shp = FakeLinkedShape("tqq_x", "C:\t\a.xlsx"): shp.Tags.Add PPTSYNC_TAG_AUTO, "1"
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_x"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 256
    Set tg = New TestPptTags
    Set pres = New TestPptFake: Set pres.Slides = slides: Set pres.Tags = tg
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    d.Add "C:\t\a.xlsx", True
    eA = PptIdx_Entry(256, "tqq_x", "C:\t\a.xlsx")
    eB = PptIdx_Entry(256, "tqq_y", "C:\t\b.xlsx")                 ' pasta que nao mudou (a forma nem existe: so a reconstrucao completa a tira)
    tg.Preset PPTSYNC_TAG_INDEX, eA & vbLf & eB
    AssertEqual 1, PptAuto_ItemsFor(pres, d).Count
    AssertEqual 0, tg.Writes + tg.Deletes, "nenhuma entrada saiu: o indice nao e regravado"
    AssertEqual 1, slides.FindBySlideIDCalls, "so a entrada da pasta suja e resolvida"
    tg.Preset PPTSYNC_TAG_INDEX, eA & vbLf & eB & vbLf & eB & vbLf & eA & vbLf & PptIdx_Entry(256, "TQQ_X", "C:\t\a.xlsx")
    AssertEqual 1, PptAuto_ItemsFor(pres, d).Count, "a mesma forma uma vez so"
    AssertEqual 1, tg.Writes, "repetidas: o indice e regravado uma vez"
    AssertEqual eA & vbLf & eB, tg.Item(PPTSYNC_TAG_INDEX), "sem as repetidas; a entrada da outra pasta fica"
    tg.Writes = 0
    tg.Preset PPTSYNC_TAG_INDEX, eB & vbLf & PptIdx_Entry(256, "sumiu", "C:\t\a.xlsx")
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count
    AssertEqual 1, tg.Writes
    AssertEqual eB, tg.Item(PPTSYNC_TAG_INDEX), "a entrada velha da pasta suja sai"
    tg.Preset PPTSYNC_TAG_INDEX, PptIdx_Entry(256, "sumiu", "C:\t\a.xlsx")
    AssertEqual 0, PptAuto_ItemsFor(pres, d).Count
    AssertEqual 1, tg.Deletes, "nada sobrou: a etiqueta e apagada"
End Sub

' Leitura recusada pelo PowerPoint ocupado (gancho com o numero do erro) no indice, na TQQ_AUTO ou no vinculo da forma: sobe como ocupado
' (nova tentativa), nunca vira "entrada velha" nem "pasta de outra origem".
Public Sub Test_PptAuto_ItemsFor_BusyReads_Raise_Fake()
    Dim shp As TestPptFake, pres As TestPptFake, d As Object, pts As Variant, i As Long, n As Long, en As Long, ed As String
    On Error GoTo Cleanup
    Set shp = FakeLinkedShape("tqq_x", "C:\t\a.xlsx")
    Set pres = FakeAutoPres(shp, "C:\t\a.xlsx")
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    d.Add "C:\t\a.xlsx", True
    pts = Array("autoidx", "autotag", "autobook")
    For i = 0 To 2
        PptSync_SetFailNumberForTest CStr(pts(i)), -2147418111
        On Error Resume Next
        n = PptAuto_ItemsFor(pres, d).Count
        en = Err.Number
        On Error GoTo Cleanup
        AssertEqual -2147418111, en, "ocupado sobe (" & pts(i) & ")"
        AssertTrue InStr(PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "tqq_x") > 0, "a entrada continua no indice (" & pts(i) & ")"
    Next
    PptSync_ResetTestHooks
    AssertEqual 1, PptAuto_ItemsFor(pres, d).Count, "sem gancho: a forma entra"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 5: o ciclo nunca abre pasta. Com a pasta aberta o ciclo chega a forma; fechada depois da mudanca, sai da fila sem ser aberta.
Public Sub Test_PptAuto_Tick_NeverOpensClosedBook_Fake()
    Dim wb As Workbook, rng As Range, p As String, shp As TestPptFake, pres As TestPptFake, app As TestPptFake, n0 As Long, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Set wb = AutoBook("pptauto-closed.xlsx", rng)
    p = wb.FullName
    Set shp = FakeLinkedShape("tqq_x", p)
    Set pres = FakeAutoPres(shp, p)
    Set app = FakeApp(pres)
    Ppt_SetAppForTest app
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertTrue InStr(PptAuto_LastResultForTest(), "Gráficos do ThinQQ atualizados: 0.") = 1, "pasta aberta: o ciclo chegou a forma (" & PptAuto_LastResultForTest() & ")"
    AssertEqual 1, PptAuto_LogCountForTest(), "grafico nao encontrado: registrado"
    PptAuto_OnChange wb
    wb.Close SaveChanges:=False
    Set wb = Nothing
    n0 = Application.Workbooks.Count
    PptAuto_Tick
    AssertEqual "", PptAuto_LastResultForTest(), "pasta fechada: nada sincronizado (nem aberta)"
    AssertEqual 0, PptAuto_PendingCountForTest()
    AssertEqual n0, Application.Workbooks.Count
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    Application.StatusBar = False
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Importante 1 (revisao): PowerPoint aberto sem nenhuma forma automatica: o ciclo so le ReadOnly e TQQ_AUTOIDX das apresentacoes; nenhuma
' leitura da janela (guarda de texto em edicao), nenhuma nova tentativa nem registro, fila esvaziada. Controle: com uma forma automatica da
' pasta e texto "em edicao" (a selecao do duble e um grafico do Excel: Type = 3), a guarda roda e adia.
Public Sub Test_PptAuto_Tick_NoAutoShapes_NoPptWindowCalls_Fake()
    Dim wb As Workbook, rng As Range, s As Shape, sel As Shape, win As TestPptFake, presNo As TestPptFake, presIdx As TestPptFake, app As TestPptFake
    Dim shp As TestPptFake, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Set wb = AutoBook("pptauto-noidx.xlsx", rng)
    For Each s In rng.Worksheet.Shapes
        If s.Type = 3 Then Set sel = s: Exit For
    Next
    AssertTrue Not sel Is Nothing, "grafico do Excel como selecao do duble"
    Set win = New TestPptFake: Set win.Selection = sel
    Set presNo = New TestPptFake: Set presNo.Tags = CreateObject("Scripting.Dictionary")
    Set app = FakeApp(presNo): Set app.ActiveWindow = win
    app.ActiveWindowReads = 0
    Ppt_SetAppForTest app
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertEqual 0, app.ActiveWindowReads, "sem formas automaticas: nenhuma leitura da janela do PowerPoint"
    AssertEqual 0, PptAuto_RetriesForTest(), "nenhuma nova tentativa"
    AssertEqual 0, PptAuto_PendingCountForTest(), "fila esvaziada"
    AssertFalse PptAuto_IsScheduledForTest()
    AssertEqual 0, PptAuto_LogCountForTest(), "nada registrado"
    Set shp = FakeLinkedShape("tqq_x", wb.FullName)
    Set presIdx = FakeAutoPres(shp, wb.FullName)
    app.Presentations.Add presIdx
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertTrue app.ActiveWindowReads > 0, "com forma automatica a guarda de texto em edicao roda"
    AssertEqual 1, PptAuto_RetriesForTest(), "texto em edicao: nova tentativa"
    AssertEqual 1, PptAuto_PendingCountForTest()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Ppt_ResetTestHooks
    Set win = Nothing: Set sel = Nothing: Set s = Nothing
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Menor 2 (revisao): pasta que recalcula sem parar (uma mudanca por "segundo" do relogio do gancho durante 25 s, como RTD): o ciclo dispara
' mesmo assim, uns 10 s depois da primeira marca (o OnTime vivo deixa de ser adiado). Mesmo segundo: nenhum cancelamento/agendamento novo.
Public Sub Test_PptAuto_Storm_CycleStillRuns()
    Dim wb As Workbook, rng As Range, t0 As Date, now_ As Date, s As Long, fired As Long, d As Date
    AutoBegin
    Ppt_ResetTestHooks
    Set wb = AutoBook("pptauto-storm.xlsx", rng)
    t0 = DateSerial(2026, 1, 5) + TimeSerial(9, 0, 0)
    fired = -1
    For s = 0 To 25
        now_ = t0 + TimeSerial(0, 0, s)
        PptAuto_SetNowForTest now_
        If PptAuto_IsScheduledForTest() Then
            If DateDiff("s", PptAuto_DueForTest(), now_) >= 0 Then
                PptAuto_Tick
                If fired < 0 Then fired = s
            End If
        End If
        PptAuto_OnChange wb
    Next
    AssertTrue fired >= 10 And fired <= 12, "o ciclo disparou em " & fired & " s"
    PptAuto_SetNowForTest t0
    PptAuto_Cancel: PptAuto_SetNoTimerForTest True
    PptAuto_SetNowForTest t0
    PptAuto_OnChange wb
    d = PptAuto_DueForTest()
    PptAuto_OnChange wb
    AssertEqual CDbl(d), CDbl(PptAuto_DueForTest()), "mesmo segundo: o horario vivo fica"
    Teardown wb
    AutoEnd
End Sub

' Menor 3 (revisao): o texto que outro recurso pos na barra de status volta depois do ciclo que sincronizou.
Public Sub Test_PptAuto_Tick_RestoresOtherStatusText_Fake()
    Dim wb As Workbook, rng As Range, shp As TestPptFake, pres As TestPptFake, en As Long, ed As String, v As Variant
    On Error GoTo Cleanup
    AutoBegin
    Set wb = AutoBook("pptauto-status.xlsx", rng)
    Set shp = FakeLinkedShape("tqq_x", wb.FullName)
    Set pres = FakeAutoPres(shp, wb.FullName)
    Ppt_SetAppForTest FakeApp(pres)
    Application.StatusBar = "outro recurso (teste)"
    PptAuto_OnChange wb
    PptAuto_Tick
    AssertTrue Len(PptAuto_LastResultForTest()) > 0, "o ciclo sincronizou"
    AssertEqual "outro recurso (teste)", CStr(Application.StatusBar), "texto de outro recurso devolvido"
    Application.StatusBar = False
    PptAuto_OnChange wb
    PptAuto_Tick
    v = Application.StatusBar                           ' o Excel invisivel do teste le False como o texto "FALSE" (TestDeltaPick)
    AssertTrue VarType(v) = vbBoolean Or CStr(v) = "FALSE", "sem texto antes: a barra volta ao Excel (" & CStr(v) & ")"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.StatusBar = False
    AutoReset
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Importante 2 da revisao final: dentro do ciclo automatico (OnTime) Core_Notify e Core_Confirm so registram (nenhuma caixa modal no Excel
' do usuario); a mensagem continua visivel para os testes (Core_LastNotice). Fora do ciclo a caixa abriria (contador do build Test).
Public Sub Test_PptAuto_Tick_NoMessageBoxInsideTick()
    Dim wb As Workbook, rng As Range, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Ppt_ResetTestHooks
    Set wb = AutoBook("pptauto-nobox.xlsx", rng)
    Core_ClearLastNotice
    PptAuto_OnChange wb
    PptAuto_SetNotifyInTickForTest "aviso de dentro do ciclo (teste)"
    PptAuto_Tick
    AssertFalse PptAuto_InTick(), "o ciclo terminou"
    AssertEqual "aviso de dentro do ciclo (teste)", Core_LastNotice(), "a mensagem foi registrada"
    AssertEqual "aviso de dentro do ciclo (teste)", Core_LastConfirm()
    AssertEqual 1, Core_NoticeCount()
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa dentro do ciclo"
    Core_Notify "aviso fora do ciclo (teste)"
    Core_Confirm "pergunta fora do ciclo (teste)"
    AssertEqual 2, Core_ShownCountForTest(), "fora do ciclo as caixas abririam"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Core_ClearLastNotice
    AutoReset
    Ppt_ResetTestHooks
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Menor (a) da revisao final: Ppt_GetApp(False) com o GetObject recusado pelo PowerPoint ocupado (caixa de dialogo aberta) levanta o erro
' em vez de dizer "sem PowerPoint": o ciclo tenta de novo (fila mantida) e o botao do Excel diz que o PowerPoint esta ocupado. Outro erro
' do GetObject continua sendo "sem PowerPoint"; com create = True nada muda.
Public Sub Test_Ppt_GetApp_BusyRaises_Fake()
    Dim wb As Workbook, rng As Range, app As TestPptFake, o As Object, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Set app = FakeApp(Nothing)
    Ppt_SetAppForTest app
    PptSync_SetFailNumberForTest "getapp", -2147418111
    On Error Resume Next
    Set o = Ppt_GetApp(False)
    en = Err.Number
    On Error GoTo Cleanup
    AssertEqual -2147418111, en, "ocupado sobe"
    PptSync_SetFailForTest "getapp"
    Set o = Ppt_GetApp(False)
    AssertTrue o Is Nothing, "outro erro: sem PowerPoint, sem erro"
    PptSync_SetFailNumberForTest "getapp", -2147417846
    Set o = Ppt_GetApp(True)
    AssertTrue o Is Nothing, "create = True: como antes (sem erro)"
    PptSync_ResetTestHooks
    AssertTrue Ppt_GetApp(False) Is app, "sem gancho: a instancia injetada"
    Set wb = AutoBook("pptauto-getapp.xlsx", rng)
    PptAuto_OnChange wb
    PptSync_SetFailNumberForTest "getapp", -2147418111
    PptAuto_Tick
    AssertEqual 1, PptAuto_RetriesForTest(), "o ciclo tenta de novo"
    AssertEqual 1, PptAuto_PendingCountForTest(), "fila mantida"
    Core_ClearLastNotice
    PptSync_SetFailNumberForTest "getapp", -2147418111
    Core_PptSyncOpen
    AssertEqual PPT_ERR_BUSY, Core_LastNotice(), "botao do Excel: PowerPoint ocupado (nao 'abra a apresentacao')"
    AssertFalse Ppt_InOperation()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set o = Nothing
    Core_ClearLastNotice
    AutoReset
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Menor (c) da revisao final: "Atualizar uma vez" com entradas que nao acham nenhuma forma (renomeadas, apagadas) diz isso; sem entradas,
' o aviso de apresentacao sem graficos do ThinQQ.
Public Sub Test_Core_PptSyncShapes_NotFoundMessage_Fake()
    Dim shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake
    Set shp = FakeLinkedShape("tqq_x", "C:\nao\existe.xlsx")
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_x"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 256
    Set pres = New TestPptFake: Set pres.Slides = slides
    AssertEqual PPTSYNC_NOT_FOUND, Core_PptSyncShapes(pres, PptIdx_Entry(256, "renomeada", "")), "entradas sem forma"
    AssertEqual "Os gráficos selecionados não foram encontrados na apresentação (renomeados?): selecione-os de novo.", PPTSYNC_NOT_FOUND
    AssertEqual PPTSYNC_NONE, Core_PptSyncShapes(pres, ""), "sem entradas"
    AssertFalse Ppt_InOperation()
End Sub

' Menor 5 (revisao): PptSync_Items com noOpen nunca abre a pasta: fechada, ela entra no aviso de pasta nao encontrada; sem noOpen (padrao) a
' pasta e aberta somente leitura (o grafico do duble nao existe nela).
Public Sub Test_PptSync_Items_NoOpen_ListsMissing_Fake()
    Dim wb As Workbook, p As String, items As New Collection, n0 As Long, s As String, en As Long, ed As String
    On Error GoTo Cleanup
    p = TempDir() & "\pptauto-noopen.xlsx"
    Set wb = NewTempWorkbook()
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing
    items.Add FakeLinkedShape("tqq_x", p)
    n0 = Application.Workbooks.Count
    s = PptSync_Items(items, False, True)
    AssertTrue InStr(s, "Pasta de origem não encontrada") > 0 And InStr(s, "pptauto-noopen.xlsx") > 0, s
    AssertEqual n0, Application.Workbooks.Count
    s = PptSync_Items(items, False)
    AssertTrue InStr(s, "Gráfico não encontrado") > 0, "padrao: abriu a pasta (" & s & ")"
    AssertEqual n0, Application.Workbooks.Count, "e fechou"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.StatusBar = False
    PptSync_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Menor 7 (revisao): o ciclo nao roda no meio de uma operacao do core com o PowerPoint (colagem, sincronizacao pedida): reagenda. As
' entradas do core fecham a marca em todo caminho.
Public Sub Test_PptAuto_Tick_DuringCoreOperation_Reschedules()
    Dim wb As Workbook, rng As Range, pres As TestPptFake
    AutoBegin
    Ppt_ResetTestHooks
    Set wb = AutoBook("pptauto-op.xlsx", rng)
    PptAuto_OnChange wb
    Ppt_BeginOp
    PptAuto_Tick
    AssertEqual 1, PptAuto_PendingCountForTest(), "operacao em andamento: o ciclo espera"
    AssertTrue PptAuto_IsScheduledForTest(), "e fica agendado"
    Ppt_EndOp
    AssertFalse Ppt_InOperation()
    PptAuto_Tick
    AssertEqual 0, PptAuto_PendingCountForTest()
    rng.Worksheet.Range("E10").Select
    Core_CopyToPowerPoint "chart"
    AssertFalse Ppt_InOperation(), "colagem sem grafico ativo: marca fechada"
    Core_PptSyncOpen
    AssertFalse Ppt_InOperation(), "Atualizar graficos sem PowerPoint: marca fechada"
    Set pres = New TestPptFake
    Core_PptSyncPresentation pres, True
    AssertFalse Ppt_InOperation(), "API 1 com erro: marca fechada"
    Core_PptSyncShapes pres, PptIdx_Entry(256, "x", ""), True
    AssertFalse Ppt_InOperation(), "API 2 com erro: marca fechada"
    Teardown wb
    AutoEnd
End Sub

' Review Focus 1: so a forma automatica muda; entrada de forma apagada sai do indice; o refresh do proprio ciclo nao reagenda.
Public Sub Test_PptAuto_Tick_UpdatesOnlyAutoShapes_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, s3 As Object, own As New Collection
    Dim en As Long, ed As String, v2 As Double
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-tick.xlsx")
    Set s2 = PasteAgain(app, pres)
    Set s3 = PasteAgain(app, pres)
    PptSync_SetAuto s1, True
    PptSync_SetAuto s3, True
    s3.Delete: Set s3 = Nothing
    v2 = PptSeriesValue(s2.Chart, "Sul", 1)
    Watchers_Start
    rng.Worksheet.Range("B2").Value = 55                 ' evento real: marca a pasta (pasta aberta, nao salva: o vinculo le o valor novo)
    AssertEqual 1, PptAuto_PendingCountForTest()
    Ppt_SetAppForTest app
    PptAuto_Tick
    AssertAlmost 55, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001, "automatica atualizada"
    AssertAlmost v2, PptSeriesValue(s2.Chart, "Sul", 1), 0.0001, "manual ficou"
    AssertEqual 1, PptIdx_Parse(PptSync_Tag(pres, PPTSYNC_TAG_INDEX)).Count, "entrada da forma apagada saiu do indice"
    AssertEqual 0, PptAuto_PendingCountForTest(), "o refresh do proprio ciclo nao marcou a pasta de novo"
    AssertFalse PptAuto_IsScheduledForTest(), "nada agendado depois do ciclo"
    AssertTrue InStr(PptAuto_LastResultForTest(), "Gráficos do ThinQQ atualizados: 1.") = 1, PptAuto_LastResultForTest()
    AssertEqual 1, PptAuto_LogCountForTest(), "atualizacao registrada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set s2 = Nothing: Set s3 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Sem mudanca nos dados (so recalculo): a impressao digital pula; nada reescrito e nada registrado (menor 6).
Public Sub Test_PptAuto_Tick_SkipsUnchanged_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptauto-same.xlsx")
    PptSync_SetAuto s1, True
    PptAuto_OnChange wb
    Ppt_SetAppForTest app
    PptSync_ResetTestHooks
    PptAuto_Tick
    AssertTrue InStr(PptAuto_LastResultForTest(), "Sem mudança: 1.") > 0, PptAuto_LastResultForTest()
    AssertEqual 0, PptSync_UpdateCountForTest(), "nenhum LinkFormat.Update"
    AssertEqual 0, PptAuto_LogCountForTest(), "sem mudanca: nada registrado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Parte D2: inserir a partir do intervalo direto no slide ----

Public Sub Test_InsertToPpt_Unsaved_NothingCreated()
    Dim wb As Workbook, rng As Range, n0 As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    n0 = rng.Worksheet.ChartObjects.Count
    rng.Select
    Core_ClearLastNotice
    Core_InsertToPowerPoint "stacked"
    AssertEqual PPT_ERR_INSERT_UNSAVED, Core_LastNotice()
    AssertEqual n0, rng.Worksheet.ChartObjects.Count, "nenhum grafico criado"
    AssertEqual 0, rng.Worksheet.Shapes.Count - rng.Worksheet.ChartObjects.Count, "nenhuma moldura"
    Teardown wb
End Sub

Public Sub Test_InsertToPpt_NoPowerPoint_FrameRemoved()
    Dim wb As Workbook, rng As Range
    Ppt_ResetTestHooks                                   ' build Test sem instancia injetada: Ppt_GetApp devolve Nothing
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\pptins-noppt.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_ClearLastNotice
    Core_InsertToPowerPoint "clustered"
    AssertEqual 1, rng.Worksheet.ChartObjects.Count, "o grafico fica no Excel"
    AssertFalse Core_LastPptPasteOk()
    AssertTrue LinkFrame_Find(rng.Worksheet.ChartObjects(1).Chart) Is Nothing, "sem colagem, sem moldura"
    AssertEqual PPT_ERR_NOT_INSTALLED, Core_LastNotice()
    Teardown wb
End Sub

Public Sub Test_InsertToPpt_ChartFrameAndSlide_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, shp As Object, c As Chart
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Ppt_SetAppForTest app
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    wb.SaveAs TempDir() & "\pptins-ok.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_InsertToPowerPoint "stacked"                    ' o teste confere a integracao, nao o tipo
    Set c = rng.Worksheet.ChartObjects(1).Chart
    AssertTrue Core_LastPptPasteOk()
    AssertEqual 1, slide.Shapes.Count
    Set shp = slide.Shapes(1)
    AssertEqual Registry_IdFromChart(c), PptSync_Tag(shp, PPTSYNC_TAG_CHART)
    AssertEqual 1, Ppt_LinkState(shp)
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "moldura no intervalo"
    ' X2 (1.13.0): moldura com estado
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB, "verde: a apresentacao aberta usa o grafico"
    AssertEqual 1, Core_RefreshFrames(wb), "uma verde"
    shp.Delete: Set shp = Nothing
    AssertEqual 0, Core_RefreshFrames(wb), "forma apagada: nenhuma apresentacao usa"
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB, "vermelha"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Parte G: grafico criado pelo PowerPoint ----

Public Sub Test_PptInsertChart_UnsavedPres_NothingCreated_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    AssertEqual PPTDATA_MSG_UNSAVED, Core_PptInsertChart(slide, "stacked")
    AssertEqual 0, slide.Shapes.Count
    AssertEqual "Tipo de gráfico desconhecido: xyz", Core_PptInsertChart(slide, "xyz")
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 6: criar pelo PowerPoint, digitar no "datasheet" (planilha de dados), o ciclo leva ao slide e salva a pasta.
Public Sub Test_PptInsertChart_FromPpt_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim shp As Object, s2 As Object, p As String, pp As String, wb As Workbook, ws As Worksheet
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    pp = TempDir()                                       ' TempDir() cria uma pasta nova a cada chamada: as duas na MESMA pasta
    p = pp & "\Vendas G - dados ThinQQ.xlsx": pp = pp & "\Vendas G.pptx"
    If Len(Dir(p)) > 0 Then Kill p
    pres.SaveAs pp
    PptSync_ResetTestHooks                               ' contador de LinkFormat.Update zerado aqui (nao depende de outro teste)
    msg = Core_PptInsertChart(slide, "clustered")
    AssertEqual "", msg
    AssertEqual 0, PptSync_UpdateCountForTest(), "a colagem nunca chama LinkFormat.Update (PG4: origem ausente trava o PowerPoint)"
    AssertEqual 1, slide.Shapes.Count
    Set shp = slide.Shapes(1)
    AssertEqual "1", PptSync_Tag(shp, PPTDATA_TAG)
    AssertEqual "", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "nasce manual, como o think-cell (1.11.0)"
    AssertEqual 1, Ppt_LinkState(shp)
    AssertEqual LCase$(p), LCase$(PptSync_ItemBook(shp))
    AssertEqual "Gráfico 1", PptSync_Tag(shp, PPTSYNC_TAG_SHEET)
    AssertEqual LCase$(p), LCase$(PptSync_Tag(pres, PPTDATA_MARK))
    AssertTrue InStr(1, PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CStr(shp.Name), vbBinaryCompare) > 0, "no indice de todos os vinculos"
    AssertEqual "", PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "nenhum automatico"
    Set wb = Workbooks(Dir(p))
    AssertTrue PptData_IsDataBook(wb), "pasta de dados marcada"
    AssertTrue wb.Saved, "pasta de dados salva"
    Set ws = wb.Worksheets("Gráfico 1")
    AssertTrue Not LinkFrame_Find(ws.ChartObjects(1).Chart) Is Nothing, "dados destacados pela moldura"
    Watchers_Start
    ws.Range("B2").Value = 999
    AssertTrue PptAuto_PendingCountForTest() >= 1
    Ppt_SetAppForTest app
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(shp, PPTSTATE_TAG), "bandeira: o slide nao acompanha sozinho"
    AssertTrue wb.Saved, "o ciclo salvou a pasta de dados (G3)"
    AssertTrue InStr(Core_PptSyncShapes(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""), True), "Gráficos do ThinQQ atualizados: 1.") = 1
    AssertAlmost 999, PptSeriesValue(shp.Chart, "Sul", 1), 0.0001, "Atualizar traz o dado"
    AssertEqual "", Core_PptInsertChart(slide, "pie")
    AssertEqual 2, slide.Shapes.Count
    Set s2 = slide.Shapes(2)
    AssertEqual "Gráfico 2", PptSync_Tag(s2, PPTSYNC_TAG_SHEET)
    ' G3: Editar dados e Salvar
    PptSync_ResetTestHooks                               ' o ciclo acima atualizou o grafico (arquivo existe); conta so o que vem agora
    Core_ClearLastNotice
    msg = Core_PptEditData(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""))
    AssertTrue Left$(msg, 1) <> "!", msg
    AssertEqual CStr(Application.ActiveWindow.Caption), msg, "devolve a legenda da janela"
    AssertFalse Core_IsQuiet(), "modo silencioso fechado depois de Editar dados"
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa durante a chamada do PowerPoint"
    AssertEqual 0, PptSync_UpdateCountForTest(), "Editar dados nunca chama LinkFormat.Update"
    AssertEqual "Gráfico 1", ActiveSheet.Name, "Editar dados ativa a planilha do grafico"
    AssertEqual "$A$1:$D$5", Selection.Address, "e seleciona os dados"
    ws.Range("C2").Value = 7
    AssertFalse wb.Saved
    AssertEqual 1, Core_PptDataSave(pres), "Salvar no PowerPoint salva a pasta de dados"
    AssertTrue wb.Saved
    AssertEqual 0, Core_PptDataSave(pres), "nada alterado: nada a salvar"
    AssertFalse Core_IsQuiet()
    AssertEqual "!" & PPTSYNC_NONE, Core_PptEditData(pres, PptIdx_Entry(999999, "nada", ""))
    ' M9 (revisao da G2): Alterar dados num grafico criado pelo PowerPoint recria o grafico com id novo; a moldura faz o aviso aparecer
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(ws.ChartObjects(1).Chart, ws.Range("A1:D5")), "recriado"
    AssertTrue InStr(1, Core_LastNotice(), LINKFRAME_NOTE_REBUILT, vbBinaryCompare) > 0, "aviso de grafico recriado: " & Core_LastNotice()
    AssertEqual "!Gráfico não encontrado na pasta de dados (apagado ou recriado).", Core_PptEditData(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""))
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set shp = Nothing: Set s2 = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))      ' falha antes de o teste pegar a pasta de dados: fecha mesmo assim
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' G3: pasta de dados sumiu (apresentacao enviada sem ela, movida): Editar dados avisa sem abrir nada e sem LinkFormat.Update (PG4: Update com a
' origem ausente abre uma caixa modal que trava o PowerPoint).
Public Sub Test_PptEditData_MissingDataBook_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim shp As Object, td As String, p As String, pp As String, moved As String, wb As Workbook
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    td = TempDir()
    pp = td & "\Sumiu G.pptx": p = td & "\Sumiu G - dados ThinQQ.xlsx": moved = td & "\outro nome.xlsx"
    pres.SaveAs pp
    AssertEqual "", Core_PptInsertChart(slide, "stacked")
    Set shp = slide.Shapes(1)
    Set wb = Workbooks(Dir(p))
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    Name p As moved
    PptSync_ResetTestHooks
    Core_ClearLastNotice
    msg = Core_PptEditData(pres, PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), ""))
    AssertEqual "!Pasta de dados não encontrada: " & p & ". Se ela foi movida, ponha-a na mesma pasta da apresentação.", msg
    AssertEqual 0, PptSync_UpdateCountForTest(), "sem LinkFormat.Update na origem ausente"
    AssertEqual "", Dir(p), "nada criado no caminho antigo"
    AssertFalse Core_IsQuiet()
    AssertEqual 0, Core_ShownCountForTest()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp: Kill moved
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' G3: Editar dados num grafico colado do Excel (pasta do USUARIO, .xlsm, fechada): abre para edicao com as macros desligadas, ativa a planilha
' do grafico e o aviso das macros fica so no registro/Core_LastNotice (modo silencioso: nenhuma caixa enquanto o PowerPoint espera). Grafico
' que nao existe mais na pasta: recusa e a pasta que o ThinQQ abriu e fechada. Sem PowerPoint (duble).
Public Sub Test_PptEditData_UserXlsm_NoBox_ClosesWhenChartMissing()
    Dim wb As Workbook, rng As Range, p As String, id As String, shp As TestPptFake, lf As TestPptFake, pres As TestPptFake, msg As String
    Dim w As Workbook, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    id = Registry_IdFromChart(rng.Worksheet.ChartObjects(1).Chart)
    p = TempDir() & "\usuario-g3.xlsm"
    wb.SaveAs p, 52                                      ' xlOpenXMLWorkbookMacroEnabled
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    Set shp = New TestPptFake: Set shp.Tags = CreateObject("Scripting.Dictionary")
    shp.Tags.Add PPTSYNC_TAG_CHART, id
    shp.Tags.Add PPTSYNC_TAG_BOOK, p
    Set lf = New TestPptFake: lf.SourceFullName = p: Set shp.LinkFormat = lf
    Set pres = FakeAutoPres(shp, p)
    Core_ClearLastNotice
    msg = Core_PptEditData(pres, PptIdx_Entry(256, "tqq_x", ""))
    AssertTrue Left$(msg, 1) <> "!", msg
    Set wb = Workbooks(Dir(p))
    AssertFalse wb.ReadOnly, "aberta para edicao"
    AssertTrue ActiveWorkbook Is wb
    AssertEqual "Dados", ActiveSheet.Name, "planilha do grafico ativa"
    AssertTrue InStr(1, Core_LastNotice(), "macros desta pasta", vbTextCompare) > 0, "aviso das macros registrado: " & Core_LastNotice()
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa durante a chamada do PowerPoint"
    AssertFalse Core_IsQuiet()
    AssertEqual 0, lf.UpdateCount, "sem LinkFormat.Update"
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    shp.Tags(PPTSYNC_TAG_CHART) = "naoexiste"
    AssertEqual "!Gráfico não encontrado na pasta de origem (apagado ou recriado).", Core_PptEditData(pres, PptIdx_Entry(256, "tqq_x", ""))
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, p, vbTextCompare) <> 0, "a pasta aberta pelo ThinQQ foi fechada"
    Next
    AssertFalse Core_IsQuiet()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Core_ClearLastNotice
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then Teardown wb
    Kill p
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' G7 (1.9.0): comandos do grafico pelo PowerPoint (API 4) num grafico criado pelo PowerPoint: o mesmo codigo da faixa do Excel na pasta de
' dados, que fica salva; o slide acompanha (setas clonadas e removidas); TQQ_KIND/TQQ_OPTS para a faixa do PowerPoint; recusa da matriz em
' texto; pasta de dados FECHADA e aberta so para o comando e fechada de novo; sem caixas e com o modo silencioso e a operacao fechados.
Public Sub Test_PptCommand_FromPpt_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Dim shp As Object, td As String, p As String, pp As String, wb As Workbook, e As String, n0 As Long, w As Workbook, o0 As String
    Dim nA As Long, id0 As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    td = TempDir()                                       ' TempDir() cria uma pasta nova a cada chamada: as duas na MESMA pasta
    pp = td & "\Cmd G.pptx": p = td & "\Cmd G - dados ThinQQ.xlsx"
    If Len(Dir(p)) > 0 Then Kill p
    pres.SaveAs pp
    AssertEqual "", Core_PptInsertChart(slide, "clustered")
    Set shp = slide.Shapes(1)
    Set wb = Workbooks(Dir(p))
    e = PptIdx_Entry(CLng(slide.SlideID), CStr(shp.Name), "")
    Ppt_SetAppForTest app
    AssertEqual "clustered", PptSync_Tag(shp, PPTSYNC_TAG_KIND), "TQQ_KIND gravado na colagem"
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "caps=labels,") > 0, PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertEqual 0, InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "totals,"), "agrupadas nao tem totais na faixa"
    Core_ClearLastNotice
    AssertEqual "", Core_PptCommand(pres, e, "numfmt", "pct0")
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "numfmt=pct0") > 0, PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertTrue wb.Saved, "pasta de dados salva depois do comando"
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
    AssertFalse Ppt_InOperation(), "operacao do PowerPoint fechada"
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa durante a chamada do PowerPoint"
    AssertEqual "2022" & Chr$(31) & "2023" & Chr$(31) & "2024" & Chr$(31) & "2025", Core_PptCategories(pres, e)
    n0 = shp.Chart.Shapes.Count
    AssertEqual "", Core_PptCommand(pres, e, "cagr", "1,4")
    AssertTrue shp.Chart.Shapes.Count > n0, "seta CAGR clonada no slide"
    AssertEqual PPTCMD_MSG_CATS, Core_PptCommand(pres, e, "diff", "3,2"), "de depois de ate"
    AssertEqual PPTCMD_MSG_CATS, Core_PptCommand(pres, e, "diff", "1,9"), "fora do grafico"
    AssertEqual PPTCMD_MSG_CATS, Core_PptCommand(pres, e, "diff", "x"), "sem as duas categorias"
    nA = CountAnn(Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART)).Chart)
    AssertTrue nA > 0, "seta CAGR na pasta de dados"
    AssertEqual "", Core_PptCommand(pres, e, "clear", "")
    AssertEqual 0, CountAnn(shp.Chart), "setas removidas no slide"
    ' revisao final G-II (M2): alternancia com o valor do botao ("labels:1"): repetir nao desfaz. (M1, duas formas do mesmo grafico, fica no
    ' teste com dubles abaixo: duplicar um grafico vinculado no PowerPoint chama o servidor OLE deste mesmo Excel no meio da suite)
    AssertEqual "", Core_PptCommand(pres, e, "toggle", "labels:1")
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "labels=1;") = 1, PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertEqual "", Core_PptCommand(pres, e, "toggle", "labels:1")
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "labels=1;") = 1, "valor absoluto: de novo continua ligado"
    AssertEqual "", Core_PptCommand(pres, e, "toggle", "labels:0")
    AssertTrue InStr(PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "labels=0;") = 1, PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertEqual PPTCMD_MSG_UNKNOWN & "toggle labels:x", Core_PptCommand(pres, e, "toggle", "labels:x")
    ' (M7) grafico apagado da pasta de DADOS: o texto da pasta de dados
    id0 = PptSync_Tag(shp, PPTSYNC_TAG_CHART)
    shp.Tags.Add PPTSYNC_TAG_CHART, "naoexiste"
    AssertEqual "Gráfico não encontrado na pasta de dados (apagado ou recriado).", Core_PptCommand(pres, e, "clear", "")
    shp.Tags.Add PPTSYNC_TAG_CHART, id0
    AssertEqual Kinds_NotApplicableText("totals"), Core_PptCommand(pres, e, "toggle", "totals"), "agrupadas nao tem totais: recusa da matriz em texto"
    AssertEqual PPTCMD_MSG_UNKNOWN & "toggle palette", Core_PptCommand(pres, e, "toggle", "palette")
    AssertEqual "Formato numérico desconhecido: abc", Core_PptCommand(pres, e, "numfmt", "abc")
    AssertEqual PPTCMD_MSG_UNKNOWN & "xyz", Core_PptCommand(pres, e, "xyz", "")
    AssertEqual PPTSYNC_NOT_FOUND, Core_PptCommand(pres, PptIdx_Entry(999999, "nada", ""), "clear", "")
    ' pasta de dados FECHADA: aberta para o comando (macros desligadas), salva e fechada de novo
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    o0 = PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    AssertEqual "", Core_PptCommand(pres, e, "toggle", "labels")
    AssertTrue o0 <> PptSync_Tag(shp, PPTSYNC_TAG_OPTS), "rotulos invertidos no slide: " & PptSync_Tag(shp, PPTSYNC_TAG_OPTS)
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, p, vbTextCompare) <> 0, "a pasta de dados aberta pelo comando foi fechada"
    Next
    AssertEqual "2022" & Chr$(31) & "2023" & Chr$(31) & "2024" & Chr$(31) & "2025", Core_PptCategories(pres, e), "categorias com a pasta fechada"
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, p, vbTextCompare) <> 0, "e fechada de novo depois das categorias"
    Next
    Set wb = Application.Workbooks.Open(p)
    AssertEqual Mid$(o0, InStr(o0, "labels=") + 7, 1) = "1", Not (Registry_GetOption(Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART)).Chart, "labels", "0") = "1"), "o comando foi salvo na pasta de dados"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing: Set w = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Private Function CountAnn(pc As Object) As Long
    Dim i As Long
    For i = 1 To pc.Shapes.Count
        If Left$(CStr(pc.Shapes(i).Name), Len(ANN_PREFIX)) = ANN_PREFIX Then CountAnn = CountAnn + 1
    Next
End Function

' G7 (seguranca): grafico colado de uma pasta do USUARIO. Fechada: o comando recusa (PPTCMD_MSG_OPEN_SOURCE) e o ThinQQ NUNCA a abre (nem
' para as categorias). Aberta pelo usuario: o comando aplica e a pasta fica alterada e NAO salva (o usuario decide). Sem PowerPoint (duble:
' sem grafico no slide, a sincronizacao conta falha e o resumo volta como mensagem).
Public Sub Test_PptCommand_UserBook_OnlyWhenOpen_NeverSaved()
    Dim wb As Workbook, rng As Range, p As String, id As String, shp As TestPptFake, lf As TestPptFake, pres As TestPptFake, msg As String
    Dim w As Workbook, en As Long, ed As String, t0 As Date, e As String, v0 As String, shp2 As TestPptFake, lf2 As TestPptFake, v1 As String
    On Error GoTo Cleanup
    AutoBegin
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    id = Registry_IdFromChart(rng.Worksheet.ChartObjects(1).Chart)
    p = TempDir() & "\usuario-g7.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    t0 = FileDateTime(p)
    Set shp = New TestPptFake: Set shp.Tags = CreateObject("Scripting.Dictionary")
    shp.Tags.Add PPTSYNC_TAG_CHART, id
    shp.Tags.Add PPTSYNC_TAG_BOOK, p
    Set lf = New TestPptFake: lf.SourceFullName = p: Set shp.LinkFormat = lf
    Set pres = FakeAutoPres(shp, p)
    e = PptIdx_Entry(256, "tqq_x", "")
    Core_ClearLastNotice
    AssertEqual PPTCMD_MSG_OPEN_SOURCE, Core_PptCommand(pres, e, "toggle", "totals")
    AssertEqual "", Core_PptCategories(pres, e), "categorias: pasta do usuario fechada nunca e aberta"
    shp.Tags(PPTDATA_TAG) = "1"                           ' etiqueta forjada: o nome nao e de pasta de dados => nada e aberto
    AssertEqual PPTDATA_MSG_NOTDATANAME & p, Core_PptCommand(pres, e, "toggle", "totals")
    shp.Tags.Remove PPTDATA_TAG
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, p, vbTextCompare) <> 0, "a pasta do usuario nao foi aberta"
    Next
    AssertTrue FileDateTime(p) = t0, "a pasta do usuario nao foi gravada"
    AssertEqual 0, lf.UpdateCount, "sem LinkFormat.Update"
    Set wb = Application.Workbooks.Open(p)                ' o usuario abre a pasta
    v0 = Registry_GetOption(Registry_FindChartObject(wb, id).Chart, "totals", "0")
    msg = Core_PptCommand(pres, e, "toggle", "totals")
    AssertTrue InStr(msg, "Falha ao atualizar: 1") > 0, "o slide (duble sem grafico) nao acompanhou: o resumo volta: " & msg
    AssertEqual IIf(v0 = "1", "0", "1"), Registry_GetOption(Registry_FindChartObject(wb, id).Chart, "totals", "0"), "comando aplicado (invertido) na pasta aberta"
    ' revisao final G-II (M1): duas formas do MESMO grafico (pasta + TQQ_CHART; copia colada no slide) selecionadas: o comando vale UMA vez
    ' (um toggle nao desfaz o outro) e as duas formas vao para a sincronizacao (o duble sem grafico conta 2 falhas)
    Set shp2 = New TestPptFake: Set shp2.Tags = CreateObject("Scripting.Dictionary")
    shp2.Tags.Add PPTSYNC_TAG_CHART, id
    shp2.Tags.Add PPTSYNC_TAG_BOOK, p
    Set lf2 = New TestPptFake: lf2.SourceFullName = p: Set shp2.LinkFormat = lf2
    pres.Slides.Slide.Shapes.Add shp2, "tqq_y"
    AssertEqual 2, PptSync_ResolveEntries(pres, e & vbLf & PptIdx_Entry(256, "tqq_y", "")).Count, "duas formas"
    v1 = Registry_GetOption(Registry_FindChartObject(wb, id).Chart, "totals", "0")
    msg = Core_PptCommand(pres, e & vbLf & PptIdx_Entry(256, "tqq_y", ""), "toggle", "totals")
    AssertEqual IIf(v1 = "1", "0", "1"), Registry_GetOption(Registry_FindChartObject(wb, id).Chart, "totals", "0"), "invertido UMA vez com as duas formas: " & msg
    AssertTrue InStr(msg, "Falha ao atualizar: 2") > 0, "as duas formas sincronizadas: " & msg
    msg = Core_PptCommand(pres, e & vbLf & PptIdx_Entry(256, "tqq_y", ""), "toggle", "totals:1")
    AssertEqual "1", Registry_GetOption(Registry_FindChartObject(wb, id).Chart, "totals", "0"), "(M2) valor absoluto"
    AssertFalse wb.Saved, "pasta do usuario alterada e NAO salva"
    AssertTrue FileDateTime(p) = t0, "arquivo do usuario intacto"
    AssertEqual "T1" & Chr$(31) & "T2", Core_PptCategories(pres, e), "categorias da pasta aberta"
    AssertFalse Core_IsQuiet()
    AssertFalse Ppt_InOperation()
    AssertEqual 0, Core_ShownCountForTest()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Core_ClearLastNotice
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then Teardown wb
    Kill p
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Revisao final (I2): regra pura das extensoes aceitas e do nome da pasta de dados.
Public Sub Test_PptData_ExtensionAndDataNameRules()
    AssertTrue PptData_IsWorkbookExt("C:\a\b.xlsx"): AssertTrue PptData_IsWorkbookExt("C:\a\b.XLSM")
    AssertTrue PptData_IsWorkbookExt("\\s\x\b.xlsb"): AssertTrue PptData_IsWorkbookExt("C:\a\b.xls")
    AssertTrue PptData_IsWorkbookExt("https://e.sharepoint.com/x/b.xlsx")
    AssertFalse PptData_IsWorkbookExt("C:\a\b.xll"), "suplemento nativo (DLL)"
    AssertFalse PptData_IsWorkbookExt("C:\a\b.iqy"): AssertFalse PptData_IsWorkbookExt("C:\a\b.csv")
    AssertFalse PptData_IsWorkbookExt("C:\a\b.slk"): AssertFalse PptData_IsWorkbookExt("C:\a\b.htm")
    AssertFalse PptData_IsWorkbookExt("C:\a.xlsx\b"), "extensao so do nome do arquivo"
    AssertFalse PptData_IsWorkbookExt(""): AssertFalse PptData_IsWorkbookExt("C:\a\b")
    AssertTrue PptData_IsDataBookName("C:\a\Vendas - dados ThinQQ.xlsx")
    AssertTrue PptData_IsDataBookName("\\s\x\v - DADOS thinqq.XLSX")
    AssertFalse PptData_IsDataBookName("C:\a\ - dados ThinQQ.xlsx"), "so o sufixo"
    AssertFalse PptData_IsDataBookName("C:\a\Orcamento.xlsx")
    AssertFalse PptData_IsDataBookName("C:\a\Vendas - dados ThinQQ.xlsm")
End Sub

' Revisao final (I2, M4), duble sem PowerPoint: Editar dados numa forma cuja TQQ_BOOK e um .xll recusa SEM abrir (o arquivo existe); forma
' TQQ_DATA apontando para uma pasta que nao e de dados do ThinQQ recusa sem abrir; pasta de origem do usuario ja aberta SOMENTE LEITURA e
' usada como esta (spec 14.3).
Public Sub Test_PptEditData_RefusesNonWorkbookAndForeignDataPath_UsesOpenReadOnlyUserBook()
    Dim wb As Workbook, rng As Range, p As String, x As String, id As String, shp As TestPptFake, lf As TestPptFake, pres As TestPptFake
    Dim msg As String, k As Integer, w As Workbook, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    x = TempDir() & "\hostil.xll"
    k = FreeFile: Open x For Output As #k: Print #k, "x": Close #k
    Set shp = New TestPptFake: Set shp.Tags = CreateObject("Scripting.Dictionary")
    shp.Tags.Add PPTSYNC_TAG_CHART, "abc"
    shp.Tags.Add PPTSYNC_TAG_BOOK, x
    Set lf = New TestPptFake: lf.SourceFullName = x: Set shp.LinkFormat = lf
    Set pres = FakeAutoPres(shp, x)
    msg = Core_PptEditData(pres, PptIdx_Entry(256, "tqq_x", ""))
    AssertEqual "!" & PPTDATA_MSG_NOTBOOK & x, msg
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, x, vbTextCompare) <> 0, ".xll nunca aberto"
    Next
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    id = Registry_IdFromChart(rng.Worksheet.ChartObjects(1).Chart)
    p = TempDir() & "\usuario-final.xlsx"
    wb.SaveAs p, 51
    Watchers_Stop
    wb.Close SaveChanges:=False: Set wb = Nothing
    shp.Tags(PPTSYNC_TAG_CHART) = id
    shp.Tags(PPTSYNC_TAG_BOOK) = p
    lf.SourceFullName = p
    shp.Tags(PPTDATA_TAG) = "1"                          ' atribuicao: a leitura anterior do Dictionary ja criou a chave
    AssertEqual "!" & PPTDATA_MSG_NOTDATANAME & p, Core_PptEditData(pres, PptIdx_Entry(256, "tqq_x", "")), "TQQ_DATA so com a pasta de dados do ThinQQ"
    For Each w In Application.Workbooks
        AssertTrue StrComp(w.FullName, p, vbTextCompare) <> 0, "nao abriu a pasta do usuario"
    Next
    shp.Tags(PPTDATA_TAG) = ""
    Set wb = Workbooks.Open(Filename:=p, ReadOnly:=True)
    msg = Core_PptEditData(pres, PptIdx_Entry(256, "tqq_x", ""))
    AssertTrue Left$(msg, 1) <> "!", "pasta do usuario ja aberta somente leitura: usada como esta: " & msg
    AssertTrue ActiveWorkbook Is wb
    AssertFalse Core_IsQuiet()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then Teardown wb
    If Len(p) > 0 Then Kill p
    If Len(x) > 0 Then Kill x
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' G3: o ciclo automatico salva as pastas de DADOS (marca TQQ_DATABOOK) que mudaram; a pasta do usuario nunca e salva pelo ThinQQ. Sem
' PowerPoint (build Test: Ppt_GetApp devolve so a instancia injetada).
Public Sub Test_PptAuto_Tick_SavesOnlyDataBooks()
    Dim wb As Workbook, rng As Range, db As Workbook, co As ChartObject, drng As Range, p As String, en As Long, ed As String
    On Error GoTo Cleanup
    AutoBegin
    Ppt_ResetTestHooks
    Set wb = AutoBook("pptauto-user.xlsx", rng)
    p = TempDir() & "\Auto G - dados ThinQQ.xlsx"
    Set db = PptData_OpenOrCreate(p, "Auto G.pptx")
    Set co = PptData_NewChart(db, "stacked", drng)
    db.Save
    Watchers_Stop                                        ' o Core_Insert religa os vigias: o teste marca as pastas ele mesmo
    rng.Cells(2, 2).Value = 123
    drng.Cells(2, 2).Value = 456
    AssertFalse wb.Saved
    AssertFalse db.Saved
    PptAuto_OnChange wb
    PptAuto_OnChange db
    AssertEqual 2, PptAuto_PendingCountForTest()
    PptAuto_Tick
    AssertEqual 0, PptAuto_PendingCountForTest()
    AssertTrue db.Saved, "pasta de dados salva pelo ciclo"
    AssertFalse wb.Saved, "pasta do usuario nunca salva pelo ThinQQ"
    AssertFalse PptAuto_InTick()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set co = Nothing: Set drng = Nothing
    If Not db Is Nothing Then db.Close SaveChanges:=False
    Kill p
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' M3 (revisao da G2): falha DEPOIS da colagem (forma ja etiquetada e, com "Novos vinculos" ligada, ja no indice) desfaz tudo: sem forma, sem
' entrada no indice, sem a planilha "Gráfico 1", pasta de dados salva, operacao e modo silencioso fechados.
Public Sub Test_PptInsertChart_FailureAfterPaste_Undone_WhenAvailable()
    Dim app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String, msg As String
    Dim td As String, p As String, pp As String, wb As Workbook, ws As Worksheet
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoReset
    PptAuto_SetNewLinksAuto True                         ' a colagem ja indexa: o desfazer precisa tirar a entrada
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    If UserWindowAppeared(app) Then GoTo Cleanup
    td = TempDir()
    pp = td & "\Desfaz G.pptx": p = td & "\Desfaz G - dados ThinQQ.xlsx"
    pres.SaveAs pp
    PptData_SetFailForTest "tagged"
    msg = Core_PptInsertChart(slide, "stacked")
    AssertTrue InStr(1, msg, "Falha simulada (teste) tagged", vbBinaryCompare) > 0, "mensagem devolvida: " & msg
    AssertEqual 0, slide.Shapes.Count, "forma colada removida"
    AssertEqual "", PptSync_Tag(pres, PPTSYNC_TAG_INDEX), "sem entrada no indice"
    Set wb = Workbooks(Dir(p))
    AssertFalse wb Is Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Gráfico 1")
    On Error GoTo Cleanup
    AssertTrue ws Is Nothing, "planilha do grafico apagada"
    AssertTrue wb.Saved, "pasta de dados salva depois do desfazer"
    AssertFalse Ppt_InOperation(), "operacao fechada"
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
    AssertEqual "", Core_PptInsertChart(slide, "stacked"), "a proxima insercao funciona"
    AssertEqual "Gráfico 1", PptSync_Tag(slide.Shapes(1), PPTSYNC_TAG_SHEET), "e reaproveita o nome livre"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptData_ResetTestHooks
    AutoReset
    Set ws = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    Watchers_Stop
    If wb Is Nothing And Len(p) > 0 Then Set wb = Workbooks(Dir(p))
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill p: Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- Parte E1: texto vinculado ----

Private Function NewTextBox(pres As Object, txt As String) As Object
    Dim sl As Object
    Set sl = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Set NewTextBox = sl.Shapes.AddTextbox(1, 20, 20, 400, 40)
    NewTextBox.TextFrame.TextRange.Text = txt
End Function

Public Sub Test_PptText_LinkAndSync_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, tb As Object, own As New Collection, en As Long, ed As String, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("E1").Value = 123
    wb.SaveAs TempDir() & "\ppttext-sync.xlsx", xlOpenXMLWorkbook
    Set tb = NewTextBox(pres, "Receita de X mil")
    tb.TextFrame.TextRange.Characters(12, 1).Font.Bold = -1
    AssertEqual "", PptText_Link(tb, 12, 1, ws.Range("E1"))
    AssertEqual "Receita de 123 mil", tb.TextFrame.TextRange.Text
    AssertEqual "text", PptSync_ItemKind(tb)
    ws.Range("E1").Value = 4567
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual "Receita de 4567 mil", tb.TextFrame.TextRange.Text
    AssertEqual -1, CLng(tb.TextFrame.TextRange.Characters(12, 4).Font.Bold), "formatacao do campo mantida"
    AssertEqual 0, CLng(tb.TextFrame.TextRange.Characters(1, 7).Font.Bold), "resto do texto intacto"
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, s
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set tb = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 4: o usuario editou o texto fora do campo: nada e sobrescrito; o resumo conta.
Public Sub Test_PptText_Sync_EditedOutsideField_NotTouched_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, tb As Object, own As New Collection, en As Long, ed As String, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1").Value = 10
    wb.SaveAs TempDir() & "\ppttext-edit.xlsx", xlOpenXMLWorkbook
    Set tb = NewTextBox(pres, "Total: X")
    AssertEqual "", PptText_Link(tb, 8, 1, ws.Range("A1"))
    tb.TextFrame.TextRange.Text = "Outro texto do usuario"
    ws.Range("A1").Value = 20
    s = PptSync_Presentation(pres, True)
    AssertEqual "Outro texto do usuario", tb.TextFrame.TextRange.Text
    AssertTrue InStr(s, "Texto editado no PowerPoint fora do campo vinculado (vincule de novo): 1.") > 0, s
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set tb = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Revisao final E (M3, M10), PowerPoint real: emoji no prefixo (par substituto: Characters/Start contam unidades UTF-16 como Len?), vinculo no
' cursor (tamanho 0, InsertBefore), campo que esvazia (celula vazia) e volta a ter valor.
Public Sub Test_PptText_CursorEmptyEmoji_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, tb As Object, own As New Collection, en As Long, ed As String, s As String
    Dim emo As String, pre As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1").Value = "abc"
    wb.SaveAs TempDir() & "\ppttext-cursor.xlsx", xlOpenXMLWorkbook
    emo = ChrW$(&HD83D) & ChrW$(&HDE00)
    pre = emo & " Total: ["
    Set tb = NewTextBox(pres, pre & "] fim")
    AssertEqual Len(pre) + 5, Len(CStr(tb.TextFrame.TextRange.Text)), "premissa: o PowerPoint guarda o emoji como 2 unidades UTF-16"
    AssertEqual "", PptText_Link(tb, Len(pre) + 1, 0, ws.Range("A1"))
    AssertEqual pre & "abc] fim", CStr(tb.TextFrame.TextRange.Text), "cursor depois do emoji: no lugar certo"
    ws.Range("A1").ClearContents
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual pre & "] fim", CStr(tb.TextFrame.TextRange.Text), "celula vazia: campo vazio"
    ws.Range("A1").Value = "zz"
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual pre & "zz] fim", CStr(tb.TextFrame.TextRange.Text), "o campo vazio volta a ter valor"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set tb = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Adendo do controlador: as etiquetas do texto vinculado (prefixo com quebra de paragrafo, quebra de linha e Tab; reserva planilha|endereco)
' sobrevivem a Salvar como .pptx, fechar e reabrir: a sincronizacao depois de reabrir diz "sem mudanca" (nao "editado"), atualiza quando a
' celula muda, e a forma automatica continua no indice que o ciclo automatico le (PptAuto_ItemsFor).
Public Sub Test_PptText_TagsPersistSaveReopen_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, p2 As Object, tb As Object, tb2 As Object, own As New Collection
    Dim en As Long, ed As String, s As String, tmp As String, cur0 As String, st As Long, k As Variant, saved As Object, items As Collection
    Dim dirty As Object
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_ResetTestHooks
    PptAuto_SetNewLinksAuto True
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("B2").Value = 10
    tmp = TempDir()
    wb.SaveAs tmp & "\ppttext-persist.xlsx", xlOpenXMLWorkbook
    Set tb = NewTextBox(pres, "Linha 1" & vbCr & "A" & Chr$(11) & "Total:" & vbTab & "X fim")
    cur0 = CStr(tb.TextFrame.TextRange.Text)
    st = InStr(cur0, "X")
    AssertTrue st > 0, "premissa: o X esta no texto (" & cur0 & ")"
    AssertEqual "", PptText_Link(tb, st, 1, ws.Range("B2"))
    AssertEqual Left$(cur0, st - 1) & "10" & Mid$(cur0, st + 1), CStr(tb.TextFrame.TextRange.Text)
    AssertEqual "1", PptSync_Tag(tb, PPTSYNC_TAG_AUTO), "novo vinculo automatico"
    Set saved = CreateObject("Scripting.Dictionary")
    For Each k In Array(PPTTEXT_TAG, PPTSYNC_TAG_BOOK, PPTLINK_TAG_NAME, PPTLINK_TAG_REF, PPTTEXT_TAG_PRE, PPTTEXT_TAG_SUF, PPTTEXT_TAG_LAST)
        saved.Add CStr(k), PptSync_Tag(tb, CStr(k))
    Next
    AssertEqual Left$(cur0, st - 1), PptText_TagDec(CStr(saved(PPTTEXT_TAG_PRE))), "prefixo gravado"
    pres.SaveAs tmp & "\ppttext-persist.pptx"
    Set tb = Nothing
    pres.Close
    Set pres = Nothing
    Set own = New Collection
    Set p2 = app.Presentations.Open(tmp & "\ppttext-persist.pptx", -1, 0, 0): own.Add p2
    Set tb2 = p2.Slides(1).Shapes(1)
    For Each k In saved.Keys
        AssertEqual CStr(saved(k)), PptSync_Tag(tb2, CStr(k)), "etiqueta " & k & " igual depois de reabrir"
    Next
    AssertEqual "text", PptSync_ItemKind(tb2)
    s = PptSync_Presentation(p2, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, "depois de reabrir: " & s
    AssertTrue InStr(s, "Texto editado") = 0, s
    Set dirty = CreateObject("Scripting.Dictionary")
    dirty.CompareMode = 1
    dirty.Add wb.FullName, True
    Set items = PptAuto_ItemsFor(p2, dirty)
    AssertEqual 1, items.Count, "o ciclo automatico acha o texto vinculado pelo indice"
    ws.Range("B2").Value = 20
    s = PptSync_Items(items, False, True)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual Left$(cur0, st - 1) & "20" & Mid$(cur0, st + 1), CStr(tb2.TextFrame.TextRange.Text)
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set items = Nothing: Set tb = Nothing: Set tb2 = Nothing: Set pres = Nothing: Set p2 = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Review Focus 5 (mesmo padrao de Test_PptSync_HostileSourceMacros_NeverRun_Fake): uma forma de TEXTO vinculado (duble com texto) citando uma
' .xlsm hostil fechada com Workbook_Open e Workbook_BeforeClose que gravam um marcador. A sincronizacao abre a pasta (prova: o campo recebe o
' valor da celula), le e fecha COM OS EVENTOS LIGADOS no fechamento (gancho de teste), e mesmo assim nenhuma macro roda: so as macros
' desligadas (AutomationSecurity 3) seguram. Controle positivo no fim: a mesma pasta aberta com as macros LIGADAS roda o BeforeClose (o
' teste nao passa a toa). A pasta nunca e salva (data do arquivo igual).
Public Sub Test_PptText_HostileSourceMacros_NeverRun_Fake()
    Dim wb As Workbook, dir_ As String, xlsm As String, marker As String, cm As Object, pres As TestPptFake, sl As TestPptFake, t As TestPptTextFake
    Dim tg As TestPptTags, sec0 As Long, prevEv As Boolean, s As String, n0 As Long, code As String, sheetName As String, d0 As Date, w As Workbook
    dir_ = TempDir()
    xlsm = dir_ & "\ppttext-hostil.xlsm": marker = dir_ & "\marcador-texto.txt"
    Set wb = NewTempWorkbook()
    sheetName = wb.Worksheets(1).Name
    wb.Worksheets(1).Range("A1").Value = 42
    PptLink_DefineName wb.Worksheets(1).Range("A1"), "TQQ_X_"         ' a pasta tem vinculos do ThinQQ (marca): a reserva TQQ_TREF vale (I1)
    code = "Private Sub Workbook_Open()" & vbCrLf
    code = code & "    Dim f As Integer: f = FreeFile: Open """ & marker & """ For Output As #f: Print #f, ""macro rodou (open)"": Close #f" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    code = code & "    Dim f As Integer: f = FreeFile: Open """ & marker & """ For Output As #f: Print #f, ""macro rodou"": Close #f" & vbCrLf
    code = code & "End Sub"
    Set cm = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    cm.AddFromString code
    Set cm = Nothing
    wb.SaveAs xlsm, 52
    prevEv = Application.EnableEvents: Application.EnableEvents = False
    wb.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Set wb = Nothing
    If Len(Dir(marker)) > 0 Then Kill marker
    d0 = FileDateTime(xlsm)
    Set tg = New TestPptTags
    tg.Preset PPTTEXT_TAG, "1"
    tg.Preset PPTSYNC_TAG_BOOK, xlsm
    tg.Preset PPTLINK_TAG_NAME, "TQQ_X_naoexiste"
    tg.Preset PPTLINK_TAG_REF, sheetName & "|$A$1"
    tg.Preset PPTTEXT_TAG_PRE, PptText_TagEnc("v=")
    tg.Preset PPTTEXT_TAG_SUF, ""
    tg.Preset PPTTEXT_TAG_LAST, "X"
    Set t = New TestPptTextFake: t.HasTextFrame = -1: Set t.Tags = tg: t.Preset "v=X"
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add t
    Set pres = New TestPptFake: Set pres.Slides = New Collection: pres.Slides.Add sl
    PptSync_ResetTestHooks
    PptSync_KeepEventsOnCloseForTest
    Application.EnableEvents = True
    sec0 = Application.AutomationSecurity
    n0 = Application.Workbooks.Count
    s = PptSync_Presentation(pres, True)
    Application.EnableEvents = prevEv
    PptSync_ResetTestHooks
    AssertEqual 0, Len(Dir(marker)), "a macro da pasta de origem NAO rodou (" & s & ")"
    AssertEqual "v=42", t.Text, "a pasta foi aberta e lida (" & s & ")"
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual sec0, CLng(Application.AutomationSecurity), "AutomationSecurity restaurada"
    AssertEqual n0, Application.Workbooks.Count, "a pasta aberta pela sincronizacao foi fechada"
    AssertTrue d0 = FileDateTime(xlsm), "a pasta de origem nunca e salva"
    ' controle positivo: com as macros ligadas o BeforeClose roda e grava o marcador
    On Error Resume Next
    Application.AutomationSecurity = 1                              ' msoAutomationSecurityLow (so este Open; restaurada em seguida)
    Set w = Application.Workbooks.Open(xlsm, 0, True)
    Application.AutomationSecurity = sec0
    Application.EnableEvents = True
    If Not w Is Nothing Then w.Close SaveChanges:=False
    Application.EnableEvents = prevEv
    Err.Clear
    On Error GoTo 0
    Set w = Nothing
    AssertEqual sec0, CLng(Application.AutomationSecurity), "AutomationSecurity restaurada depois do controle"
    AssertTrue Len(Dir(marker)) > 0, "controle positivo: com as macros ligadas a macro roda (senao o teste acima passaria a toa)"
    If Len(Dir(marker)) > 0 Then Kill marker
End Sub

' ---- Parte F2: tabela vinculada ----

Public Sub Test_PptTable_InsertAndSync_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Dim shp As Object, t As Object, s As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1:C1").Value = Array("Região", "2025", "2026"): ws.Range("A1:C1").Font.Bold = True
    ws.Range("A2:C2").Value = Array("Sul", 10, 12): ws.Range("A3:C3").Value = Array("Norte", 5, 7)
    ws.Range("C3").Interior.Color = RGB(255, 200, 0): ws.Range("A3").Font.Color = RGB(200, 0, 0)
    wb.SaveAs TempDir() & "\ppttable.xlsx", xlOpenXMLWorkbook
    Set shp = PptTable_Insert(ws.Range("A1:C3"), slide)
    Set t = shp.Table
    AssertEqual 3, t.Rows.Count: AssertEqual 3, t.Columns.Count
    AssertEqual "Norte", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    AssertEqual -1, CLng(t.Cell(1, 2).Shape.TextFrame.TextRange.Font.Bold), "negrito do Excel"
    AssertEqual 0, CLng(t.Cell(2, 1).Shape.TextFrame.TextRange.Font.Bold), "sem negrito no corpo"
    AssertEqual 3, CLng(t.Cell(2, 2).Shape.TextFrame.TextRange.ParagraphFormat.Alignment), "numero a direita (alinhamento Geral)"
    AssertEqual 1, CLng(t.Cell(2, 1).Shape.TextFrame.TextRange.ParagraphFormat.Alignment), "texto a esquerda"
    AssertEqual RGB(255, 200, 0), CLng(t.Cell(3, 3).Shape.Fill.ForeColor.RGB), "preenchimento do Excel"
    AssertEqual RGB(200, 0, 0), CLng(t.Cell(3, 1).Shape.TextFrame.TextRange.Font.Color.RGB), "cor da fonte do Excel"
    AssertEqual CLng(ws.Range("A2").Font.Size), CLng(t.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size), "tamanho da fonte do Excel"
    AssertEqual "table", PptSync_ItemKind(shp)
    AssertEqual "TQQ_T_", Left$(PptSync_Tag(shp, PPTLINK_TAG_NAME), 6), "nome oculto TQQ_T_"
    AssertTrue PptLink_HasMark(wb), "a marca TQQ_XLINK"
    t.Cell(2, 2).Shape.TextFrame.TextRange.Font.Italic = -1          ' formatacao feita no PowerPoint
    ws.Range("B2").Value = 99
    ws.Rows(4).Insert: ws.Range("A4:C4").Value = Array("Leste", 1, 2)   ' fora do nome: nao entra
    ws.Rows(3).Insert: ws.Range("A3:C3").Value = Array("Oeste", 3, 4)   ' dentro: entra
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual "99", t.Cell(2, 2).Shape.TextFrame.TextRange.Text
    AssertEqual -1, CLng(t.Cell(2, 2).Shape.TextFrame.TextRange.Font.Italic), "formatacao do PowerPoint mantida"
    AssertEqual 4, t.Rows.Count, "linha inserida dentro do intervalo"
    AssertEqual "Oeste", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    AssertEqual "Norte", t.Cell(4, 1).Shape.TextFrame.TextRange.Text
    ws.Rows(3).Delete
    s = PptSync_Presentation(pres, False)
    AssertEqual 3, t.Rows.Count, "linha apagada"
    AssertEqual "Norte", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, s
    t.Cell(3, 1).Shape.TextFrame.TextRange.Text = "editado no slide"   ' texto mudado no PowerPoint, origem igual: fica (sem forcar)
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, s
    AssertEqual "editado no slide", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    s = PptSync_Presentation(pres, True)                              ' forcar: o texto volta ao do Excel
    AssertEqual "Norte", t.Cell(3, 1).Shape.TextFrame.TextRange.Text
    ' revisao final F (M11): coluna inserida dentro do intervalo e apagada
    ws.Columns(2).Insert: ws.Range("B1:B3").Value = Application.Transpose(Array("Meta", "1", "2"))
    s = PptSync_Presentation(pres, False)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual 4, t.Columns.Count, "coluna inserida dentro do intervalo"
    AssertEqual "Meta", t.Cell(1, 2).Shape.TextFrame.TextRange.Text
    AssertEqual "2025", t.Cell(1, 3).Shape.TextFrame.TextRange.Text
    AssertEqual "2026", t.Cell(1, 4).Shape.TextFrame.TextRange.Text
    ws.Columns(2).Delete
    s = PptSync_Presentation(pres, False)
    AssertEqual 3, t.Columns.Count, "coluna apagada"
    AssertEqual "2025", t.Cell(1, 2).Shape.TextFrame.TextRange.Text
    AssertEqual "2026", t.Cell(1, 3).Shape.TextFrame.TextRange.Text
    AssertEqual 1, slide.Shapes.Count, "nenhuma forma criada ou apagada"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set t = Nothing: Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_PptInsertTable_Unsaved_NothingCreated()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.Worksheets(1).Range("A1:B2").Value = 1
    wb.Worksheets(1).Range("A1:B2").Select
    Core_ClearLastNotice
    Core_PptInsertTable
    AssertEqual PPTTEXT_MSG_UNSAVED, Core_LastNotice()
    AssertEqual 0, wb.Names.Count, "nenhum nome criado"
    AssertFalse Ppt_InOperation(), "operacao fechada"
    CloseTemp wb
End Sub

' O botao do Excel (Core_PptInsertTable) com o PowerPoint do teste e o slide fixado: mescladas recusadas sem nada criado; falha depois da
' tabela desfeita (sem tabela, sem nome oculto, sem marca); 30 x 10 criada em poucos segundos (tempo medido no texto do assert).
Public Sub Test_PptInsertTable_Core_RefusalUndoAndSize_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, slide As Object, own As New Collection, en As Long, ed As String
    Dim t0 As Single, dt As Single, shp As Object, n0 As Long, n1 As Long
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Ppt_SetAppForTest app
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1:J30").Formula = "=ROW()*10+COLUMN()"
    ws.Range("A1:J1").Value = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j")
    ws.Range("L1:M1").Merge
    wb.SaveAs TempDir() & "\ppttable-core.xlsx", xlOpenXMLWorkbook
    n0 = wb.Names.Count
    wb.Activate: ws.Activate
    ws.Range("K1:M2").Select
    Core_ClearLastNotice
    Core_PptInsertTable
    AssertTrue InStr(Core_LastNotice(), "mescladas") > 0, Core_LastNotice()
    AssertEqual 0, slide.Shapes.Count, "mescladas: nada criado"
    ws.Range("A1:B2").Select
    Util_FailThinSyncForTest = "ppttable-tags:1"
    Core_PptInsertTable
    Util_FailThinSyncForTest = ""
    AssertEqual 0, slide.Shapes.Count, "falha depois da tabela: tabela apagada"
    AssertEqual n0, wb.Names.Count, "falha: nome oculto e marca apagados"
    AssertFalse Ppt_InOperation(), "operacao fechada na falha"
    ' revisao final F (M2, M11): falha DEPOIS de criar a tabela, numa apresentacao nova criada pela chamada: tudo desfeito
    Ppt_PinTargetForTest Nothing, True
    n1 = app.Presentations.Count
    Util_FailThinSyncForTest = "ppttable-after:1"
    Core_PptInsertTable
    Util_FailThinSyncForTest = ""
    AssertTrue Len(Ppt_LastCreatedPresNameForTest()) > 0, "a chamada criou uma apresentacao"
    AssertEqual n1, app.Presentations.Count, "a apresentacao nova foi fechada"
    AssertEqual n0, wb.Names.Count, "falha depois da tabela: nome oculto e marca apagados"
    AssertFalse Ppt_InOperation(), "operacao fechada na falha depois da tabela"
    Ppt_PinTargetForTest slide
    ws.Range("A1:J30").Select
    t0 = Timer
    Core_PptInsertTable
    dt = Timer - t0
    AssertTrue InStr(Core_LastNotice(), "Tabela vinculada criada no slide 1") = 1, Core_LastNotice()
    AssertEqual 1, slide.Shapes.Count
    Set shp = slide.Shapes(1)
    AssertEqual 30, shp.Table.Rows.Count: AssertEqual 10, shp.Table.Columns.Count
    AssertEqual ws.Range("J30").Text, shp.Table.Cell(30, 10).Shape.TextFrame.TextRange.Text
    AssertEqual 1, CLng(shp.Table.Cell(1, 10).Shape.TextFrame.TextRange.ParagraphFormat.Alignment), "cabecalho (texto) a esquerda"
    AssertEqual 3, CLng(shp.Table.Cell(30, 10).Shape.TextFrame.TextRange.ParagraphFormat.Alignment), "numero a direita na ultima linha (herdada)"
    AssertEqual CLng(ws.Range("J30").Font.Size), CLng(shp.Table.Cell(30, 10).Shape.TextFrame.TextRange.Font.Size), "tamanho na ultima linha (herdado)"
    AssertEqual 0, CLng(shp.Table.Cell(30, 10).Shape.TextFrame.TextRange.Font.Bold), "sem negrito"
    AssertEqual n0 + 2, wb.Names.Count, "o nome oculto TQQ_T_ e a marca"
    AssertFalse Ppt_InOperation(), "operacao fechada"
    AssertTrue dt < 20, "30 x 10 em " & Format$(dt, "0.0") & " s"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Util_FailThinSyncForTest = ""
    Set shp = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' Adendo do controlador: as etiquetas da tabela vinculada (celulas com "|", "\", quebra de linha, Tab; reserva planilha|endereco) sobrevivem a
' Salvar como .pptx, fechar e reabrir: a sincronizacao depois de reabrir diz "sem mudanca", atualiza quando a celula muda, e a forma automatica
' continua no indice que o ciclo automatico le (PptAuto_ItemsFor).
Public Sub Test_PptTable_TagsPersistSaveReopen_WhenAvailable()
    Dim wb As Workbook, ws As Worksheet, app As Object, pres As Object, p2 As Object, shp As Object, shp2 As Object, own As New Collection
    Dim en As Long, ed As String, s As String, tmp As String, k As Variant, saved As Object, items As Collection, dirty As Object, slide As Object
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptAuto_ResetTestHooks
    PptAuto_SetNewLinksAuto True
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    If UserWindowAppeared(app) Then GoTo Cleanup
    ws.Range("A1:B1").Value = Array("a|b\c", "linha 1" & vbLf & "linha 2")
    ws.Range("A2:B2").Value = Array("x" & vbTab & "y", 10)
    tmp = TempDir()
    wb.SaveAs tmp & "\ppttable-persist.xlsx", xlOpenXMLWorkbook
    Set shp = PptTable_Insert(ws.Range("A1:B2"), slide)
    AssertEqual "1", PptSync_Tag(shp, PPTSYNC_TAG_AUTO), "novo vinculo automatico"
    AssertEqual "a|b\c", shp.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text
    AssertEqual "linha 1" & Chr$(11) & "linha 2", shp.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text, "quebra de linha da celula"
    Set saved = CreateObject("Scripting.Dictionary")
    For Each k In Array(PPTTABLE_TAG, PPTSYNC_TAG_BOOK, PPTLINK_TAG_NAME, PPTLINK_TAG_REF, PPTTABLE_TAG_LAST)
        saved.Add CStr(k), PptSync_Tag(shp, CStr(k))
    Next
    AssertTrue InStr(CStr(saved(PPTTABLE_TAG_LAST)), vbLf) > 0, "premissa: duas linhas na etiqueta"
    pres.SaveAs tmp & "\ppttable-persist.pptx"
    Set shp = Nothing: Set slide = Nothing
    pres.Close
    Set pres = Nothing
    Set own = New Collection
    Set p2 = app.Presentations.Open(tmp & "\ppttable-persist.pptx", -1, 0, 0): own.Add p2
    Set shp2 = p2.Slides(1).Shapes(1)
    For Each k In saved.Keys
        AssertEqual CStr(saved(k)), PptSync_Tag(shp2, CStr(k)), "etiqueta " & k & " igual depois de reabrir"
    Next
    AssertEqual "table", PptSync_ItemKind(shp2)
    s = PptSync_Presentation(p2, False)
    AssertTrue InStr(s, "Sem mudança: 1.") > 0, "depois de reabrir: " & s
    Set dirty = CreateObject("Scripting.Dictionary")
    dirty.CompareMode = 1
    dirty.Add wb.FullName, True
    Set items = PptAuto_ItemsFor(p2, dirty)
    AssertEqual 1, items.Count, "o ciclo automatico acha a tabela pelo indice"
    ws.Range("B2").Value = 20
    s = PptSync_Items(items, False, True)
    AssertTrue InStr(s, "Gráficos do ThinQQ atualizados: 1.") = 1, s
    AssertEqual "20", shp2.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptAuto_ResetTestHooks
    Set items = Nothing: Set shp = Nothing: Set shp2 = Nothing: Set slide = Nothing: Set pres = Nothing: Set p2 = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' ---- 1.11.0 (U1): estado (bandeira TQQ_STATE), indice de todos os vinculos (TQQ_LINKIDX), motor marca os manuais, API 6 ----

Public Sub Test_PptState_SetTag_OnlyWhenChanged_Fake()
    Dim f As TestPptFake, t As TestPptTags
    Set f = New TestPptFake: Set t = New TestPptTags: Set f.Tags = t
    PptState_ResetTestHooks
    AssertFalse PptState_SetTag(f, False), "sem bandeira e sem dados novos: nada gravado"
    AssertTrue PptState_SetTag(f, True)
    AssertEqual PPTSTATE_NEW, t.Item(PPTSTATE_TAG)
    AssertFalse PptState_SetTag(f, True), "ja marcada: nao regrava"
    AssertEqual 1, t.Writes
    AssertTrue PptState_SetTag(f, False)
    AssertEqual "", t.Item(PPTSTATE_TAG)
    AssertEqual 1, t.Deletes
    t.Preset PPTSTATE_TAG, "lixo"
    AssertFalse PptState_SetTag(f, False), "valor estranho (deck alheio) nao e bandeira: nada a apagar"
    t.FailOnKey = PPTSTATE_TAG
    AssertFalse PptState_SetTag(f, True), "falha comum ao gravar: registrada, sem erro"
    AssertEqual 2, PptState_MarksForTest()
End Sub

' Sem indice TQQ_LINKIDX: nenhum manual (nunca varre as formas); com o indice, so as formas da pasta suja, nao automaticas, do ThinQQ.
Public Sub Test_PptState_ManualFor_IndexOnly_Fake()
    Dim shp As TestPptFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, dirty As Object, lk As TestPptFake
    Set shp = New TestPptFake: Set shp.Tags = CreateObject("Scripting.Dictionary")
    shp.Tags.Add PPTSYNC_TAG_CHART, "abc"
    shp.Tags.Add PPTSYNC_TAG_BOOK, "C:\x\a.xlsx"
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add shp, "tqq_m"
    Set slides = New TestPptFake: Set slides.Slide = sl: slides.SlideIdValue = 300
    Set pres = New TestPptFake: Set pres.Slides = slides: Set pres.Tags = CreateObject("Scripting.Dictionary")
    Set dirty = CreateObject("Scripting.Dictionary"): dirty.CompareMode = 1
    dirty.Add "C:\x\a.xlsx", True
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "sem indice"
    AssertEqual 0, slides.FindBySlideIDCalls, "sem indice: nenhuma chamada ao PowerPoint"
    pres.Tags(PPTSTATE_TAG_LINKIDX) = PptIdx_Entry(300, "tqq_m", "C:\x\a.xlsx")    ' o duble (Dictionary) ja tem a chave vazia da leitura acima
    AssertEqual 1, PptState_ManualFor(pres, dirty).Count
    shp.Tags(PPTSYNC_TAG_AUTO) = "1"
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "automatica fica com o ciclo automatico"
    shp.Tags.Remove PPTSYNC_TAG_AUTO
    dirty.RemoveAll: dirty.Add "C:\x\outra.xlsx", True
    AssertEqual 0, PptState_ManualFor(pres, dirty).Count, "pasta que nao mudou"
    AssertEqual 2, slides.FindBySlideIDCalls, "entrada de pasta limpa nao chega ao PowerPoint"
End Sub

' Processa mensagens por secs segundos (DoEvents e uma chamada ao PowerPoint de teste, como o Excel ocioso faria antes do OnTime do ciclo).
Private Sub PumpMessages(app As Object, secs As Single)
    Dim t0 As Single, n As Long
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + secs
        DoEvents
        n = CLng(app.Presentations.Count)
    Loop
End Sub

' Bandeira (Review Focus 1): colado manual, B2 muda, o ciclo marca TQQ_STATE sem sincronizar o slide nem tocar no Excel, com a apresentacao salva
' continuando salva; Atualizar tira a bandeira; o dado voltando ao valor do slide tira a bandeira sozinho.
Public Sub Test_PptAuto_Tick_FlagsManual_NoExcelChange_KeepsSaved_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, own As New Collection, en As Long, ed As String
    Dim fp0 As String, pp As String, w0 As Double, msg As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    PptState_ResetTestHooks
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    Set s1 = PasteLinked(app, pres, wb, rng, TempDir() & "\pptstate-flag.xlsx")
    AssertEqual "", PptSync_Tag(s1, PPTSYNC_TAG_AUTO), "nasce manual"
    AssertTrue InStr(1, PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CStr(s1.Name), vbBinaryCompare) > 0, "colagem entra no indice de todos os vinculos"
    pp = TempDir() & "\pptstate-flag.pptx"
    pres.SaveAs pp
    AssertTrue CLng(pres.Saved) <> 0, "apresentacao salva"
    fp0 = PptSync_Tag(s1, PPTSYNC_TAG_FP)
    w0 = rng.Worksheet.ChartObjects(1).Width
    Watchers_Start
    rng.Worksheet.Range("B2").Value = 77
    AssertEqual 1, PptAuto_PendingCountForTest()
    ' MEDIDO (U1): com a pasta aberta neste Excel, o grafico colado (LinkFormat.AutoUpdate = 1, manual) recebe os valores novos pelo vinculo
    ' OLE do proprio Office em ~0,3 s quando as mensagens sao processadas, e o Excel marca a pasta (e o PowerPoint a apresentacao) como
    ' alteradas. Isso nao e o ThinQQ: o teste deixa esse efeito acontecer (o ciclo real so roda 2 s depois da ultima mudanca), marca as duas
    ' como salvas e prova que o CICLO nao altera nada.
    PumpMessages app, 2.5
    wb.Saved = True                                     ' qualquer coisa que o ciclo fizer no Excel apareceria como alteracao
    pres.Saved = -1
    AssertTrue CLng(pres.Saved) <> 0, "apresentacao marcada como salva"
    Ppt_SetAppForTest app
    PptSync_ResetTestHooks
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG), "bandeira"
    AssertEqual 1, PptAuto_LastMarkedForTest()
    AssertEqual fp0, PptSync_Tag(s1, PPTSYNC_TAG_FP), "manual: o ThinQQ nao sincronizou o slide (TQQ_FP intacta)"
    AssertEqual 0, PptSync_UpdateCountForTest(), "nenhum LinkFormat.Update"
    AssertTrue wb.Saved, "o ciclo nao alterou a pasta do Excel (o Desfazer do usuario fica)"
    AssertAlmost w0, rng.Worksheet.ChartObjects(1).Width, 0.01, "o grafico do Excel nao foi redimensionado"
    AssertTrue CLng(pres.Saved) <> 0, "a bandeira nao deixa a apresentacao alterada (P-U1)"
    msg = Core_PptSyncShapes(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue InStr(msg, "Gráficos do ThinQQ atualizados: 1.") = 1, msg
    AssertAlmost 77, PptSeriesValue(s1.Chart, "Sul", 1), 0.0001
    AssertEqual "", PptSync_Tag(s1, PPTSTATE_TAG), "atualizar tira a bandeira"
    rng.Worksheet.Range("B2").Value = 78
    PptAuto_Tick
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG), "de novo"
    rng.Worksheet.Range("B2").Value = 77
    PptAuto_Tick
    AssertEqual "", PptSync_Tag(s1, PPTSTATE_TAG), "o dado voltou ao valor do slide: a bandeira sai"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    If Len(pp) > 0 Then Kill pp
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

' API 6: linhas de estado (manual e automatica), bandeira gravada, pasta fechada "unknown" sem verificar e verificada com abertura somente
' leitura (fechada no fim), indice reconstruido com entries = "".
Public Sub Test_Core_PptLinkStates_Lines_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, s1 As Object, s2 As Object, own As New Collection, en As Long, ed As String
    Dim lines As String, a() As String, f() As String, p As String, nBooks As Long, i As Long, l1 As String, l2 As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    AutoBegin
    Set pres = app.Presentations.Add(0): own.Add pres
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    If UserWindowAppeared(app) Then GoTo Cleanup
    p = TempDir() & "\pptstate-lines.xlsx"
    Set s1 = PasteLinked(app, pres, wb, rng, p)
    Set s2 = PasteAgain(app, pres)
    PptSync_SetAuto s2, True
    PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX
    rng.Worksheet.Range("B2").Value = 66
    lines = Core_PptLinkStates(pres, "", False)
    AssertTrue Left$(lines, 1) <> "!", lines
    a = Split(lines, vbLf)
    AssertEqual 2, UBound(a) + 1, lines
    For i = 0 To 1
        f = Split(a(i), "|")
        AssertEqual 5, UBound(f), a(i)
        AssertTrue Left$(f(3), 6) = "chart:", a(i)
        ' a colagem repetida pode ter o MESMO nome em outro slide: a linha e achada pelo slide e pelo nome
        If f(0) = CStr(s1.Parent.SlideID) And f(1) = CStr(s1.Name) Then l1 = a(i)
        If f(0) = CStr(s2.Parent.SlideID) And f(1) = CStr(s2.Name) Then l2 = a(i)
    Next
    AssertTrue Right$(l1, Len("|manual|new")) = "|manual|new", l1
    AssertTrue Right$(l2, Len("|auto|new")) = "|auto|new", l2
    AssertEqual PPTSTATE_NEW, PptSync_Tag(s1, PPTSTATE_TAG)
    AssertEqual 2, PptIdx_Parse(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX)).Count, "indice de todos os vinculos reconstruido"
    Core_PptSyncShapes pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), False)
    AssertTrue Right$(lines, Len("|manual|ok")) = "|manual|ok", lines
    wb.Save
    nBooks = Workbooks.Count
    wb.Close SaveChanges:=False: Set wb = Nothing
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), False)
    AssertTrue Right$(lines, Len("|unknown")) = "|unknown", "pasta fechada sem verificar: " & lines
    lines = Core_PptLinkStates(pres, PptIdx_Entry(CLng(s1.Parent.SlideID), CStr(s1.Name), ""), True)
    AssertTrue Right$(lines, Len("|ok")) = "|ok", "verificada (aberta somente leitura): " & lines
    AssertEqual nBooks - 1, Workbooks.Count, "a pasta aberta para verificar foi fechada"
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    AutoReset
    Set s1 = Nothing: Set s2 = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    AutoEnd
    If Len(p) > 0 Then Kill p
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub

Public Sub Test_Core_PptLinkStates_NeverRaises()
    Dim s As String
    s = Core_PptLinkStates(Nothing, "", False)
    AssertTrue Left$(s, 1) = "!", s
    AssertFalse Core_IsQuiet()
End Sub

' Revisao final M3: sem "Verificar" (abertura, abrir o dialogo) o disco nao e tocado: pasta fechada local inexistente = "unknown", nenhuma
' sonda; com "Verificar", "missing". M2: numa apresentacao somente leitura nada e gravado (indice, bandeira) nem renomeado.
Public Sub Test_PptState_Lines_NoProbeWithoutVerify_ReadOnlyWritesNothing_WhenAvailable()
    Dim app As Object, pres As Object, ro As Object, s As Object, own As New Collection, en As Long, ed As String, lines As String, p As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    PptState_ResetTestHooks
    Set pres = app.Presentations.Add(0): own.Add pres
    If UserWindowAppeared(app) Then GoTo Cleanup
    pres.Slides.Add 1, 12
    Set s = pres.Slides(1).Shapes.AddShape(1, 10, 10, 50, 50)
    s.Name = "txt_m3": s.Tags.Add PPTTEXT_TAG, "1": s.Tags.Add PPTSYNC_TAG_BOOK, TempDir() & "\nao-existe-m3.xlsx"
    lines = Core_PptLinkStates(pres, "", False)
    AssertTrue Right$(lines, Len("|text|manual|unknown")) = "|text|manual|unknown", "sem Verificar: " & lines
    AssertEqual 0, PptState_ProbesForTest(), "sem Verificar: nenhuma sonda do disco"
    lines = Core_PptLinkStates(pres, "", True)
    AssertTrue Right$(lines, Len("|missing")) = "|missing", "com Verificar: " & lines
    AssertEqual 1, PptState_ProbesForTest(), "Verificar sonda a pasta fechada"
    ' somente leitura: duas formas com o mesmo nome, sem indice gravado
    Set s = pres.Slides(1).Shapes.AddShape(1, 70, 10, 50, 50)
    s.Name = "txt_m3": s.Tags.Add PPTTEXT_TAG, "1": s.Tags.Add PPTSYNC_TAG_BOOK, TempDir() & "\nao-existe-m3.xlsx"
    PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX
    p = TempDir() & "\m2-somente-leitura.pptx"
    pres.SaveAs p
    pres.Close
    Set pres = Nothing
    Set ro = app.Presentations.Open(p, -1, 0, 0): own.Add ro
    AssertTrue CLng(ro.ReadOnly) <> 0, "premissa: somente leitura"
    lines = Core_PptLinkStates(ro, "", False)
    AssertTrue Left$(lines, 1) <> "!", lines
    AssertEqual 2, UBound(Split(lines, vbLf)) + 1, "os estados sao calculados: " & lines
    AssertEqual "", PptSync_Tag(ro, PPTSTATE_TAG_LINKIDX), "somente leitura: indice nao gravado"
    AssertEqual "txt_m3", CStr(ro.Slides(1).Shapes(2).Name), "somente leitura: nada renomeado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set s = Nothing: Set pres = Nothing: Set ro = Nothing
    StopScratchPowerPoint app, own
    If Len(p) > 0 Then Kill p
    PptState_ResetTestHooks
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptSync", ed
End Sub
' Sem "Verificar" (abertura da apresentacao), PptState_Lines nunca sonda caminho de rede/web vindo do deck (mesma regra do .ppam).
Public Sub Test_PptState_IsRemotePath_Pure()
    AssertTrue PptState_IsRemotePath("\\servidor\pasta\a.xlsx")
    AssertTrue PptState_IsRemotePath("//servidor/pasta/a.xlsx")
    AssertTrue PptState_IsRemotePath("\\?\UNC\servidor\a.xlsx")
    AssertTrue PptState_IsRemotePath("https://empresa.sharepoint.com/a.xlsx")
    AssertTrue PptState_IsRemotePath("file://servidor/a.xlsx")
    AssertFalse PptState_IsRemotePath("C:\dados\a.xlsx")
    AssertFalse PptState_IsRemotePath("\\?\C:\dados\a.xlsx")
    AssertFalse PptState_IsRemotePath("")
End Sub