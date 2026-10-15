Attribute VB_Name = "TestPptSync"
Option Explicit
' Parte A (1.4.0): motor de sincronizacao dos graficos ThinQQ colados no PowerPoint (modPptSync). Os testes com PowerPoint real so rodam sem
' nenhum POWERPNT aberto (aplicativo de instancia unica: CreateObject anexaria ao do usuario); senao saem verdes registrando o motivo.

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
    AssertEqual 2, Core_PptApiVersion(), "contrato com o suplemento do PowerPoint (spec 4.9; 2 = 1.6.0)"
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

' Contrato com o .ppam (TestPptAddin.Test_PptIdx_Literal tem as MESMAS cadeias): entradas por Chr(30), campos por Chr(31).
Public Sub Test_PptIdx_Literal()
    Dim s As String, a As String, b As String
    a = "256" & Chr$(31) & "tqq_abc" & Chr$(31) & "C:\a\b.xlsx"
    b = "257" & Chr$(31) & "tqq_def" & Chr$(31) & "C:\a\c.xlsx"
    s = PptIdx_Set("", 256, "tqq_abc", "C:\a\b.xlsx", True)
    AssertEqual a, s
    s = PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True)
    AssertEqual a & Chr$(30) & b, s
    AssertEqual a & Chr$(30) & b, PptIdx_Set(s, 257, "tqq_def", "C:\a\c.xlsx", True), "sem repeticao"
    AssertEqual b, PptIdx_Set(s, 256, "tqq_abc", "", False), "desligar tira a entrada"
    AssertEqual "", PptIdx_Set(a, 256, "tqq_abc", "", False)
    AssertEqual 2, PptIdx_Parse(s & Chr$(30) & "x" & Chr$(31) & "y" & Chr$(31) & "z" & Chr$(30) & Chr$(30)).Count, "entrada invalida e vazia ignoradas"
    AssertEqual a, PptIdx_Entry(256, "tqq_abc", "C:\a\b.xlsx")
End Sub

Public Sub Test_PptSync_ItemKind_ByTags()
    Dim f As TestPptFake
    Set f = New TestPptFake
    AssertEqual "", PptSync_ItemKind(f), "duble sem Tags: nenhum tipo (sem erro)"
End Sub

Public Sub Test_PptAuto_Flags_TestBuildInMemory()
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
    AssertEqual PPTSYNC_NONE, Core_PptSyncShapes(pres, PptIdx_Entry(999999, "nada", ""), True), "entrada que nao existe"
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
