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

Private Function PowerPointProcessCount() As Long
    PowerPointProcessCount = GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count
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

' So cria o PowerPoint se nenhum estiver aberto (ROT nem processo). Nothing (e o motivo no log) se pulou.
Private Function StartScratchPowerPoint() As Object
    Dim app As Object
    Ppt_ResetTestHooks
    If Ppt_IsRunning() Then Errors_Log "TestPptSync: pulado (ha um PowerPoint aberto)": Exit Function
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPptSync: pulado (processo POWERPNT existente)": Exit Function
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
        Do While PowerPointProcessCount() > 0
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
    Dim wb As Workbook, rng As Range, c As Chart, shp As TestPptFake, w0 As Double, h0 As Double, pt As Variant, en As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    w0 = c.Parent.Width: h0 = c.Parent.Height
    Set shp = New TestPptFake: shp.Width = 709.7: shp.Height = 432      ' sem Chart: Ppt_LinkState = -1 (sem Update)
    For Each pt In Array("snapshot", "apply")
        wb.Saved = True
        PptSync_SetFailForTest CStr(pt)
        On Error Resume Next
        PptSync_SyncShape shp, c.Parent
        en = Err.Number
        On Error GoTo 0
        AssertTrue en <> 0, "a falha injetada sobe (" & pt & ")"
        AssertAlmost w0, c.Parent.Width, 0.01, "largura restaurada (" & pt & ")"
        AssertAlmost h0, c.Parent.Height, 0.01, "altura restaurada (" & pt & ")"
        AssertTrue wb.Saved, "pasta continua salva (" & pt & ")"
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
    Set shp = slide.Shapes(1)
    Set seq = slide.TimeLine.MainSequence
    seq.AddEffect shp, 10                                       ' msoAnimEffectFade
    rng.Worksheet.Range("B3").Value = 40                        ' sem salvar
    Application.Calculate
    Ann_RemoveAll c
    Ann_AddDifference c, 1, 2
    PptSync_SyncShape shp, c.Parent
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
