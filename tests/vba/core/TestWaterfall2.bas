Attribute VB_Name = "TestWaterfall2"
Option Explicit
' Fase 2: cascata com varias series, build-down ('e' na primeira linha) e cascata horizontal.

Private Function MultiFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("Inicio", 100, 50)
    ws.Range("A3:C3").Value = Array("Ajuste", -30, 10)
    ws.Range("A4:C4").Value = Array("Subtotal", "e", "e")
    ws.Range("A5:C5").Value = Array("Queda", -200, -20)
    ws.Range("A6:B6").Value = Array("Fim", "e")
    Set MultiFixture = ChartData_FromRange(ws.Range("A1:C6"))
End Function

' Review Focus 2: coluna com sinais misturados que cruza o zero (Queda: inicio 130, -200 e -20).
Public Sub Test_WaterfallMulti_Block_MixedSignsCrossingZero()
    Dim wb As Workbook, body As Range, blk As Range
    Set wb = NewTempWorkbook()
    Set blk = Waterfall_BuildBlock(wb, "wfm001", MultiFixture(wb)): Set body = Block_Body(blk)
    AssertEqual 40, blk.Columns.Count, "20 + 10 por serie"
    ' End: 150, 130, 130, -90, -90
    AssertEqual 150, body.Cells(2, WF_COL_END).Value: AssertEqual 130, body.Cells(3, WF_COL_END).Value
    AssertEqual 130, body.Cells(4, WF_COL_END).Value: AssertEqual -90, body.Cells(5, WF_COL_END).Value
    AssertEqual -90, body.Cells(6, WF_COL_END).Value
    ' Inicio: duas series positivas empilhadas a partir de 0
    AssertEqual 100, body.Cells(2, Waterfall_SeriesCol(1, "UP")).Value: AssertEqual 50, body.Cells(2, Waterfall_SeriesCol(2, "UP")).Value
    AssertEqual 50, body.Cells(2, Waterfall_SeriesCol(1, "M")).Value: AssertEqual 125, body.Cells(2, Waterfall_SeriesCol(2, "M")).Value
    ' Ajuste: Loja desce 30 a partir de 150 (120..150), Online sobe 10 (150..160); Base = 120
    AssertEqual 120, body.Cells(3, Waterfall_SeriesCol(1, "A")).Value: AssertEqual 150, body.Cells(3, Waterfall_SeriesCol(1, "B")).Value
    AssertEqual 30, body.Cells(3, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 0, body.Cells(3, Waterfall_SeriesCol(1, "UP")).Value
    AssertEqual 150, body.Cells(3, Waterfall_SeriesCol(2, "A")).Value: AssertEqual 160, body.Cells(3, Waterfall_SeriesCol(2, "B")).Value
    AssertEqual 10, body.Cells(3, Waterfall_SeriesCol(2, "UP")).Value
    AssertEqual 120, body.Cells(3, WF_COL_BASE).Value: AssertEqual 120, body.Cells(3, 7).Value: AssertEqual 160, body.Cells(3, 8).Value
    ' Subtotal: total 130, sem segmentos
    AssertEqual True, body.Cells(4, 3).Value: AssertEqual 130, body.Cells(4, WF_COL_TOTAL).Value
    AssertEqual 0, body.Cells(4, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 130, body.Cells(4, WF_COL_LABELVAL).Value
    ' Queda: Loja -200 cruza o zero (-70..130): DN 130 acima, DNN -70 abaixo; Online -20 (-90..-70): DNN -20
    AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(1, "A")).Value: AssertEqual 130, body.Cells(5, Waterfall_SeriesCol(1, "B")).Value
    AssertEqual 130, body.Cells(5, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(1, "DNN")).Value
    AssertEqual -90, body.Cells(5, Waterfall_SeriesCol(2, "A")).Value: AssertEqual -70, body.Cells(5, Waterfall_SeriesCol(2, "B")).Value
    AssertEqual 0, body.Cells(5, Waterfall_SeriesCol(2, "DN")).Value: AssertEqual -20, body.Cells(5, Waterfall_SeriesCol(2, "DNN")).Value
    AssertEqual 0, body.Cells(5, WF_COL_BASE).Value: AssertEqual -90, body.Cells(5, 7).Value: AssertEqual 130, body.Cells(5, 8).Value
    AssertEqual -220, body.Cells(5, 4).Value, "delta liquido"
    AssertTrue IsError(body.Cells(5, WF_COL_LABELVAL).Value), "linha sem total: rotulo de total NA"
    ' Fim: total -90, rotulo em Lo
    AssertEqual True, body.Cells(6, 3).Value: AssertEqual -90, body.Cells(6, WF_COL_TOTAL).Value: AssertEqual -90, body.Cells(6, WF_COL_LABELPOS).Value
    AssertEqual 4.5, body.Cells(2, WF_COL_POS).Value: AssertEqual False, body.Cells(2, WF_COL_BD).Value
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Insert_SeriesOrderColorsLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, pal As Variant, names As String, s As Series
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertEqual 2, Waterfall_SeriesCount(cht)
    AssertEqual 13, cht.SeriesCollection.Count, "base + 4 por serie + total + 2 rotulos de serie + rotulo de total"
    For Each s In cht.SeriesCollection: names = names & s.Name & ";": Next
    AssertEqual "tqq_base;tqq_dn_2;tqq_dn_1;Loja;Online;tqq_upn_2;tqq_upn_1;tqq_dnn_1;tqq_dnn_2;tqq_total;tqq_lbl_1;tqq_lbl_2;tqq_lbl;", names
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "tqq_dnn_1").Format.Fill.ForeColor.RGB, "partes da serie 1 com a mesma cor"
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "tqq_dn_2").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(3)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB, "totais em cinza"
    AssertEqual msoFalse, Chart_FindSeries(cht, "tqq_base").Format.Fill.Visible
    AssertTrue cht.HasLegend
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Loja e Online"
    AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertEqual "5", Registry_GetOption(cht, "n")
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Labels_SegmentsAndTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, l1 As Series, l2 As Series, lt As Series
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2"): Set lt = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual "100", l1.Points(1).DataLabel.Text
    AssertEqual "-30", l1.Points(2).DataLabel.Text
    AssertEqual "10", l2.Points(2).DataLabel.Text
    AssertEqual xlLabelPositionCenter, l1.Points(2).DataLabel.Position
    AssertEqual "-200", l1.Points(4).DataLabel.Text: AssertEqual "-20", l2.Points(4).DataLabel.Text
    AssertEqual "130", lt.Points(3).DataLabel.Text: AssertEqual xlLabelPositionAbove, lt.Points(3).DataLabel.Position
    AssertEqual "-90", lt.Points(5).DataLabel.Text: AssertEqual xlLabelPositionBelow, lt.Points(5).DataLabel.Position
    Waterfall_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing And Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    AssertEqual 10, cht.SeriesCollection.Count
    Waterfall_SetLabels cht, True
    AssertEqual 13, cht.SeriesCollection.Count
    ' auxiliares com pontos NA() entram na varredura da limpeza (Tarefa 0): o grafico novo nao pode bloquear nem falhar
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com a cascata multi remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos rotulos nao sao falha de leitura"
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Connectors_FollowNetEnd()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, it As Double, ih As Double
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight: Set ax = cht.Axes(xlValue)
    AssertAlmost it + ih * (ax.MaximumScale - 150) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_1").Top, 1
    AssertAlmost it + ih * (ax.MaximumScale - 130) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_2").Top, 1, "conector no liquido (-30+10)"
    AssertAlmost it + ih * (ax.MaximumScale + 90) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_4").Top, 1
    CloseTemp wb
End Sub

' 'e' na primeira linha + numero na ultima: o primeiro total e calculado para tras (think-cell), a ultima linha vira total.
Public Sub Test_Waterfall_BuildDown_FirstRowE_LastRowNumber()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, body As Range, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:B3").Value = Array("Custos", -700)
    ws.Range("A4:B4").Value = Array("Despesas", -250)
    ws.Range("A5:B5").Value = Array("Lucro", 250)
    Set d = ChartData_FromRange(ws.Range("A1:B5"))
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfbd01", d))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual True, body.Cells(2, 3).Value: AssertEqual 1200, body.Cells(2, WF_COL_END).Value, "Receita = 250 - (-700 - 250)"
    AssertEqual 1200, body.Cells(2, WF_COL_TOTAL).Value: AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual 500, body.Cells(3, WF_COL_END).Value: AssertEqual 250, body.Cells(4, WF_COL_END).Value
    AssertEqual True, body.Cells(5, 3).Value, "ultima linha vira total": AssertEqual 250, body.Cells(5, WF_COL_TOTAL).Value
    AssertEqual 250, body.Cells(5, WF_COL_LABELVAL).Value
    ws.Range("B5").Value = 300
    AssertEqual 1250, body.Cells(2, WF_COL_END).Value, "recalcula ao mudar o total final"
    Set cht = Waterfall_Insert(ws, d).Chart
    AssertEqual "1" & Application.International(xlThousandsSeparator) & "250", Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text, "formato int com separador de milhar"
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_TooManySeries_Refused()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:H1").Value = Array("Etapa", "S1", "S2", "S3", "S4", "S5", "S6", "S7")
    ws.Range("A2:H2").Value = Array("x", 1, 1, 1, 1, 1, 1, 1)
    ws.Range("A3:H3").Value = Array("y", 1, 1, 1, 1, 1, 1, 1)
    Set d = ChartData_FromRange(ws.Range("A1:H3"))
    On Error Resume Next
    Waterfall_Insert ws, d
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 6007, en
    AssertEqual "A cascata aceita no máximo 6 séries.", ed
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nenhum bloco alocado"
    Next
    ws.Range("A1:H3").Select
    Core_ClearLastNotice
    Core_Insert "waterfall"
    AssertEqual "A cascata aceita no máximo 6 séries.", Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_WaterfallMulti_Insert_RollsBackOnFailure()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Waterfall_Insert ws, d
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub

' ---------- Tarefa 4: legenda, auxiliares, bloco ----------

Private Function SingleFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Inicio", 100)
    ws.Range("A3:B3").Value = Array("A", -150)
    ws.Range("A4:B4").Value = Array("B", 30)
    ws.Range("A5:B5").Value = Array("Subtotal", "e")
    ws.Range("A6:B6").Value = Array("C", 70)
    ws.Range("A7:B7").Value = Array("Fim", "e")
    Set SingleFixture = ChartData_FromRange(ws.Range("A1:B7"))
End Function

' Cores das chaves da legenda, na ordem em que a legenda lista as entradas.
Private Function LegendKeys(cht As Chart) As String
    Dim k As Long
    For k = 1 To cht.Legend.LegendEntries.Count
        LegendKeys = LegendKeys & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB & ";"
    Next
End Function

' Numero de conectores da cascata desenhados no grafico (formas tqq_ann_conn_*).
Private Function ConnectorCount(cht As Chart) As Long
    Dim shp As Shape
    For Each shp In cht.Shapes
        If Left$(shp.Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then ConnectorCount = ConnectorCount + 1
    Next
End Function

' Regressao da Fase 1: com uma serie o bloco tem 20 colunas (as 18 da Fase 1 + Pos + BD), o grafico tem as 6 series e o
' rotulo da Fase 1, sem legenda, e a meta registra o tipo.
Public Sub Test_Waterfall_SingleSeries_BlockAndChartUnchanged()
    Dim wb As Workbook, d As clsChartData, blk As Range, body As Range, cht As Chart, names As String, s As Series, i As Long
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set blk = Waterfall_BuildBlock(wb, "wfs001", d): Set body = Block_Body(blk)
    AssertEqual 20, blk.Columns.Count
    AssertEqual "waterfall", CStr(Block_Meta(blk).Cells(1, 2).Value), "a meta registra o tipo"
    AssertEqual "BD", CStr(body.Cells(1, WF_COL_BD).Value): AssertEqual "Pos", CStr(body.Cells(1, WF_COL_POS).Value)
    For i = 1 To 6
        AssertEqual 6.5 - i, body.Cells(i + 1, WF_COL_POS).Value, "Pos = n - i + 0,5 (linha " & i & ")"
        AssertEqual False, body.Cells(i + 1, WF_COL_BD).Value, "BD falso (a primeira linha e numero)"
    Next
    AssertEqual 50, body.Cells(7, WF_COL_END).Value
    wb.Names("tqq_wfs001").Delete
    Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    AssertEqual 1, Waterfall_SeriesCount(cht)
    For Each s In cht.SeriesCollection: names = names & s.Name & ";": Next
    AssertEqual "tqq_base;tqq_uppos;tqq_upneg;tqq_downpos;tqq_downneg;tqq_total;tqq_lbl;", names
    AssertFalse cht.HasLegend: AssertEqual "", Registry_GetOption(cht, "legend")
    AssertEqual 20, Registry_Block(cht).Columns.Count
    CloseTemp wb
End Sub

' O tipo gravado na meta do bloco vem do argumento (Tarefa 5: barwaterfall).
Public Sub Test_Waterfall_BuildBlock_RecordsKind()
    Dim wb As Workbook, blk As Range
    Set wb = NewTempWorkbook()
    Set blk = Waterfall_BuildBlock(wb, "wfk001", SingleFixture(wb), "barwaterfall")
    AssertEqual "barwaterfall", CStr(Block_Meta(blk).Cells(1, 2).Value)
    AssertEqual "barwaterfall", CStr(wb.Names("tqq_wfk001").RefersToRange.Cells(1, 2).Value)
    CloseTemp wb
End Sub

' A legenda da cascata so existe com varias series (design secao 7): a matriz e estatica, entao a pergunta leva o numero de series.
Public Sub Test_Waterfall_LegendSupport_FollowsSeriesCount()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    AssertFalse Kinds_Supports("waterfall", "legend"), "sem informar o numero de series: sem legenda"
    AssertFalse Kinds_Supports("waterfall", "legend", 1): AssertTrue Kinds_Supports("waterfall", "legend", 2)
    AssertFalse Kinds_Supports("barwaterfall", "legend", 1): AssertTrue Kinds_Supports("barwaterfall", "legend", 6)
    AssertTrue Kinds_Supports("stacked", "legend"): AssertTrue Kinds_Supports("stacked", "legend", 1): AssertFalse Kinds_Supports("gantt", "legend", 3)
    AssertTrue Kinds_Supports("waterfall", "connectors", 1), "o numero de series so afeta a legenda"
    Set wb = NewTempWorkbook(): Set ws = MultiFixture(wb).Source.Worksheet
    ws.Range("A1:B6").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual 1, Waterfall_SeriesCount(cht)
    AssertFalse Core_Supports("legend"), "cascata de uma serie: sem legenda"
    AssertFalse Kinds_SupportsChart(cht, "legend"), "a pergunta por grafico (comandos de opcao) responde igual"
    AssertTrue Core_Supports("connectors"): AssertTrue Core_Supports("labels")
    ws.Range("A1:C6").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual 2, Waterfall_SeriesCount(cht)
    AssertTrue Core_Supports("legend"), "cascata de varias series: legenda"
    AssertTrue Kinds_SupportsChart(cht, "legend")
    ' "Alterar dados": de 2 series para 1 e de volta; a legenda e a opcao acompanham
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B6"))
    Set cht = Core_ActiveChart()
    AssertEqual 1, Waterfall_SeriesCount(cht)
    AssertFalse Core_Supports("legend"): AssertFalse cht.HasLegend
    AssertEqual "right", Registry_GetOption(cht, "legend"), "a escolha salva (da legenda de 2 series) fica: a cascata de 1 serie nao tem legenda, mas a preferencia nao se perde"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:C6"))
    Set cht = Core_ActiveChart()
    AssertTrue Core_Supports("legend"): AssertTrue cht.HasLegend
    AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Loja e Online"
    AssertFalse Core_ChartSourceMismatch(cht)
    Watchers_Stop
    CloseTemp wb
End Sub

' A legenda "nenhuma" escolhida numa cascata de varias series e uma ESCOLHA: "Alterar dados" a reaplica, e ela sobrevive a
' passar por uma tabela de 1 serie (sem legenda possivel) e voltar a varias (carry-over da Tarefa 4: antes a legenda voltava).
Public Sub Test_Waterfall_Rebuild_KeepsLegendNoneChoice_AcrossSingleSeriesAndBack()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = MultiFixture(wb).Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertTrue cht.HasLegend
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend: AssertEqual "none", Core_GetOption("legend")
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:C6"))
    Set cht = Core_ActiveChart()
    AssertFalse cht.HasLegend, "a escolha 'nenhuma' volta a valer na reconstrucao"
    AssertEqual "none", Registry_GetOption(cht, "legend")
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B6"))
    Set cht = Core_ActiveChart()
    AssertFalse Core_Supports("legend"): AssertEqual "none", Registry_GetOption(cht, "legend"), "1 serie: a escolha fica gravada"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:C6"))
    Set cht = Core_ActiveChart()
    AssertTrue Core_Supports("legend")
    AssertFalse cht.HasLegend, "de volta a varias series a escolha 'nenhuma' continua valendo"
    Core_SetOption "legend", "right"
    AssertTrue cht.HasLegend: AssertEqual 2, cht.Legend.LegendEntries.Count
    Core_ClearLastNotice
    ws.Range("A1:B6").Select
    Core_Insert "waterfall"
    Core_SetOption "legend", "right"
    AssertEqual "Legenda não se aplica a este tipo de gráfico.", Core_LastNotice(), "cascata de 1 serie: legenda recusada com o aviso do brief da Tarefa 11"
    AssertFalse Core_ActiveChart().HasLegend
    Watchers_Stop
    CloseTemp wb
End Sub

' Auxiliares tqq_* ficam INTERCALADAS com as series de dados no mesmo grupo empilhado: a legenda volta a esconder so as
' auxiliares depois de mudar a posicao, de religar a legenda e de religar os rotulos, e a ordem das entradas reais e a
' que o Excel lista (verificada tambem por exportacao em PNG): pilha invertida com a legenda a direita/esquerda/canto,
' ordem das series em cima/embaixo.
Public Sub Test_WaterfallMulti_Legend_ReHidesAfterPositionChange()
    Dim lp As Variant, wb As Workbook, d As clsChartData, cht As Chart, pal As Variant, fwd As String, rev As String, want As String
    pal = Style_Palette("padrao")
    fwd = CLng(pal(0)) & ";" & CLng(pal(1)) & ";"        ' Loja; Online (ordem das series)
    rev = CLng(pal(1)) & ";" & CLng(pal(0)) & ";"        ' Online; Loja (pilha invertida)
    For Each lp In Array(xlLegendPositionBottom, xlLegendPositionTop, xlLegendPositionLeft, xlLegendPositionCorner, xlLegendPositionRight)
        want = IIf(lp = xlLegendPositionBottom Or lp = xlLegendPositionTop, fwd, rev)
        Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
        Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
        cht.Legend.Position = lp
        AssertEqual 2, cht.Legend.LegendEntries.Count, "mudar a posicao mantem as auxiliares fora (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem das entradas (" & lp & ")"
        cht.HasLegend = False: cht.HasLegend = True: cht.Legend.Position = lp
        AssertEqual 13, cht.Legend.LegendEntries.Count, "legenda religada volta com todas as entradas (" & lp & ")"
        Waterfall_SetLabels cht, True                    ' reaplica o esconder
        AssertEqual 2, cht.Legend.LegendEntries.Count, "reescondida (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem apos reesconder (" & lp & ")"
        Chart_HideHelperLegendEntries cht
        AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem apos a segunda chamada (" & lp & ")"
        CloseTemp wb
    Next
End Sub

' Alternar rotulos e mudar paleta/formato pela API: a legenda continua so com as series de dados, sem rotulo "#N/D".
Public Sub Test_WaterfallMulti_CoreToggles_LabelsConnectorsPaletteFormat()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, pal As Variant, l1 As Series, lt As Series, i As Variant, sep As String
    Set wb = NewTempWorkbook(): Set ws = MultiFixture(wb).Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    Core_SetToggle "labels", False
    AssertEqual 10, cht.SeriesCollection.Count: AssertEqual "0", Core_GetOption("labels")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "rotulos desligados: legenda intacta"
    Core_SetToggle "labels", True
    AssertEqual 13, cht.SeriesCollection.Count: AssertEqual "1", Core_GetOption("labels")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "rotulos religados: auxiliares novas fora da legenda"
    Core_SetToggle "connectors", False
    AssertEqual 0, ConnectorCount(cht)
    Core_SetToggle "connectors", True
    AssertEqual 4, ConnectorCount(cht)
    ' formato dec1: segmentos e totais acompanham; linhas fora de total ficam sem texto (nunca "#N/D")
    sep = Application.International(xlDecimalSeparator)
    Core_SetOption "numfmt", "dec1"
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set lt = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual "100" & sep & "0", l1.Points(1).DataLabel.Text
    AssertEqual "-30" & sep & "0", l1.Points(2).DataLabel.Text
    AssertEqual "130" & sep & "0", lt.Points(3).DataLabel.Text
    AssertEqual "-90" & sep & "0", lt.Points(5).DataLabel.Text
    For Each i In Array(1, 2, 4)
        AssertEqual "", lt.Points(CLng(i)).DataLabel.Text, "linha " & i & " nao e total: rotulo de total vazio"
    Next
    ' paleta
    Core_SetOption "palette", "cinza"
    pal = Style_Palette("cinza")
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "Online").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "tqq_dnn_1").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(3)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB
    AssertEqual 2, cht.Legend.LegendEntries.Count
    ' a fonte dos rotulos de segmento acompanha a cor atual da serie (paleta do tema: as cores vem do documento)
    Core_SetOption "palette", "tema"
    AssertEqual Style_ContrastFont(CLng(Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB)), CLng(Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Font.Color)
    AssertEqual Style_ContrastFont(CLng(Chart_FindSeries(cht, "Online").Format.Fill.ForeColor.RGB)), CLng(Chart_FindSeries(cht, "tqq_lbl_2").Points(1).DataLabel.Font.Color)
    AssertEqual 2, cht.Legend.LegendEntries.Count
    Watchers_Stop
    CloseTemp wb
End Sub

' As auxiliares de total tem pontos NA() (linhas que nao sao total): ler e ligar rotulos neles nao falha, e o texto do
' rotulo nesses pontos e vazio (a coluna Total vale 0 e o formato esconde o zero).
Public Sub Test_WaterfallMulti_TotalHelper_NaPoints_ReadableAndBlank()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lt As Series, i As Long, v As Variant
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    Set lt = Chart_FindSeries(cht, "tqq_lbl")
    v = lt.Values
    ' Series.Values devolve Empty para um ponto #N/D (medido no Excel 2019)
    AssertTrue IsEmpty(v(1)) And IsEmpty(v(2)) And IsEmpty(v(4)), "pontos de linhas sem total nao sao plotados (#N/D)"
    AssertFalse IsEmpty(v(3)) Or IsEmpty(v(5))
    For i = 1 To 5
        AssertTrue lt.Points(i).HasDataLabel, "ler HasDataLabel num ponto NA nao falha (" & i & ")"
        AssertTrue InStr(lt.Points(i).DataLabel.Formula, CALC_SHEET) > 0, "ler Formula num ponto NA nao falha (" & i & ")"
    Next
    AssertEqual "", lt.Points(1).DataLabel.Text: AssertEqual "", lt.Points(2).DataLabel.Text: AssertEqual "", lt.Points(4).DataLabel.Text
    CloseTemp wb
End Sub

' Limpeza (Tarefa 0) com pontos #N/D: serie que NAO cita a planilha interna (so os rotulos ligados citam o bloco) e tem
' pontos NA(): a leitura de HasDataLabel/DataLabel.Formula desses pontos nao e falha, o bloco citado pelos rotulos
' sobrevive e so o orfao e removido.
Public Sub Test_Cleanup_NaPointLabels_AreNotReadFailure()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, keep As Range, i As Long, nm As Name, alive As Boolean
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:A4").Value = Application.Transpose(Array("a", "b", "c", "d"))
    ws.Range("B1").Value = 10: ws.Range("B2").Formula = "=NA()": ws.Range("B3").Value = 30: ws.Range("B4").Formula = "=NA()"
    Set keep = CalcSheet_AllocateBlock(wb, "keep01", "stacked", "x", 4, 3)
    CalcSheet_AllocateBlock wb, "orph01", "stacked", "x", 4, 3
    For i = 1 To 4: keep.Cells(i + 1, 1).Value = i * 10: Next
    Set co = ws.ChartObjects.Add(10, 10, 300, 200)
    Do While co.Chart.SeriesCollection.Count > 0: co.Chart.SeriesCollection(1).Delete: Loop
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Values = "=Dados!$B$1:$B$4": s.XValues = "=Dados!$A$1:$A$4"
    co.Chart.ChartType = xlLine
    s.HasDataLabels = True
    For i = 1 To 4
        s.Points(i).HasDataLabel = True
        s.Points(i).DataLabel.Formula = "='" & CALC_SHEET & "'!" & keep.Cells(i + 1, 1).Address(True, True)
    Next
    AssertTrue InStr(s.Formula, CALC_SHEET) = 0, "premissa: a serie nao cita a planilha interna"
    AssertEqual 1, CalcSheet_Cleanup(wb), "so o orfao e removido"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA() com rotulo nao sao falha de leitura"
    For Each nm In wb.Names
        If nm.Name = "tqq_keep01" Then alive = True
        AssertTrue nm.Name <> "tqq_orph01", "orfao removido"
    Next
    AssertTrue alive, "o bloco citado pelos rotulos ligados sobrevive"
    CloseTemp wb
End Sub

' Alterar o dado: tudo acompanha (blocos, rotulos, posicao do rotulo de total e conectores no liquido).
Public Sub Test_WaterfallMulti_DataChange_LabelsAndConnectorsFollow()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, ax As Axis, it As Double, ih As Double, v() As Double
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Waterfall_Insert(ws, d).Chart
    ws.Range("B5").Value = -100             ' Queda: Loja -100, Online -20 => liquido -120 => 130 -> 10
    Core_SafeRefresh cht
    v = Waterfall_EndValues(cht)
    AssertAlmost 10, v(4), 0.0001: AssertAlmost 10, v(5), 0.0001
    AssertEqual "-100", Chart_FindSeries(cht, "tqq_lbl_1").Points(4).DataLabel.Text
    AssertEqual "10", Chart_FindSeries(cht, "tqq_lbl").Points(5).DataLabel.Text
    AssertEqual xlLabelPositionAbove, Chart_FindSeries(cht, "tqq_lbl").Points(5).DataLabel.Position, "o total agora e positivo"
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight: Set ax = cht.Axes(xlValue)
    AssertAlmost it + ih * (ax.MaximumScale - 10) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_4").Top, 1, "conector no novo liquido"
    CloseTemp wb
End Sub

' Limite: 6 series (a paleta tem 6 cores) e a cor de cada serie e de suas partes.
Public Sub Test_WaterfallMulti_SixSeries_Boundary()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, pal As Variant, j As Long, blk As Range, ends() As Double, want As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Seis"
    ws.Range("A1:G1").Value = Array("Etapa", "S1", "S2", "S3", "S4", "S5", "S6")
    ws.Range("A2:G2").Value = Array("x", 1, 2, 3, 4, 5, 6)
    ws.Range("A3:G3").Value = Array("y", -1, 1, -1, 1, -1, 1)
    ws.Range("A4:B4").Value = Array("z", "e")
    Set d = ChartData_FromRange(ws.Range("A1:G4"))
    Set cht = Waterfall_Insert(ws, d).Chart
    Set blk = Registry_Block(cht)
    AssertEqual 80, blk.Columns.Count: AssertEqual 6, Waterfall_SeriesCount(cht)
    AssertEqual 33, cht.SeriesCollection.Count, "base + 24 partes + total + 6 rotulos de segmento + rotulo de total"
    pal = Style_Palette("padrao")
    want = Array(0, 1, 2, 4, 5, 0)          ' as series pulam pal(3) (cor dos totais): ciclo 0, 1, 2, 4, 5 (Tarefa 11)
    For j = 1 To 6
        AssertEqual CLng(pal(want(j - 1))), Chart_FindSeries(cht, "S" & j).Format.Fill.ForeColor.RGB, "serie " & j
        AssertEqual CLng(pal(want(j - 1))), Chart_FindSeries(cht, "tqq_upn_" & j).Format.Fill.ForeColor.RGB
        AssertEqual CLng(pal(want(j - 1))), Chart_FindSeries(cht, "tqq_dn_" & j).Format.Fill.ForeColor.RGB
    Next
    AssertEqual 6, cht.Legend.LegendEntries.Count
    ' x soma 21; y soma 0 (liquido); z e total: 21
    ends = Waterfall_EndValues(cht)
    AssertAlmost 21, ends(1), 0.0001: AssertAlmost 21, ends(2), 0.0001: AssertAlmost 21, ends(3), 0.0001
    CloseTemp wb
End Sub

' Build-down com varias series: o 'e' na primeira linha (em qualquer coluna) e numeros na ultima.
Public Sub Test_WaterfallMulti_BuildDown()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, body As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:C3").Value = Array("Custos", -700, -100)
    ws.Range("A4:C4").Value = Array("Despesas", -250, 0)
    ws.Range("A5:C5").Value = Array("Lucro", 250, 50)
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfmbd1", d))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual 1350, body.Cells(2, WF_COL_END).Value, "300 - (-700 - 100 - 250)"
    AssertEqual 1350, body.Cells(2, WF_COL_TOTAL).Value: AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual 550, body.Cells(3, WF_COL_END).Value: AssertEqual 300, body.Cells(4, WF_COL_END).Value
    AssertEqual 700, body.Cells(3, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 100, body.Cells(3, Waterfall_SeriesCol(2, "DN")).Value
    AssertEqual 550, body.Cells(3, WF_COL_BASE).Value
    AssertEqual True, body.Cells(5, 3).Value: AssertEqual 300, body.Cells(5, WF_COL_TOTAL).Value: AssertEqual 300, body.Cells(5, WF_COL_END).Value
    AssertEqual 0, body.Cells(5, Waterfall_SeriesCol(1, "UP")).Value, "a ultima linha vira total: sem segmentos"
    ws.Range("B5").Value = 400
    AssertEqual 1500, body.Cells(2, WF_COL_END).Value, "recalcula ao mudar o total final"
    ' 'e' na ultima linha tambem: nada de build-down (como na Fase 1)
    ws.Range("B5").Value = "e": ws.Range("C5").ClearContents
    AssertEqual False, body.Cells(2, WF_COL_BD).Value: AssertEqual 0, body.Cells(2, WF_COL_END).Value
    ' erro na ultima linha: nao conta como numero
    ws.Range("B5").Formula = "=NA()": ws.Range("C5").ClearContents
    AssertEqual False, body.Cells(2, WF_COL_BD).Value
    CloseTemp wb
End Sub

' Tabelas de 1 e de 2 categorias: sem referencia circular (BD usa as celulas de origem, nao IsTotal) e sem erro.
Public Sub Test_Waterfall_TinyTables_NoErrorsNoCircularity()
    Dim wb As Workbook, ws As Worksheet, body As Range, c As Long, d As clsChartData, k As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Mini"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("So", 5, 3)
    For k = 1 To 2                                     ' 1 = so a Loja (S = 1), 2 = Loja e Online
        Set d = ChartData_FromRange(ws.Range(ws.Cells(1, 1), ws.Cells(2, 1 + k)))
        Set body = Block_Body(Waterfall_BuildBlock(wb, "wft0" & k, d))
        For c = 1 To body.Columns.Count
            If Not (k = 2 And (c = WF_COL_LABELVAL Or c = WF_COL_LABELPOS)) Then      ' com varias series, rotulo de total fora das linhas de total e NA() de proposito
                AssertFalse IsError(body.Cells(2, c).Value), "n = 1, S = " & k & ", coluna " & c
            End If
        Next
        AssertEqual False, body.Cells(2, WF_COL_BD).Value
        AssertEqual 5 + IIf(k = 2, 3, 0), body.Cells(2, WF_COL_END).Value
    Next
    ws.Range("A2:C2").Value = Array("So", "e", "e")
    AssertEqual True, body.Cells(2, 3).Value: AssertEqual 0, body.Cells(2, WF_COL_END).Value
    AssertEqual False, body.Cells(2, WF_COL_BD).Value, "uma unica linha nunca e build-down"
    ws.Range("A2:C2").Value = Array("Um", "e", "e")
    ws.Range("A3:C3").Value = Array("Dois", 7, 2)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wft03", d))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual 9, body.Cells(2, WF_COL_END).Value, "n = 2: o total da primeira linha e o numero da ultima"
    AssertEqual True, body.Cells(3, 3).Value: AssertEqual 9, body.Cells(3, WF_COL_END).Value: AssertEqual 9, body.Cells(3, WF_COL_TOTAL).Value
    CloseTemp wb
End Sub

' Rotulos de segmento: fonte clara sobre cor escura e escura sobre cor clara (a 3.a cor da paleta padrao e azul claro).
Public Sub Test_WaterfallMulti_SegmentLabelFont_ContrastsWithSeriesColor()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Tres"
    ws.Range("A1:D1").Value = Array("Etapa", "A", "B", "C")
    ws.Range("A2:D2").Value = Array("x", 10, 20, 30)
    ws.Range("A3:D3").Value = Array("y", -1, -2, -3)
    Set d = ChartData_FromRange(ws.Range("A1:D3"))
    Set cht = Waterfall_Insert(ws, d).Chart
    AssertEqual RGB(255, 255, 255), CLng(Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Font.Color), "azul escuro: fonte branca"
    AssertEqual RGB(64, 64, 64), CLng(Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Font.Color), "azul claro: fonte cinza-escuro"
    CloseTemp wb
End Sub

' ---------- Tarefa 5: cascata horizontal (barwaterfall) ----------

Public Sub Test_BarWaterfall_Insert_SeriesConnectorsAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, iw As Double, c As Shape
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual xlSecondary, lbl.AxisGroup
    AssertEqual "100", lbl.Points(1).DataLabel.Text: AssertEqual xlLabelPositionRight, lbl.Points(1).DataLabel.Position
    AssertEqual "-150", lbl.Points(2).DataLabel.Text: AssertEqual xlLabelPositionLeft, lbl.Points(2).DataLabel.Position
    AssertEqual "-20", lbl.Points(4).DataLabel.Text, "subtotal"
    ' conectores verticais em x = End_i (regra de tres independente): End = 100, -50, -20, -20, 50
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: Set ax = cht.Axes(xlValue, xlPrimary)
    Set c = cht.Shapes("tqq_ann_conn_1")
    AssertAlmost 0, c.Width, 0.01, "conector vertical"
    AssertAlmost il + iw * (100 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), c.Left, 1
    AssertAlmost il + iw * (-50 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_2").Left, 1
    AssertTrue c.Top > cht.PlotArea.InsideTop And c.Top + c.Height < cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight / 3, "entre as barras 1 e 2"
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    CloseTemp wb
End Sub

Public Sub Test_BarWaterfall_LabelHelper_AlignedWithBarEnds()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, it As Double, iw As Double, ih As Double
    Dim x As Double, dl As DataLabel
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue, xlPrimary)
    x = il + iw * (100 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' Hi da barra 1
    Set dl = lbl.Points(1).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 14, "rotulo 1 no fim da barra: " & dl.Left & " vs " & x
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + ih / 12)) <= ih / 12, "rotulo 1 na primeira faixa (6 categorias)"
    x = il + iw * (-50 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' Lo da barra 2 (rotulo a esquerda)
    Set dl = lbl.Points(2).DataLabel
    AssertTrue dl.Left + dl.Width <= x + 2 And dl.Left + dl.Width >= x - 14, "rotulo 2 a esquerda do Lo: " & (dl.Left + dl.Width) & " vs " & x
    CloseTemp wb
End Sub

Public Sub Test_BarWaterfall_Multi_And_Core()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual 13, cht.SeriesCollection.Count
    AssertEqual "-30", Chart_FindSeries(cht, "tqq_lbl_1").Points(2).DataLabel.Text
    AssertTrue Core_GetToggle("connectors")
    Core_SetToggle "connectors", False
    AssertEqual 0, cht.Shapes.Count
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "Loja").Format.Fill.ForeColor.RGB
    Core_AddAnnotation "diff", 1, 5
    AssertEqual "-240 (-160%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    Watchers_Stop
    CloseTemp wb
End Sub

' x (pt) de um valor no eixo horizontal: regra de tres sobre a area interna e a escala reais (independente de Geom_*).
Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

' Centro vertical da faixa i, lido da barra (tqq_base) e nao calculado: Point.Top/Height sao sempre frescos.
Private Function YCenterOf(cht As Chart, i As Long) As Double
    With Chart_FindSeries(cht, "tqq_base").Points(i)
        YCenterOf = .Top + .Height / 2
    End With
End Function

' Barra horizontal: os conectores sao verticais em x = End_i, da borda inferior da barra i ao topo da barra i + 1 (geometria
' lida dos pontos nativos das barras, nao das funcoes de producao).
Public Sub Test_BarWaterfall_Connectors_FollowBarEdges()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ends As Variant, i As Long, c As Shape, bs As Series
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    ends = Array(100, -50, -20, -20, 50)
    Set bs = Chart_FindSeries(cht, "tqq_base")
    For i = 1 To 5
        Set c = cht.Shapes("tqq_ann_conn_" & i)
        AssertAlmost 0, c.Width, 0.01, "conector " & i & " vertical"
        AssertAlmost XOf(cht, CDbl(ends(i - 1))), c.Left, 1, "conector " & i & " em x = End"
        AssertAlmost bs.Points(i).Top + bs.Points(i).Height, c.Top, 1.5, "conector " & i & " parte da borda inferior da barra " & i
        AssertAlmost bs.Points(i + 1).Top, c.Top + c.Height, 1.5, "conector " & i & " termina no topo da barra " & (i + 1)
    Next
    AssertEqual 5, ConnectorCount(cht)
    CloseTemp wb
End Sub

' Rotulos (auxiliar XY): o ponto i coincide com o fim da barra (Hi, direita; ou Lo, esquerda) na faixa i, e o rotulo sai do
' lado do sinal. Posicoes esperadas escritas a mao: LabelPos = 100, -50, -20, -20, 50, 50; sinais + - + - + +.
Public Sub Test_BarWaterfall_LabelHelper_PointGeometryAndSides()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, i As Long, xs As Variant, ps As Variant, txt As Variant
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    xs = Array(100, -50, -20, -20, 50, 50)
    ps = Array(xlLabelPositionRight, xlLabelPositionLeft, xlLabelPositionRight, xlLabelPositionLeft, xlLabelPositionRight, xlLabelPositionRight)
    txt = Array("100", "-150", "30", "-20", "70", "50")
    AssertEqual 6, lbl.Points.Count
    For i = 1 To 6
        AssertAlmost XOf(cht, CDbl(xs(i - 1))), lbl.Points(i).Left, 1.5, "ponto " & i & " no fim da barra (x)"
        AssertAlmost YCenterOf(cht, i), lbl.Points(i).Top, 1.5, "ponto " & i & " no centro da faixa (y)"
        AssertEqual CLng(ps(i - 1)), lbl.Points(i).DataLabel.Position, "lado do rotulo " & i
        AssertEqual CStr(txt(i - 1)), lbl.Points(i).DataLabel.Text, "texto do rotulo " & i
    Next
    CloseTemp wb
End Sub

' Segmentos visiveis lidos dos pontos nativos das barras x regra de tres sobre os valores esperados a mao:
' Inicio sobe 0..100; A desce 100 -> -50 (parte positiva 0..100, negativa -50..0); B sobe -50..-20; Fim e total 0..50.
Public Sub Test_BarWaterfall_SegmentGeometry_RuleOfThree()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    With Chart_FindSeries(cht, "tqq_uppos").Points(1)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 100) - XOf(cht, 0), .Width, 1, "Inicio 0..100"
    End With
    With Chart_FindSeries(cht, "tqq_downpos").Points(2)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 100) - XOf(cht, 0), .Width, 1, "A: parte positiva 0..100"
    End With
    With Chart_FindSeries(cht, "tqq_downneg").Points(2)
        AssertAlmost XOf(cht, -50), .Left, 1: AssertAlmost XOf(cht, 0) - XOf(cht, -50), .Width, 1, "A: parte negativa -50..0"
    End With
    With Chart_FindSeries(cht, "tqq_upneg").Points(3)
        AssertAlmost XOf(cht, -50), .Left, 1: AssertAlmost XOf(cht, -20) - XOf(cht, -50), .Width, 1, "B: -50..-20"
    End With
    With Chart_FindSeries(cht, "tqq_total").Points(6)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 50) - XOf(cht, 0), .Width, 1, "Fim: total 0..50"
    End With
    With Chart_FindSeries(cht, "tqq_total").Points(4)
        AssertAlmost XOf(cht, -20), .Left, 1: AssertAlmost XOf(cht, 0) - XOf(cht, -20), .Width, 1, "Subtotal: -20..0"
    End With
    ' a primeira categoria no topo (eixo invertido): a barra 1 acima da barra 6
    AssertTrue Chart_FindSeries(cht, "tqq_base").Points(1).Top < Chart_FindSeries(cht, "tqq_base").Points(6).Top, "categoria 1 no topo"
    CloseTemp wb
End Sub

' Setas em cascata horizontal: CAGR a direita das pontas (o mesmo afastamento do fim da barra nas duas pontas, depois dos rotulos de ponta:
' Tarefa 15b), ponta na categoria 'to'; diferenca abaixo da ultima barra com a linha de nivel vertical saindo do fim da barra inicial.
' Posicoes lidas dos pontos nativos das barras e dos rotulos renderizados.
Public Sub Test_BarWaterfall_CagrAndDiff_Geometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, lbl As Shape, lvl As Shape, bs As Series
    Dim x1 As Double, x6 As Double, yBottom As Double, tip As DataLabel
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    Set ln = Ann_AddCagr(cht, 1, 6)
    Set bs = Chart_FindSeries(cht, "tqq_base")
    ForceLayout cht
    AssertAlmost XOf(cht, 100) - XOf(cht, 50), ln.Width, 1.5, "origem e ponta com o mesmo afastamento do fim da barra (End 100 e 50)"
    AssertTrue ln.Left >= XOf(cht, 50) + 10 - 0.5, "ponta (categoria 6) a pelo menos 10 pt do fim da barra"
    Set tip = Chart_FindSeries(cht, "tqq_lbl").Points(6).DataLabel
    AssertTrue ln.Left >= tip.Left + tip.Width + 2, "ponta depois do rotulo da ponta da categoria 6 (" & ln.Left & " x " & (tip.Left + tip.Width) & ")"
    x1 = ln.Left + ln.Width: x6 = ln.Left
    AssertAlmost YCenterOf(cht, 1), ln.Top, 1.5, "a seta parte do centro da faixa 1"
    AssertAlmost YCenterOf(cht, 6), ln.Top + ln.Height, 1.5, "e termina no centro da faixa 6"
    AssertTrue ln.HorizontalFlip = msoTrue And ln.VerticalFlip = msoFalse, "a ponta fica na categoria 6 (x menor, y maior)"
    Set lbl = cht.Shapes(ln.Name & "_lbl")
    AssertEqual "CAGR " & Format$((50 / 100) ^ (1 / 5) - 1, "+0.0%;-0.0%;0.0%"), lbl.TextFrame2.TextRange.Text
    AssertTrue lbl.Left >= x1, "rotulo do CAGR a direita da seta"
    AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width + 0.5, "rotulo dentro do grafico"
    Set ln = Ann_AddDifference(cht, 1, 6)
    Set bs = Chart_FindSeries(cht, "tqq_base")
    x1 = XOf(cht, 100): x6 = XOf(cht, 50)
    yBottom = bs.Points(6).Top + bs.Points(6).Height
    AssertAlmost 0, ln.Height, 0.01, "seta de diferenca horizontal"
    AssertAlmost yBottom + 10, ln.Top, 1.5, "10 pt abaixo da ultima barra"
    AssertAlmost x6, ln.Left, 1.5: AssertAlmost x1 - x6, ln.Width, 1.5, "de End 1 a End 6"
    Set lvl = cht.Shapes(ln.Name & "_lvl")
    AssertAlmost 0, lvl.Width, 0.01, "linha de nivel vertical"
    AssertAlmost x1, lvl.Left, 1.5, "no fim da barra inicial"
    AssertAlmost bs.Points(1).Top + bs.Points(1).Height, lvl.Top, 1.5, "saindo da borda inferior da barra 1"
    AssertAlmost ln.Top + 4, lvl.Top + lvl.Height, 1.5, "ate um pouco abaixo da seta"
    Set lbl = cht.Shapes(ln.Name & "_lbl")
    AssertEqual Format$(-50, "+#,##0;-#,##0;0") & " (" & Format$(-0.5, "+0%;-0%;0%") & ")", lbl.TextFrame2.TextRange.Text
    AssertTrue lbl.Top >= ln.Top And lbl.Top + lbl.Height <= cht.ChartArea.Height + 0.5, "rotulo abaixo da seta e dentro do grafico"
    ' a area de plotagem cedeu espaco: a seta de CAGR nao saiu do grafico
    AssertTrue cht.Shapes("tqq_ann_cagr_1").Left + cht.Shapes("tqq_ann_cagr_1").Width < cht.ChartArea.Width, "seta de CAGR dentro do grafico"
    CloseTemp wb
End Sub

' Mudar dados e tamanho: o ponto XY, o lado do rotulo e os conectores acompanham (leitura dos pontos nativos, sempre frescos).
' A passa de -150 a +40 (rotulo de Left para Right); C passa de 70 a -300 (End: -130; o total Fim vira negativo: Right -> Left).
Public Sub Test_BarWaterfall_DataChangeAndResize_LabelsAndConnectorsFollow()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, co As ChartObject, cht As Chart, lbl As Series, i As Long
    Dim xs As Variant, ps As Variant, ends As Variant
    Set wb = NewTempWorkbook(): Set d = SingleFixture(wb): Set ws = d.Source.Worksheet
    Set co = Waterfall_Insert(ws, d, "barwaterfall"): Set cht = co.Chart
    ws.Range("B3").Value = 40: ws.Range("B6").Value = -300
    co.Width = 640: co.Height = 360
    Core_SafeRefresh cht
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    xs = Array(100, 140, 170, 170, -130, -130)          ' LabelPos: Hi (>= 0) ou Lo (< 0) de cada barra
    ps = Array(xlLabelPositionRight, xlLabelPositionRight, xlLabelPositionRight, xlLabelPositionRight, xlLabelPositionLeft, xlLabelPositionLeft)
    ends = Array(100, 140, 170, 170, -130)
    For i = 1 To 6
        AssertAlmost XOf(cht, CDbl(xs(i - 1))), lbl.Points(i).Left, 1.5, "ponto " & i & " acompanha a nova escala"
        AssertAlmost YCenterOf(cht, i), lbl.Points(i).Top, 1.5, "ponto " & i & " na faixa"
        AssertEqual CLng(ps(i - 1)), lbl.Points(i).DataLabel.Position, "lado do rotulo " & i
    Next
    For i = 1 To 5
        AssertAlmost XOf(cht, CDbl(ends(i - 1))), cht.Shapes("tqq_ann_conn_" & i).Left, 1.5, "conector " & i & " no novo End"
    Next
    AssertEqual "40", lbl.Points(2).DataLabel.Text: AssertEqual "-300", lbl.Points(5).DataLabel.Text: AssertEqual "-130", lbl.Points(6).DataLabel.Text
    AssertAlmost 640, cht.ChartArea.Width, 6, "o grafico foi redimensionado"
    CloseTemp wb
End Sub

' Varias series na horizontal: ordem das series (a mesma da vertical), auxiliares XY no grupo secundario, pontos do centro do
' segmento coincidindo com o centro das barras nativas (inclusive o segmento que cruza o zero) e rotulos de total por sinal.
Public Sub Test_BarWaterfallMulti_Structure_AndGeometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, names As String, s As Series, l1 As Series, l2 As Series, lt As Series
    Dim lo As Double, hi As Double, v As Variant
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual 2, Waterfall_SeriesCount(cht)
    For Each s In cht.SeriesCollection: names = names & s.Name & ";": Next
    AssertEqual "tqq_base;tqq_dn_2;tqq_dn_1;Loja;Online;tqq_upn_2;tqq_upn_1;tqq_dnn_1;tqq_dnn_2;tqq_total;tqq_lbl_1;tqq_lbl_2;tqq_lbl;", names
    AssertEqual xlBarStacked, Chart_FindSeries(cht, "Loja").ChartType
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2"): Set lt = Chart_FindSeries(cht, "tqq_lbl")
    For Each v In Array("tqq_lbl_1", "tqq_lbl_2", "tqq_lbl")
        Set s = Chart_FindSeries(cht, CStr(v))
        AssertEqual xlSecondary, s.AxisGroup, s.Name & " no grupo secundario"
        AssertTrue IsScatterType(s.ChartType), s.Name & " e XY (tipo " & s.ChartType & ")"
    Next
    ' centro do segmento Loja da linha 1 (0..100) e Online (100..150): ponto XY = centro da barra nativa
    With Chart_FindSeries(cht, "Loja").Points(1)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 100) - XOf(cht, 0), .Width, 1, "Loja Inicio 0..100"
        AssertAlmost .Left + .Width / 2, l1.Points(1).Left, 1.5, "rotulo da Loja no centro do segmento"
    End With
    With Chart_FindSeries(cht, "Online").Points(1)
        AssertAlmost XOf(cht, 100), .Left, 1: AssertAlmost XOf(cht, 150) - XOf(cht, 100), .Width, 1, "Online Inicio 100..150"
        AssertAlmost .Left + .Width / 2, l2.Points(1).Left, 1.5, "rotulo do Online no centro do segmento"
    End With
    ' Queda: Loja -200 cruza o zero (de 130 a -70): partes dn_1 (0..130) e dnn_1 (-70..0); o rotulo fica no centro (30)
    lo = Chart_FindSeries(cht, "tqq_dnn_1").Points(4).Left
    hi = Chart_FindSeries(cht, "tqq_dn_1").Points(4).Left + Chart_FindSeries(cht, "tqq_dn_1").Points(4).Width
    AssertAlmost XOf(cht, -70), lo, 1, "borda esquerda do segmento que cruza o zero"
    AssertAlmost XOf(cht, 130), hi, 1, "borda direita do segmento que cruza o zero"
    AssertAlmost (lo + hi) / 2, l1.Points(4).Left, 1.5, "rotulo no centro do segmento que cruza o zero"
    AssertAlmost XOf(cht, 30), l1.Points(4).Left, 1.5
    AssertAlmost YCenterOf(cht, 4), l1.Points(4).Top, 1.5, "na faixa 4"
    AssertEqual "-200", l1.Points(4).DataLabel.Text: AssertEqual xlLabelPositionCenter, l1.Points(4).DataLabel.Position
    ' totais: 130 (Subtotal, a direita) e -90 (Fim, a esquerda), no fim da barra de total
    AssertEqual "130", lt.Points(3).DataLabel.Text: AssertEqual xlLabelPositionRight, lt.Points(3).DataLabel.Position
    AssertAlmost XOf(cht, 130), lt.Points(3).Left, 1.5: AssertAlmost YCenterOf(cht, 3), lt.Points(3).Top, 1.5
    AssertEqual "-90", lt.Points(5).DataLabel.Text: AssertEqual xlLabelPositionLeft, lt.Points(5).DataLabel.Position
    AssertAlmost XOf(cht, -90), lt.Points(5).Left, 1.5
    AssertEqual "", lt.Points(1).DataLabel.Text, "linha sem total: rotulo de total vazio"
    ' conectores no liquido: End = 150, 130, 130, -90, -90
    AssertAlmost XOf(cht, 150), cht.Shapes("tqq_ann_conn_1").Left, 1.5
    AssertAlmost XOf(cht, 130), cht.Shapes("tqq_ann_conn_2").Left, 1.5, "liquido -30 + 10"
    AssertAlmost XOf(cht, -90), cht.Shapes("tqq_ann_conn_4").Left, 1.5
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Loja e Online"
    AssertEqual "right", Registry_GetOption(cht, "legend")
    ' limpeza (Tarefa 0): pontos NA() nos X das auxiliares XY nao sao falha de leitura e o bloco citado sobrevive
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "so o orfao e removido"
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertEqual "130", lt.Points(3).DataLabel.Text, "o bloco da cascata sobreviveu"
    CloseTemp wb
End Sub

' Legenda das barras (carry-over d): medido no Excel (PNG e chaves das entradas) em 5 posicoes, a lista sai em ordem normal
' (Loja, Online) em TODAS, diferente das pilhas de colunas (invertida a direita/esquerda/canto); religar a legenda e os
' rotulos volta a esconder so as auxiliares, e a chamada repetida e idempotente.
Public Sub Test_BarWaterfallMulti_Legend_OrderAtAllPositions()
    Dim lp As Variant, wb As Workbook, d As clsChartData, cht As Chart, pal As Variant, want As String
    pal = Style_Palette("padrao")
    want = CLng(pal(0)) & ";" & CLng(pal(1)) & ";"        ' Loja; Online
    For Each lp In Array(xlLegendPositionBottom, xlLegendPositionTop, xlLegendPositionLeft, xlLegendPositionCorner, xlLegendPositionRight)
        Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
        Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
        AssertEqual want, LegendKeys(cht), "ordem inicial"
        cht.Legend.Position = lp
        AssertEqual 2, cht.Legend.LegendEntries.Count, "mudar a posicao mantem as auxiliares fora (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem das entradas (" & lp & ")"
        cht.HasLegend = False: cht.HasLegend = True: cht.Legend.Position = lp
        AssertEqual 13, cht.Legend.LegendEntries.Count, "legenda religada volta com todas as entradas (" & lp & ")"
        Waterfall_SetLabels cht, True
        AssertEqual 2, cht.Legend.LegendEntries.Count, "reescondida (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem apos reesconder (" & lp & ")"
        Chart_HideHelperLegendEntries cht
        AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente (" & lp & ")"
        AssertEqual want, LegendKeys(cht), "ordem apos a segunda chamada (" & lp & ")"
        CloseTemp wb
    Next
End Sub

' "De cima para baixo" na horizontal: o primeiro total vem do numero da ultima linha, com 1 e com 2 series. Segmentos lidos
' dos pontos nativos: Receita 0..1200 (1 serie); com 2 series Custos empilha Loja 650..1350 (mais perto do inicio) e Online
' 550..650, e o total da ultima linha 0..300.
Public Sub Test_BarWaterfall_BuildDown_Horizontal_SingleAndMulti()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:B3").Value = Array("Custos", -700)
    ws.Range("A4:B4").Value = Array("Despesas", -250)
    ws.Range("A5:B5").Value = Array("Lucro", 250)
    Set d = ChartData_FromRange(ws.Range("A1:B5"))
    Set cht = Waterfall_Insert(ws, d, "barwaterfall").Chart
    With Chart_FindSeries(cht, "tqq_total").Points(1)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 1200) - XOf(cht, 0), .Width, 1, "Receita = 250 - (-700 - 250) = 1200"
    End With
    AssertEqual "1" & Application.International(xlThousandsSeparator) & "200", Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    AssertEqual xlLabelPositionRight, Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Position
    AssertAlmost XOf(cht, 1200), Chart_FindSeries(cht, "tqq_lbl").Points(1).Left, 1.5
    AssertAlmost XOf(cht, 250), Chart_FindSeries(cht, "tqq_lbl").Points(4).Left, 1.5, "a ultima linha vira total"
    AssertAlmost XOf(cht, 1200), cht.Shapes("tqq_ann_conn_1").Left, 1.5: AssertAlmost XOf(cht, 500), cht.Shapes("tqq_ann_conn_2").Left, 1.5
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:C3").Value = Array("Custos", -700, -100)
    ws.Range("A4:C4").Value = Array("Despesas", -250, 0)
    ws.Range("A5:C5").Value = Array("Lucro", 250, 50)
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Waterfall_Insert(ws, d, "barwaterfall").Chart
    With Chart_FindSeries(cht, "tqq_total").Points(1)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 1350) - XOf(cht, 0), .Width, 1, "Receita = 300 - (-700 - 100 - 250)"
    End With
    With Chart_FindSeries(cht, "tqq_dn_1").Points(2)
        AssertAlmost XOf(cht, 650), .Left, 1: AssertAlmost XOf(cht, 1350) - XOf(cht, 650), .Width, 1, "Loja -700: 1350 -> 650"
    End With
    With Chart_FindSeries(cht, "tqq_dn_2").Points(2)
        AssertAlmost XOf(cht, 550), .Left, 1: AssertAlmost XOf(cht, 650) - XOf(cht, 550), .Width, 1, "Online -100: 650 -> 550"
    End With
    With Chart_FindSeries(cht, "tqq_total").Points(4)
        AssertAlmost XOf(cht, 0), .Left, 1: AssertAlmost XOf(cht, 300) - XOf(cht, 0), .Width, 1, "Lucro total 300"
    End With
    AssertEqual "-700", Chart_FindSeries(cht, "tqq_lbl_1").Points(2).DataLabel.Text
    AssertAlmost (XOf(cht, 650) + XOf(cht, 1350)) / 2, Chart_FindSeries(cht, "tqq_lbl_1").Points(2).Left, 1.5, "rotulo no centro do segmento"
    ' o numero da ultima linha muda: o total inicial acompanha
    ws.Range("B5").Value = 400
    Core_SafeRefresh cht
    AssertAlmost XOf(cht, 1500), Chart_FindSeries(cht, "tqq_lbl").Points(1).Left, 1.5, "build-down recalculado (400 + 50 + 1050)"
    CloseTemp wb
End Sub

' Rollback: falha depois de alocar o bloco desfaz grafico e bloco; tipo que nao e cascata e recusado antes de alocar.
Public Sub Test_BarWaterfall_Insert_RollsBack_AndRefusesOtherKinds()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, n As Long, nm As Name, k As Variant
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Waterfall_Insert ws, d, "barwaterfall"
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    For Each k In Array("stacked", "xyz", "")
        n = 0
        On Error Resume Next
        Waterfall_Insert ws, d, CStr(k)
        n = Err.Number
        On Error GoTo 0
        AssertEqual 5, n, "tipo '" & k & "' nao e cascata"
        AssertEqual 0, ws.ChartObjects.Count
    Next
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nenhum bloco alocado " & nm.Name
    Next
    CloseTemp wb
End Sub

' API do core na cascata horizontal de uma serie: rotulos (a auxiliar XY some e volta alinhada), conectores, formato, paleta; o
' toggle de totais nao se aplica a cascata; legenda inexistente.
Public Sub Test_BarWaterfall_CoreToggles_SingleSeries()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, lbl As Series, sep As String, pal As Variant, n As Long
    Set wb = NewTempWorkbook(): Set ws = SingleFixture(wb).Source.Worksheet
    ws.Range("A1:B7").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual 7, cht.SeriesCollection.Count, "6 series da cascata + rotulo"
    AssertFalse cht.HasLegend: AssertFalse Core_Supports("legend")
    AssertTrue Core_Supports("connectors") And Core_Supports("labels")
    n = cht.SeriesCollection.Count
    Core_SetToggle "totals", True
    AssertEqual n, cht.SeriesCollection.Count, "totais nao se aplicam a cascata"
    Core_SetToggle "labels", False
    AssertEqual 6, cht.SeriesCollection.Count: AssertEqual "0", Core_GetOption("labels")
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    Core_SetToggle "labels", True
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual 7, cht.SeriesCollection.Count: AssertEqual "1", Core_GetOption("labels")
    AssertEqual xlSecondary, lbl.AxisGroup
    AssertAlmost XOf(cht, 100), lbl.Points(1).Left, 1.5, "rotulo religado alinhado ao fim da barra"
    AssertAlmost XOf(cht, -50), lbl.Points(2).Left, 1.5
    AssertEqual xlLabelPositionRight, lbl.Points(2).DataLabel.Position, "-150 com o ponto perto do minimo do eixo: do outro lado (Tarefa 15b)"
    Core_SetToggle "connectors", False
    AssertEqual 0, ConnectorCount(cht): AssertEqual "0", Core_GetOption("connectors")
    Core_SetToggle "connectors", True
    AssertEqual 5, ConnectorCount(cht)
    sep = Application.International(xlDecimalSeparator)
    Core_SetOption "numfmt", "dec1"
    AssertEqual "100" & sep & "0", lbl.Points(1).DataLabel.Text
    AssertEqual "-150" & sep & "0", lbl.Points(2).DataLabel.Text
    AssertEqual xlLabelPositionRight, lbl.Points(2).DataLabel.Position, "rotulo mais largo, mesmo lado (Tarefa 15b)"
    Core_SetOption "palette", "cinza"
    pal = Style_Palette("cinza")
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "tqq_uppos").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(2)), Chart_FindSeries(cht, "tqq_downpos").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB
    Core_SetOption "gap", "20"
    AssertEqual 20, cht.ChartGroups(1).GapWidth
    Watchers_Stop
    CloseTemp wb
End Sub

' "Alterar dados" na cascata horizontal: o tipo e as opcoes se mantem, o grafico novo nasce horizontal (nao de colunas), com os
' conectores e os rotulos do novo numero de categorias; Ann_Refresh refaz os conectores que faltam (familia da cascata).
Public Sub Test_BarWaterfall_RebuildFromRange_KeepsKindAndConnectors()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = SingleFixture(wb).Source.Worksheet
    ws.Range("A1:B7").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    Core_SetOption "numfmt", "dec1"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B6"))
    Set cht = Core_ActiveChart()
    AssertEqual "barwaterfall", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType, "grafico novo em barras"
    AssertEqual "dec1", Registry_GetOption(cht, "numfmt")
    AssertEqual "5", Registry_GetOption(cht, "n")
    AssertEqual 4, ConnectorCount(cht), "5 categorias: 4 conectores"
    AssertEqual 5, Chart_FindSeries(cht, "tqq_lbl").Points.Count
    AssertAlmost XOf(cht, 100), Chart_FindSeries(cht, "tqq_lbl").Points(1).Left, 1.5
    AssertFalse Core_ChartSourceMismatch(cht)
    cht.Shapes("tqq_ann_conn_2").Delete
    Ann_Refresh cht
    AssertEqual 4, ConnectorCount(cht), "a familia da cascata refaz os conectores"
    AssertAlmost XOf(cht, -50), cht.Shapes("tqq_ann_conn_2").Left, 1.5, "conector refeito no lugar certo (End 2 = -50)"
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function IsScatterType(ct As Long) As Boolean
    IsScatterType = (ct = xlXYScatter Or ct = xlXYScatterLines Or ct = xlXYScatterLinesNoMarkers Or ct = xlXYScatterSmooth Or ct = xlXYScatterSmoothNoMarkers)
End Function

' Setas na cascata horizontal de varias series com a legenda a direita: a area de plotagem cede espaco e o rotulo do CAGR
' fica antes da legenda (nada desenhado por cima dela).
Public Sub Test_BarWaterfallMulti_Cagr_StaysLeftOfLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set d = MultiFixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d, "barwaterfall").Chart
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set lbl = cht.Shapes(ln.Name & "_lbl")
    AssertTrue cht.Legend.Left > cht.ChartArea.Width / 2, "premissa: legenda a direita (esq=" & cht.Legend.Left & ")"
    AssertTrue lbl.Left + lbl.Width <= cht.Legend.Left + 0.5, "rotulo do CAGR antes da legenda (fim=" & (lbl.Left + lbl.Width) & ", legenda=" & cht.Legend.Left & ")"
    AssertTrue ln.Left + ln.Width <= cht.Legend.Left, "seta antes da legenda"
    AssertAlmost XOf(cht, 150) - XOf(cht, 130), ln.Width, 1.5, "origem (End 150) e ponta (End 130) com o mesmo afastamento"
    AssertTrue ln.Left >= XOf(cht, 130) + 10 - 0.5, "ponta a pelo menos 10 pt do End 130"
    ForceLayout cht
    With Chart_FindSeries(cht, "tqq_lbl").Points(3).DataLabel
        AssertTrue ln.Left >= .Left + .Width + 2, "ponta depois do rotulo do total 130 (Tarefa 15b): " & ln.Left & " x " & (.Left + .Width)
    End With
    CloseTemp wb
End Sub

' Tarefa 11 (ruling de cor da Tarefa 4): os totais usam SEMPRE pal(3) e as SERIES da cascata pulam o indice 3 (ciclo 0, 1, 2, 4, 5), para
' nenhuma serie ter a cor das barras de total (com 4 ou mais series a 4a caia em pal(3)). Com 6 series a 6a volta ao inicio do ciclo.
' Expectativas escritas a mao; com 1 serie as cores seguem as da Fase 1 (alta pal(1), queda pal(2), total pal(0)): ver os outros testes.
Public Sub Test_WaterfallMulti_SeriesColors_SkipTheTotalColor()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, pal As Variant, want As Variant, j As Long, p As Variant, pn As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:G1").Value = Array("Etapa", "S1", "S2", "S3", "S4", "S5", "S6")
    ws.Range("A2:G2").Value = Array("Inicio", 10, 20, 30, 40, 50, 60)
    ws.Range("A3:G3").Value = Array("Ajuste", 1, -2, 3, -4, 5, -6)
    ws.Range("A4:G4").Value = Array("Fim", "e", "e", "e", "e", "e", "e")
    ws.Range("A1:G4").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual 6, Waterfall_SeriesCount(cht)
    want = Array(0, 1, 2, 4, 5, 0)                       ' indice da paleta de S1..S6
    For Each pn In Array("padrao", "cinza")
        If pn = "cinza" Then Core_SetOption "palette", "cinza"
        pal = Style_Palette(CStr(pn))
        AssertEqual CLng(pal(3)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB, pn & ": totais em pal(3)"
        For j = 1 To 6
            AssertEqual CLng(pal(want(j - 1))), Chart_FindSeries(cht, "S" & j).Format.Fill.ForeColor.RGB, pn & ": serie S" & j
            If j <= 5 Then AssertTrue Chart_FindSeries(cht, "S" & j).Format.Fill.ForeColor.RGB <> Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB, pn & ": S" & j & " nao tem a cor dos totais"
            For Each p In Array("tqq_dn_", "tqq_upn_", "tqq_dnn_")
                AssertEqual CLng(pal(want(j - 1))), Chart_FindSeries(cht, p & j).Format.Fill.ForeColor.RGB, pn & ": " & p & j & " com a cor da serie"
            Next
        Next
    Next
    Teardown wb
End Sub

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Tarefa 15b (deferred da T4): com varias series a linha e de total quando QUALQUER celula dela e um marcador (SUMPRODUCT/MATCH sobre a lista
' WF_MARKERS): cada marcador (com maiusculas e espacos) numa coluna so, o marcador ao lado de um numero e de um erro, e os nao marcadores.
Public Sub Test_WaterfallMulti_MarkerList_AnyCellOfTheRow()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, body As Range, m As Variant, i As Long, n As Long, want As Boolean
    m = Array("e", "t", "=", "total", "subtotal", "E", "TOTAL", " Subtotal ", "x", "tt", "")
    n = UBound(m) + 1
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("B2:C" & n + 3).NumberFormat = "@"
    For i = 0 To n - 1
        ws.Cells(i + 2, 1).Value = "L" & (i + 1)
        If i Mod 2 = 0 Then ws.Cells(i + 2, 3).Value = CStr(m(i)) Else ws.Cells(i + 2, 2).Value = CStr(m(i))     ' o marcador ora na 1a, ora na 2a serie
    Next
    ws.Cells(n + 2, 1).Value = "Numero e marcador": ws.Cells(n + 2, 2).NumberFormat = "General": ws.Cells(n + 2, 2).Value = 10: ws.Cells(n + 2, 3).Value = "t"
    ws.Cells(n + 3, 1).Value = "Erro e marcador": ws.Cells(n + 3, 2).NumberFormat = "General": ws.Cells(n + 3, 2).Formula = "=NA()": ws.Cells(n + 3, 3).Value = "e"
    Set d = ChartData_FromRange(ws.Range("A1:C" & n + 3))
    AssertEqual 2, d.SeriesCount
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfmk01", d))
    For i = 0 To n - 1
        want = Waterfall_IsTotalMarker(CStr(m(i)))
        AssertEqual want, CBool(body.Cells(i + 2, 3).Value), "linha " & (i + 1) & " marcador '" & m(i) & "'"
    Next
    AssertEqual True, CBool(body.Cells(n + 2, 3).Value), "numero na 1a serie e 't' na 2a: linha de total"
    AssertEqual True, CBool(body.Cells(n + 3, 3).Value), "#N/D na 1a serie nao propaga e 'e' na 2a marca o total"
    AssertFalse IsError(body.Cells(n + 3, WF_COL_END).Value), "sem erro no End da linha com #N/D"
    CloseTemp wb
End Sub

' Tarefa 15b (deferred da T4): build-down ('e' na primeira linha e numero na ultima) com um subtotal 'e' no MEIO, com 1 e com 2 series: o
' primeiro total e calculado para tras pelos deltas das linhas do meio (o subtotal tem delta 0), o subtotal mostra o acumulado e a ultima linha
' vira total.
Public Sub Test_Waterfall_BuildDown_WithSubtotalInTheMiddle()
    Dim wb As Workbook, ws As Worksheet, body As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", "e")
    ws.Range("A3:B3").Value = Array("Custos", -700)
    ws.Range("A4:B4").Value = Array("Subtotal", "e")
    ws.Range("A5:B5").Value = Array("Despesas", -250)
    ws.Range("A6:B6").Value = Array("Lucro", 250)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfbm01", ChartData_FromRange(ws.Range("A1:B6"))))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value, "build-down"
    AssertEqual 1200, body.Cells(2, WF_COL_END).Value, "Receita = 250 - (-700 + 0 - 250)"
    AssertEqual 1200, body.Cells(2, WF_COL_TOTAL).Value
    AssertEqual 500, body.Cells(3, WF_COL_END).Value
    AssertEqual True, body.Cells(4, 3).Value, "subtotal do meio": AssertEqual 0, body.Cells(4, 4).Value, "subtotal sem delta"
    AssertEqual 500, body.Cells(4, WF_COL_END).Value: AssertEqual 500, body.Cells(4, WF_COL_TOTAL).Value: AssertEqual 500, body.Cells(4, WF_COL_LABELVAL).Value
    AssertEqual 250, body.Cells(5, WF_COL_END).Value
    AssertEqual True, body.Cells(6, 3).Value, "ultima linha vira total": AssertEqual 250, body.Cells(6, WF_COL_TOTAL).Value
    ws.Range("B6").Value = 300
    AssertEqual 1250, body.Cells(2, WF_COL_END).Value, "recalcula ao mudar o total final": AssertEqual 550, body.Cells(4, WF_COL_TOTAL).Value
    ' duas series
    ws.Range("D1:F1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("D2:E2").Value = Array("Receita", "e")
    ws.Range("D3:F3").Value = Array("Custos", -700, -100)
    ws.Range("D4:F4").Value = Array("Subtotal", "e", "e")
    ws.Range("D5:F5").Value = Array("Despesas", -250, 0)
    ws.Range("D6:F6").Value = Array("Lucro", 250, 50)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wfbm02", ChartData_FromRange(ws.Range("D1:F6"))))
    AssertEqual True, body.Cells(2, WF_COL_BD).Value
    AssertEqual 1350, body.Cells(2, WF_COL_END).Value, "300 - (-700 - 100 + 0 - 250)"
    AssertEqual 550, body.Cells(3, WF_COL_END).Value
    AssertEqual True, body.Cells(4, 3).Value: AssertEqual 550, body.Cells(4, WF_COL_TOTAL).Value
    AssertEqual 0, body.Cells(4, Waterfall_SeriesCol(1, "DN")).Value: AssertEqual 0, body.Cells(4, Waterfall_SeriesCol(2, "UP")).Value, "subtotal sem segmentos"
    AssertEqual 300, body.Cells(5, WF_COL_END).Value
    AssertEqual True, body.Cells(6, 3).Value: AssertEqual 300, body.Cells(6, WF_COL_TOTAL).Value
    CloseTemp wb
End Sub

' Tarefa 15b (A4): cascata horizontal. (1) Um rotulo "a esquerda" (barra que desce / total negativo) cujo ponto fica perto do minimo do eixo nao
' cabe entre o ponto e a borda esquerda da area interna (onde comecam os nomes das categorias): ele passa para o outro lado do ponto (sobre a
' propria barra, com fundo branco); os que cabem continuam a esquerda. (2) A seta de CAGR corre a direita dos rotulos de ponta (nao os cruza): a
' folga a direita conta a largura do maior rotulo. (3) A seta de diferenca (abaixo da ultima barra) nao cruza nenhum rotulo. Geometria dos
' rotulos medida depois de renderizar (ForceLayout), nunca pelas funcoes de producao.
Public Sub Test_BarWaterfall_LeftLabelNearAxisMin_And_ArrowsClearTipLabels()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, lbl As Series, i As Long, dl As DataLabel, il As Double, ln As Shape, xa As Double, ya As Double
    Dim y1 As Double, y6 As Double, x1 As Double, x6 As Double
    Set wb = NewTempWorkbook(): Set ws = SingleFixture(wb).Source.Worksheet
    ws.Range("A1:B7").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    ForceLayout cht
    il = cht.PlotArea.InsideLeft
    For i = 1 To 6
        Set dl = lbl.Points(i).DataLabel
        AssertTrue dl.Left >= il - 0.5, "rotulo " & i & " (" & dl.Text & ") nao invade os nomes das categorias: " & dl.Left & " < " & il
        AssertTrue dl.Left >= 0 And dl.Left + dl.Width <= cht.ChartArea.Width + 0.5, "rotulo " & i & " dentro do grafico"
    Next
    AssertEqual xlLabelPositionRight, lbl.Points(2).DataLabel.Position, "-150 (Lo = -50, perto do minimo): do outro lado do ponto"
    AssertTrue lbl.Points(2).DataLabel.Left >= XOf(cht, -50) - 0.5, "rotulo 2 a direita do Lo"
    AssertEqual msoTrue, lbl.Points(2).DataLabel.Format.Fill.Visible, "sobre a barra: fundo branco"
    AssertEqual RGB(255, 255, 255), lbl.Points(2).DataLabel.Format.Fill.ForeColor.RGB
    AssertEqual xlLabelPositionLeft, lbl.Points(4).DataLabel.Position, "-20 tem espaco: continua a esquerda"
    AssertTrue lbl.Points(4).DataLabel.Format.Fill.Visible <> msoTrue, "a esquerda: sem fundo"
    ' dados mudam e o ponto se afasta do minimo: o rotulo volta para a esquerda, sem o fundo
    ws.Range("B2").Value = 400
    Core_SafeRefresh cht
    AssertEqual xlLabelPositionLeft, lbl.Points(2).DataLabel.Position, "com espaco o rotulo volta a esquerda"
    AssertTrue lbl.Points(2).DataLabel.Format.Fill.Visible <> msoTrue, "e perde o fundo"
    ws.Range("B2").Value = 100
    Core_SafeRefresh cht
    ' (2) CAGR 1 -> 6: a seta (das pontas, End 100 e 50) passa a direita do rotulo da ponta das categorias 1 e 6
    Core_AddAnnotation "cagr", 1, 6
    ForceLayout cht
    Set ln = cht.Shapes("tqq_ann_cagr_1")
    y1 = YCenterOf(cht, 1): y6 = YCenterOf(cht, 6)
    If ln.HorizontalFlip = msoTrue Then x1 = ln.Left + ln.Width: x6 = ln.Left Else x1 = ln.Left: x6 = ln.Left + ln.Width
    Set dl = lbl.Points(1).DataLabel
    AssertTrue x1 >= dl.Left + dl.Width + 2, "origem da seta (" & x1 & ") depois do rotulo da categoria 1 (termina em " & (dl.Left + dl.Width) & ")"
    Set dl = lbl.Points(6).DataLabel
    AssertTrue x6 >= dl.Left + dl.Width + 2, "ponta da seta (" & x6 & ") depois do rotulo da categoria 6 (termina em " & (dl.Left + dl.Width) & ")"
    AssertTrue cht.Shapes("tqq_ann_cagr_1_lbl").Left + cht.Shapes("tqq_ann_cagr_1_lbl").Width <= cht.ChartArea.Width + 0.5, "rotulo do CAGR dentro do grafico"
    ' (3) diferenca 1 -> 6: a seta horizontal abaixo da ultima barra nao atravessa nenhum rotulo de ponta
    Core_AddAnnotation "diff", 1, 6
    ForceLayout cht
    Set ln = cht.Shapes("tqq_ann_diff_1")
    ya = ln.Top
    For i = 1 To 6
        Set dl = lbl.Points(i).DataLabel
        If Len(dl.Text) > 0 Then
            AssertFalse (ya > dl.Top And ya < dl.Top + dl.Height And ln.Left < dl.Left + dl.Width And ln.Left + ln.Width > dl.Left), "seta de diferenca cruza o rotulo " & i
        End If
    Next
    Watchers_Stop
    CloseTemp wb
End Sub

' Tarefa 15b (B3): cascata com varias series. O rotulo de um segmento mais fino que o texto (a altura do segmento, pela regra de tres no eixo e
' na area interna REAIS, menor que ~12 pt nas colunas; nas barras, mais estreito que o texto) encostaria nas bordas: fica oculto; quando o
' segmento passa a caber (dados mudaram) o rotulo volta, ligado a celula. Os demais segmentos continuam com rotulo.
Public Sub Test_WaterfallMulti_ThinSegmentLabel_HiddenUntilItFits()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, ax As Axis, k As Double, l1 As Series, l2 As Series, kind As Variant, tag As String
    For Each kind In Array("waterfall", "barwaterfall")
        tag = CStr(kind) & ": "
        Set wb = NewTempWorkbook(): Set ws = MultiFixture(wb).Source.Worksheet
        ws.Range("A1:C6").Select
        Core_Insert CStr(kind)
        Set cht = Core_ActiveChart()
        Set ax = cht.Axes(xlValue, xlPrimary)
        If kind = "waterfall" Then k = cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale) Else k = cht.PlotArea.InsideWidth / (ax.MaximumScale - ax.MinimumScale)
        Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2")
        AssertTrue 10 * k < 12, tag & "pre-condicao: o segmento Online do Ajuste (10) tem " & (10 * k) & " pt"
        AssertFalse l2.Points(2).HasDataLabel, tag & "segmento fino (Online, Ajuste = 10): sem rotulo"
        AssertTrue l1.Points(2).HasDataLabel, tag & "segmento largo (Loja, Ajuste = -30) mantem o rotulo"
        AssertEqual "-30", l1.Points(2).DataLabel.Text
        AssertTrue l2.Points(1).HasDataLabel And l2.Points(4).HasDataLabel, tag & "os outros segmentos do Online mantem o rotulo"
        ws.Range("C3").Value = 60                                   ' o segmento passa a caber
        Core_SafeRefresh cht
        AssertTrue l2.Points(2).HasDataLabel, tag & "segmento que passou a caber: rotulo de volta"
        AssertEqual "60", l2.Points(2).DataLabel.Text, tag & "ligado a celula"
        AssertTrue Left$(l2.Points(2).DataLabel.Formula, 1) = "=" And InStr(1, l2.Points(2).DataLabel.Formula, CALC_SHEET, vbTextCompare) > 0, tag & "religado como ligacao ao bloco (nao texto fixo): " & l2.Points(2).DataLabel.Formula
        AssertEqual xlLabelPositionCenter, l2.Points(2).DataLabel.Position
        ws.Range("C3").Value = 10
        Core_SafeRefresh cht
        AssertFalse l2.Points(2).HasDataLabel, tag & "fino de novo: oculto"
        Core_SetToggle "labels", False
        Core_SetToggle "labels", True
        Set l2 = Chart_FindSeries(cht, "tqq_lbl_2")
        AssertFalse l2.Points(2).HasDataLabel, tag & "religar os rotulos mantem o segmento fino sem rotulo"
        Watchers_Stop
        CloseTemp wb
    Next
End Sub

' Revisao da 15b (I-1): o ThinQQ so religa os rotulos de segmento que ELE escondeu (registrados na opcao "thinlbl", pares serie.ponto); um rotulo
' que o usuario apagou continua apagado depois de atualizar. A opcao vive na meta do bloco: sobrevive a mover o grafico, a uma copia dele e ao
' Atualizar; religar os rotulos (Rotulos desligado e ligado) recomeca a lista.
Public Sub Test_WaterfallMulti_UserDeletedSegmentLabel_StaysDeleted_ThinListPersists()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, co As ChartObject, l1 As Series, l2 As Series, dup As Shape
    Set wb = NewTempWorkbook(): Set ws = MultiFixture(wb).Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart(): Set co = cht.Parent
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2")
    AssertFalse l2.Points(2).HasDataLabel, "pre-condicao: o segmento fino (Online, Ajuste) escondido pelo ThinQQ"
    AssertEqual "2.2", Registry_GetOption(cht, "thinlbl", ""), "o par escondido fica registrado na opcao"
    l1.Points(1).DataLabel.Delete                                  ' o usuario apaga o rotulo de um segmento largo (Loja, Inicio = 100)
    AssertFalse l1.Points(1).HasDataLabel, "pre-condicao: rotulo apagado"
    Core_SafeRefresh cht
    ws.Calculate
    Core_Refresh
    AssertFalse l1.Points(1).HasDataLabel, "rotulo apagado pelo usuario continua apagado depois de atualizar"
    ws.Range("C3").Value = 60                                      ' o segmento escondido pelo ThinQQ passa a caber
    Core_SafeRefresh cht
    AssertTrue l2.Points(2).HasDataLabel, "o ThinQQ religa so o que ele escondeu"
    AssertEqual "", Registry_GetOption(cht, "thinlbl", ""), "religado: sai da opcao"
    AssertFalse l1.Points(1).HasDataLabel, "o apagado pelo usuario continua apagado"
    ' persistencia: mover o grafico, copiar o grafico e Atualizar
    ws.Range("C3").Value = 10
    Core_SafeRefresh cht
    AssertEqual "2.2", Registry_GetOption(cht, "thinlbl", "")
    co.Left = co.Left + 30: co.Top = co.Top + 20
    Set dup = co.Duplicate
    co.Select
    Core_Refresh
    AssertEqual "2.2", Registry_GetOption(cht, "thinlbl", ""), "a lista sobrevive a mover, copiar o grafico e Atualizar"
    AssertFalse l2.Points(2).HasDataLabel
    ws.Range("C3").Value = 60
    Core_SafeRefresh cht
    AssertTrue l2.Points(2).HasDataLabel, "depois de mover/copiar o par registrado ainda e religado"
    AssertFalse l1.Points(1).HasDataLabel
    ' religar os rotulos recomeca: todos ligados e o fino escondido de novo
    ws.Range("C3").Value = 10
    co.Select
    Core_SetToggle "labels", False
    AssertEqual "", Registry_GetOption(cht, "thinlbl", ""), "desligar os rotulos limpa a lista"
    Core_SetToggle "labels", True
    Set l1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set l2 = Chart_FindSeries(cht, "tqq_lbl_2")
    AssertTrue l1.Points(1).HasDataLabel, "religar os rotulos recria tambem o que o usuario apagou"
    AssertFalse l2.Points(2).HasDataLabel
    AssertEqual "2.2", Registry_GetOption(cht, "thinlbl", "")
    Watchers_Stop
    CloseTemp wb
End Sub

' Revisao da 15b (menor 1): o fundo dos rotulos de ponta da cascata horizontal. O ThinQQ so liga o branco no rotulo que ele levou para dentro da
' barra e so desliga o branco que ele mesmo pos; um preenchimento escolhido pelo usuario (num rotulo de fora ou num de dentro) fica.
Public Sub Test_BarWaterfall_TipLabelBackground_KeepsUserFill()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, lbl As Series
    Set wb = NewTempWorkbook(): Set ws = SingleFixture(wb).Source.Worksheet
    ws.Range("A1:B7").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual xlLabelPositionRight, lbl.Points(2).DataLabel.Position, "pre-condicao: -150 levado para dentro"
    AssertEqual msoTrue, lbl.Points(2).DataLabel.Format.Fill.Visible, "pre-condicao: com o branco do ThinQQ"
    With lbl.Points(1).DataLabel.Format.Fill                     ' o usuario pinta o rotulo "100" (fora da barra) de amarelo
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
    End With
    With lbl.Points(2).DataLabel.Format.Fill                     ' e o "-150" (dentro da barra) de verde-claro
        .Visible = msoTrue
        .ForeColor.RGB = RGB(200, 255, 200)
    End With
    Core_SafeRefresh cht
    Core_Refresh
    AssertEqual msoTrue, lbl.Points(1).DataLabel.Format.Fill.Visible, "o preenchimento do usuario num rotulo de fora fica"
    AssertEqual RGB(255, 255, 0), lbl.Points(1).DataLabel.Format.Fill.ForeColor.RGB
    AssertEqual RGB(200, 255, 200), lbl.Points(2).DataLabel.Format.Fill.ForeColor.RGB, "o preenchimento do usuario num rotulo de dentro nao vira branco"
    ' o branco do proprio ThinQQ ainda sai quando o rotulo volta para a esquerda
    lbl.Points(2).DataLabel.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ws.Range("B2").Value = 400
    Core_SafeRefresh cht
    AssertEqual xlLabelPositionLeft, lbl.Points(2).DataLabel.Position
    AssertTrue lbl.Points(2).DataLabel.Format.Fill.Visible <> msoTrue, "o branco do ThinQQ sai"
    Watchers_Stop
    CloseTemp wb
End Sub

' Revisao da 15b (A4 com varias series): na cascata horizontal de varias series a auxiliar dos totais (tqq_lbl) tem pontos #N/D (linhas que nao
' sao total); medido: Series.HasDataLabels continua True (os pontos #N/D tem rotulo vazio), entao a saida antecipada de Labels_RefreshPositions nao
' pula a verificacao do lado. Um total negativo cujo rotulo nao cabe antes dos nomes das categorias vai para dentro da barra, com fundo branco.
Public Sub Test_BarWaterfallMulti_NegativeTotalNearAxisMin_GoesInside()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, lt As Series, dl As DataLabel
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("Inicio", 100000, 20000)
    ws.Range("A3:C3").Value = Array("Queda", -179000, -15000)
    ws.Range("A4:B4").Value = Array("Fim", "e")
    ws.Range("A1:C4").Select
    Core_Insert "barwaterfall"
    Set cht = Core_ActiveChart()
    Set lt = Chart_FindSeries(cht, "tqq_lbl")
    AssertTrue lt.HasDataLabels, "medido: HasDataLabels da auxiliar de totais continua True com pontos #N/D"
    ForceLayout cht
    Set dl = lt.Points(3).DataLabel
    AssertEqual "-74" & Application.International(xlThousandsSeparator) & "000", dl.Text
    AssertEqual xlLabelPositionRight, dl.Position, "total negativo perto do minimo do eixo: para dentro da barra"
    AssertTrue dl.Left >= cht.PlotArea.InsideLeft - 0.5, "nao invade os nomes das categorias: " & dl.Left & " x " & cht.PlotArea.InsideLeft
    AssertEqual msoTrue, dl.Format.Fill.Visible, "com fundo branco"
    Watchers_Stop
    CloseTemp wb
End Sub

' Onda final (I-1): a lista "thinlbl" e transacional. (c) uma falha numa serie posterior nao perde o registro do que ja foi escondido na passada;
' (a) um ponto cuja leitura falha fica como esta (nao herda o estado do ponto anterior); (b) um par so sai da lista quando religar deu certo.
' Falhas injetadas pelo gancho de teste Util_FailThinSyncForTest.
Private Function ThinHas(cht As Chart, key As String) As Boolean
    ThinHas = (InStr("," & Registry_GetOption(cht, "thinlbl", "") & ",", "," & key & ",") > 0)
End Function

Public Sub Test_WaterfallMulti_ThinList_IsTransactional()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, l2 As Series, en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:D1").Value = Array("Etapa", "Loja", "Online", "Extra")
    ws.Range("A2:D2").Value = Array("Inicio", 100, 50, 20)
    ws.Range("A3:D3").Value = Array("Ajuste", -30, 10, 5)
    ws.Range("A4:D4").Value = Array("Subtotal", "e", "e", "e")
    ws.Range("A5:D5").Value = Array("Queda", -200, -20, 4)
    ws.Range("A6:B6").Value = Array("Fim", "e")
    ws.Range("A1:D6").Select
    On Error GoTo Cleanup
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertTrue ThinHas(cht, "2.2") And ThinHas(cht, "3.2"), "pre-condicao: finos registrados (" & Registry_GetOption(cht, "thinlbl", "") & ")"
    ' (c) falha ao comecar a serie 3: o 2.2 escondido antes da falha fica registrado
    Core_SetToggle "labels", False
    Util_FailThinSyncForTest = "series:3"
    Core_SetToggle "labels", True
    Util_FailThinSyncForTest = ""
    Set l2 = Chart_FindSeries(cht, "tqq_lbl_2")
    AssertFalse l2.Points(2).HasDataLabel, "(c) 2.2 escondido antes da falha"
    AssertTrue ThinHas(cht, "2.2"), "(c) e registrado mesmo com a falha na serie 3: <" & Registry_GetOption(cht, "thinlbl", "") & ">"
    Core_SafeRefresh cht
    ' (a) leitura do ponto 2.2 falha: o rotulo (ja religado e visivel) fica como esta, sem herdar o estado do ponto 2.1
    ws.Range("C3").Value = 60
    Core_SafeRefresh cht
    AssertTrue l2.Points(2).HasDataLabel, "pre-condicao: 2.2 religado"
    AssertFalse ThinHas(cht, "2.2")
    Util_FailThinSyncForTest = "read:2.2"
    ws.Range("C3").Value = 10
    Core_SafeRefresh cht
    Util_FailThinSyncForTest = ""
    AssertTrue l2.Points(2).HasDataLabel, "(a) ponto ilegivel fica como esta"
    AssertFalse ThinHas(cht, "2.2"), "(a) e nao entra na lista"
    ' (b) religar falha: o par continua na lista e e religado na proxima atualizacao
    Core_SafeRefresh cht
    AssertFalse l2.Points(2).HasDataLabel, "pre-condicao: 2.2 escondido de novo"
    AssertTrue ThinHas(cht, "2.2")
    Util_FailThinSyncForTest = "link:2.2"
    ws.Range("C3").Value = 60
    Core_SafeRefresh cht
    Util_FailThinSyncForTest = ""
    AssertTrue ThinHas(cht, "2.2"), "(b) religar falhou: o par continua na lista"
    Core_SafeRefresh cht
    AssertTrue l2.Points(2).HasDataLabel, "(b) religado na atualizacao seguinte"
    AssertFalse ThinHas(cht, "2.2")
    Util_FailThinSyncForTest = ""
    Watchers_Stop
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Watchers_Stop
    CloseTemp wb
    Err.Raise en, es, ed
End Sub
