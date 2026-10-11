Attribute VB_Name = "TestBars"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Regiao", "Loja", "Online")
    ws.Range("A2:C2").Value = Array("Sul", 10, 5)
    ws.Range("A3:C3").Value = Array("Sudeste", 25, 10)
    ws.Range("A4:C4").Value = Array("Norte", 40, 20)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
End Function

Private Function IsScatter(ct As Long) As Boolean
    IsScatter = (ct = xlXYScatter Or ct = xlXYScatterLines Or ct = xlXYScatterLinesNoMarkers Or ct = xlXYScatterSmooth Or ct = xlXYScatterSmoothNoMarkers)
End Function

Public Sub Test_Bars_Insert_KindsAndAxes()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    AssertEqual "barstacked", Registry_ChartType(cht)
    AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder, "primeira categoria no topo"
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue Not tot Is Nothing
    AssertTrue IsScatter(tot.ChartType), "totais em XY (barras nao combinam com linha)"
    AssertEqual xlSecondary, tot.AxisGroup
    AssertEqual 0, cht.Axes(xlValue, xlSecondary).MinimumScale
    AssertEqual 3, cht.Axes(xlValue, xlSecondary).MaximumScale
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "tqq_total fora da legenda"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barclustered").Chart
    AssertEqual xlBarClustered, cht.SeriesCollection(1).ChartType
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    AssertEqual "80", Registry_GetOption(cht, "gap")
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    AssertEqual xlBarStacked100, cht.SeriesCollection(1).ChartType
    AssertEqual "67%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "segmento 100% mostra a participacao (10/15)"
    AssertEqual "15", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Stacked_SetConnectors cht, True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    CloseTemp wb
End Sub

Public Sub Test_Bars_Block_PosAndVLineColumns()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 5, blk.Columns.Count
    AssertEqual 2.5, body.Cells(2, Stacked_PosCol(blk)).Value: AssertEqual 0.5, body.Cells(4, Stacked_PosCol(blk)).Value
    AssertAlmost 36.666667, body.Cells(2, Stacked_VLineCol(blk)).Value, 0.0001, "media dos totais 15, 35, 60"
    AssertAlmost 36.666667, body.Cells(4, Stacked_VLineCol(blk)).Value, 0.0001
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 7, blk.Columns.Count, "3 + 2 participacoes + Pos + VLine"
    AssertTrue IsError(body.Cells(2, Stacked_VLineCol(blk)).Value), "100%: sem linha de valor"
    CloseTemp wb
End Sub

' Sonda e regra de ouro das barras: o rotulo do total i fica no fim da barra i (eixo de valores primario) e na faixa i.
Public Sub Test_Bars_TotalsHelper_AlignedWithBarsAndSlots()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series, ax As Axis, i As Long
    Dim il As Double, it As Double, iw As Double, ih As Double, xEnd As Double, yc As Double, lbl As DataLabel, totals As Variant
    totals = Array(15, 35, 60)
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    AssertTrue iw > 50 And ih > 50, "area interna real"
    Set ax = cht.Axes(xlValue, xlPrimary)
    For i = 1 To 3
        xEnd = il + iw * (totals(i - 1) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
        yc = it + (i - 0.5) * ih / 3
        Set lbl = tot.Points(i).DataLabel
        AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "rotulo " & i & " a direita do fim da barra: left=" & lbl.Left & " fim=" & xEnd
        AssertTrue Abs((lbl.Top + lbl.Height / 2) - yc) <= ih / 6, "rotulo " & i & " na faixa da categoria: centro=" & (lbl.Top + lbl.Height / 2) & " esperado=" & yc
    Next
    CloseTemp wb
End Sub

' Review Focus 1: depois de mudar dados (nova escala) e redimensionar, os totais continuam no fim das barras.
Public Sub Test_Bars_TotalsHelper_StaysAlignedAfterDataChangeAndResize()
    Dim wb As Workbook, d As clsChartData, co As ChartObject, cht As Chart, tot As Series, ax As Axis
    Dim il As Double, iw As Double, xEnd As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set co = Stacked_Insert(d.Source.Worksheet, d, "barstacked")
    Set cht = co.Chart
    d.Source.Worksheet.Range("C4").Value = 200           ' total Norte = 240: eixo rescala
    co.Width = 620
    Core_SafeRefresh cht
    Set tot = Chart_FindSeries(cht, "tqq_total")
    Set ax = cht.Axes(xlValue, xlPrimary)
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    xEnd = il + iw * (240 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set lbl = tot.Points(3).DataLabel
    AssertEqual "240", lbl.Text
    AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "total 3 segue o fim da barra apos mudanca: left=" & lbl.Left & " fim=" & xEnd
    xEnd = il + iw * (15 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set lbl = tot.Points(1).DataLabel
    AssertTrue lbl.Left >= xEnd - 2 And lbl.Left <= xEnd + 14, "total 1 segue a nova escala"
    CloseTemp wb
End Sub

Public Sub Test_Bars_TotalLabels_RightLeftBySign()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B2:C2").Value = Array(-10, -5)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual xlLabelPositionLeft, tot.Points(1).DataLabel.Position, "total negativo a esquerda"
    AssertEqual xlLabelPositionRight, tot.Points(2).DataLabel.Position
    d.Source.Worksheet.Range("B2:C2").Value = Array(10, 5)
    Labels_RefreshPositions cht
    AssertEqual xlLabelPositionRight, tot.Points(1).DataLabel.Position
    CloseTemp wb
End Sub

Public Sub Test_Bars_Annotations_UseTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Ann_AddCagr cht, 1, 3
    Ann_AddDifference cht, 1, 3
    AssertEqual "CAGR +100,0%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text     ' (60/15)^(1/2)-1
    AssertEqual "+45 (+300%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    AssertTrue cht.Shapes("tqq_ann_cagr_1").Left + cht.Shapes("tqq_ann_cagr_1").Width <= cht.ChartArea.Width, "CAGR dentro do grafico (folga a direita)"
    CloseTemp wb
End Sub

Public Sub Test_Bars_Insert_RollsBackOnFailure()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Stacked_Insert ws, d, "barclustered"
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

' ---------- extras da Tarefa 3 ----------

' Tres linhas com totais t1..t3 (duas series: ~2/3 e o resto), categorias C1..C3.
Private Function FixtureTotals(wb As Workbook, t1 As Double, t2 As Double, t3 As Double) As clsChartData
    Dim ws As Worksheet, v As Variant, i As Long, a As Double
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Regiao", "Loja", "Online")
    v = Array(t1, t2, t3)
    For i = 1 To 3
        a = Int(v(i - 1) * 2 / 3)
        ws.Cells(i + 1, 1).Value = "C" & i
        ws.Cells(i + 1, 2).Value = a
        ws.Cells(i + 1, 3).Value = v(i - 1) - a
    Next
    Set FixtureTotals = ChartData_FromRange(ws.Range("A1:C4"))
End Function

' Equivalente horizontal de Test_Ann_Cagr_ClearsTotalLabels_WhenTotalsOn (colunas): com os totais ligados, as pontas da
' seta de CAGR ficam depois do fim REAL do rotulo do total (Left + Width do DataLabel) de cada categoria, inclusive com
' totais de 4 e 7 digitos (o afastamento de 18 pt foi calibrado para a ALTURA do rotulo das colunas).
Public Sub Test_Bars_Cagr_ClearsTotalLabels_WhenTotalsOn()
    Dim sc As Variant, wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, tot As Series, lbl As Shape
    Dim xb As Double, xe As Double, rf As Double, rt As Double, tag As String
    For Each sc In Array(Array(15, 35, 60), Array(1500, 3500, 6000), Array(123456, 234567, 1234567))
        tag = "totais " & sc(0) & "/" & sc(1) & "/" & sc(2)
        Set wb = NewTempWorkbook(): Set d = FixtureTotals(wb, CDbl(sc(0)), CDbl(sc(1)), CDbl(sc(2)))
        Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
        Set ln = Ann_AddCagr(cht, 1, 3)
        ForceLayout cht
        Set tot = Chart_FindSeries(cht, "tqq_total")
        If ln.HorizontalFlip = msoTrue Then xb = ln.Left + ln.Width: xe = ln.Left Else xb = ln.Left: xe = ln.Left + ln.Width
        rf = tot.Points(1).DataLabel.Left + tot.Points(1).DataLabel.Width
        rt = tot.Points(3).DataLabel.Left + tot.Points(3).DataLabel.Width
        AssertTrue xb >= rf + 2, tag & ": inicio da seta depois do rotulo do total 1: seta=" & xb & " fim do rotulo=" & rf
        AssertTrue xe >= rt + 2, tag & ": ponta da seta depois do rotulo do total 3: seta=" & xe & " fim do rotulo=" & rt
        Set lbl = cht.Shapes(ln.Name & "_lbl")
        AssertTrue lbl.Left >= IIf(xb > xe, xb, xe), tag & ": rotulo do CAGR a direita da seta"
        AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width + 0.5, tag & ": rotulo do CAGR dentro do grafico: " & (lbl.Left + lbl.Width) & "/" & cht.ChartArea.Width
        CloseTemp wb
    Next
End Sub

' Total negativo: o rotulo fica a esquerda e termina junto ao fim da barra (extensao negativa da pilha).
Public Sub Test_Bars_Totals_NegativeBars_LabelLeftOfBarEnd()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series, ax As Axis, xEnd As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B2:C2").Value = Array(-10, -5)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    Set ax = cht.Axes(xlValue, xlPrimary)
    xEnd = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (-15 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set lbl = tot.Points(1).DataLabel
    AssertEqual "-15", lbl.Text
    AssertTrue lbl.Left + lbl.Width <= xEnd + 2 And lbl.Left + lbl.Width >= xEnd - 14, "rotulo negativo termina no fim da barra: fim do rotulo=" & (lbl.Left + lbl.Width) & " fim da barra=" & xEnd
    CloseTemp wb
End Sub

' Carry-over da Tarefa 1: a legenda de barras com a auxiliar XY secundaria volta a esconder a entrada auxiliar
' depois de mudar a posicao da legenda e de religar a legenda, em empilhadas e 100%.
Public Sub Test_Bars_Legend_ReHidesHelper_AfterPositionChange()
    Dim kind As Variant, lp As Variant, wb As Workbook, d As clsChartData, cht As Chart, k As Long, keys As String, pal As Variant
    pal = Style_Palette("padrao")
    For Each kind In Array("barstacked", "barstacked100")
        For Each lp In Array(xlLegendPositionBottom, xlLegendPositionTop, xlLegendPositionLeft, xlLegendPositionCorner, xlLegendPositionRight)
            Set wb = NewTempWorkbook(): Set d = Fixture(wb)
            Set cht = Stacked_Insert(d.Source.Worksheet, d, CStr(kind)).Chart
            cht.Legend.Position = lp
            AssertEqual 2, cht.Legend.LegendEntries.Count, kind & ": mudar a posicao mantem a auxiliar fora (" & lp & ")"
            cht.HasLegend = False: cht.HasLegend = True: cht.Legend.Position = lp
            AssertEqual 3, cht.Legend.LegendEntries.Count, kind & ": legenda religada volta com a auxiliar (" & lp & ")"
            Stacked_SetTotals cht, True                      ' reaplica o esconder
            AssertEqual 2, cht.Legend.LegendEntries.Count, kind & ": reescondida (" & lp & ")"
            keys = ";"
            For k = 1 To 2: keys = keys & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB & ";": Next
            AssertTrue InStr(keys, ";" & CLng(pal(0)) & ";") > 0 And InStr(keys, ";" & CLng(pal(1)) & ";") > 0, kind & ": entradas restantes = Loja e Online (" & lp & "): " & keys
            CloseTemp wb
        Next
    Next
End Sub

' Carry-over da Tarefa 1: em 100% so as colunas Part_j recebem o formato de participacao; TotalPos, Pos e VLine ficam
' como estao e o Total segue o formato absoluto.
Public Sub Test_Bars_ApplyCellFormat_OnlyPartColumnsIn100()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, n As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk): n = body.Rows.Count - 1
    Registry_SetOption cht, "numfmt", "pct1"
    Style_ApplyNumberFormat cht, Style_NumberFormat("pct1")
    AssertEqual Style_LocalizeCode(Style_NumberFormat("pct1")), body.Cells(2, 4).NumberFormatLocal, "Part1 em pct1"
    AssertEqual Style_LocalizeCode(Style_NumberFormat("pct1")), body.Cells(n + 1, 5).NumberFormatLocal, "Part2 em pct1"
    AssertEqual Style_LocalizeCode(Style_NumberFormat("int")), body.Cells(2, 2).NumberFormatLocal, "Total absoluto em 100%"
    AssertEqual "General", body.Cells(2, 3).NumberFormat, "TotalPos intacto"
    AssertEqual "General", body.Cells(2, Stacked_PosCol(blk)).NumberFormat, "Pos intacto"
    AssertEqual "General", body.Cells(2, Stacked_VLineCol(blk)).NumberFormat, "VLine intacto"
    AssertEqual "General", body.Cells(n + 1, Stacked_VLineCol(blk)).NumberFormat, "VLine intacto (ultima linha)"
    CloseTemp wb
    ' empilhada: Pos e VLine nao recebem o formato de total
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    Style_ApplyNumberFormat cht, Style_NumberFormat("dec1")
    AssertEqual Style_LocalizeCode(Style_NumberFormat("dec1")), body.Cells(2, 2).NumberFormatLocal, "Total em dec1"
    AssertEqual "General", body.Cells(2, Stacked_PosCol(blk)).NumberFormat, "Pos intacto"
    AssertEqual "General", body.Cells(2, Stacked_VLineCol(blk)).NumberFormat, "VLine intacto"
    CloseTemp wb
End Sub

' Bloco das colunas e da linha (Fase 1): as colunas 1..3 (e Part_j) sao as mesmas; Pos/VLine vem depois, em toda a familia.
Public Sub Test_Bars_Block_ColumnsAndLine_AlsoGetPosAndVLine()
    Dim kind As Variant, wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range
    For Each kind In Array("stacked", "clustered", "line")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Stacked_Insert(d.Source.Worksheet, d, CStr(kind)).Chart
        Set blk = Registry_Block(cht): Set body = Block_Body(blk)
        AssertEqual 5, blk.Columns.Count, CStr(kind)
        AssertEqual "Pos", body.Cells(1, Stacked_PosCol(blk)).Value: AssertEqual "VLine", body.Cells(1, Stacked_VLineCol(blk)).Value
        AssertEqual 2.5, body.Cells(2, 4).Value, kind & ": Pos da 1a categoria = n - 0,5"
        If kind = "stacked" Then
            AssertAlmost 36.666667, body.Cells(3, 5).Value, 0.0001, kind & ": media dos totais (15, 35, 60... = 110 / 3)"
            AssertTrue Left$(body.Cells(2, 5).Formula, 9) = "=AVERAGE(", kind & ": " & body.Cells(2, 5).Formula
        Else
            AssertAlmost 18.333333, body.Cells(3, 5).Value, 0.0001, kind & ": media de todos os valores plotados (110 / 6), nao a media das somas"
            AssertTrue InStr(body.Cells(2, 5).Formula, "AGGREGATE(1,6,") > 0, kind & ": " & body.Cells(2, 5).Formula
        End If
        AssertEqual "=2.5", body.Cells(2, 4).Formula, kind & ": Pos e formula em ingles"
        AssertTrue InStr(body.Cells(2, 5).Formula, "$") > 0, kind & ": referencias absolutas"
        CloseTemp wb
    Next
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    Set blk = Registry_Block(cht)
    AssertEqual 7, blk.Columns.Count, "stacked100: 3 + 2 participacoes + Pos + VLine"
    AssertEqual "Part2", Block_Body(blk).Cells(1, 5).Value
    CloseTemp wb
End Sub

' O gap segue o tipo (80 nas agrupadas, 50 nas demais).
Public Sub Test_Bars_Gap_ByKind()
    Dim kind As Variant, wb As Workbook, d As clsChartData, cht As Chart, want As Long
    For Each kind In Array("barstacked", "barclustered", "barstacked100", "clustered", "stacked")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Stacked_Insert(d.Source.Worksheet, d, CStr(kind)).Chart
        want = IIf(CStr(kind) = "barclustered" Or CStr(kind) = "clustered", 80, 50)
        AssertEqual want, cht.ChartGroups(1).GapWidth, CStr(kind)
        AssertEqual CStr(want), Registry_GetOption(cht, "gap"), CStr(kind)
        CloseTemp wb
    Next
End Sub

' Conectores (linhas de serie) existem nas barras empilhadas e 100%; nas agrupadas o pedido e recusado (opcao 0).
Public Sub Test_Bars_Connectors_StackedAndRefusedInClustered()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Stacked_SetConnectors cht, True
    AssertTrue cht.ChartGroups(1).HasSeriesLines, "barstacked"
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    Stacked_SetConnectors cht, False
    AssertFalse cht.ChartGroups(1).HasSeriesLines
    AssertEqual "0", Registry_GetOption(cht, "connectors")
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barclustered").Chart
    Stacked_SetConnectors cht, True
    AssertEqual "0", Registry_GetOption(cht, "connectors"), "agrupadas nao tem linhas de serie"
    CloseTemp wb
End Sub

' Core_SafeRefresh em barras chama Chart_SyncSecondaryX: o eixo X secundario (se existir) acompanha o primario.
Public Sub Test_Bars_SafeRefresh_SyncsSecondaryX()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    cht.HasAxis(xlCategory, xlSecondary) = True
    cht.Axes(xlCategory, xlSecondary).MinimumScale = 0
    cht.Axes(xlCategory, xlSecondary).MaximumScale = 5
    cht.Axes(xlValue, xlPrimary).MaximumScale = 120
    Core_SafeRefresh cht
    AssertEqual 120, cht.Axes(xlCategory, xlSecondary).MaximumScale, "Core_SafeRefresh chama Chart_SyncSecondaryX nas horizontais"
    AssertEqual 0, cht.Axes(xlCategory, xlSecondary).MinimumScale
    CloseTemp wb
End Sub

' Barras com totais ligados: o rotulo do total mais a direita nao invade a legenda (a direita) nem sai do grafico, em 100%
' (a barra termina no fim do eixo), com o maximo do eixo fixo no maior total e com totais de 7 digitos.
Public Sub Test_Bars_TotalLabels_StayLeftOfLegend()
    Dim sc As Variant, wb As Workbook, d As clsChartData, cht As Chart, tot As Series, i As Long, r As Double, tag As String
    For Each sc In Array("100", "fixmax", "7dig")
        tag = CStr(sc)
        Set wb = NewTempWorkbook()
        If tag = "7dig" Then Set d = FixtureTotals(wb, 123456, 234567, 1234567) Else Set d = FixtureTotals(wb, 15, 35, 60)
        Set cht = Stacked_Insert(d.Source.Worksheet, d, IIf(tag = "100", "barstacked100", "barstacked")).Chart
        If tag = "fixmax" Then
            Stacked_SetTotals cht, False
            cht.Axes(xlValue, xlPrimary).MaximumScale = 60       ' a maior barra termina no fim do eixo
            Stacked_SetTotals cht, True
        End If
        ForceLayout cht
        Set tot = Chart_FindSeries(cht, "tqq_total")
        For i = 1 To 3
            r = tot.Points(i).DataLabel.Left + tot.Points(i).DataLabel.Width
            AssertTrue r <= cht.Legend.Left + 0.5, tag & ": rotulo do total " & i & " termina antes da legenda: fim=" & r & " legenda=" & cht.Legend.Left
            AssertTrue r <= cht.ChartArea.Width, tag & ": rotulo do total " & i & " dentro do grafico"
        Next
        CloseTemp wb
    Next
End Sub

' A estimativa de largura (usada no afastamento da seta de CAGR e na folga dos totais, sem ler a geometria do rotulo)
' cobre a largura REAL do rotulo (DataLabel.Width, medida no Excel) com pouca sobra, de 1 a 9 digitos.
Public Sub Test_Labels_EstimateWidth_CoversRealLabelWidth()
    Dim sc As Variant, wb As Workbook, d As clsChartData, cht As Chart, tot As Series, real As Double, est As Double
    For Each sc In Array(Array(5, 8, 9), Array(15, 35, 60), Array(150, 350, 600), Array(1500, 3500, 6000), Array(15000, 35000, 60000), Array(123456, 234567, 1234567), Array(12345678, 23456789, 123456789))
        Set wb = NewTempWorkbook(): Set d = FixtureTotals(wb, CDbl(sc(0)), CDbl(sc(1)), CDbl(sc(2)))
        Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
        Set tot = Chart_FindSeries(cht, "tqq_total")
        est = Labels_EstimateWidth(tot.Points(3).DataLabel.Text)
        real = tot.Points(3).DataLabel.Width
        AssertTrue est >= real, "estimativa cobre o rotulo '" & tot.Points(3).DataLabel.Text & "': estimado=" & est & " real=" & real
        AssertTrue est <= real + 9, "estimativa sem excesso para '" & tot.Points(3).DataLabel.Text & "': estimado=" & est & " real=" & real
        AssertAlmost est, Labels_MaxTotalWidth(cht), 0.001, "maior largura entre os totais"
        CloseTemp wb
    Next
    AssertAlmost 6, Labels_EstimateWidth(""), 0.001
End Sub

' Limpeza de blocos (Tarefa 0): uma copia de grafico de barras (perdeu o nome tqq_<id>) continua apontando para o bloco pela
' auxiliar XY (X = TotalPos, Y = Pos) e pelos rotulos ligados; o bloco nao e apagado. Com o grafico removido, e apagado.
Public Sub Test_Bars_Cleanup_KeepsBlockReferencedByXYHelperAndLabels()
    Dim kind As Variant, wb As Workbook, d As clsChartData, co As ChartObject, id As String
    For Each kind In Array("barstacked", "barstacked100")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set co = Stacked_Insert(d.Source.Worksheet, d, CStr(kind))
        id = Mid$(co.Name, Len(ID_PREFIX) + 1)
        co.Name = "copia_de_grafico"                        ' como uma copia colada: sem o nome tqq_<id>
        AssertEqual 0, CalcSheet_Cleanup(wb), kind & ": nada removido"
        AssertTrue CalcSheet_BlockExists(wb, id), kind & ": bloco mantido pelas referencias da copia"
        co.Delete
        AssertEqual 1, CalcSheet_Cleanup(wb), kind & ": sem grafico o bloco e removido"
        AssertFalse CalcSheet_BlockExists(wb, id), CStr(kind)
        CloseTemp wb
    Next
End Sub
