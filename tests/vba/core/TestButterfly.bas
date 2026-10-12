Attribute VB_Name = "TestButterfly"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Faixa", "Mulheres", "Homens")
    ws.Range("A2:C2").Value = Array("0-14", 120, 130)
    ws.Range("A3:C3").Value = Array("15-29", 200, 210)
    ws.Range("A4:C4").Value = Array("30-59", 310, 290)
    ws.Range("A5:C5").Value = Array("60+", 90, 60)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C5"))
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function TqqNameCount(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then TqqNameCount = TqqNameCount + 1
    Next
End Function

' x (pt) de um valor do eixo de valores, por regra de tres sobre a area interna e a escala REAL (independe da producao).
Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

Public Sub Test_Bfly_Insert_SeriesBlockOptions()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, ax As Axis, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Bfly_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "butterfly", Registry_ChartType(cht)
    AssertEqual xlBarStacked, Bfly_SideSeries(cht, 1).ChartType: AssertEqual xlBarStacked, cht.SeriesCollection(1).ChartType, "barras empilhadas (o grafico le xlCombination com as auxiliares XY)"
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 10, blk.Columns.Count: AssertEqual 5, body.Rows.Count
    AssertEqual "Pos", body.Cells(1, BF_COL_POS).Value: AssertEqual 3.5, body.Cells(2, BF_COL_POS).Value
    AssertEqual 120, body.Cells(2, BF_COL_A).Value: AssertEqual 130, body.Cells(2, BF_COL_B).Value
    AssertEqual 250, body.Cells(2, BF_COL_TOTAL).Value
    AssertEqual -120, body.Cells(2, BF_COL_NEGA).Value
    AssertTrue body.Cells(2, BF_COL_GAP).Value > 0, "vao positivo"
    AssertAlmost body.Cells(2, BF_COL_GAP).Value, body.Cells(5, BF_COL_GAP).Value, 0.000001, "mesmo vao em todas as linhas"
    AssertAlmost CDbl(body.Cells(2, BF_COL_GAP).Value) + 130, CDbl(body.Cells(2, BF_COL_TOTALPOS).Value), 0.000001, "TotalPos = Gap + B"
    AssertAlmost -CDbl(blk.Cells(1, 7).Value), CDbl(blk.Cells(1, 6).Value), 0.000001, "SymL = -SymR"
    AssertAlmost CDbl(body.Cells(2, BF_COL_GAP).Value) + 310, CDbl(blk.Cells(1, 7).Value), 0.000001, "SymR = Gap + maior valor"
    AssertEqual 2, CDbl(blk.Cells(1, 8).Value), "Y dos pontos de simetria = n/2"
    AssertEqual "tqq_gapL", cht.SeriesCollection(1).Name: AssertEqual "Mulheres", cht.SeriesCollection(2).Name
    AssertEqual "tqq_gapR", cht.SeriesCollection(3).Name: AssertEqual "Homens", cht.SeriesCollection(4).Name
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible: AssertEqual msoFalse, cht.SeriesCollection(3).Format.Fill.Visible
    For Each s In cht.SeriesCollection
        If s.Name = "tqq_sym" Or s.Name = "tqq_cat" Or s.Name = "tqq_total" Then AssertEqual xlSecondary, s.AxisGroup, s.Name
    Next
    AssertTrue Not Chart_FindSeries(cht, "tqq_sym") Is Nothing: AssertTrue Not Chart_FindSeries(cht, "tqq_cat") Is Nothing
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertTrue ax.MinimumScaleIsAuto And ax.MaximumScaleIsAuto, "escala automatica"
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo simetrico"
    AssertEqual 4, cht.Axes(xlValue, xlSecondary).MaximumScale, "eixo Y secundario 0..n"
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlCategory).TickLabelPosition, "nomes das categorias so no centro"
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    AssertEqual msoFalse, cht.Axes(xlCategory).Format.Line.Visible, "sem a linha do eixo de categorias atravessando os nomes do centro"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com Mulheres e Homens"
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "4", Registry_GetOption(cht, "n"): AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertTrue Util_ToDouble(blk.Cells(1, 5).Value) >= 0.15, "GapK >= minimo"
    CloseTemp wb
End Sub

Public Sub Test_Bfly_Geometry_CenterLabelsAndMirroredBars()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, g As Double, pA As Point, pB As Point, lbl As DataLabel, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Bfly_Insert(d.Source.Worksheet, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    For i = 1 To 4
        Set pA = Bfly_SideSeries(cht, 1).Points(i): Set pB = Bfly_SideSeries(cht, 2).Points(i)
        AssertAlmost XOf(cht, -g), pA.Left + pA.Width, 1.5, "fim da barra esquerda " & i & " em x(-gap)"
        AssertAlmost XOf(cht, -g - CDbl(body.Cells(i + 1, BF_COL_A).Value)), pA.Left, 1.5, "inicio da barra esquerda " & i
        AssertAlmost XOf(cht, g), pB.Left, 1.5, "inicio da barra direita " & i & " em x(gap)"
        Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(i).DataLabel
        AssertAlmost XOf(cht, 0), lbl.Left + lbl.Width / 2, 1.5, "nome " & i & " centrado em x(0)"
        AssertTrue lbl.Top + lbl.Height / 2 > pA.Top - 1 And lbl.Top + lbl.Height / 2 < pA.Top + pA.Height + 1, "nome " & i & " na faixa da categoria"
        AssertEqual Util_CellLabel(body.Cells(i + 1, 1)), lbl.Text
        Set lbl = Chart_FindSeries(cht, "tqq_total").Points(i).DataLabel
        AssertTrue lbl.Left >= pB.Left + pB.Width - 1, "total " & i & " a direita da barra direita"
        AssertEqual Format$(body.Cells(i + 1, BF_COL_TOTAL).Value, "#,##0"), lbl.Text
    Next
    CloseTemp wb
End Sub

' Review Focus 1: nome longo e lados desiguais: o nome cabe no vao central e o eixo e simetrico.
Public Sub Test_Bfly_LongName_GapFitsAndAxisSymmetric()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, g As Double, lbl As DataLabel, ax As Axis, gapPt As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A4").Value = "Trinta a cinquenta e nove": ws.Range("B4").Value = 600: ws.Range("C4").Value = 60
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    gapPt = XOf(cht, g) - XOf(cht, -g)
    Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(3).DataLabel
    AssertTrue gapPt >= lbl.Width + 6, "vao (" & gapPt & " pt) cabe o nome (" & lbl.Width & " pt)"
    AssertAlmost XOf(cht, 0), lbl.Left + lbl.Width / 2, 1.5, "nome longo centrado"
    AssertAlmost Chart_FindSeries(cht, "tqq_cat").Points(1).DataLabel.Height, lbl.Height, 1, "nome longo numa linha so (sem quebra automatica do rotulo)"
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo simetrico com 600 x 60"
    AssertTrue ax.MaximumScale >= 600 + g, "o lado direito tambem alcanca 600 + vao"
    AssertEqual Format$(660, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    CloseTemp wb
End Sub

' O vao cresce no refresh quando um nome fica longo DEPOIS de inserir (Core_SafeRefresh -> Bfly_RefreshAfterDataChange -> Bfly_RefreshGap).
Public Sub Test_Bfly_NameChange_GapGrowsOnRefresh()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, g0 As Double, g As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    g0 = CDbl(body.Cells(2, BF_COL_GAP).Value)
    ws.Range("A4").Value = "Trinta a cinquenta e nove"
    Core_SafeRefresh cht
    ForceLayout cht
    g = CDbl(body.Cells(2, BF_COL_GAP).Value)
    AssertTrue g > g0, "o vao cresceu no refresh (" & g0 & " -> " & g & ")"
    Set lbl = Chart_FindSeries(cht, "tqq_cat").Points(3).DataLabel
    AssertTrue XOf(cht, g) - XOf(cht, -g) >= lbl.Width, "o nome longo cabe no vao depois do refresh (" & lbl.Width & " pt)"
    Teardown wb
End Sub

Public Sub Test_Bfly_Labels_AbsFormat_And_NumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, dec As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C5").Select
    Core_Insert "butterfly"
    Set cht = Core_ActiveChart()
    AssertEqual "butterfly", Registry_ChartType(cht)
    AssertEqual "120", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "lado esquerdo sem sinal"
    AssertEqual "130", Bfly_SideSeries(cht, 2).Points(1).DataLabel.Text
    AssertEqual xlLabelPositionInsideEnd, Bfly_SideSeries(cht, 1).DataLabels.Position
    AssertEqual "#,##0;#,##0;", Bfly_AbsFormat("#,##0;-#,##0;")
    AssertEqual "0.0%;0.0%;", Bfly_AbsFormat("0.0%;-0.0%;")
    Core_SetOption "numfmt", "dec1"
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "120" & dec & "0", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "dec1 no lado esquerdo, ainda sem sinal"
    AssertEqual "250" & dec & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total com dec1"
    Core_SetToggle "labels", False
    AssertFalse Bfly_SideSeries(cht, 1).HasDataLabels: AssertFalse Bfly_SideSeries(cht, 2).HasDataLabels
    AssertTrue Chart_FindSeries(cht, "tqq_cat").Points(1).HasDataLabel, "nomes das categorias nao sao 'Rotulos'"
    Core_SetToggle "labels", True
    AssertEqual "120" & dec & "0", Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "religar mantem o formato"
    Teardown wb
End Sub

Public Sub Test_Bfly_Totals_Legend_Palette_Gap()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C5").Select
    Core_Insert "butterfly"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "totals")
    Core_SetToggle "totals", True
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 2, cht.Legend.LegendEntries.Count, "religar os totais nao devolve auxiliares a legenda"
    AssertTrue Chart_FindSeries(cht, "tqq_total").DataLabels.Font.Bold
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), Bfly_SideSeries(cht, 1).Format.Fill.ForeColor.RGB
    AssertEqual RGB(127, 127, 127), Bfly_SideSeries(cht, 2).Format.Fill.ForeColor.RGB
    AssertEqual RGB(255, 255, 255), Bfly_SideSeries(cht, 1).DataLabels.Font.Color, "fonte por contraste no escuro"
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "o vao continua invisivel apos a paleta"
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_ClearLastNotice
    Core_SetOption "legend", "lineend"
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", Core_LastNotice()
    Teardown wb
End Sub

Public Sub Test_Bfly_Refusals_And_Rollback()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, ed As String, nm As Name, coll As New Collection, skipped As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("D1").Value = "Outro": ws.Range("D2:D5").Value = 1
    ws.Range("A1:D5").Select
    Core_ClearLastNotice
    Core_Insert "butterfly"
    AssertEqual "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1:D5").ClearContents
    Set cht = Bfly_Insert(ws, d).Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    coll.Add cht
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 0, skipped, "borboleta tem escala livre (um grafico so nao fixa)"
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Bfly_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Bfly_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    CloseTemp wb
End Sub

' Sem o ThinQQ (nenhum refresh): mudar um valor mantem o eixo simetrico e o vao proporcional (so formulas e series nativas).
Public Sub Test_Bfly_DataChange_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ax As Axis, body As Range, g0 As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Bfly_Insert(ws, d).Chart
    Set body = Block_Body(Registry_Block(cht)): Set ax = cht.Axes(xlValue, xlPrimary)
    g0 = CDbl(body.Cells(2, BF_COL_GAP).Value)
    ws.Range("B4").Value = 620
    cht.Refresh
    AssertAlmost -ax.MaximumScale, ax.MinimumScale, 0.000001, "eixo continua simetrico"
    AssertTrue ax.MaximumScale >= 620, "escala acompanha o dado novo"
    AssertTrue CDbl(body.Cells(2, BF_COL_GAP).Value) > g0, "vao cresce com o maximo (formula)"
    AssertEqual Format$(910, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    CloseTemp wb
End Sub

' Nomes no centro com o TEXTO da categoria de origem: data no formato da celula (nao o numero serial), ano como ano, vazio sem "0".
Public Sub Test_Bfly_CategoryText_DateYearBlank()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Periodo", "A", "B")
    ws.Range("A2:C2").Value = Array(2022, 1500, 1300)
    ws.Range("A3").Value = DateSerial(2026, 3, 2): ws.Range("A3").NumberFormat = "dd/mm/yyyy": ws.Range("B3:C3").Value = Array(10, 20)
    ws.Range("B4:C4").Value = Array(5, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set cht = Bfly_Insert(ws, d).Chart
    Set s = Chart_FindSeries(cht, "tqq_cat")
    AssertEqual "2022", s.Points(1).DataLabel.Text, "ano"
    AssertEqual Format$(ws.Range("A3").Value, "dd/mm/yyyy"), s.Points(2).DataLabel.Text, "data no formato da celula de origem"
    AssertEqual "", s.Points(3).DataLabel.Text, "categoria vazia sem 0"
    AssertEqual Format$(1500, "#,##0"), Bfly_SideSeries(cht, 1).Points(1).DataLabel.Text, "lado esquerdo absoluto com milhar"
    CloseTemp wb
End Sub

' Revisao 1 (importante): nome longo demais para o grafico (o vao em pontos nunca chega a largura interna). O vao tem teto (metade da area
' interna) e o refresh repetido (Core_SafeRefresh roda a cada calculo) nao multiplica GapK a cada chamada.
Public Sub Test_Bfly_NarrowLongName_GapCappedAndStable()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, k1 As Double, g As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A4").Value = "Trinta a cinquenta e nove anos completos"
    Set cht = Bfly_Insert(ws, d).Chart
    Set blk = Registry_Block(cht)
    cht.Parent.Width = 220
    Core_SafeRefresh cht
    ForceLayout cht
    k1 = Util_ToDouble(blk.Cells(1, 5).Value)
    Core_SafeRefresh cht
    ForceLayout cht
    AssertAlmost k1, Util_ToDouble(blk.Cells(1, 5).Value), 0.000001, "GapK estavel entre dois refresh"
    g = CDbl(Block_Body(blk).Cells(2, BF_COL_GAP).Value)
    AssertTrue XOf(cht, g) - XOf(cht, -g) <= 0.5 * cht.PlotArea.InsideWidth + 1, "vao (" & (XOf(cht, g) - XOf(cht, -g)) & " pt) <= metade da area interna (" & cht.PlotArea.InsideWidth & " pt)"
    AssertTrue XOf(cht, g) - XOf(cht, -g) > 0.25 * cht.PlotArea.InsideWidth, "o vao chega perto do teto (nome longo)"
    Teardown wb
End Sub

' Revisao 1: o vao que cresce DEPOIS de inserir (nome longo no refresh) empurra os totais para a direita; a folga reservada na insercao valia
' para o vao antigo. O total mais a direita continua antes da legenda (a direita) e dentro do grafico.
Public Sub Test_Bfly_GapGrowth_KeepsTotalsRoom()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, i As Long, lbl As DataLabel, k0 As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("C4").Value = 895                                   ' o lado direito tem o maior valor (medido: sem a folga refeita o total 1.205 passa 1,1 pt da legenda)
    Set cht = Bfly_Insert(ws, d).Chart
    k0 = Util_ToDouble(Registry_Block(cht).Cells(1, 5).Value)
    ws.Range("A4").Value = "Trinta a cinquenta e nove anos completos"
    Core_SafeRefresh cht
    ForceLayout cht
    AssertTrue Util_ToDouble(Registry_Block(cht).Cells(1, 5).Value) > k0, "o vao cresceu"
    Set s = Chart_FindSeries(cht, "tqq_total")
    For i = 1 To s.Points.Count
        Set lbl = s.Points(i).DataLabel
        AssertTrue lbl.Left + lbl.Width <= cht.Legend.Left + 0.5, "total " & i & " (" & lbl.Text & ") termina em " & (lbl.Left + lbl.Width) & " pt, antes da legenda (" & cht.Legend.Left & " pt)"
        AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width, "total " & i & " dentro do grafico"
    Next
    Teardown wb
End Sub

' Duas borboletas na mesma planilha: a da fixture (A1:C5) e outra em E1:G5 (lado direito maior).
Private Function TwoButterflies(wb As Workbook, ByRef c1 As Chart, ByRef c2 As Chart) As Worksheet
    Dim d As clsChartData, ws As Worksheet
    Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Bfly_Insert(ws, d).Chart
    ws.Range("E1:G1").Value = Array("Faixa", "Antes", "Depois")
    ws.Range("E2:G2").Value = Array("A", 50, 400): ws.Range("E3:G3").Value = Array("B", 80, 700)
    ws.Range("E4:G4").Value = Array("C", 40, 300): ws.Range("E5:G5").Value = Array("D", 10, 100)
    Set c2 = Bfly_Insert(ws, ChartData_FromRange(ws.Range("E1:G5"))).Chart
    Set TwoButterflies = ws
End Function

' Revisao 1: com Mesma escala, o lado ESQUERDO acima do maximo fixo tambem corta as barras (o aviso le Gap + o maior dos dois lados).
Public Sub Test_Bfly_SameScale_LeftSideAboveMax_IsClipped()
    Dim wb As Workbook, ws As Worksheet, c1 As Chart, c2 As Chart, coll As New Collection, skipped As Long, vmax As Double
    Set wb = NewTempWorkbook()
    Set ws = TwoButterflies(wb, c1, c2)
    coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    vmax = Val(Registry_GetOption(c1, "vmax", ""))
    AssertFalse Deco_ScaleClipped(c1), "sem corte logo depois de fixar"
    ws.Range("B4").Value = vmax * 2                              ' Mulheres (lado esquerdo) passa do maximo fixo; o lado direito fica pequeno
    AssertTrue Deco_ScaleClipped(c1), "lado esquerdo acima do maximo fixo (" & vmax & ") e corte"
    Teardown wb
End Sub

' Revisao 1: com o calculo MANUAL a celula Gap nao recalcula depois de gravar GapK; as passadas de Bfly_RefreshGap reliam o mesmo vao e
' multiplicavam GapK. A pasta em manual deve dar o mesmo vao que em automatico. O modo de calculo volta ao anterior tambem em erro.
Public Sub Test_Bfly_ManualCalc_SameGapAsAutomatic()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, prevCalc As Long
    Dim kAuto As Double, kMan As Double, gMan As Double, symMan As Double, maxAB As Double, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A4").Value = "Cinquenta+"
    Set cht = Bfly_Insert(ws, d).Chart
    kAuto = Util_ToDouble(Registry_Block(cht).Cells(1, 5).Value)
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    prevCalc = Application.Calculation
    On Error GoTo Restore
    Application.Calculation = xlCalculationManual
    Set cht = Bfly_Insert(ws, ChartData_FromRange(ws.Range("A1:C5"))).Chart
    Set blk = Registry_Block(cht)
    kMan = Util_ToDouble(blk.Cells(1, 5).Value)
    gMan = Util_ToDouble(Block_Body(blk).Cells(2, BF_COL_GAP).Value)
    symMan = Util_ToDouble(blk.Cells(1, 7).Value)
Restore:
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    Application.Calculation = prevCalc
    If en <> 0 Then CloseTemp wb: Err.Raise en, "Test_Bfly_ManualCalc_SameGapAsAutomatic", ed
    maxAB = 310
    AssertTrue kAuto > 0.15, "o nome pede vao maior que o minimo (" & kAuto & ")"
    AssertAlmost kAuto, kMan, 0.03 * kAuto, "GapK em manual = em automatico"
    AssertAlmost kMan * maxAB, gMan, 0.000001, "a celula Gap foi recalculada em manual"
    AssertAlmost gMan + maxAB, symMan, 0.000001, "SymR recalculado em manual"
    Teardown wb
End Sub
