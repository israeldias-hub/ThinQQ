Attribute VB_Name = "TestWaterfallPair"
Option Explicit
' Cascata comparativa (waterfallpair). Geometria por regra de tres sobre a area interna e a escala REAL do eixo (XOfX, YOf), independente das
' formulas do bloco; textos com Format$ (localidade).

' A imagem do think-cell: Segmento 1 | Segmento 2, Bruto 27 | 29, tres componentes que descem, Liquido = total.
Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "Segmento 1", "Segmento 2")
    ws.Range("A2:C2").Value = Array("Bruto", 27, 29)
    ws.Range("A3:C3").Value = Array("Componente 1", -4, -7)
    ws.Range("A4:C4").Value = Array("Componente 2", -6, -2)
    ws.Range("A5:C5").Value = Array("Componente 3", -6, -10)
    ws.Range("A6:C6").Value = Array("Líquido", "e", "e")
    Set Fixture = ChartData_FromRange(ws.Range("A1:C6"))
End Function

' Cruza o zero: a serie A desce abaixo de zero e termina negativa; a B cruza para cima.
Private Function FixtureZero(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "Loja A", "Loja B")
    ws.Range("A2:C2").Value = Array(2024, 10, 5)
    ws.Range("A3:C3").Value = Array("Queda", -25, -12)
    ws.Range("A4:C4").Value = Array("Alta", 8, 20)
    ws.Range("A5:C5").Value = Array("Fim", "t", "total")
    Set FixtureZero = ChartData_FromRange(ws.Range("A1:C5"))
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

Private Function SlotN(cht As Chart) As Long
    SlotN = cht.SeriesCollection(1).Points.Count
End Function

' x (pt) da fatia (ou do X de uma auxiliar XY do grupo principal: o mesmo numero), eixo de categorias "entre" as marcas.
Private Function XOfX(cht As Chart, x As Double) As Double
    XOfX = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (x - 0.5) / SlotN(cht)
End Function

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function BodyOf(cht As Chart) As Range
    Set BodyOf = Block_Body(Registry_Block(cht))
End Function

Private Function SideVal(cht As Chart, i As Long, j As Long, part As String) As Double
    SideVal = Util_ToDouble(BodyOf(cht).Cells(i + 1, Wfp_SideCol(j, part)).Value)
End Function

' "27,23,17,11,11": o fim da ponte do lado j em cada etapa.
Private Function EndsText(cht As Chart, j As Long) As String
    Dim i As Long
    For i = 1 To Wfp_StepCount(cht)
        EndsText = EndsText & IIf(i > 1, ",", "") & CStr(SideVal(cht, i, j, "End"))
    Next
End Function

Public Sub Test_Wfp_Insert_SeriesBlockOptions()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfallpair", Registry_ChartType(cht)
    AssertEqual xlColumnStacked, cht.SeriesCollection(2).ChartType, "colunas empilhadas (o grafico le xlCombination por causa das auxiliares XY)"
    AssertEqual "tqq_base", cht.SeriesCollection(1).Name: AssertEqual "Segmento 1", cht.SeriesCollection(2).Name
    AssertEqual "tqq_aneg", cht.SeriesCollection(3).Name: AssertEqual "Segmento 2", cht.SeriesCollection(4).Name
    AssertEqual "tqq_bneg", cht.SeriesCollection(5).Name
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "base invisivel"
    AssertEqual 14, SlotN(cht), "3n - 1 fatias (A, B e o vao; a ultima etapa sem vao)"
    AssertEqual 5, Wfp_StepCount(cht)
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 58, blk.Columns.Count: AssertEqual 31, body.Rows.Count, "cabecalho + 6n linhas"
    AssertEqual "27,23,17,11,11", EndsText(cht, 1): AssertEqual "29,22,20,10,10", EndsText(cht, 2)
    AssertTrue CBool(body.Cells(2, Wfp_SideCol(1, "Big")).Value), "a 1a barra e grande (inicio)"
    AssertFalse CBool(body.Cells(3, Wfp_SideCol(1, "Big")).Value): AssertTrue CBool(body.Cells(6, Wfp_SideCol(2, "Big")).Value), "o total e grande"
    AssertEqual "Bruto", body.Cells(2, WFP_COL_OUTER).Text, "fatia A leva o nome da etapa"
    AssertTrue IsEmpty(body.Cells(3, WFP_COL_OUTER).Value), "fatia B sem nome (o grupo do eixo de 2 niveis cobre A e B)"
    AssertEqual " ", CStr(body.Cells(4, WFP_COL_OUTER).Value), "vao"
    AssertEqual " ", CStr(body.Cells(2, WFP_COL_INNER).Value), "nivel interno"
    AssertEqual 27, body.Cells(2, WFP_COL_APOS).Value: AssertEqual 0, body.Cells(2, WFP_COL_BPOS).Value
    AssertEqual 29, body.Cells(3, WFP_COL_BPOS).Value: AssertEqual 0, body.Cells(3, WFP_COL_APOS).Value
    AssertEqual 23, body.Cells(5, WFP_COL_SBASE).Value, "Componente 1 de A: base 23": AssertEqual 4, body.Cells(5, WFP_COL_APOS).Value
    AssertEqual 0, Application.WorksheetFunction.Sum(body.Cells(4, WFP_COL_SBASE).Resize(1, 5)), "vao vazio"
    AssertEqual "int", Registry_GetOption(cht, "numfmt"): AssertEqual "padrao", Registry_GetOption(cht, "palette")
    AssertEqual "30", Registry_GetOption(cht, "gap"): AssertEqual "5", Registry_GetOption(cht, "n")
    AssertEqual "1", Registry_GetOption(cht, "totals"): AssertEqual "inbar", Registry_GetOption(cht, "legend")
    AssertEqual 30, cht.ChartGroups(1).GapWidth
    AssertAlmost 0.5 / 1.3, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "meia-largura em fatias"
    AssertTrue cht.Axes(xlCategory).TickLabels.MultiLevel, "eixo de categorias de 2 niveis"
    AssertEqual msoFalse, cht.Axes(xlCategory).Format.Line.Visible, "sem a linha do eixo (separadores)"
    AssertEqual xlNotPlotted, cht.DisplayBlanksAs
    AssertTrue Not Chart_FindSeries(cht, "tqq_zero") Is Nothing: AssertTrue Not Chart_FindSeries(cht, "tqq_names") Is Nothing
    AssertEqual xlPrimary, Chart_FindSeries(cht, "tqq_zero").AxisGroup, "auxiliar XY no grupo principal"
    AssertFalse Core_ChartSourceMismatch(cht)
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Geometry_SlotsAndBars()
    Dim wb As Workbook, d As clsChartData, cht As Chart, i As Long, j As Long, k As Long, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    For i = 1 To 5
        For j = 1 To 2
            k = Wfp_SlotOf(i, j)
            If j = 1 Then Set p = cht.SeriesCollection(2).Points(k) Else Set p = cht.SeriesCollection(4).Points(k)
            AssertAlmost XOfX(cht, CDbl(k)), p.Left + p.Width / 2, 1, "centro da barra " & i & "/" & j
            AssertAlmost cht.PlotArea.InsideWidth / SlotN(cht) / 1.3, p.Width, 1, "largura (espaco 30%) " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "Hi")), p.Top, 1.5, "topo da barra " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "Lo")), p.Top + p.Height, 1.5, "base da barra " & i & "/" & j
        Next
    Next
    CloseTemp wb
End Sub

Public Sub Test_Wfp_CrossesZero_NegativeTotal()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, pp As Point, pn As Point, z As Series
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    Set body = BodyOf(cht)
    AssertEqual "10,-15,-7,-7", EndsText(cht, 1): AssertEqual "5,-7,13,13", EndsText(cht, 2)
    AssertEqual "2024", body.Cells(2, WFP_COL_OUTER).Text, "ano como categoria"
    ' A2 (fatia 4) cruza o zero: base 0, +10 e -15
    AssertEqual 0, body.Cells(5, WFP_COL_SBASE).Value: AssertEqual 10, body.Cells(5, WFP_COL_APOS).Value: AssertEqual -15, body.Cells(5, WFP_COL_ANEG).Value
    ' A3 (fatia 7) inteira abaixo de zero: base invisivel -7 e a parte visivel -8 (de -7 a -15)
    AssertEqual -7, body.Cells(8, WFP_COL_SBASE).Value: AssertEqual 0, body.Cells(8, WFP_COL_APOS).Value: AssertEqual -8, body.Cells(8, WFP_COL_ANEG).Value
    ' B3 (fatia 8) cruza para cima: -7 a 13
    AssertEqual 13, body.Cells(9, WFP_COL_BPOS).Value: AssertEqual -7, body.Cells(9, WFP_COL_BNEG).Value
    ForceLayout cht
    Set pp = cht.SeriesCollection(2).Points(4): Set pn = cht.SeriesCollection(3).Points(4)
    AssertAlmost YOf(cht, 10), pp.Top, 1.5, "A2 parte positiva: topo em 10": AssertAlmost YOf(cht, 0), pp.Top + pp.Height, 1.5
    AssertAlmost YOf(cht, 0), pn.Top, 1.5, "A2 parte negativa: de 0": AssertAlmost YOf(cht, -15), pn.Top + pn.Height, 1.5, "a -15"
    Set pn = cht.SeriesCollection(3).Points(7)
    AssertAlmost YOf(cht, -7), pn.Top, 1.5, "A3 de -7": AssertAlmost YOf(cht, -15), pn.Top + pn.Height, 1.5, "a -15"
    Set pn = cht.SeriesCollection(3).Points(10)
    AssertAlmost YOf(cht, 0), pn.Top, 1.5, "total negativo de 0": AssertAlmost YOf(cht, -7), pn.Top + pn.Height, 1.5, "a -7"
    Set z = Chart_FindSeries(cht, "tqq_zero")
    AssertAlmost YOf(cht, 0), z.Points(1).Top, 1, "linha do zero na altura do zero"
    AssertAlmost cht.PlotArea.InsideLeft, z.Points(1).Left, 1, "da borda esquerda"
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, z.Points(2).Left, 1, "a borda direita"
    CloseTemp wb
End Sub

' Marcadores de total por celula (cada serie tem a sua ponte), texto e vazio valem 0, marcador na 1a linha = total de nada (0).
Public Sub Test_Wfp_Markers_TextAndFirstRow()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 10, "abc")
    ws.Range("A3:C3").Value = Array("Sub", "t", "'=")          ' "'=": texto (um "=" sozinho seria lido como formula)
    ws.Range("A4").Value = "Mais": ws.Range("B4").Value = "x": ws.Range("C4").Value = 5
    ws.Range("A5:C5").Value = Array("Fim", " Total ", "SUBTOTAL")
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Wfp_Insert(ws, d).Chart
    AssertEqual "10,10,10,10", EndsText(cht, 1), "t e Total (com espacos) sao totais; texto vale 0"
    AssertEqual "0,0,5,5", EndsText(cht, 2), "texto na 1a linha = 0; '=' e SUBTOTAL sao totais"
    AssertTrue CBool(BodyOf(cht).Cells(3, Wfp_SideCol(1, "Tot")).Value): AssertFalse CBool(BodyOf(cht).Cells(4, Wfp_SideCol(1, "Tot")).Value)
    ws.Range("B2").Value = "e"
    AssertEqual 0, SideVal(cht, 1, 1, "End"), "marcador na 1a linha: total de nada"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Refusals_And_Rollback()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, ed As String, nm As Name, d2 As clsChartData, cs As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B6").Select
    Core_ClearLastNotice
    Core_Insert "waterfallpair"
    AssertEqual "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1.", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1").Value = "Segmento 3": ws.Range("D2:D6").Value = 1
    ws.Range("A1:D6").Select
    Core_ClearLastNotice
    Core_Insert "waterfallpair"
    AssertEqual "A cascata comparativa compara exatamente 2 séries; a tabela tem 3 colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("D1:D6").ClearContents
    Set cht = Wfp_Insert(ws, d).Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "linha de valor: fora da matriz"
    ' "Nome dentro das barras" em outro tipo: recusa 7002
    ws.Range("H1:J1").Value = Array("Ano", "A", "B"): ws.Range("H2:J2").Value = Array(2024, 1, 2): ws.Range("H3:J3").Value = Array(2025, 2, 3)
    Set d2 = ChartData_FromRange(ws.Range("H1:J3"))
    Set cs = Stacked_Insert(ws, d2, "stacked").Chart
    On Error Resume Next
    Deco_SetLegend cs, "inbar"
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en: AssertEqual "'Nome dentro das barras' não se aplica a este tipo de gráfico.", ed
    ' rollback
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    CalcSheet_Cleanup wb
    AssertEqual 0, TqqNameCount(wb), "blocos orfaos limpos"
    On Error GoTo Cleanup
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Wfp_Insert ws, d
    en = Err.Number
    On Error GoTo Cleanup
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Wfp_Insert ws, d
    en = Err.Number
    On Error GoTo Cleanup
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    Teardown wb
    Exit Sub
Cleanup:
    en = Err.Number: ed = Err.Description
    Util_FailAfterBlockForTest = False: Util_FailAfterChartForTest = False
    Teardown wb
    Err.Raise en, "TestWaterfallPair", ed
End Sub

' Review Focus 4: nomes verticais dentro da 1a barra quando cabem; Direita / Nenhuma / Nome dentro das barras; barra baixa -> legenda.
Public Sub Test_Wfp_NamesInsideBars_And_LegendModes()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C6").Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Set s = Chart_FindSeries(cht, "tqq_names")
    cht.Refresh
    AssertEqual "Segmento 1", s.Points(1).DataLabel.Text: AssertEqual "Segmento 2", s.Points(2).DataLabel.Text
    AssertEqual xlUpward, s.Points(1).DataLabel.Orientation
    AssertFalse cht.HasLegend, "os nomes cabem: sem legenda"
    AssertEqual 1, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_FIT).Value)
    ForceLayout cht
    Set lbl = s.Points(1).DataLabel: Set p = cht.SeriesCollection(2).Points(1)
    AssertAlmost p.Left + p.Width / 2, lbl.Left + lbl.Width / 2, 1.5, "nome centrado na 1a barra"
    AssertTrue lbl.Top >= p.Top - 1 And lbl.Top + lbl.Height <= p.Top + p.Height + 1, "nome dentro da barra"
    AssertEqual RGB(255, 255, 255), lbl.Font.Color, "fonte branca na barra escura"
    AssertEqual RGB(64, 64, 64), s.Points(2).DataLabel.Font.Color, "fonte escura na barra clara"
    Core_SetOption "legend", "right"
    cht.Refresh
    AssertTrue cht.HasLegend: AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda so com as duas series"
    AssertEqual "", s.Points(1).DataLabel.Text, "Direita: sem nomes nas barras"
    Core_SetOption "legend", "none"
    cht.Refresh
    AssertFalse cht.HasLegend: AssertEqual "", s.Points(1).DataLabel.Text
    Core_SetOption "legend", "inbar"
    cht.Refresh
    AssertFalse cht.HasLegend: AssertEqual "Segmento 1", s.Points(1).DataLabel.Text
    ' nao cabe: 1a barra de A baixa -> os dois nomes somem e a legenda aparece (troca automatica no refresh)
    ws.Range("B2").Value = 2
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual "", s.Points(1).DataLabel.Text: AssertEqual "", s.Points(2).DataLabel.Text
    AssertTrue cht.HasLegend, "legenda no lugar dos nomes": AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "inbar", Registry_GetOption(cht, "legend"), "a escolha do usuario continua sendo 'dentro das barras'"
    ws.Range("B2").Value = 27
    Core_SafeRefresh cht
    cht.Refresh
    AssertFalse cht.HasLegend, "voltou a caber": AssertEqual "Segmento 1", s.Points(1).DataLabel.Text
    Teardown wb
End Sub

Public Sub Test_Wfp_CategoryLabels_TwoLevels_DateAndBlank()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A3").Value = DateSerial(2026, 1, 15): ws.Range("A3").NumberFormat = "dd/mm/yyyy"
    ws.Range("A4").ClearContents
    Set cht = Wfp_Insert(ws, d).Chart
    Set body = BodyOf(cht)
    AssertEqual ws.Range("A3").Text, body.Cells(5, WFP_COL_OUTER).Text, "data com o formato da origem (fatia A2)"
    AssertEqual " ", CStr(body.Cells(8, WFP_COL_OUTER).Value), "categoria vazia vira um espaco (fatia A3)"
    AssertTrue IsEmpty(body.Cells(6, WFP_COL_OUTER).Value), "fatia B2 sem nome"
    AssertEqual " ", CStr(body.Cells(7, WFP_COL_OUTER).Value), "vao 2"
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlCategory).TickLabelPosition
    AssertEqual xlHorizontal, cht.Axes(xlCategory).TickLabels.Orientation
    CloseTemp wb
End Sub

' Review Focus 5 (parte das barras): sem o ThinQQ (nenhum refresh), mudar a origem move a ponte.
Public Sub Test_Wfp_DataChange_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, v As Variant, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C3").Value = -12
    cht.Refresh
    AssertEqual "29,17,15,5,5", EndsText(cht, 2)
    v = cht.SeriesCollection(4).Values: AssertEqual 12, v(5), "B2: parte visivel 12"
    v = cht.SeriesCollection(1).Values: AssertEqual 17, v(5), "B2: base 17"
    ForceLayout cht
    Set p = cht.SeriesCollection(4).Points(5)
    AssertAlmost YOf(cht, 29), p.Top, 1.5: AssertAlmost YOf(cht, 17), p.Top + p.Height, 1.5
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Gap_SyncsHalfWidth()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart(): Set blk = Registry_Block(cht)
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    AssertAlmost 0.5 / 1.8, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "HW acompanha o Espaco"
    cht.ChartGroups(1).GapWidth = 120                 ' mudado a mao no Excel
    Core_SafeRefresh cht
    AssertAlmost 0.5 / 2.2, Util_ToDouble(blk.Cells(1, WFP_META_HW).Value), 0.000001, "o refresh le o espaco real"
    Teardown wb
End Sub

Public Sub Test_Wfp_SourceSizeChange_Detected()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    AssertEqual 5, Wfp_SourceRowsSpanned(cht)
    AssertFalse Core_ChartSourceMismatch(cht)
    ws.Rows(4).Insert
    ws.Range("A4:C4").Value = Array("Extra", -1, -1)
    AssertEqual 6, Wfp_SourceRowsSpanned(cht)
    AssertTrue Core_ChartSourceMismatch(cht), "linha inserida no meio da tabela"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Palette_And_Orphan()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    AssertEqual RGB(31, 56, 100), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB, "serie 1: pal(0)"
    AssertEqual RGB(157, 195, 230), cht.SeriesCollection(4).Format.Fill.ForeColor.RGB, "serie 2: pal(2)"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB: AssertEqual RGB(64, 64, 64), cht.SeriesCollection(3).Format.Fill.ForeColor.RGB
    AssertEqual RGB(166, 166, 166), cht.SeriesCollection(4).Format.Fill.ForeColor.RGB: AssertEqual RGB(166, 166, 166), cht.SeriesCollection(5).Format.Fill.ForeColor.RGB
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible, "a base continua invisivel"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_names").Points(2).DataLabel.Font.Color, "contraste com o cinza claro"
    ws.ChartObjects(1).Delete
    AssertEqual 1, CalcSheet_Cleanup(wb), "o bloco do grafico apagado e removido"
    Teardown wb
End Sub
